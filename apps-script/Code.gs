/**
 * Ragdoll ナビ — サーバー（Google Apps Script）
 *
 * 「00.AION2 DB_日課ツール」に紐づくスクリプトとして使う（拡張機能 → Apps Script）。
 * 秘密の値はコードに書かず、プロジェクトの設定 → スクリプト プロパティで管理する。
 *
 *   AUTH_CODE        認証コード（必須）
 *   TOKEN_SECRET     トークン署名用の鍵（setup() を一度実行すると自動で作成）
 *   IMAGE_FOLDER_ID  画像フォルダのID（必須）
 *   LOG_LAST_LOGIN   "false" にするとメンバー認証の「最終ログイン日」を記録しない（未設定なら記録する。1日1回）
 *   ADMIN_CODE       管理者コード（任意）。ニックネーム＋管理者コードでログインすると管理者メニューが使える
 *   GUIDE_SHEET_ID   DB_ガイドのスプレッドシートID（ガイド.gs）
 *   NOTIFY_EMAIL     通知メールの送り先（任意）
 *
 * 通信は text/plain の POST（JSON文字列）で受け取る。CORSの事前リクエストを避けるため。
 */

var TZ = 'Asia/Tokyo';
var TOKEN_DAYS = 30;          // ログイン保持期間
var REFRESH_AFTER_DAYS = 1;   // 発行から1日以上たっていたら自動延長
var FAIL_LIMIT = 5;           // 連続失敗の上限
var LOCK_SEC = 600;           // ロック時間（10分）
var CACHE_SEC = 60;           // シート読み込みのキャッシュ

var SH = {
  settings: '設定',
  items: '項目',
  timers: '定時イベント',
  schedule: '日程',
  legion: 'レギオン日程',
  members: 'メンバー認証',
  official: '公式候補'
};

var MSG = {
  empty: 'ニックネームと認証コードを入力してください。',
  wrong: 'ニックネームまたは認証コードが正しくありません。',
  locked: 'ログインに続けて失敗したため、10分後にもう一度お試しください。',
  blocked: 'このキャラクターは現在ご利用いただけません。管理者にお問い合わせください。',
  closed: 'Ragdoll ナビの公開は終了しました。',
  expired: 'ログインの有効期限が切れました。もう一度ログインしてください。',
  error: 'サーバーでエラーが発生しました。しばらくしてからもう一度お試しください。',
  unregistered: 'このニックネームはまだ登録されていません。利用申請を送ると、管理者が確認します。',
  pending: '利用申請を確認中です。管理者の承認をお待ちください。',
  applied: '利用申請を送りました。管理者の承認後にログインできます。'
};

/* ========== 入口 ========== */

function doPost(e) {
  var res;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (req.action) {
      case 'login':    res = actLogin(req); break;
      case 'data':     res = actData(req); break;
      case 'image':    res = actImage(req); break;
      case 'withdraw': res = actWithdraw(req); break;
      case 'ver':      res = actVer(req); break;
      case 'adm':      res = actAdmin(req); break;   // 管理者用（管理.gs）
      case 'apply':    res = actApply(req); break;   // 利用申請（管理.gs）
      case 'guide':    res = actGuide(req); break;   // ガイド（ガイド.gs）
      default:         res = { ok: false, reason: 'bad_request' };
    }
  } catch (err) {
    console.error(err);
    res = { ok: false, reason: 'error', message: MSG.error };
  }
  res.serverTime = Date.now();
  return json(res);
}

function doGet() {
  return json({ ok: true, service: 'ragdoll-navi', serverTime: Date.now() });
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ========== 初期設定（エディタから一度だけ実行） ========== */

function setup() {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('TOKEN_SECRET')) {
    p.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  }
  var lack = ['AUTH_CODE', 'IMAGE_FOLDER_ID'].filter(function (k) { return !p.getProperty(k); });
  console.log(lack.length ? '未設定のスクリプト プロパティ: ' + lack.join(', ') : '設定OK');
}

/* ========== アクション ========== */

function actLogin(req) {
  var name = String(req.name || '').trim(), code = String(req.code || '');
  if (!name || !code) return { ok: false, reason: 'empty', message: MSG.empty };

  var key = norm(name), cache = CacheService.getScriptCache(), fk = 'fail:' + hash(key).slice(0, 20);
  var fails = Number(cache.get(fk) || 0);
  if (fails >= FAIL_LIMIT) return { ok: false, reason: 'locked', message: MSG.locked };

  var adminOk = !!prop('ADMIN_CODE') && norm(code) === norm(prop('ADMIN_CODE'));
  var codeOk = adminOk || norm(code) === norm(prop('AUTH_CODE'));
  var m = findMember(key);
  if (!codeOk) {
    cache.put(fk, String(fails + 1), LOCK_SEC);
    return { ok: false, reason: 'wrong', message: MSG.wrong };
  }
  cache.remove(fk);
  if (!adminOk) {                                   // 管理者は公開停止中・名簿の状態に関係なく入れる（元に戻せるように）
    if (!isPublic()) return closedRes();
    // 認証コードが正しい人にだけ、未登録・申請中であることを伝える（コードを知らない人には区別しない）
    if (!m || m.status === '却下') return { ok: false, reason: 'unregistered', message: MSG.unregistered, canApply: true };
    if (m.status === '申請中') return { ok: false, reason: 'pending', message: MSG.pending };
    if (m.status !== '有効') return { ok: false, reason: 'blocked', message: MSG.blocked };
  } else if (!m) {
    return { ok: false, reason: 'wrong', message: MSG.wrong };
  }

  touchLogin(m);
  var res = { ok: true, token: issueToken(m.name, adminOk), name: m.name, data: buildData() };
  if (adminOk) res.admin = true;
  return res;
}

function actData(req) {
  var v = checkSession(req.token);
  if (!v.ok) return v;
  var res = { ok: true, token: v.token, name: v.name, data: buildData() };
  if (v.admin) res.admin = true;
  return res;
}

/* データの版だけを返す（軽い確認用。更新ボタンの赤い点に使う） */
function actVer(req) {
  var v = checkSession(req.token);
  if (!v.ok) return v;
  return { ok: true, token: v.token, ver: dataVer() };
}

function dataVer() { return Number(prop('DATA_VER')) || 0; }

function actImage(req) {
  var v = checkSession(req.token);
  if (!v.ok) return v;
  var file = String(req.file || '');
  if (!file || /[\/\\]/.test(file)) return { ok: false, reason: 'bad_request' };
  var it = DriveApp.getFolderById(prop('IMAGE_FOLDER_ID')).getFilesByName(file);
  if (!it.hasNext()) return { ok: false, reason: 'not_found' };
  var blob = it.next().getBlob();
  return { ok: true, token: v.token, file: file, mime: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) };
}

function actWithdraw(req) {
  var v = checkSession(req.token);
  if (!v.ok) return v;
  var m = findMember(norm(v.name));
  if (m) {
    writeMemberCell(m.row, 2, '退会');
    writeMemberCell(m.row, 3, today());
    CacheService.getScriptCache().remove('members');
  }
  return { ok: true, reason: 'withdrawn' };
}

/* セッション確認：トークン → 公開状態 → 名簿の状態。必要なら延長したトークンを返す */
function checkSession(token) {
  var t = readToken(token);
  if (!t) return { ok: false, reason: 'expired', message: MSG.expired };
  var admin = t.a === 1;
  if (!admin && !isPublic()) return closedRes();
  var m = findMember(norm(t.n));
  if (!m || (!admin && m.status !== '有効')) return { ok: false, reason: 'blocked', message: MSG.blocked };
  touchLogin(m);
  var fresh = (Date.now() - t.iat > REFRESH_AFTER_DAYS * 864e5) ? issueToken(m.name, admin) : null;
  return { ok: true, name: m.name, token: fresh, admin: admin };
}

function closedRes() {
  var s = readSettings();
  return { ok: false, reason: 'closed', message: s['終了メッセージ'] || MSG.closed };
}

/* ========== トークン（署名付き） ========== */
/* 認証コードを変えると、それまでのトークンはすべて無効になる（署名に含めているため）
 * 管理者トークンは管理者コードも署名に含める（管理者コードを変えると管理者ログインだけが無効になる） */

function issueToken(name, admin) {
  var now = Date.now(), o = { n: name, iat: now, exp: now + TOKEN_DAYS * 864e5 };
  if (admin) o.a = 1;
  var body = b64(JSON.stringify(o));
  return body + '.' + sign(body, !!admin);
}

function readToken(token) {
  if (!token || typeof token !== 'string') return null;
  var p = token.split('.');
  if (p.length !== 2) return null;
  try {
    var o = JSON.parse(unb64(p[0]));
    if (sign(p[0], o.a === 1) !== p[1]) return null;
    return o.exp > Date.now() ? o : null;
  } catch (e) { return null; }
}

function sign(s, admin) {
  var key = prop('TOKEN_SECRET') + '|' + norm(prop('AUTH_CODE')) + (admin ? '|A|' + norm(prop('ADMIN_CODE')) : '');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, key)).replace(/=+$/, '');
}

function b64(s) { return Utilities.base64EncodeWebSafe(Utilities.newBlob(s).getBytes()).replace(/=+$/, ''); }
function hash(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)).replace(/=+$/, ''); }
function unb64(s) { while (s.length % 4) s += '='; return Utilities.newBlob(Utilities.base64DecodeWebSafe(s)).getDataAsString(); }

/* ========== 名簿 ========== */

function findMember(key) {
  var list = cached('members', function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SH.members);
    var n = sh.getLastRow() - 1;
    if (n < 1) return [];
    return sh.getRange(2, 1, n, 4).getDisplayValues().map(function (r, i) {
      return { name: r[0].trim(), status: r[1].trim(), last: r[3].trim(), row: i + 2 };
    }).filter(function (r) { return r.name; });
  });
  for (var i = 0; i < list.length; i++) if (norm(list[i].name) === key) return list[i];
  return null;
}

/* 最終ログイン日（自動ログインでも1日1回記録） */
function touchLogin(m) {
  if (prop('LOG_LAST_LOGIN') === 'false') return;
  var d = today();
  if (m.last === d) return;
  try { writeMemberCell(m.row, 4, d); CacheService.getScriptCache().remove('members'); } catch (e) { console.error('touchLogin', e); }
}

function writeMemberCell(row, col, value) {
  SpreadsheetApp.getActive().getSheetByName(SH.members).getRange(row, col).setValue(value);
}

/* ========== 入力のゆれをそろえる ========== */
/* 全角/半角・ひらがな/カタカナ・長音記号・空白・英字の大小を同じ扱いにする */

function norm(s) {
  s = String(s || '').normalize('NFKC').toLowerCase();
  s = s.replace(/[ぁ-ゖ]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) + 0x60); });
  s = s.replace(/[‐-―−ｰ〜~-]/g, 'ー');
  return s.replace(/\s+/g, '');
}

/* ========== ツール用データ ========== */

function buildData() {
  return cached('data', function () {
    return {
      settings: readSettings(),
      items: readItems(),
      timers: readTimers(),
      schedule: readSchedule(),
      legion: readLegion(),
      ver: dataVer(),
      guideVer: guideVer(),
      updated: Utilities.formatDate(dataVer() ? new Date(dataVer()) : new Date(), TZ, 'yyyy-MM-dd HH:mm')
    };
  });
}

function rows(sheetName, cols) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var n = sh ? sh.getLastRow() - 1 : 0;
  return n < 1 ? [] : sh.getRange(2, 1, n, cols).getDisplayValues();
}

function on(v) { return String(v).toUpperCase() === 'TRUE'; }

function readSettings() {
  return cached('settings', function () {
    var o = {};
    rows(SH.settings, 2).forEach(function (r) { if (r[0]) o[r[0].trim()] = r[1].trim(); });
    return o;
  });
}

function isPublic() { return readSettings()['公開状態'] !== '停止'; }

// 項目: ID / 区分 / 単位 / 名称 / 回数 / 説明 / 表示 / 並び順
function readItems() {
  return rows(SH.items, 8).filter(function (r) { return r[0] && on(r[6]); })
    .sort(function (a, b) { return (Number(a[7]) || 999) - (Number(b[7]) || 999); })
    .map(function (r) {
      var o = { id: r[0], kind: r[1], unit: r[2], name: r[3] };
      if (r[4]) o.count = Number(r[4]) || r[4];
      if (r[5]) o.desc = r[5];
      return o;
    });
}

// 定時イベント: ID / 名称 / 種類 / 開始時刻 / 間隔 / 曜日 / 入場可能(分) / 表示文 / 表示 / 初期通知 / 補足
function readTimers() {
  var DOW = '日月火水木金土';
  return rows(SH.timers, 11).filter(function (r) { return r[0] && on(r[8]); }).map(function (r) {
    var hm = (r[3] || '0:00').split(':'), h = Number(hm[0]) || 0, m = Number(hm[1]) || 0;
    var o = { id: r[0], name: r[1], h: h, m: m, openMin: Number(r[6]) || 3, rule: r[7] };
    if (r[2] === '毎時') { o.h = 0; o.stepMin = 60; }
    else if (r[2] === '間隔') o.stepMin = minutes(r[4]) || 60;
    else if (r[2] === '曜日') o.dows = r[5].split('').map(function (c) { return DOW.indexOf(c); }).filter(function (d) { return d >= 0; });
    if (!on(r[9])) o.def = false;
    if (r[10]) o.tip = r[10];
    return o;
  });
}

function minutes(s) {
  var m = String(s).match(/(\d+)\s*(時間|分)/);
  return m ? Number(m[1]) * (m[2] === '時間' ? 60 : 1) : 0;
}

// 日程: 開始日 / 終了日 / 名称 / 種類 / メモ / 表示
function readSchedule() {
  var T = { 'シーズン': 'season', 'イベント': 'event', 'アップデート': 'update', 'メンテナンス': 'maint' };
  return rows(SH.schedule, 6).filter(function (r) { return r[0] && r[2] && on(r[5]); }).map(function (r, i) {
    var o = { id: 's' + (i + 1), start: ymd(r[0]), end: ymd(r[1] || r[0]), name: r[2], type: T[r[3]] || 'event' };
    if (r[4]) o.memo = r[4];
    if (o.type === 'update' || o.type === 'maint') o.keep = true; else o.endAt = endAtOf(o.end, '');
    return o;
  }).concat(readOfficial());
}

// 公式候補（承認済みのみ）: 承認 / 状態 / 種別 / 名称 / 開始日 / 終了日 / 開始時刻 / 終了時刻 / 公式URL / 掲載日時 / … / 完了時刻(Q)
// アップデート・メンテナンスは同じ日に複数あれば掲載日時が新しいものだけを使う（keep: 期限で消さない）
function readOfficial() {
  var T = { 'イベント': 'event', 'アップデート': 'update', 'メンテナンス': 'maint' };
  var list = rows(SH.official, 17).filter(function (r) { return on(r[0]) && r[1] !== '除外' && T[r[2]] && r[3] && r[4]; });
  var latest = {};
  list.forEach(function (r) {
    if (r[2] === 'イベント') return;
    var k = r[2] + '|' + ymd(r[4]);
    if (!latest[k] || String(r[9]) > String(latest[k][9])) latest[k] = r;
  });
  var maint = {};
  Object.keys(latest).forEach(function (k) { var r = latest[k]; if (r[2] === 'メンテナンス') maint[ymd(r[4])] = { start: r[6], end: r[16] || r[7] }; });
  var out = [];
  Object.keys(latest).forEach(function (k) {
    var r = latest[k], o = { id: 'o' + (r[11] || r[3]), start: ymd(r[4]), end: ymd(r[5] || r[4]), name: r[3], type: T[r[2]], keep: true };
    if (r[8]) o.url = r[8];
    if (r[9]) o.posted = r[9];
    if (o.type === 'maint') { if (r[6]) o.time = r[6] + (r[7] ? '〜' + r[7] : ''); if (r[16]) o.done = r[16]; }
    out.push(o);
  });
  list.forEach(function (r) {
    if (r[2] !== 'イベント') return;
    var o = { id: 'o' + (r[11] || r[3]), start: ymd(r[4]), end: ymd(r[5] || r[4]), name: r[3], type: 'event' };
    if (r[8]) o.url = r[8];
    var m1 = maint[o.start] || {}, m2 = maint[o.end] || {};
    var st = r[6] === 'メンテ後' ? (m1.end || '') : r[6], et = r[7] === 'メンテ前' ? (m2.start || '') : r[7];
    if (/^\d{1,2}:\d{2}$/.test(st || '')) o.startAt = new Date(o.start + 'T' + ('0' + st).slice(-5) + ':00+09:00').getTime();
    o.endAt = endAtOf(o.end, et);
    out.push(o);
  });
  return out;
}

/* 終了時刻（ミリ秒）：時刻があればその時刻、なければ終了日の翌日0:00（日本時間） */
function endAtOf(end, time) {
  if (/^\d{1,2}:\d{2}$/.test(time || '')) return new Date(end + 'T' + ('0' + time).slice(-5) + ':00+09:00').getTime();
  return new Date(end + 'T00:00:00+09:00').getTime() + 864e5;
}

// レギオン日程: 日付 / 時刻 / 名称 / メモ / 表示
function readLegion() {
  return rows(SH.legion, 5).filter(function (r) { return r[0] && r[2] && on(r[4]); }).map(function (r) {
    var o = { date: ymd(r[0]), time: r[1], name: r[2] };
    if (r[3]) o.memo = r[3];
    return o;
  });
}

function ymd(s) { return String(s).trim().replace(/\//g, '-'); }
function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/* ========== 共通 ========== */

function prop(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }

function cached(key, fn) {
  var c = CacheService.getScriptCache(), hit = c.get(key);
  if (hit) return JSON.parse(hit);
  var v = fn();
  try { c.put(key, JSON.stringify(v), CACHE_SEC); } catch (e) {}
  return v;
}

/* シートを編集したらキャッシュを消す（すぐ反映させるため）。トリガー不要の単純トリガー */
/* 名簿以外を編集した時はデータの版も上げる（ツールの「更新」に赤い点が出る） */
function onEdit(e) {
  CacheService.getScriptCache().removeAll(['data', 'settings', 'members']);
  var name = e && e.range ? e.range.getSheet().getName() : '';
  if (name !== SH.members) PropertiesService.getScriptProperties().setProperty('DATA_VER', String(Date.now()));
}
