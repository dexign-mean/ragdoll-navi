/**
 * Ragdoll ナビ — サーバー（Google Apps Script）
 *
 * 「00.AION2 DB_日課ツール」に紐づくスクリプトとして使う（拡張機能 → Apps Script）。
 * 秘密の値はコードに書かず、プロジェクトの設定 → スクリプト プロパティで管理する。
 *
 *   AUTH_CODE        認証コード（必須）
 *   TOKEN_SECRET     トークン署名用の鍵（setup() を一度実行すると自動で作成）
 *   IMAGE_FOLDER_ID  画像フォルダのID（必須）
 *   LOG_LAST_LOGIN   "true" にするとメンバー認証の「最終ログイン日」を記録（任意）
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
  members: 'メンバー認証'
};

var MSG = {
  empty: 'ニックネームと認証コードを入力してください。',
  wrong: 'ニックネームまたは認証コードが正しくありません。',
  locked: 'ログインに続けて失敗したため、10分後にもう一度お試しください。',
  blocked: 'このキャラクターは現在ご利用いただけません。管理者にお問い合わせください。',
  closed: 'Ragdoll ナビの公開は終了しました。',
  expired: 'ログインの有効期限が切れました。もう一度ログインしてください。',
  error: 'サーバーでエラーが発生しました。しばらくしてからもう一度お試しください。'
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
  if (!isPublic()) return closedRes();

  var key = norm(name), cache = CacheService.getScriptCache(), fk = 'fail:' + hash(key).slice(0, 20);
  var fails = Number(cache.get(fk) || 0);
  if (fails >= FAIL_LIMIT) return { ok: false, reason: 'locked', message: MSG.locked };

  var codeOk = norm(code) === norm(prop('AUTH_CODE'));
  var m = findMember(key);
  if (!codeOk || !m) {
    cache.put(fk, String(fails + 1), LOCK_SEC);
    return { ok: false, reason: 'wrong', message: MSG.wrong };
  }
  cache.remove(fk);
  if (m.status !== '有効') return { ok: false, reason: 'blocked', message: MSG.blocked };

  if (prop('LOG_LAST_LOGIN') === 'true') writeMemberCell(m.row, 4, today());
  return { ok: true, token: issueToken(m.name), name: m.name, data: buildData() };
}

function actData(req) {
  var v = checkSession(req.token);
  if (!v.ok) return v;
  return { ok: true, token: v.token, name: v.name, data: buildData() };
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
  if (!isPublic()) return closedRes();
  var m = findMember(norm(t.n));
  if (!m || m.status !== '有効') return { ok: false, reason: 'blocked', message: MSG.blocked };
  var fresh = (Date.now() - t.iat > REFRESH_AFTER_DAYS * 864e5) ? issueToken(m.name) : null;
  return { ok: true, name: m.name, token: fresh };
}

function closedRes() {
  var s = readSettings();
  return { ok: false, reason: 'closed', message: s['終了メッセージ'] || MSG.closed };
}

/* ========== トークン（署名付き） ========== */
/* 認証コードを変えると、それまでのトークンはすべて無効になる（署名に含めているため） */

function issueToken(name) {
  var now = Date.now();
  var body = b64(JSON.stringify({ n: name, iat: now, exp: now + TOKEN_DAYS * 864e5 }));
  return body + '.' + sign(body);
}

function readToken(token) {
  if (!token || typeof token !== 'string') return null;
  var p = token.split('.');
  if (p.length !== 2 || sign(p[0]) !== p[1]) return null;
  try {
    var o = JSON.parse(unb64(p[0]));
    return o.exp > Date.now() ? o : null;
  } catch (e) { return null; }
}

function sign(s) {
  var key = prop('TOKEN_SECRET') + '|' + norm(prop('AUTH_CODE'));
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
    return sh.getRange(2, 1, n, 2).getDisplayValues().map(function (r, i) {
      return { name: r[0].trim(), status: r[1].trim(), row: i + 2 };
    }).filter(function (r) { return r.name; });
  });
  for (var i = 0; i < list.length; i++) if (norm(list[i].name) === key) return list[i];
  return null;
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
    return o;
  });
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
