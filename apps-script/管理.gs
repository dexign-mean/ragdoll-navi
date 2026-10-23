/**
 * Ragdoll ナビ — 管理者用の操作（ツールの「設定 → 管理者」から使う）
 *
 * ニックネーム＋管理者コード（スクリプト プロパティ ADMIN_CODE）でログインした時だけ使える。
 * 権限の確認はすべてサーバー側で行う（画面のコードは公開されているため）。
 * 変更はすべてシートに書き込まれるので、シートを見ればいつでも確認・修正できる。
 *
 * op: list / fetch / cand / setting / member
 */

var ADM_SETTINGS = { '自動承認': ['ON', 'OFF'], '公開状態': ['公開', '停止'], 'お知らせ': null };
var ADM_STATES = ['新規', '要確認', '期間変更あり', '確認済み', '除外'];
var ADM_MEMBER = ['有効', '停止', '退会', '申請中', '却下'];

function actAdmin(req) {
  var v = checkSession(req.token);
  if (!v.ok) return v;
  if (!v.admin) return { ok: false, reason: 'forbidden', message: '管理者のみ使えます。' };
  var r;
  switch (req.op) {
    case 'list':    r = { ok: true }; break;
    case 'fetch':
      try { var f = fetchCore(); r = { ok: true, message: fetchSummary(f) }; }
      catch (e) { r = { ok: false, reason: 'fetch', message: String(e.message || e) + '（シートは変更していません）' }; }
      break;
    case 'cand': case 'setting': case 'member':
      try { r = req.op === 'cand' ? admCand(req) : req.op === 'setting' ? admSetting(req) : admMember(req); }
      catch (e) { r = { ok: false, message: String(e.message || e) }; }
      break;
    default:        return { ok: false, reason: 'bad_request' };
  }
  if (r.ok || req.op === 'fetch') { var l = admList(); for (var k in l) r[k] = l[k]; }
  r.token = v.token;
  return r;
}

/* 管理画面に出す内容：公式候補・設定・メンバー */
function admList() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SH.official), cands = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 17).getDisplayValues().forEach(function (r, i) {
      if (!r[11] && !r[3]) return;
      cands.push({ row: i + 2, approve: on(r[0]), state: r[1], type: r[2], name: r[3], start: r[4], end: r[5], startTime: r[6], endTime: r[7],
        url: r[8], posted: r[9], id: r[11], detected: r[12], note: r[13], done: r[16] });
    });
  }
  var st = readSettingsFresh(), settings = {};
  Object.keys(ADM_SETTINGS).forEach(function (k) { settings[k] = st[k] || ''; });
  if (!settings['自動承認']) settings['自動承認'] = 'ON';
  if (!settings['公開状態']) settings['公開状態'] = '公開';
  var ms = ss.getSheetByName(SH.members), members = [];
  if (ms && ms.getLastRow() > 1) {
    ms.getRange(2, 1, ms.getLastRow() - 1, 7).getDisplayValues().forEach(function (r, i) {
      if (r[0].trim()) members.push({ row: i + 2, name: r[0].trim(), status: r[1], quit: r[2], last: r[3], memo: r[4], joined: r[5], applied: r[6] });
    });
  }
  var pending = cands.filter(function (c) { return c.state === '要確認' || c.state === '期間変更あり'; }).length;
  var applies = members.filter(function (m) { return m.status === '申請中'; }).length;
  return { cands: cands, settings: settings, members: members, pending: pending + applies, applies: applies };
}

/* 候補の更新（行番号と記事IDの両方が一致する時だけ書き込む） */
function admCand(req) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.official), row = Number(req.row);
  if (!sh || !(row >= 2) || row > sh.getLastRow()) return { ok: false, message: '候補が見つかりません。' };
  var cur = sh.getRange(row, 1, 1, 17).getDisplayValues()[0];
  if ((cur[11] || cur[3]) !== String(req.key || '')) return { ok: false, message: 'シートが変更されています。一覧を読み込み直してください。' };
  var d = req.data || {};
  if ('approve' in d) sh.getRange(row, 1).setValue(!!d.approve);
  if ('state' in d && ADM_STATES.indexOf(d.state) >= 0) sh.getRange(row, 2).setValue(d.state);
  if ('name' in d) sh.getRange(row, 4).setValue(String(d.name).slice(0, 100));
  [['start', 5], ['end', 6]].forEach(function (f) {
    if (f[0] in d) { var x = String(d[f[0]]).trim(); if (x && !/^\d{4}-\d{2}-\d{2}$/.test(x)) throw new Error('日付の形式が正しくありません'); sh.getRange(row, f[1]).setValue(x); }
  });
  [['startTime', 7], ['endTime', 8]].forEach(function (f) {
    if (f[0] in d) { var x = String(d[f[0]]).trim(); if (x && !/^(\d{1,2}:\d{2}|メンテ後|メンテ前)$/.test(x)) throw new Error('時刻の形式が正しくありません'); sh.getRange(row, f[1]).setValue(x); }
  });
  if ('url' in d) { var u = String(d.url).trim(); if (u && !/^https:\/\//.test(u)) throw new Error('URLは https:// で始めてください'); sh.getRange(row, 9).setValue(u); }
  bumpData();
  return { ok: true, message: '保存しました。' };
}

/* 設定の変更（決められたキーのみ） */
function admSetting(req) {
  var key = String(req.key || ''), val = String(req.value == null ? '' : req.value).trim();
  if (!(key in ADM_SETTINGS)) return { ok: false, message: 'この設定は変更できません。' };
  if (ADM_SETTINGS[key] && ADM_SETTINGS[key].indexOf(val) < 0) return { ok: false, message: '値が正しくありません。' };
  if (key === 'お知らせ' && val.length > 200) return { ok: false, message: 'お知らせは200文字以内にしてください。' };
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.settings), last = sh.getLastRow(), row = 0;
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(function (r, i) { if (!row && r[0].trim() === key) row = i + 2; });
  if (row) sh.getRange(row, 2).setValue(val);
  else sh.getRange(last + 1, 1, 1, 2).setValues([[key, val]]);
  bumpData();
  return { ok: true, message: '保存しました。' };
}

/* メンバーの追加・状態変更 */
function admMember(req) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.members);
  if (req.sub === 'add') {
    // 1行に1人。改行・カンマ・読点区切りで複数まとめて追加できる（最大100人）
    var names = String(req.name || '').split(/[\r\n,、，\t]+/).map(function (x) { return x.trim(); }).filter(Boolean).slice(0, 100);
    if (!names.length) return { ok: false, message: 'ニックネームを入力してください。' };
    CacheService.getScriptCache().remove('members');
    var seen = {}, add = [], dup = [];
    names.forEach(function (n) {
      var k = norm(n);
      if (seen[k] || findMember(k)) { dup.push(n); return; }
      seen[k] = 1; add.push([n, '有効', '', '', '', today()]);
    });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 6).setValues(add);
    CacheService.getScriptCache().remove('members');
    if (!add.length) return { ok: false, message: 'すべて登録済みです（' + dup.join('、') + '）。' };
    return { ok: true, message: add.length + '人を追加しました。' + (dup.length ? '\n登録済みのためスキップ：' + dup.join('、') : '') };
  }
  if (req.sub === 'status') {
    var row = Number(req.row), st = String(req.status || '');
    if (ADM_MEMBER.indexOf(st) < 0 || !(row >= 2) || row > sh.getLastRow()) return { ok: false, message: '値が正しくありません。' };
    var cur = sh.getRange(row, 1).getDisplayValue().trim();
    if (cur !== req.name) return { ok: false, message: 'シートが変更されています。一覧を読み込み直してください。' };
    sh.getRange(row, 2).setValue(st);
    if (st === '退会') sh.getRange(row, 3).setValue(today());
    if (st === '有効' && !sh.getRange(row, 6).getDisplayValue()) sh.getRange(row, 6).setValue(today());   // 承認した日を登録日に
    CacheService.getScriptCache().remove('members');
    return { ok: true, message: cur + ' を「' + st + '」にしました。' };
  }
  return { ok: false, reason: 'bad_request' };
}

/* ========== 利用申請（ログイン画面から。認証コードが正しい人だけ） ========== */

function actApply(req) {
  var name = String(req.name || '').trim(), code = String(req.code || '');
  if (!name || !code) return { ok: false, reason: 'empty', message: MSG.empty };
  if (!isPublic()) return closedRes();
  if (norm(code) !== norm(prop('AUTH_CODE'))) return { ok: false, reason: 'wrong', message: MSG.wrong };
  var cache = CacheService.getScriptCache(), ck = 'apply:' + hash(norm(name)).slice(0, 20);
  if (cache.get(ck)) return { ok: true, reason: 'applied', message: MSG.applied };   // 連打しても1回だけ
  if (name.length > 30) return { ok: false, message: 'ニックネームが長すぎます。' };

  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var m = findMemberFresh(norm(name)), sh = SpreadsheetApp.getActive().getSheetByName(SH.members);
    if (m && m.status === '有効') return { ok: false, reason: 'registered', message: 'すでに登録されています。ログインしてください。' };
    if (m && m.status === '停止') return { ok: false, reason: 'blocked', message: MSG.blocked };
    if (m && m.status === '申請中') { cache.put(ck, '1', 3600); return { ok: true, reason: 'applied', message: MSG.pending }; }
    if (m) { sh.getRange(m.row, 2).setValue('申請中'); sh.getRange(m.row, 7).setValue(today()); }   // 却下・退会からの再申請
    else sh.getRange(sh.getLastRow() + 1, 1, 1, 7).setValues([[name, '申請中', '', '', '', '', today()]]);
    CacheService.getScriptCache().remove('members');
  } finally { lock.releaseLock(); }
  cache.put(ck, '1', 3600);
  notifyApply(name);
  return { ok: true, reason: 'applied', message: MSG.applied };
}

/* 管理者へメールで知らせる（同じ人は1日1回まで）。送り先はスクリプトの持ち主、または NOTIFY_EMAIL */
function notifyApply(name) {
  try {
    var cache = CacheService.getScriptCache(), k = 'mail:' + hash(norm(name)).slice(0, 20);
    if (cache.get(k)) return;
    var to = prop('NOTIFY_EMAIL') || Session.getEffectiveUser().getEmail();
    if (!to) return;
    MailApp.sendEmail(to, '[Ragdoll ナビ] 利用申請：' + name,
      name + ' さんから利用申請がありました。\n\nツールの管理者メニュー（盾のアイコン）→「メンバー」で承認・却下できます。\nhttps://ragdoll-navi.pages.dev/');
    cache.put(k, '1', 21600);
  } catch (e) { console.error('notify failed', e); }
}

function findMemberFresh(key) { CacheService.getScriptCache().remove('members'); return findMember(key); }

/* データの版を上げ、キャッシュを消す（ツールの「更新」に赤い点が出る） */
function bumpData() {
  CacheService.getScriptCache().removeAll(['data', 'settings', 'members']);
  PropertiesService.getScriptProperties().setProperty('DATA_VER', String(Date.now()));
}
