/**
 * Ragdoll ナビ — 公式情報の候補取得（管理者が手動で実行）
 *
 * シートのメニュー「Ragdoll ナビ → 公式情報を取得（Steam）」から実行する。
 * - Steam 公式ニュースRSS（日本語）を1回読み、イベント・アップデート・メンテナンスの候補を「公式候補」タブに追加
 * - リンク先は公式サイトを優先（アップデート・メンテナンスは公式掲示板の記事、イベントは公式イベント一覧）
 * - 本文は日付・時刻を探すためだけに使い、シートには保存しない
 * - 候補は「承認」にチェックしたものだけがツールに表示される（自動承認はしない）
 * - 定期実行はしない。promotion.plaync.com（robots.txtで禁止）にはアクセスしない
 */

var OF = {
  steamApp: '3393110',
  steamHost: 'https://store.steampowered.com',
  board: 'https://api-global-community.plaync.com/aion2_global/board/',
  site: 'https://aion2.plaync.com/ja-jp/',
  eventList: 'https://aion2.plaync.com/ja-jp/eventon',
  updateList: 'https://aion2.plaync.com/ja-jp/board/update/list',
  noticeList: 'https://aion2.plaync.com/ja-jp/board/notice/list',
  head: ['承認', '状態', '種別', '名称', '開始日', '終了日', '開始時刻', '終了時刻', '公式URL', '掲載日時', '出典', '記事ID', '検出内容', '備考', '取得日時', 'Steam URL', '完了時刻']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ragdoll ナビ')
    .addItem('公式情報を取得（Steam）', 'fetchOfficial')
    .addToUi();
}

/* ========== メイン ========== */

function fetchOfficial() {
  var ui = SpreadsheetApp.getUi(), sh = ensureOfficialSheet();
  var items;
  try { items = readSteamFeed(); }
  catch (e) { ui.alert('取得できませんでした', String(e.message || e) + '\n\nシートの内容は変更していません。しばらくしてからもう一度お試しください。', ui.ButtonSet.OK); return; }

  var boards = readOfficialBoards();   // 公式リンク探し用（失敗しても続行）
  var last = sh.getLastRow(), existing = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 17).getDisplayValues().forEach(function (r, i) { if (r[11]) existing[r[11]] = { row: i + 2, detected: r[12], state: r[1], type: r[2], start: r[4], done: r[16] }; });

  var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), add = [], changed = 0, skipped = 0, done = 0;
  items.forEach(function (it) {
    var c = classify(it);
    if (!c) { skipped++; return; }
    var ex = existing[c.id];
    if (ex && ex.type === 'メンテナンス') {                         // 完了時刻は公式タイトルから自動で反映（承認し直し不要）
      var dn = doneTime(ex.start, boards);
      if (dn && dn !== ex.done) { sh.getRange(ex.row, 17).setValue(dn); done++; }
    }
    if (ex) {
      if (ex.detected !== c.detected && ex.state !== '除外') {
        sh.getRange(ex.row, 2).setValue('期間変更あり');
        sh.getRange(ex.row, 13).setValue(c.detected);
        sh.getRange(ex.row, 14).setValue('検出内容が変わりました（前回：' + ex.detected + '）');
        sh.getRange(ex.row, 15).setValue(now);
        changed++;
      }
      return;
    }
    add.push([false, c.state, c.type, c.name, c.start, c.end, c.startTime, c.endTime, officialUrl(c, boards),
      c.posted, 'Steam', c.id, c.detected, c.note, now, c.steamUrl, c.type === 'メンテナンス' ? doneTime(c.start, boards) : '']);
  });
  if (sh.getRange(1, 17).getValue() !== '完了時刻') sh.getRange(1, 17).setValue('完了時刻').setBackground('#1F3863').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 17).setValues(add);

  ui.alert('公式情報を取得しました',
    '新しい候補：' + add.length + '件\n内容が変わった候補：' + changed + '件\n完了時刻を反映：' + done + '件\n対象外（英語・期間なしなど）：' + skipped + '件\n\n' +
    '「公式候補」タブで内容を確認し、表示するものの「承認」にチェックしてください。' +
    (boards.ok ? '' : '\n\n※ 公式掲示板を読めなかったため、リンクは公式の一覧ページにしています。'), ui.ButtonSet.OK);
}

/* ========== Steam RSS ========== */

function readSteamFeed() {
  var path = '/feeds/news/app/' + OF.steamApp + '/';
  var rb = UrlFetchApp.fetch(OF.steamHost + '/robots.txt', { muteHttpExceptions: true });
  if (rb.getResponseCode() === 200 && !robotsAllows(rb.getContentText(), path)) throw new Error('Steam の robots.txt で取得が禁止されています。');

  var url = OF.steamHost + path + '?l=japanese', res = UrlFetchApp.fetch(url, { muteHttpExceptions: true }), code = res.getResponseCode();
  if (code === 403 || code === 429 || code >= 500) { Utilities.sleep(3000); res = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); code = res.getResponseCode(); }
  if (code !== 200) throw new Error('Steam から取得できませんでした（HTTP ' + code + '）。');

  var ch;
  try { ch = XmlService.parse(res.getContentText()).getRootElement().getChild('channel'); }
  catch (e) { throw new Error('Steam の応答を読み取れませんでした（形式エラー）。'); }
  if (!ch) throw new Error('Steam の応答に記事がありませんでした。');
  return ch.getChildren('item').map(function (it) {
    return { title: it.getChildText('title') || '', link: it.getChildText('link') || '', pub: it.getChildText('pubDate') || '',
      text: (it.getChildText('description') || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&') };
  });
}

/* robots.txt の簡易判定（User-agent: * のグループで最長一致） */
function robotsAllows(text, path) {
  var groups = [], cur = null, ua = false;
  text.split(/\r?\n/).forEach(function (l) {
    l = l.replace(/#.*/, '').trim(); var m = l.match(/^([A-Za-z-]+)\s*:\s*(.*)$/); if (!m) return;
    var k = m[1].toLowerCase(), v = m[2].trim();
    if (k === 'user-agent') { if (!ua) { cur = { ua: [], r: [] }; groups.push(cur); } cur.ua.push(v.toLowerCase()); ua = true; return; }
    ua = false; if (cur && (k === 'allow' || k === 'disallow') && v) cur.r.push({ allow: k === 'allow', p: v });
  });
  var g = groups.filter(function (x) { return x.ua.indexOf('google-apps-script') >= 0; })[0] || groups.filter(function (x) { return x.ua.indexOf('*') >= 0; })[0];
  if (!g) return true;
  var best = null;
  g.r.forEach(function (r) {
    var re = new RegExp('^' + r.p.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\\\$$/, '$'));
    if (re.test(path) && (!best || r.p.length > best.p.length || (r.p.length === best.p.length && r.allow))) best = r;
  });
  return !best || best.allow;
}

/* ========== 判別 ========== */

var RX_RANGE = /(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日([^～〜~]{0,40})[～〜~]([^0-9]{0,6})(?:(\d{4})年\s*)?(\d{1,2})月\s*(\d{1,2})日((?:[^\n。■、\d]|\d{1,2}[:：]\d{2}){0,30})/g;
var RX_TIME = /(\d{1,2})[:：](\d{2})\s*[～〜~\-－]\s*(\d{1,2})[:：](\d{2})/;

function classify(it) {
  if (!/[぀-ヿ]/.test(it.text)) return null;           // 日本語でない記事は対象外
  var idm = it.link.match(/\/view\/(\d+)/); if (!idm) return null;
  var pub = new Date(it.pub), posted = isNaN(pub) ? '' : Utilities.formatDate(pub, TZ, 'yyyy-MM-dd HH:mm');
  var baseY = isNaN(pub) ? new Date().getFullYear() : Number(Utilities.formatDate(pub, TZ, 'yyyy'));
  var o = { id: idm[1], steamUrl: it.link, posted: posted, start: '', end: '', startTime: '', endTime: '', note: '', state: '新規' };
  var r = classifyBody(it, o, baseY);
  if (r && /延期|延長|変更|中止/.test(it.title)) { r.state = '要確認'; r.note = (r.note ? r.note + ' / ' : '') + 'タイトルに「' + it.title.match(/延期|延長|変更|中止/)[0] + '」があります'; }
  return r;
}

function classifyBody(it, o, baseY) {

  if (/アップデートノート/.test(it.title)) {
    var d = titleDate(it.title, baseY); if (!d) return null;
    o.type = 'アップデート'; o.name = 'アップデートノート'; o.start = o.end = d; o.detected = d; o.md = mdKey(it.title);
    return o;
  }
  if (/メンテナンス/.test(it.title)) {
    var d2 = titleDate(it.title, baseY); if (!d2) return null;
    var t = it.text.match(RX_TIME);
    o.type = 'メンテナンス'; o.name = (/臨時/.test(it.title) ? '臨時メンテナンス' : '定期メンテナンス') + (/延長/.test(it.title) ? '（延長）' : '');
    o.start = o.end = d2; o.md = mdKey(it.title);
    if (t) { o.startTime = pad2(t[1]) + ':' + t[2]; o.endTime = pad2(t[3]) + ':' + t[4]; }
    else o.note = '時刻を検出できませんでした';
    o.detected = d2 + ' ' + (t ? o.startTime + '〜' + o.endTime : '時刻なし');
    return o;
  }

  var found = [], seen = {}, m; RX_RANGE.lastIndex = 0;
  while ((m = RX_RANGE.exec(it.text))) {
    var y = Number(m[1]), s = y + '-' + pad2(m[2]) + '-' + pad2(m[3]);
    var ey = m[6] ? Number(m[6]) : y, e = ey + '-' + pad2(m[7]) + '-' + pad2(m[8]);
    if (e < s) e = (ey + 1) + '-' + pad2(m[7]) + '-' + pad2(m[8]);
    var head = m[4] || '', hm = head.match(/(\d{1,2})[:：](\d{2})/);
    var st = hm ? pad2(hm[1]) + ':' + hm[2] : (/メンテナンス(終了)?後/.test(head) ? 'メンテ後' : '');
    var tail = m[9] || '', tm = tail.match(/(\d{1,2})[:：](\d{2})/);
    var et = tm ? pad2(tm[1]) + ':' + tm[2] : (/メンテナンス(開始)?前/.test(tail) ? 'メンテ前' : '');
    var k = s + '〜' + e;
    if (!seen[k]) { seen[k] = 1; found.push({ s: s, e: e, st: st, et: et }); }
  }
  if (!found.length) return null;                                // 期間の記載がない記事は対象外
  o.type = 'イベント'; o.name = it.title.replace(/^\s*\[イベント\]\s*/, '').trim();
  if (found.length === 1) { o.start = found[0].s; o.end = found[0].e; o.startTime = found[0].st; o.endTime = found[0].et; }
  else { o.state = '要確認'; o.note = '期間が複数あります：' + found.map(function (f) { return f.s + '〜' + f.e; }).join(' / '); }
  o.detected = found.map(function (f) { return f.s + (f.st ? ' ' + f.st : '') + '〜' + f.e + (f.et ? ' ' + f.et : ''); }).join(' / ');
  return o;
}

function titleDate(title, baseY) {
  var m = title.match(/(\d{1,2})月\s*(\d{1,2})日/); if (!m) return '';
  return baseY + '-' + pad2(m[1]) + '-' + pad2(m[2]);
}
function mdKey(title) { var m = title.match(/(\d{1,2})月\s*(\d{1,2})日/); return m ? Number(m[1]) + '月' + Number(m[2]) + '日' : ''; }
function pad2(n) { return ('0' + Number(n)).slice(-2); }

/* ========== 公式リンク ========== */

function readOfficialBoards() {
  var out = { ok: true, notice: [], update: [] };
  [['notice', 'notice_ja'], ['update', 'update_ja']].forEach(function (b) {
    try {
      var r = UrlFetchApp.fetch(OF.board + b[1] + '/article/search/moreArticle?isVote=true&moreSize=10&moreDirection=BEFORE&previousArticleId=0', { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) { out.ok = false; return; }
      out[b[0]] = (JSON.parse(r.getContentText()).contentList || []).map(function (a) { var raw = String(a.title || '').normalize('NFKC'); return { id: a.id, title: raw.replace(/\s+/g, ''), raw: raw }; });
    } catch (e) { out.ok = false; }
  });
  return out;
}

/* 公式お知らせのタイトル「(10/21 21:00 完了)」から完了時刻を取る。開始日 yyyy-MM-dd のメンテナンスが対象 */
function doneTime(start, boards) {
  var p = String(start || '').split('-'); if (p.length < 3) return '';
  var key = Number(p[1]) + '月' + Number(p[2]) + '日';
  var n = (boards.notice || []).filter(function (a) { return a.title.indexOf(key) >= 0 && /メンテナンス/.test(a.title); });
  for (var i = 0; i < n.length; i++) { var m = n[i].raw.match(/(\d{1,2}):(\d{2})\s*完了/); if (m) return pad2(m[1]) + ':' + m[2]; }
  return '';
}

function officialUrl(c, boards) {
  if (c.type === 'イベント') return OF.eventList;
  var key = (c.md || '').normalize('NFKC');
  if (c.type === 'アップデート') {
    var u = boards.update.filter(function (a) { return key && a.title.indexOf(key) >= 0 && /アップデートノート/.test(a.title); })[0];
    return u ? OF.site + 'board/update/view?articleId=' + u.id : OF.updateList;
  }
  var n = boards.notice.filter(function (a) { return key && a.title.indexOf(key) >= 0 && /メンテナンス/.test(a.title); })[0];
  return n ? OF.site + 'board/notice/view?articleId=' + n.id : OF.noticeList;
}

/* ========== 「公式候補」タブ ========== */

function ensureOfficialSheet() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SH.official);
  if (sh) return sh;
  sh = ss.insertSheet(SH.official);
  sh.getRange(1, 1, 1, 17).setValues([OF.head]).setBackground('#1F3863').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  var n = 500;
  sh.getRange(2, 1, n, 1).insertCheckboxes();
  sh.getRange(2, 2, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['新規', '要確認', '期間変更あり', '確認済み', '除外'], true).build());
  sh.getRange(2, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['イベント', 'アップデート', 'メンテナンス'], true).build());
  sh.getRange(2, 5, n, 4).setNumberFormat('@');
  sh.getRange(2, 1, n, 17).setVerticalAlignment('middle');
  sh.getRange(2, 17, n, 1).setNumberFormat('@');
  [60, 90, 90, 220, 95, 95, 70, 70, 260, 120, 60, 150, 220, 260, 120, 260, 80].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  var rng = sh.getRange(2, 1, n, 17);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="除外"').setBackground('#EDEDED').setFontColor('#808080').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR($B2="要確認",$B2="期間変更あり")').setBackground('#FFF4CC').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A2=TRUE').setBackground('#E3F2E8').setRanges([rng]).build()
  ]);
  return sh;
}
