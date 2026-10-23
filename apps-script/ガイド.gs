/**
 * Ragdoll ナビ — ガイド（「00.AION2 DB_ガイド」を読み、ツールのガイドタブ用に変換して渡す）
 *
 * スクリプト プロパティ
 *   GUIDE_SHEET_ID   DB_ガイドのスプレッドシートID（必須）
 *   GUIDE_VER        ガイドの版（自動。変換が終わった時に更新）
 *   IMG_MAP          採集マップの元画像ID → 画像フォルダのファイル名（prepareMapImages が作成）
 *
 * 初回だけエディタで setupGuide を実行する（DB_ガイド編集時のトリガー設定＋採集マップ画像の作成＋最初の変換）。
 * DB_ガイドを編集すると、数分後にまとめて1回だけ変換し、画像フォルダの guide-data.json に保存する。
 * ツールはこの保存済みの結果を受け取るだけなので、開く時に待たされない。
 * 内容（文章・表）はシートの値と書式から作るだけで、シートやリンクのURLは渡さない。
 */

var GUIDE_SKIP = ['目次'];
var GUIDE_CHUNK = 90000;      // CacheService の1件あたり上限（100KB）より少し小さく
var GUIDE_CACHE_SEC = 21600;  // 6時間（版が変わるとキーが変わる）

var GUIDE_FILE = 'guide-data.json';
var GUIDE_DELAY_MIN = 3;      // 編集が続いても、最後の編集から数分後に1回だけ変換

function actGuide(req) {
  var v = checkSession(req.token);
  if (!v.ok) return v;
  var ver = guideVer();
  if (req.ver && Number(req.ver) === ver) return { ok: true, token: v.token, ver: ver, same: true };
  var g = loadGuide(ver);
  return { ok: true, token: v.token, ver: guideVer(), guide: g };
}

function guideVer() { return Number(prop('GUIDE_VER')) || 1; }

/* 保存済みの変換結果を返す（キャッシュ → 画像フォルダのファイル → なければその場で変換） */
function loadGuide(ver) {
  var s = cacheGetBig('guide:' + ver);
  if (!s) {
    var it = DriveApp.getFolderById(prop('IMAGE_FOLDER_ID')).getFilesByName(GUIDE_FILE);
    if (it.hasNext()) s = it.next().getBlob().getDataAsString('UTF-8');
    if (!s) return saveGuide(readGuide()).guide;
    cachePutBig('guide:' + ver, s);
  }
  return JSON.parse(s);
}

/* 変換して保存し、版を上げる（ツールの「更新」に赤い点が出る） */
function saveGuide(g) {
  var s = JSON.stringify(g), folder = DriveApp.getFolderById(prop('IMAGE_FOLDER_ID')), it = folder.getFilesByName(GUIDE_FILE);
  if (it.hasNext()) it.next().setContent(s); else folder.createFile(GUIDE_FILE, s, 'application/json');
  var now = String(Date.now()), P = PropertiesService.getScriptProperties();
  P.setProperty('GUIDE_VER', now);
  P.setProperty('DATA_VER', now);
  CacheService.getScriptCache().removeAll(['data']);
  cachePutBig('guide:' + now, s);
  return { ver: Number(now), guide: g };
}

/* CacheService は1件100KBまでなので分割して保存 */
function cachePutBig(key, s) {
  var put = {}, m = 0;
  for (var k = 0; k < s.length; k += GUIDE_CHUNK, m++) put[key + ':' + m] = s.slice(k, k + GUIDE_CHUNK);
  put[key + ':n'] = String(m);
  try { CacheService.getScriptCache().putAll(put, GUIDE_CACHE_SEC); } catch (e) { console.error('guide cache', e); }
}
function cacheGetBig(key) {
  var c = CacheService.getScriptCache(), n = Number(c.get(key + ':n'));
  if (!n) return '';
  var keys = []; for (var i = 0; i < n; i++) keys.push(key + ':' + i);
  var got = c.getAll(keys), parts = [];
  for (var j = 0; j < n; j++) { if (!got[key + ':' + j]) return ''; parts.push(got[key + ':' + j]); }
  return parts.join('');
}

/* ========== 読み込み・変換 ========== */
/* 1行 = { r: 行番号, c: [セル] }。セル = { x: 列(0始まり), s: 横の結合数, t: 文字,
 *   h: 見出し色, n: 章見出し色, b: 太字, e: 強調(全体=1 / 一部=[[開始,終了],…]), l: 外部リンク, j: シート内リンク, i: 画像ファイル名 } */
function readGuide() {
  try { return readGuideApi(); }
  catch (e) { console.error('Sheets API で読めなかったため通常の方法で読みます', e); return readGuideSlow(); }
}

/* Sheets API で全タブの値・書式・結合を1回のリクエストで読む（速い） */
function readGuideApi() {
  var id = prop('GUIDE_SHEET_ID'), imgMap = {};
  try { imgMap = JSON.parse(prop('IMG_MAP') || '{}'); } catch (e) {}
  var fields = 'sheets(properties(sheetId,title),merges,data(rowData(values(formattedValue,hyperlink,userEnteredValue/formulaValue,' +
    'effectiveFormat(backgroundColor,textFormat(bold,foregroundColor)),textFormatRuns))))';
  var res = UrlFetchApp.fetch('https://sheets.googleapis.com/v4/spreadsheets/' + id + '?includeGridData=true&fields=' + encodeURIComponent(fields),
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Sheets API HTTP ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
  var book = JSON.parse(res.getContentText()), tabs = [], intro = [];
  (book.sheets || []).forEach(function (sh) {
    var name = sh.properties.title, rows = ((sh.data || [])[0] || {}).rowData || [];
    if (GUIDE_SKIP.indexOf(name) >= 0) {
      intro = readIntroValues(rows.map(function (r) { return (r.values || []).map(function (c) { return c.formattedValue || ''; }); }));
      return;
    }
    var span = {}, cover = {};
    (sh.merges || []).forEach(function (m) {
      span[m.startRowIndex + ':' + m.startColumnIndex] = m.endColumnIndex - m.startColumnIndex;
      for (var r = m.startRowIndex; r < m.endRowIndex; r++) for (var c = m.startColumnIndex; c < m.endColumnIndex; c++)
        if (r !== m.startRowIndex || c !== m.startColumnIndex) cover[r + ':' + c] = 1;
    });
    var out = [];
    rows.forEach(function (row, r) {
      var cells = [];
      (row.values || []).forEach(function (cv, c) {
        if (cover[r + ':' + c]) return;
        var t = String(cv.formattedValue || '').replace(/\s+$/, '');
        if (!t) return;
        var ef = cv.effectiveFormat || {}, tf = ef.textFormat || {}, o = { x: c, t: t }, sp = span[r + ':' + c];
        if (sp > 1) o.s = sp;
        var tone = bgTone(rgbHex(ef.backgroundColor));
        if (tone === 'n') o.n = 1; else if (tone === 'h') o.h = 1;
        if (tf.bold) o.b = 1;
        if (isRed(rgbHex(tf.foregroundColor))) o.e = 1;
        else if (cv.textFormatRuns) {
          var em = [], runs = cv.textFormatRuns;
          runs.forEach(function (run, i) {
            var end = i + 1 < runs.length ? runs[i + 1].startIndex : t.length, st = run.startIndex || 0;
            if (run.format && isRed(rgbHex(run.format.foregroundColor)) && t.slice(st, end).trim()) em.push([st, Math.min(end, t.length)]);
          });
          if (em.length) o.e = em;
        }
        var link = cv.hyperlink || hyperlinkOf((cv.userEnteredValue || {}).formulaValue, null);
        if (link) applyLink(o, link, imgMap);
        o.t = o.t.replace(/\s*🔍$/, '');
        cells.push(o);
      });
      if (cells.length) out.push({ r: r + 1, c: cells });
    });
    tabs.push({ name: name, gid: sh.properties.sheetId, rows: out });
  });
  return { tabs: tabs, intro: intro, updated: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') };
}

function rgbHex(c) {
  if (!c) return '';
  if (c.red === undefined && c.green === undefined && c.blue === undefined) return '#000000';
  var h = function (v) { return ('0' + Math.round((v || 0) * 255).toString(16)).slice(-2); };
  return '#' + h(c.red) + h(c.green) + h(c.blue);
}

function applyLink(o, link, imgMap) {
  var m1 = link.match(/^#gid=(\d+)(?:&range=[A-Z]+(\d+))?/), m2 = link.match(/\/d\/([\w-]{20,})/);
  if (m1) o.j = m1[1] + (m1[2] ? ':' + m1[2] : '');
  else if (m2 && /drive\.google\.com/.test(link)) { if (imgMap[m2[1]]) o.i = imgMap[m2[1]]; }   // 元画像のURLは渡さない
  else if (/^https:\/\//.test(link) && !/docs\.google\.com|drive\.google\.com/.test(link)) o.l = link;
}

/* 予備：SpreadsheetApp で1タブずつ読む（Sheets API が使えない場合） */
function readGuideSlow() {
  var ss = SpreadsheetApp.openById(prop('GUIDE_SHEET_ID')), imgMap = {};
  try { imgMap = JSON.parse(prop('IMG_MAP') || '{}'); } catch (e) {}
  var tabs = [], intro = [];
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (GUIDE_SKIP.indexOf(name) >= 0) { intro = readIntroValues(sh.getDataRange().getDisplayValues()); return; }
    tabs.push({ name: name, gid: sh.getSheetId(), rows: readSheetRows(sh, imgMap) });
  });
  return { tabs: tabs, intro: intro, updated: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') };
}

/* 目次の前書き（作成者・注意書き）と参考（[KR]/[Global] の説明）だけを使う */
function readIntroValues(v) {
  var out = [], ref = false;
  v.forEach(function (r) {
    var c = r.filter(function (x) { return String(x).trim(); });
    if (!c.length) return;
    if (c[0] === '目次') { ref = null; return; }
    if (c[0] === '参考') { ref = true; return; }
    if (ref === null) return;
    out.push(c.map(function (x) { return String(x).trim(); }));
  });
  return out;
}

function readSheetRows(sh, imgMap) {
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return [];
  var rg = sh.getRange(1, 1, lr, lc);
  var val = rg.getDisplayValues(), bg = rg.getBackgrounds(), fw = rg.getFontWeights(), fc = rg.getFontColors(),
      fm = rg.getFormulas(), rich = rg.getRichTextValues();
  var span = {}, cover = {};
  rg.getMergedRanges().forEach(function (m) {
    var r0 = m.getRow() - 1, c0 = m.getColumn() - 1;
    span[r0 + ':' + c0] = m.getNumColumns();
    for (var r = r0; r < r0 + m.getNumRows(); r++) for (var c = c0; c < c0 + m.getNumColumns(); c++) if (r !== r0 || c !== c0) cover[r + ':' + c] = 1;
  });
  var out = [];
  for (var r = 0; r < lr; r++) {
    var cells = [];
    for (var c = 0; c < lc; c++) {
      if (cover[r + ':' + c]) continue;
      var t = String(val[r][c] || '').replace(/\s+$/, '');
      var link = hyperlinkOf(fm[r][c], rich[r][c]);
      if (!t) continue;
      var o = { x: c, t: t }, s = span[r + ':' + c];
      if (s > 1) o.s = s;
      var tone = bgTone(bg[r][c]);
      if (tone === 'n') o.n = 1; else if (tone === 'h') o.h = 1;
      if (fw[r][c] === 'bold') o.b = 1;
      if (isRed(fc[r][c])) o.e = 1; else { var em = redRuns(rich[r][c]); if (em.length) o.e = em; }
      if (link) applyLink(o, link, imgMap);
      o.t = o.t.replace(/\s*🔍$/, '');
      cells.push(o);
    }
    if (cells.length) out.push({ r: r + 1, c: cells });
  }
  return out;
}

function hyperlinkOf(formula, rich) {
  var m = String(formula || '').match(/^=HYPERLINK\(\s*"([^"]+)"/i);
  if (m) return m[1];
  try { var u = rich && rich.getLinkUrl(); if (u) return u; } catch (e) {}
  return '';
}

/* 背景色：濃い色 = 章見出し(n)、白以外の薄い色 = 表の見出し(h) */
function bgTone(hex) {
  var h = String(hex || '').replace('#', '');
  if (h.length !== 6 || /^f{6}$/i.test(h)) return '';
  var r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) < 110 ? 'n' : 'h';
}

function isRed(hex) {
  var h = String(hex || '').replace('#', '');
  if (h.length !== 6) return false;
  var r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
  return r > 150 && g < 90 && b < 90;
}

function redRuns(rich) {
  var out = [];
  if (!rich) return out;
  try {
    rich.getRuns().forEach(function (run) {
      var st = run.getTextStyle(), col = st && st.getForegroundColor && st.getForegroundColor();
      if (col && isRed(col) && run.getText().trim()) out.push([run.getStartIndex(), run.getEndIndex()]);
    });
  } catch (e) {}
  return out;
}

/* ========== 変更の検知 ========== */

/* DB_ガイドを編集すると呼ばれる（setupGuide で設定するトリガー）。数分後の変換を予約する */
function onGuideChange() {
  var P = PropertiesService.getScriptProperties();
  P.setProperty('GUIDE_DIRTY', String(Date.now()));
  var pending = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'rebuildGuide'; });
  if (!pending) ScriptApp.newTrigger('rebuildGuide').timeBased().after(GUIDE_DELAY_MIN * 60000).create();
}

/* 予約した変換（エディタから手動で実行しても良い） */
function rebuildGuide() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'rebuildGuide') ScriptApp.deleteTrigger(t); });
  var P = PropertiesService.getScriptProperties(), dirty = Number(P.getProperty('GUIDE_DIRTY')) || 0;
  if (dirty && Date.now() - dirty < GUIDE_DELAY_MIN * 60000 - 5000) {   // まだ編集中 → もう少し待つ
    ScriptApp.newTrigger('rebuildGuide').timeBased().after(GUIDE_DELAY_MIN * 60000).create();
    return;
  }
  var t0 = Date.now(), r = saveGuide(readGuide());
  P.deleteProperty('GUIDE_DIRTY');
  console.log('ガイドを変換しました：' + r.guide.tabs.length + 'タブ（' + Math.round((Date.now() - t0) / 1000) + '秒）');
}

/* エディタで1回実行：DB_ガイド編集時のトリガー＋採集マップ画像の作成 */
function setupGuide() {
  var id = prop('GUIDE_SHEET_ID');
  if (!id) throw new Error('スクリプト プロパティに GUIDE_SHEET_ID を追加してください');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'onGuideChange') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onGuideChange').forSpreadsheet(id).onChange().create();
  console.log('トリガー：DB_ガイドの変更を検知します');
  prepareMapImages();
  PropertiesService.getScriptProperties().deleteProperty('GUIDE_DIRTY');
  rebuildGuide();
}

/* 同じ地図の古い画像（png/jpg・一覧用）をゴミ箱へ */
function cleanMapFiles(folder, base) {
  ['.png', '.jpg', '_s.png', '_s.jpg'].forEach(function (x) { var it = folder.getFilesByName(base + x); while (it.hasNext()) it.next().setTrashed(true); });
}

/* 採集マップの元画像（7MB前後）から表示用の縮小版を作り、非公開の画像フォルダに保存する */
function prepareMapImages() {
  var ss = SpreadsheetApp.openById(prop('GUIDE_SHEET_ID')), sh = ss.getSheetByName('天族採集マップ');
  if (!sh) { console.log('天族採集マップ タブがありません'); return; }
  var folder = DriveApp.getFolderById(prop('IMAGE_FOLDER_ID')), token = ScriptApp.getOAuthToken(), map = {}, n = 0;
  var fm = sh.getDataRange().getFormulas();
  fm.forEach(function (row) {
    row.forEach(function (f) {
      var m = String(f).match(/^=HYPERLINK\(\s*"[^"]*\/d\/([\w-]{20,})/i);
      if (!m || map[m[1]]) return;
      var id = m[1];
      try {
        var meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?fields=name,thumbnailLink',
          { headers: { Authorization: 'Bearer ' + token } }).getContentText());
        if (!meta.thumbnailLink) throw new Error('縮小画像を作れませんでした');
        var base = 'map_' + String(meta.name).replace(/\.[^.]+$/, ''), link = meta.thumbnailLink.replace(/=s\d+.*$/, ''), name = base + '.jpg';
        cleanMapFiles(folder, base);
        [['', 2000], ['_s', 640]].forEach(function (v) {   // 拡大表示用（2000px）と一覧用（640px）。JPEGで受け取る
          var opt = { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true },
              res = UrlFetchApp.fetch(link + '=s' + v[1] + '-rj', opt);
          if (res.getResponseCode() !== 200) res = UrlFetchApp.fetch(link + '=s' + v[1], opt);   // JPEG指定が使えない場合
          if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
          var blob = res.getBlob();
          var f = base + v[0] + (/png/.test(blob.getContentType()) ? '.png' : '.jpg');
          folder.createFile(blob.setName(f));
          if (!v[0]) name = f;
          console.log('作成：' + f + '（' + Math.round(blob.getBytes().length / 1024) + 'KB）');
        });
        map[id] = name; n++;
      } catch (e) { console.error('画像を作れませんでした：' + id, e); }
    });
  });
  PropertiesService.getScriptProperties().setProperty('IMG_MAP', JSON.stringify(map));
  console.log('採集マップ画像：' + n + '件');
}
