/**
 * Ragdoll ナビ — ガイド（「00.AION2 DB_ガイド」を読み、ツールのガイドタブ用に変換して渡す）
 *
 * スクリプト プロパティ
 *   GUIDE_SHEET_ID   DB_ガイドのスプレッドシートID（必須）
 *   GUIDE_VER        ガイドの版（自動。DB_ガイドを編集すると更新）
 *   IMG_MAP          採集マップの元画像ID → 画像フォルダのファイル名（prepareMapImages が作成）
 *
 * 初回だけエディタで setupGuide を実行する（DB_ガイド編集時のトリガー設定＋採集マップ画像の作成）。
 * 内容（文章・表）はシートの値と書式から作るだけで、シートやリンクのURLは渡さない。
 */

var GUIDE_SKIP = ['目次'];
var GUIDE_CHUNK = 90000;      // CacheService の1件あたり上限（100KB）より少し小さく
var GUIDE_CACHE_SEC = 21600;  // 6時間（版が変わるとキーが変わる）

function actGuide(req) {
  var v = checkSession(req.token);
  if (!v.ok) return v;
  var ver = guideVer();
  if (req.ver && Number(req.ver) === ver) return { ok: true, token: v.token, ver: ver, same: true };
  return { ok: true, token: v.token, ver: ver, guide: buildGuide(ver) };
}

function guideVer() { return Number(prop('GUIDE_VER')) || 1; }

/* 変換結果は分割してキャッシュ（大きいため） */
function buildGuide(ver) {
  var c = CacheService.getScriptCache(), base = 'guide:' + ver + ':', n = Number(c.get(base + 'n'));
  if (n) {
    var keys = []; for (var i = 0; i < n; i++) keys.push(base + i);
    var got = c.getAll(keys), parts = [];
    for (var j = 0; j < n; j++) { if (!got[base + j]) { parts = null; break; } parts.push(got[base + j]); }
    if (parts) return JSON.parse(parts.join(''));
  }
  var g = readGuide(), s = JSON.stringify(g), put = {};
  for (var k = 0, m = 0; k < s.length; k += GUIDE_CHUNK, m++) put[base + m] = s.slice(k, k + GUIDE_CHUNK);
  put[base + 'n'] = String(m);
  try { c.putAll(put, GUIDE_CACHE_SEC); } catch (e) { console.error('guide cache', e); }
  return g;
}

/* ========== 読み込み・変換 ========== */
/* 1行 = { r: 行番号, c: [セル] }。セル = { x: 列(0始まり), s: 横の結合数, t: 文字,
 *   h: 見出し色, n: 章見出し色, b: 太字, e: 強調(全体=1 / 一部=[[開始,終了],…]), l: 外部リンク, j: シート内リンク, i: 画像ファイル名 } */
function readGuide() {
  var ss = SpreadsheetApp.openById(prop('GUIDE_SHEET_ID')), imgMap = {};
  try { imgMap = JSON.parse(prop('IMG_MAP') || '{}'); } catch (e) {}
  var tabs = [], intro = [];
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (GUIDE_SKIP.indexOf(name) >= 0) { intro = readIntro(sh); return; }
    tabs.push({ name: name, gid: sh.getSheetId(), rows: readSheetRows(sh, imgMap) });
  });
  return { tabs: tabs, intro: intro, updated: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') };
}

/* 目次の前書き（作成者・注意書き）と参考（[KR]/[Global] の説明）だけを使う */
function readIntro(sh) {
  var v = sh.getDataRange().getDisplayValues(), out = [], ref = false;
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
      if (link) {
        var m1 = link.match(/^#gid=(\d+)(?:&range=[A-Z]+(\d+))?/), m2 = link.match(/\/d\/([\w-]{20,})/);
        if (m1) o.j = m1[1] + (m1[2] ? ':' + m1[2] : '');
        else if (m2 && /drive\.google\.com/.test(link)) { if (imgMap[m2[1]]) o.i = imgMap[m2[1]]; }   // 元画像のURLは渡さない
        else if (/^https:\/\//.test(link) && !/docs\.google\.com|drive\.google\.com/.test(link)) o.l = link;
      }
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

/* DB_ガイドを編集すると呼ばれる（setupGuide で設定するトリガー） */
function onGuideChange() {
  var P = PropertiesService.getScriptProperties(), now = String(Date.now());
  P.setProperty('GUIDE_VER', now);
  P.setProperty('DATA_VER', now);   // ツールの「更新」に赤い点を出す
  CacheService.getScriptCache().removeAll(['data']);
}

/* エディタで1回実行：DB_ガイド編集時のトリガー＋採集マップ画像の作成 */
function setupGuide() {
  var id = prop('GUIDE_SHEET_ID');
  if (!id) throw new Error('スクリプト プロパティに GUIDE_SHEET_ID を追加してください');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'onGuideChange') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onGuideChange').forSpreadsheet(id).onChange().create();
  console.log('トリガー：DB_ガイドの変更を検知します');
  prepareMapImages();
  onGuideChange();
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
