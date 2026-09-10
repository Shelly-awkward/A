/** 個案檢查表 網頁表單（Apps Script）— 工作表欄位與 bin/render_case_docs.py 的輸入活頁簿完全相同，
 *  匯出的 xlsx 直接放進 input\ 跑 3_產出檢查表.bat。
 *  安裝：以 accsfb01 開新試算表 → 擴充功能 → Apps Script → 貼上 Code.gs / Index.html / Seed.gs / appsscript.json
 *       → 執行一次 setup() → 部署 → 網頁應用程式（執行身分：我；存取：任何人）。 */
var SHEETS = {
  cases: ['公司代號','公司全名','公司簡稱','市場別','事務所','會計師A','會計師B','工作底稿冊數','電子底稿冊數','設立日期','上市櫃日期','主要營業項目','主查','覆核者','覆核日期'],
  chk2:  ['公司代號','公司簡稱','題號','章節','審閱內容(摘)','本局審查意見(V/✗/NA)','會計師工作底稿索引','備註','更新者','更新時間'],
  chk1:  ['列序','子項','章節','檢查項目','是','否','不適用','說明及發現','附件','更新者','更新時間'],
  aml:   ['表序','列序','類型','檢查項目','正常','異常','說明','更新者','更新時間'],
  findings: ['類型(個案/品質)','公司代號','公司簡稱','缺失分類','項次','會計師所涉查核疏失','會計師說明','分析意見','違反或相關之準則或法規','對應題號','IFIAR主題(選填)','id','填表人','更新時間'],
  firm:  ['項目','值']
};
var NAMES = { cases:'個案主檔', chk2:'檢查表二作答', chk1:'檢查表一作答', aml:'AML作答', findings:'缺失', firm:'事務所' };

function doGet() { return HtmlService.createHtmlOutputFromFile('Index').setTitle('個案檢查表').addMetaTag('viewport','width=device-width, initial-scale=1'); }

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh(key) {
  var s = ss().getSheetByName(NAMES[key]);
  if (!s) { s = ss().insertSheet(NAMES[key]); s.appendRow(SHEETS[key]); s.setFrozenRows(1); }
  return s;
}
function rows(key) {
  var s = sh(key), v = s.getDataRange().getValues(), h = v[0], out = [];
  for (var i = 1; i < v.length; i++) { if (v[i].join('') === '') continue; var o = { _row: i + 1 }; h.forEach(function (k, j) { o[k] = v[i][j]; }); out.push(o); }
  return out;
}
function withLock(fn) { var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return fn(); } finally { lock.releaseLock(); } }
function upsert(key, keyCols, rec) {
  return withLock(function () {
    var s = sh(key), h = SHEETS[key], v = s.getDataRange().getValues(), hit = -1;
    for (var i = 1; i < v.length; i++) { if (keyCols.every(function (k) { return String(v[i][h.indexOf(k)]) === String(rec[k]); })) { hit = i + 1; break; } }
    var line = h.map(function (k) { return rec[k] !== undefined ? rec[k] : (hit > 0 ? v[hit - 1][h.indexOf(k)] : ''); });
    if (hit > 0) s.getRange(hit, 1, 1, h.length).setValues([line]); else s.appendRow(line);
    return true;
  });
}

/** 第一次執行：建工作表、灌題庫（檢查表一、AML 為事務所層級，直接展開）、事務所基本欄。 */
function setup() {
  Object.keys(SHEETS).forEach(function (k) { sh(k); });
  var s1 = sh('chk1'); if (s1.getLastRow() < 2) s1.getRange(2, 1, SEED.chk1.length, 4).setValues(SEED.chk1.map(function (x) { return [x.row, x.sub, x.sec, x.text]; }));
  var sa = sh('aml'); if (sa.getLastRow() < 2) sa.getRange(2, 1, SEED.aml.length, 4).setValues(SEED.aml.map(function (x) { return [x.t, x.r, x.kind, x.text]; }));
  var sf = sh('firm'); if (sf.getLastRow() < 2) sf.getRange(2, 1, 5, 2).setValues([['事務所名稱',''],['檢查年度',''],['財報年度',''],['檢查者',''],['檢查日期','']]);
  var q = ss().getSheetByName('題庫二'); if (!q) { q = ss().insertSheet('題庫二'); q.appendRow(['題號','章節','審閱內容']); q.getRange(2, 1, SEED.chk2.length, 3).setValues(SEED.chk2.map(function (x) { return [x.q, x.sec, x.text]; })); }
  var r = ss().getSheetByName('人員'); if (!r) { r = ss().insertSheet('人員'); r.appendRow(['姓名','分機']); }
}

function getState() {
  var firm = {}; rows('firm').forEach(function (r) { firm[r['項目']] = r['值']; });
  var people = rows_('人員').map(function (r) { return r[0]; }).filter(String);
  return { firm: firm, cases: rows('cases'), chk2Items: SEED.chk2, chk2: rows('chk2'), chk1: rows('chk1'), aml: rows('aml'),
           findings: rows('findings'), caseCats: SEED.caseCats, qmCats: SEED.qmCats, people: people, version: SEED.version };
}
function rows_(name) { var s = ss().getSheetByName(name); if (!s || s.getLastRow() < 2) return []; return s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues(); }

/** 貼上選案工作表「建議選案」的列（Tab 分隔，欄序：代號 簡稱 市場別 會計師A 會計師B 或整列貼上皆可，程式找欄）。 */
function importCases(tsv, firmName) {
  var lines = tsv.split(/\r?\n/).map(function (l) { return l.split('\t'); }).filter(function (a) { return a.length >= 2; });
  var hdr = lines[0].map(String); var iCode = hdr.indexOf('公司代號'), iName = hdr.indexOf('公司簡稱'), iMkt = hdr.indexOf('市場別'), iA = hdr.indexOf('會計師A'), iB = hdr.indexOf('會計師B'), iOff = hdr.indexOf('承辦人');
  var body = iCode >= 0 ? lines.slice(1) : lines; if (iCode < 0) { iCode = 0; iName = 1; iMkt = 2; iA = 3; iB = 4; iOff = -1; }
  var existing = {}; rows('cases').forEach(function (c) { existing[String(c['公司代號'])] = 1; });
  var added = 0;
  body.forEach(function (a) {
    var code = String(a[iCode] || '').trim(); if (!code || existing[code] || !/^\d+$/.test(code)) return;
    upsert('cases', ['公司代號'], { '公司代號': code, '公司簡稱': a[iName] || '', '市場別': a[iMkt] || '', '事務所': firmName || '', '會計師A': a[iA] || '', '會計師B': a[iB] || '', '主查': iOff >= 0 ? a[iOff] || '' : '' });
    var s = sh('chk2'); s.getRange(s.getLastRow() + 1, 1, SEED.chk2.length, 5).setValues(SEED.chk2.map(function (x) { return [code, a[iName] || '', x.q, x.sec, x.text.split('\n')[0].slice(0, 40)]; }));
    added++;
  });
  return added;
}
function saveCase(rec) { return upsert('cases', ['公司代號'], rec); }
function saveChk2(code, q, opinion, idx, note, user) { return upsert('chk2', ['公司代號', '題號'], { '公司代號': code, '題號': q, '本局審查意見(V/✗/NA)': opinion, '會計師工作底稿索引': idx, '備註': note, '更新者': user || '', '更新時間': new Date() }); }
function saveChk1(row, sub, vals, user) { vals['列序'] = row; vals['子項'] = sub; vals['更新者'] = user || ''; vals['更新時間'] = new Date(); return upsert('chk1', ['列序', '子項'], vals); }
function saveAml(t, r, vals, user) { vals['表序'] = t; vals['列序'] = r; vals['更新者'] = user || ''; vals['更新時間'] = new Date(); return upsert('aml', ['表序', '列序'], vals); }
function saveFirm(k, v) { return upsert('firm', ['項目'], { '項目': k, '值': v }); }
function saveFinding(rec, user) {
  if (!rec['id']) rec['id'] = Utilities.getUuid();
  rec['更新時間'] = new Date(); rec['填表人'] = user || '';
  if (!rec['項次']) { var n = rows('findings').filter(function (f) { return f['類型(個案/品質)'] === rec['類型(個案/品質)'] && String(f['公司代號'] || '') === String(rec['公司代號'] || ''); }).length; rec['項次'] = n + 1; }
  upsert('findings', ['id'], rec); return rec['id'];
}
function deleteFinding(id) {
  return withLock(function () { var s = sh('findings'), v = s.getDataRange().getValues(), i = SHEETS.findings.indexOf('id');
    for (var r = v.length - 1; r >= 1; r--) if (String(v[r][i]) === String(id)) { s.deleteRow(r + 1); return true; } return false; });
}

/** 匯出整本試算表為 xlsx，檔名即產出程式要的 個案檢查_輸入_<年度><事務所2字>.xlsx；存到雲端硬碟「檢查輸入匯出」資料夾。 */
function exportXlsx() {
  var firm = {}; rows('firm').forEach(function (r) { firm[r['項目']] = r['值']; });
  var name = '個案檢查_輸入_' + (firm['檢查年度'] || '') + String(firm['事務所名稱'] || '').slice(0, 2) + '.xlsx';
  var url = 'https://docs.google.com/spreadsheets/d/' + ss().getId() + '/export?format=xlsx';
  var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName(name);
  var it = DriveApp.getFoldersByName('檢查輸入匯出'); var folder = it.hasNext() ? it.next() : DriveApp.createFolder('檢查輸入匯出');
  var old = folder.getFilesByName(name); while (old.hasNext()) old.next().setTrashed(true);
  var f = folder.createFile(blob); return { name: name, url: f.getUrl() };
}

/** 選用：把缺失同步到「缺失申報站」（同一試算表內名為 findings 的工作表）。已同步者以 id 對應不重複。 */
function syncToIntake() {
  var t = ss().getSheetByName('findings'); if (!t) return '找不到缺失申報站的 findings 工作表（需在同一試算表）';
  var h = t.getDataRange().getValues()[0], have = {}; t.getDataRange().getValues().slice(1).forEach(function (r) { have[String(r[h.indexOf('id')])] = 1; });
  var firm = {}; rows('firm').forEach(function (r) { firm[r['項目']] = r['值']; }); var n = 0;
  rows('findings').forEach(function (f) {
    if (have[String(f['id'])]) return;
    var rec = { id: f['id'], year: Number(firm['檢查年度'] || 0) + 1911, type: f['類型(個案/品質)'] === '個案' ? '審計個案缺失' : '品質管理缺失', firm: firm['事務所名稱'] || '',
      companyCode: f['公司代號'] || '', company: f['公司簡稱'] || '', content: f['會計師所涉查核疏失'] || '', localCategory: f['缺失分類'] || '', themeCode: '', themeZh: f['IFIAR主題(選填)'] || '', filler: f['填表人'] || '', createdAt: new Date().toISOString() };
    t.appendRow(h.map(function (k) { return rec[k] !== undefined ? rec[k] : ''; })); n++;
  });
  return '已同步 ' + n + ' 筆';
}
