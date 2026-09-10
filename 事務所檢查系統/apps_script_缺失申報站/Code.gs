/**
 * 缺失申報站 — Apps Script 版
 * 綁定試算表：一列一筆紀錄，欄位對應 IFIAR Part I/III/IV 分類。
 * 部署方式與欄位說明見附的部署手冊。
 */

var SHEET_NAME = "findings";
var ROSTER_SHEET_NAME = "roster";

var HEADERS = [
  "id", "year", "type", "themeCode", "themeEn", "themeZh",
  "subZh", "subOther", "firm", "companyCode", "company",
  "content", "localCategory", "handler", "filler", "createdAt"
];

function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("缺失申報站")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 讓 HTML 檔案之間可以互相 include（本專案只有一個檔案，先留著備用） */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 名冊分頁（選填）：第一欄為「填表人」下拉選單的建議名單，供前端讀取。 */
function getRoster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  return values.map(function (r) { return String(r[0]).trim(); }).filter(Boolean);
}

/**
 * 新增一筆缺失紀錄。用 LockService 避免多人同時送出時互相覆蓋或漏行。
 */
function submitFinding(record) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // 最多等 10 秒
  try {
    var sheet = getSheet_();
    var id = Utilities.getUuid();
    var row = HEADERS.map(function (key) {
      if (key === "id") return id;
      if (key === "createdAt") return new Date().toISOString();
      var v = record[key];
      return (v === undefined || v === null) ? "" : v;
    });
    sheet.appendRow(row);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

/** 刪除一筆紀錄（依 id 找列並刪除）。 */
function deleteFinding(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: false, reason: "empty" };
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        return { ok: true };
      }
    }
    return { ok: false, reason: "not_found" };
  } finally {
    lock.releaseLock();
  }
}

/** 回傳所有紀錄，最新的在前面，供前端渲染清單與統計。 */
function getAllFindings() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = HEADERS.length;
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var records = values.map(function (row) {
    var rec = {};
    HEADERS.forEach(function (key, i) { rec[key] = row[i]; });
    return rec;
  });
  records.reverse(); // newest first
  return records;
}
