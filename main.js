// ==========================================
// SAAS マスター設定
// ==========================================
const SHEET_TENANT = "企業・テナント一覧";

// ==========================================
// GETリクエスト（管理者用ダッシュボードの表示）
// ==========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("現場のミカタ - SAASマスター管理")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getHeaderMap(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}