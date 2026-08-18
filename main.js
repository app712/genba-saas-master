// ==========================================
// 設定変数（必ず一番上で定義）
// ==========================================
const SHEET_TENANT = "企業・テナント一覧";
const TEMPLATE_SHEET_ID = "1SJLWG6neWEt1xpFwfaTokNQQl92Y4rPwvgogE-wp6VA";

// ==========================================
// GETリクエスト（ダミーHTMLの返却用）
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