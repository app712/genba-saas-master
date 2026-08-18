// ==========================================
// 設定変数（変更する内容）
// ==========================================
const SHEET_TENANT = "SaaS管理マスターDB"; // 管理マスターのシート名
const TEMPLATE_SHEET_ID = "1JNDUYWZLkxF8cEW8FXiIflkAIGE6DcMbXqQjO64NgXM";
const PARENT_FOLDER_ID = "1Is1y-S5vWWjtkjha8KTXOPL3yxSLMJ_G"; // 新規追加

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