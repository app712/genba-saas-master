// ==========================================
// POSTリクエスト（各クライアントからの認証・DB接続用API）
// ==========================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "データが空です。" });
    }
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    // クライアントからの初回ログイン時、企業ID（CP-001など）をもとに企業ごとのDBIDを返す
    if (action === "getTenantInfo") {
      const companyId = payload.companyId;
      if (!companyId) return createJsonResponse({ status: "error", message: "企業IDが指定されていません。" });

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = ss.getSheetByName(SHEET_TENANT);
      if (!tenantSheet) return createJsonResponse({ status: "error", message: "テナント情報シートが存在しません。" });

      const map = getHeaderMap(tenantSheet);
      const data = tenantSheet.getRange(2, 1, tenantSheet.getLastRow() - 1, tenantSheet.getLastColumn()).getValues();

      let tenantDbId = null;
      let companyName = "";

      for (let i = 0; i < data.length; i++) {
        // "企業ID" 列と "スプレッドシートID" 列をマッピング
        if (String(data[i][map["企業ID"]]).trim().toUpperCase() === String(companyId).trim().toUpperCase()) {
          tenantDbId = String(data[i][map["スプレッドシートID"]]).trim();
          companyName = String(data[i][map["企業名"]]).trim();
          break;
        }
      }

      if (tenantDbId) {
        return createJsonResponse({ status: "success", dbId: tenantDbId, companyName: companyName });
      } else {
        return createJsonResponse({ status: "error", message: "無効な企業IDです。契約状況を確認してください。" });
      }
    }

    return createJsonResponse({ status: "error", message: "不明なアクション: " + action });

  } catch (err) {
    return createJsonResponse({ status: "error", message: "マスターサーバーエラー: " + err.message });
  }
}