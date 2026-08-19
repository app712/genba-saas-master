// ==========================================
// POSTリクエスト（登録、認証、テストデータ投入）
// ==========================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "データが空です。" });
    }
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === "register") return createNewTenant(payload.companyName, payload.email, payload.password);

    if (action === "inject_test_data") {
      const dbId = payload.dbId;
      if (!dbId) return createJsonResponse({ status: "error", message: "データベースIDが指定されていません。" });
      try {
        const clientSs = SpreadsheetApp.openById(dbId);
        let staffSheet = clientSs.getSheetByName("社員マスタ") || clientSs.insertSheet("社員マスタ");
        staffSheet.clear();
        staffSheet.getRange(1, 1, 4, 4).setValues([
          ["社員ID", "パスワード", "氏名", "権限"],
          ["EMP-001", "1234", "テスト 太郎", "staff"],
          ["EMP-002", "1234", "テスト 次郎", "staff"],
          ["ADMIN-01", "admin123", "現場 管理者", "admin"]
        ]);
        let timeSheet = clientSs.getSheetByName("出退勤記録") || clientSs.insertSheet("出退勤記録");
        if(timeSheet.getLastRow() === 0) timeSheet.appendRow(["日時", "氏名", "打刻種類", "位置情報"]);
        let reportSheet = clientSs.getSheetByName("業務日報") || clientSs.insertSheet("業務日報");
        if(reportSheet.getLastRow() === 0) reportSheet.appendRow(["送信日時", "氏名", "現場名", "報告内容"]);
        return createJsonResponse({ status: "success", message: "テストデータの投入完了。社員ID: EMP-001 / PASS: 1234 でログイン可能です。" });
      } catch (err) {
        return createJsonResponse({ status: "error", message: "投入エラー: " + err.message });
      }
    }

    const props = PropertiesService.getScriptProperties();
    const masterSsId = props.getProperty("MASTER_SS_ID");
    const sheetTenant = props.getProperty("SHEET_TENANT");
    const masterSs = SpreadsheetApp.openById(masterSsId);
    const tenantSheet = masterSs.getSheetByName(sheetTenant);
    if (!tenantSheet) return createJsonResponse({ status: "error", message: "マスターDBが存在しません。" });

    const tMap = getHeaderMap(tenantSheet);
    const tData = tenantSheet.getRange(2, 1, tenantSheet.getLastRow() - 1, tenantSheet.getLastColumn()).getValues();

    // 管理者ログイン（企業ID + メールアドレス + パスワード）
    if (action === "login_admin") {
      const compId = payload.companyId;
      const email = payload.email;
      const password = payload.password;
      if (!compId || !email || !password) return createJsonResponse({ status: "error", message: "必須項目を入力してください。" });

      let tenantInfo = null;
      for (let i = 0; i < tData.length; i++) {
        if (String(tData[i][tMap["企業ID"]]).trim().toUpperCase() === String(compId).trim().toUpperCase() &&
            String(tData[i][tMap["管理者メール"]]).trim() === String(email).trim() && 
            String(tData[i][tMap["初期パスワード"]]).trim() === String(password).trim()) {
          tenantInfo = extractTenantInfo(tData[i], tMap);
          tenantInfo.role = "admin";
          tenantInfo.userName = "テナント管理者";
          break;
        }
      }
      if (tenantInfo) return createJsonResponse({ status: "success", data: tenantInfo });
      return createJsonResponse({ status: "error", message: "企業ID、メールアドレス、またはパスワードが間違っています。" });
    }

    // 一般社員ログイン（企業ID + 個人ID + パスワード）
    if (action === "login_staff") {
      const compId = payload.companyId;
      const userId = payload.userId;
      const password = payload.password;
      if (!compId || !userId || !password) return createJsonResponse({ status: "error", message: "必須項目を入力してください。" });

      let tenantInfo = null;
      for (let i = 0; i < tData.length; i++) {
        if (String(tData[i][tMap["企業ID"]]).trim().toUpperCase() === String(compId).trim().toUpperCase()) {
          tenantInfo = extractTenantInfo(tData[i], tMap);
          break;
        }
      }
      if (!tenantInfo) return createJsonResponse({ status: "error", message: "無効な企業IDです。" });

      try {
        const clientSs = SpreadsheetApp.openById(tenantInfo.dbId);
        const staffSheet = clientSs.getSheetByName("社員マスタ");
        const sMap = getHeaderMap(staffSheet);
        const sData = staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, staffSheet.getLastColumn()).getValues();
        let isValidUser = false, userName = "", userRole = "staff";
        for (let j = 0; j < sData.length; j++) {
          if (String(sData[j][sMap["社員ID"]]).trim() === String(userId).trim() &&
              String(sData[j][sMap["パスワード"]]).trim() === String(password).trim()) {
            isValidUser = true;
            userName = String(sData[j][sMap["氏名"]]).trim();
            if (sMap["権限"] !== undefined) userRole = String(sData[j][sMap["権限"]]).trim() || "staff";
            break;
          }
        }
        if (isValidUser) {
          tenantInfo.role = userRole;
          tenantInfo.userName = userName;
          return createJsonResponse({ status: "success", data: tenantInfo });
        } else {
          return createJsonResponse({ status: "error", message: "社員IDまたはパスワードが間違っています。" });
        }
      } catch (e) {
        return createJsonResponse({ status: "error", message: "企業データベースへのアクセスに失敗しました。" });
      }
    }

    return createJsonResponse({ status: "error", message: "不明なアクション: " + action });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "マスターサーバーエラー: " + err.message });
  }
}

function createJsonResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function getHeaderMap(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}

function extractTenantInfo(row, map) {
  return {
    companyId: String(row[map["企業ID"]]).trim(),
    companyName: String(row[map["企業名"]]).trim(),
    dbId: String(row[map["企業用DB(SS)_ID"]]).trim(),
    rootFolderId: String(row[map["ルートフォルダ_ID"]]).trim(),
    meisaiFolderId: String(row[map["明細フォルダ_ID"]]).trim(),
    daichoFolderId: String(row[map["台帳フォルダ_ID"]]).trim(),
    shukkinFolderId: String(row[map["出勤簿フォルダ_ID"]]).trim(),
    meiboFolderId: String(row[map["名簿フォルダ_ID"]]).trim()
  };
}

function generateNextCompanyId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "CP-001";
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;
  for (let i = 0; i < data.length; i++) {
    const val = String(data[i][0]).trim();
    if (val.startsWith("CP-")) {
      const num = parseInt(val.replace("CP-", ""), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  return "CP-" + ("000" + (maxNum + 1)).slice(-3);
}

function createNewTenant(name, email, password) {
  try {
    const props = PropertiesService.getScriptProperties();
    const ss = SpreadsheetApp.openById(props.getProperty("MASTER_SS_ID"));
    const tenantSheet = ss.getSheetByName(props.getProperty("SHEET_TENANT"));
    const newId = generateNextCompanyId(tenantSheet);
    const initialAdminId = newId + "-admin";
    const parentFolder = DriveApp.getFolderById(props.getProperty("PARENT_FOLDER_ID"));
    const rootFolder = parentFolder.createFolder(name + "様_システム一式");
    const meisaiFolder = rootFolder.createFolder("給与明細");
    const daichoFolder = rootFolder.createFolder("賃金台帳");
    const shukkinFolder = rootFolder.createFolder("出勤簿");
    const meiboFolder = rootFolder.createFolder("労働者名簿");
    const newSheet = DriveApp.getFileById(props.getProperty("TEMPLATE_SHEET_ID")).makeCopy(name + "様_データベース", rootFolder);
    tenantSheet.appendRow([newId, name, initialAdminId, password, email, newSheet.getId(), rootFolder.getId(), meisaiFolder.getId(), daichoFolder.getId(), shukkinFolder.getId(), meiboFolder.getId(), new Date()]);
    return createJsonResponse({ status: "success", message: `構築完了: [${newId}] ${name} の環境を生成しました。` });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "環境構築エラー: " + err.message });
  }
}