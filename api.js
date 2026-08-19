// ==========================================
// 1. 変数定義（必ず一番上で定義）[過去の安定構造に復元]
// ==========================================
const SAAS_MASTER_SS_ID = "1JNDUYWZLkxF8cEW8FXiIflkAIGE6DcMbXqQjO64NgXM";
const TEMPLATE_SS_ID = "1mmQXbcUOGoKIlM6qWSGclY3G--BfkRTrKx5aG2vFNJY";
const PARENT_FOLDER_ID = "1Is1y-S5vWWjtkjha8KTXOPL3yxSLMJ_G";
const SHEET_COMPANIES = "SaaS管理マスターDB";

// ==========================================
// 2. ユーティリティ
// ==========================================
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

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

// ==========================================
// 3. APIルーティング
// ==========================================
// ★過去のコード通り、doGet で一覧を返すように復元
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return createJsonResponse({ status: "success", companies: [] });
    
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getDisplayValues();
    const companies = data.map(r => ({ companyId: r[0], companyName: r[1], createdAt: r[11] }));
    return createJsonResponse({ status: "success", companies: companies });
  } catch (err) { 
    return createJsonResponse({ status: "error", message: err.message }); 
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return createJsonResponse({ status: "error", message: "データが空です" });
    let payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === "registerCompany") return handleRegisterCompany(payload);
    if (action === "deleteCompany") return handleDeleteCompany(payload);
    if (action === "getTenantInfo") return handleGetTenantInfo(payload); 
    if (action === "login_admin") return handleLoginAdmin(payload);
    if (action === "login_staff") return handleLoginStaff(payload);

    return createJsonResponse({ status: "error", message: "不明なアクション: " + action });
  } catch (err) { 
    return createJsonResponse({ status: "error", message: "エラー: " + err.message }); 
  }
}

// ==========================================
// 4. データ操作・登録ロジック
// ==========================================
function handleDeleteCompany(payload) {
  const compId = String(payload.companyId || "").trim().toUpperCase();
  const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return createJsonResponse({ status: "error", message: "テナントが見つかりません" });

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim().toUpperCase() === compId) {
      const folderId = data[i][6]; 
      let msgExt = "";
      try { if (folderId) DriveApp.getFolderById(folderId).setTrashed(true); } catch(e) { msgExt = " (ドライブフォルダ削除スキップ)"; }
      sheet.deleteRow(i + 2); 
      return createJsonResponse({ status: "success", message: `テナント [${compId}] を削除しました。${msgExt}` });
    }
  }
  return createJsonResponse({ status: "error", message: "テナントが見つかりません" });
}

function handleGetTenantInfo(payload) {
  const compId = String(payload.companyId || "").trim().toUpperCase();
  const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return createJsonResponse({ status: "error", message: `企業ID [${compId}] が見つかりません。` });

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === compId) {
      return createJsonResponse({ status: "success", dbId: data[i][5], folderMeisai: data[i][7], folderLedger: data[i][8], folderTime: data[i][9], folderMember: data[i][10] });
    }
  }
  return createJsonResponse({ status: "error", message: `企業ID [${compId}] が見つかりません。` });
}

// ==========================================
// 5. アプリログイン用認証ロジック
// ==========================================
function handleLoginAdmin(payload) {
  const compId = payload.companyId;
  const email = payload.email;
  const password = payload.password;
  if (!compId || !email || !password) return createJsonResponse({ status: "error", message: "必須項目を入力してください。" });

  const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
  const map = getHeaderMap(sheet);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  let tenantInfo = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][map["企業ID"]]).trim().toUpperCase() === String(compId).trim().toUpperCase() &&
        String(data[i][map["管理者メール"]]).trim() === String(email).trim() && 
        String(data[i][map["初期パスワード"]]).trim() === String(password).trim()) {
      tenantInfo = extractTenantInfo(data[i], map);
      tenantInfo.role = "admin";
      tenantInfo.userName = "テナント管理者";
      break;
    }
  }
  if (tenantInfo) return createJsonResponse({ status: "success", data: tenantInfo });
  return createJsonResponse({ status: "error", message: "企業ID、メールアドレス、またはパスワードが間違っています。" });
}

function handleLoginStaff(payload) {
  const compId = payload.companyId;
  if (!compId) return createJsonResponse({ status: "error", message: "企業IDを入力してください。" });

  const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
  const map = getHeaderMap(sheet);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  let tenantInfo = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][map["企業ID"]]).trim().toUpperCase() === String(compId).trim().toUpperCase()) {
      tenantInfo = extractTenantInfo(data[i], map);
      tenantInfo.role = "staff";
      tenantInfo.userName = "従業員"; 
      break;
    }
  }
  if (tenantInfo) return createJsonResponse({ status: "success", data: tenantInfo });
  return createJsonResponse({ status: "error", message: "無効な企業IDです。" });
}

// ==========================================
// 6. 新規登録ロジック
// ==========================================
function generateCompanyId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "CP-001";
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;
  ids.forEach(idStr => {
    const match = String(idStr).match(/CP-(\d+)/);
    if (match) { const num = parseInt(match[1], 10); if (num > maxNum) maxNum = num; }
  });
  return `CP-${String(maxNum + 1).padStart(3, '0')}`;
}

function handleRegisterCompany(payload) {
  const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
  const companyId = generateCompanyId(sheet);
  const adminEmail = String(payload.adminEmail || "").trim();
  const adminPass = String(payload.adminPass || "").trim();
  const fullAdminId = adminEmail;

  const rootFolderName = `[${companyId}] ${payload.companyName}`;
  const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const companyFolder = parentFolder.createFolder(rootFolderName);
  
  const folderMeisai = companyFolder.createFolder("01_給与明細");
  const folderLedger = companyFolder.createFolder("02_賃金台帳");
  const folderTime   = companyFolder.createFolder("03_出勤簿");
  const folderMember = companyFolder.createFolder("04_労働者名簿");
  
  const newDbFile = DriveApp.getFileById(TEMPLATE_SS_ID).makeCopy(`${payload.companyName}_業務統合データベース`, companyFolder);
  const newDbId = newDbFile.getId();
  const newDb = SpreadsheetApp.openById(newDbId);
  
  const empSheet = newDb.getSheetByName("社員マスタ");
  if(empSheet) {
    if(empSheet.getLastRow() > 1) empSheet.getRange(2, 1, empSheet.getLastRow() - 1, empSheet.getLastColumn()).clearContent();
    empSheet.appendRow([`${companyId}-EMP-001`, "システム管理者", "本部", "月給", 0, 0, 0, 0, 0, "なし", "有効", adminEmail, "", "", "", "管理者", fullAdminId, adminPass, "管理者"]);
  }

  const siteSheet = newDb.getSheetByName("現場マスタ");
  if(siteSheet) {
    if(siteSheet.getLastRow() > 1) siteSheet.getRange(2, 1, siteSheet.getLastRow() - 1, siteSheet.getLastColumn()).clearContent();
    siteSheet.appendRow([`SITE-000`, "本社（基本勤務地）", "自社", 0, "2026-01-01", "2030-12-31", "システム管理者", "進行中"]);
  }

  const setSheet = newDb.getSheetByName("設定マスタ");
  if (setSheet) {
    if (setSheet.getLastRow() <= 1) setSheet.appendRow([2026, 0.0506, 0.0915, 0.007, "あり", adminEmail, 8, 160, 1.25]);
    else setSheet.getRange(2, 6).setValue(adminEmail);
  }

  sheet.appendRow([
    companyId, payload.companyName, fullAdminId, adminPass, adminEmail, newDbId, companyFolder.getId(), 
    folderMeisai.getId(), folderLedger.getId(), folderTime.getId(), folderMember.getId(), new Date()
  ]);
  
  return createJsonResponse({ 
    status: "success", 
    message: `${payload.companyName} の環境を構築しました。\n\n【管理者ログイン】\nID: ${fullAdminId}\nPass: ${adminPass}`, 
    companyId: companyId,
    dbId: newDbId
  });
}