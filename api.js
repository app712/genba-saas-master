// ==========================================
// 1. 変数定義（必ず一番上で定義）
// ==========================================
const SAAS_MASTER_SS_ID = "1JNDUYWZLkxF8cEW8FXiIflkAIGE6DcMbXqQjO64NgXM";
const TEMPLATE_SS_ID = "1mmQXbcUOGoKIlM6qWSGclY3G--BfkRTrKx5aG2vFNJY";
const PARENT_FOLDER_ID = "1Is1y-S5vWWjtkjha8KTXOPL3yxSLMJ_G";
const SHEET_COMPANIES = "SaaS管理マスターDB";

// ==========================================
// 2. システム自己診断 ＆ 権限承認用
// ==========================================
function runSelfDiagnostic() {
  Logger.log("=== システム自己診断を開始します ===");
  try {
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    Logger.log("✓ 親フォルダ接続成功: " + parentFolder.getName());

    const templateFile = DriveApp.getFileById(TEMPLATE_SS_ID);
    Logger.log("✓ テンプレートファイル接続成功: " + templateFile.getName());

    const masterSs = SpreadsheetApp.openById(SAAS_MASTER_SS_ID);
    Logger.log("✓ マスターDB接続成功: " + masterSs.getName());

    let sheet = masterSs.getSheetByName(SHEET_COMPANIES);
    if (!sheet) {
      setupSaaSBase();
    }
    Logger.log("=== すべての自己診断テストをクリアしました。エラーはありません。 ===");
  } catch (err) {
    Logger.log("✖ 診断エラー発生: " + err.message);
  }
}

// ==========================================
// 3. 初期化 ＆ 基本設定
// ==========================================
function setupSaaSBase() {
  DriveApp.getRootFolder(); 
  GmailApp.getInboxUnreadCount(); 
  const ss = SpreadsheetApp.openById(SAAS_MASTER_SS_ID);
  let sheet = ss.getSheetByName(SHEET_COMPANIES);
  if (!sheet) {
    const defaultSheet = ss.getSheetByName("シート1");
    if (defaultSheet) defaultSheet.setName(SHEET_COMPANIES);
    else sheet = ss.insertSheet(SHEET_COMPANIES);
    sheet = ss.getSheetByName(SHEET_COMPANIES);
  }
  const headers = ["企業ID", "企業名", "初期管理者ID", "初期パスワード", "管理者メール", "企業用DB(SS)_ID", "ルートフォルダ_ID", "明細フォルダ_ID", "台帳フォルダ_ID", "出勤簿フォルダ_ID", "名簿フォルダ_ID", "登録日時"];
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
}

// ==========================================
// 4. ユーティリティ
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
// 5. APIルーティング
// ==========================================
function doGet(e) {
  return handleGetCompanies(); 
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return createJsonResponse({ status: "error", message: "データが空です" });
    let payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === "getCompanies") return handleGetCompanies(); 
    if (action === "registerCompany") return handleRegisterCompany(payload);
    if (action === "deleteCompany") return handleDeleteCompany(payload);
    if (action === "getTenantInfo") return handleGetTenantInfo(payload); 
    if (action === "login_admin") return handleLoginAdmin(payload);
    if (action === "login_staff") return handleLoginStaff(payload);

    return createJsonResponse({ status: "error", message: "不明なアクション: " + action });
  } catch (err) { 
    return createJsonResponse({ status: "error", message: "POST処理エラー: " + err.message }); 
  }
}

// ==========================================
// 6. データ操作・照会ロジック
// ==========================================
function handleGetCompanies() {
  try {
    const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
    if (!sheet) return createJsonResponse({ status: "error", message: "マスターDBのシートが見つかりません" });

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return createJsonResponse({ status: "success", companies: [] });
    
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getDisplayValues();
    const companies = data.map(r => ({ companyId: r[0], companyName: r[1], createdAt: r[11] }));
    return createJsonResponse({ status: "success", companies: companies });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "一覧取得エラー: " + err.message });
  }
}

function handleGetTenantInfo(payload) {
  try {
    const compId = String(payload.companyId || "").trim().toUpperCase();
    const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
    if (!sheet) return createJsonResponse({ status: "error", message: "マスターDBのシートが見つかりません" });

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return createJsonResponse({ status: "error", message: `企業ID [${compId}] が見つかりません。` });

    const data = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === compId) {
        return createJsonResponse({ status: "success", dbId: data[i][5], folderMeisai: data[i][7], folderLedger: data[i][8], folderTime: data[i][9], folderMember: data[i][10] });
      }
    }
    return createJsonResponse({ status: "error", message: `企業ID [${compId}] が見つかりません。` });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "テナント情報取得エラー: " + err.message });
  }
}

// ==========================================
// 7. アプリ連携用ログインロジック
// ==========================================
function handleLoginAdmin(payload) {
  try {
    const compId = payload.companyId;
    const email = payload.email;
    const password = payload.password;
    if (!compId || !email || !password) return createJsonResponse({ status: "error", message: "必須項目を入力してください。" });

    const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
    if (!sheet) return createJsonResponse({ status: "error", message: "マスターDBが見つかりません" });

    const map = getHeaderMap(sheet);
    if (sheet.getLastRow() <= 1) return createJsonResponse({ status: "error", message: "登録企業がありません。" });

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
  } catch (err) {
    return createJsonResponse({ status: "error", message: "ログイン処理エラー: " + err.message });
  }
}

function handleLoginStaff(payload) {
  try {
    const compId = payload.companyId;
    if (!compId) return createJsonResponse({ status: "error", message: "企業IDを入力してください。" });

    const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
    if (!sheet) return createJsonResponse({ status: "error", message: "マスターDBが見つかりません" });

    const map = getHeaderMap(sheet);
    if (sheet.getLastRow() <= 1) return createJsonResponse({ status: "error", message: "登録企業がありません。" });

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
  } catch (err) {
    return createJsonResponse({ status: "error", message: "ログイン処理エラー: " + err.message });
  }
}

// ==========================================
// 8. 新規登録ロジック
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
  try {
    const sheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_COMPANIES);
    if (!sheet) return createJsonResponse({ status: "error", message: "マスターDBが見つかりません" });

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
  } catch (err) {
    return createJsonResponse({ status: "error", message: "企業登録エラー: " + err.message });
  }
}