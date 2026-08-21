// ==========================================
// 1. 変数定義（必ず一番上で定義）
// ==========================================
const SAAS_MASTER_SS_ID = "1trxC9oVhyCbwH8p6bD1UuBq7LWlKTpOrTmAq3UVAaoU";
const TEMPLATE_SS_ID = "1mmQXbcUOGoKIlM6qWSGclY3G--BfkRTrKx5aG2vFNJY";
const PARENT_FOLDER_ID = "1Is1y-S5vWWjtkjha8KTXOPL3yxSLMJ_G";
const SHEET_COMPANIES = "SaaS管理マスターDB";

// テナントDB（TEMPLATE_SS_IDのコピー）内のシート名。実体はpayrollMasterSetup.jsのinitPayrollTemplateDBで作成される。
const SHEET_EMPLOYEES = "社員マスタ";
const SHEET_SITES = "現場マスタ";
const SHEET_RATE_MASTER = "税率・保険料率マスタ";
const SHEET_INSURANCE_GRADE = "社会保険等級表マスタ";
const SHEET_WITHHOLDING_TAX = "源泉徴収税額表マスタ";
const SHEET_COMMUTE_LIMIT = "通勤手当非課税限度額マスタ";
const SHEET_ATTENDANCE_LOG = "打刻データ";
const SHEET_PAYROLL = "勤怠・給与データ";
const SHEET_BATCH_STATUS = "計算バッチ状態";

// SAAS_MASTER_SS_ID側のシート名。時間主導トリガーは全テナント共通の単一GASプロジェクトから発火するため、
// 「どのテナントDBのどのバッチを再開すべきか」を横断的に把握するための調停シート（payroll.js参照）。
const SHEET_PAYROLL_QUEUE = "給与計算バッチキュー";

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

// ヘッダー名をキーにしたオブジェクトから、シートの実列数に整合した配列を組み立ててappendRowする。
// 固定順配列でのappendRowは列構成変更のたびにズレるため、シート書き込みは必ずこの関数を経由する。
function appendRowByHeaderMap(sheet, valueObj) {
  const map = getHeaderMap(sheet);
  const lastCol = sheet.getLastColumn();
  const row = new Array(lastCol).fill("");
  Object.keys(valueObj).forEach(key => {
    if (map[key] !== undefined) row[map[key]] = valueObj[key];
  });
  sheet.appendRow(row);
  return row;
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
    
    // ★追加: データ取得＆テストデータ投入
    if (action === "getDashboardData") return handleGetDashboardData(payload);
    if (action === "addTestData") return handleAddTestData(payload);

    // ★追加: 給与自動計算（payroll.js）
    if (action === "runPayrollBatch") return handleRunPayrollBatch(payload);
    if (action === "getPayrollStatus") return handleGetPayrollStatus(payload);

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

function handleGetDashboardData(payload) {
  const dbId = payload.dbId;
  if (!dbId) return createJsonResponse({ status: "error", message: "DBのIDが指定されていません" });

  try {
    const db = SpreadsheetApp.openById(dbId);
    
    const empSheet = db.getSheetByName("社員マスタ");
    let empCount = 0;
    if (empSheet) {
      const lastRow = empSheet.getLastRow();
      empCount = lastRow > 1 ? lastRow - 1 : 0; 
    }

    const siteSheet = db.getSheetByName("現場マスタ");
    let siteCount = 0;
    if (siteSheet) {
      const lastRow = siteSheet.getLastRow();
      siteCount = lastRow > 1 ? lastRow - 1 : 0; 
    }

    return createJsonResponse({ 
      status: "success", 
      data: { empCount: empCount, siteCount: siteCount }
    });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "テナントDBアクセスエラー: " + err.message });
  }
}

// ★追加: テストデータを自動追加するロジック
function handleAddTestData(payload) {
  const dbId = payload.dbId;
  if (!dbId) return createJsonResponse({ status: "error", message: "DBのIDが指定されていません" });

  try {
    const db = SpreadsheetApp.openById(dbId);
    const empSheet = db.getSheetByName("社員マスタ");
    const siteSheet = db.getSheetByName("現場マスタ");
    
    let empCount = empSheet ? empSheet.getLastRow() : 0;
    let siteCount = siteSheet ? siteSheet.getLastRow() : 0;

    if (empSheet) {
      appendRowByHeaderMap(empSheet, {
        "社員No": `EMP-${String(empCount).padStart(3, '0')}`,
        "氏名": "テスト太郎",
        "所属区分": "現場",
        "給与形態": "日給",
        "基本給/日当/時給": 10000,
        "雇用保険適用": "有効",
        "ステータス": "有効",
        "メールアドレス": "test1@example.com",
        "生年月日": "1985-05-10",
        "業務内容": "一般",
        "雇用形態": "正社員",
        "通勤手段": "マイカー等",
        "通勤距離(km)": 12,
        "扶養人数": 1,
        "甲欄/乙欄区分": "甲欄",
        "社会保険加入フラグ": "加入"
      });
      appendRowByHeaderMap(empSheet, {
        "社員No": `EMP-${String(empCount + 1).padStart(3, '0')}`,
        "氏名": "テスト花子",
        "所属区分": "事務",
        "給与形態": "月給",
        "基本給/日当/時給": 250000,
        "雇用保険適用": "有効",
        "ステータス": "有効",
        "メールアドレス": "test2@example.com",
        "生年月日": "1998-11-20",
        "業務内容": "一般",
        "雇用形態": "正社員",
        "通勤手段": "公共交通機関",
        "通勤距離(km)": 0,
        "扶養人数": 0,
        "甲欄/乙欄区分": "甲欄",
        "社会保険加入フラグ": "加入"
      });
    }

    if (siteSheet) {
      siteSheet.appendRow([`SITE-${String(siteCount).padStart(3, '0')}`, "テスト開発プロジェクト", "他社", 1000000, "2026-08-01", "2026-12-31", "テスト太郎", "進行中"]);
      siteSheet.appendRow([`SITE-${String(siteCount+1).padStart(3, '0')}`, "テスト改修工事", "自社", 500000, "2026-09-01", "2026-11-30", "テスト花子", "予定"]);
    }

    return createJsonResponse({ status: "success", message: "社員マスタ・現場マスタにテストデータを各2件追加しました。" });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "データ追加エラー: " + err.message });
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
    
    // 社員マスタ・現場マスタ・税率保険料率マスタの列構成はテンプレDB側（payrollMasterSetup.js の
    // initPayrollTemplateDB）で確定済みという前提。appendRowByHeaderMap によりヘッダー名で書き込むため、
    // 将来シート列が増減してもここは影響を受けない。ログイン認証はSAAS_MASTER_SS_ID側（企業ID＋管理者
    // メール＋初期パスワード）で完結しており、テナントDBの社員マスタにログイン情報を持たせる必要はない。
    const empSheet = newDb.getSheetByName(SHEET_EMPLOYEES);
    if (empSheet) {
      if (empSheet.getLastRow() > 1) empSheet.getRange(2, 1, empSheet.getLastRow() - 1, empSheet.getLastColumn()).clearContent();
      appendRowByHeaderMap(empSheet, {
        "社員No": `${companyId}-EMP-001`,
        "氏名": "システム管理者",
        "所属区分": "本部",
        "給与形態": "月給",
        "基本給/日当/時給": 0,
        "雇用保険適用": "なし",
        "ステータス": "有効",
        "メールアドレス": adminEmail,
        "業務内容": "管理者",
        "雇用形態": "正社員",
        "通勤手段": "なし",
        "通勤距離(km)": 0,
        "扶養人数": 0,
        "甲欄/乙欄区分": "甲欄",
        "社会保険加入フラグ": "対象外"
      });
    }

    const siteSheet = newDb.getSheetByName(SHEET_SITES);
    if (siteSheet) {
      if (siteSheet.getLastRow() > 1) siteSheet.getRange(2, 1, siteSheet.getLastRow() - 1, siteSheet.getLastColumn()).clearContent();
      appendRowByHeaderMap(siteSheet, {
        "現場ID": "SITE-000",
        "現場名": "本社（基本勤務地）",
        "クライアント名": "自社",
        "契約金額": 0,
        "着工日": "2026-01-01",
        "完工日": "2030-12-31",
        "現場責任者": "システム管理者",
        "ステータス": "進行中"
      });
    }

    // 保険料率・等級表・源泉徴収税額表はテンプレDB複製時点で投入済みのため、テナント登録時は
    // テナント固有値（管理者メールアドレス）のみを税率・保険料率マスタへ反映する。
    const rateSheet = newDb.getSheetByName(SHEET_RATE_MASTER);
    if (rateSheet && rateSheet.getLastRow() > 1) {
      const rateMap = getHeaderMap(rateSheet);
      if (rateMap["管理者メールアドレス"] !== undefined) {
        rateSheet.getRange(2, rateMap["管理者メールアドレス"] + 1).setValue(adminEmail);
      }
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