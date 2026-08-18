// ==========================================
// POSTリクエスト（登録およびログイン認証API）
// ==========================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "データが空です。" });
    }
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    // 新規企業登録（自動附番・フォルダ生成・環境構築）
    if (action === "register") {
      return createNewTenant(payload.companyName, payload.email, payload.password);
    }

    // ログイン認証（メールアドレスとパスワードでDBIDとフォルダID群を取得）
    if (action === "login") {
      const email = payload.email;
      const password = payload.password;

      if (!email || !password) {
        return createJsonResponse({ status: "error", message: "メールアドレスとパスワードを入力してください。" });
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = ss.getSheetByName(SHEET_TENANT);
      if (!tenantSheet) return createJsonResponse({ status: "error", message: "マスターDBが存在しません。" });

      const map = getHeaderMap(tenantSheet);
      const data = tenantSheet.getRange(2, 1, tenantSheet.getLastRow() - 1, tenantSheet.getLastColumn()).getValues();

      let tenantInfo = null;

      for (let i = 0; i < data.length; i++) {
        if (String(data[i][map["管理者メール"]]).trim() === String(email).trim() && 
            String(data[i][map["初期パスワード"]]).trim() === String(password).trim()) {
          
          tenantInfo = {
            companyId: String(data[i][map["企業ID"]]).trim(),
            companyName: String(data[i][map["企業名"]]).trim(),
            dbId: String(data[i][map["企業用DB(SS)_ID"]]).trim(),
            rootFolderId: String(data[i][map["ルートフォルダ_ID"]]).trim(),
            meisaiFolderId: String(data[i][map["明細フォルダ_ID"]]).trim(),
            daichoFolderId: String(data[i][map["台帳フォルダ_ID"]]).trim(),
            shukkinFolderId: String(data[i][map["出勤簿フォルダ_ID"]]).trim(),
            meiboFolderId: String(data[i][map["名簿フォルダ_ID"]]).trim()
          };
          break;
        }
      }

      if (tenantInfo) {
        return createJsonResponse({ status: "success", data: tenantInfo });
      } else {
        return createJsonResponse({ status: "error", message: "メールアドレスまたはパスワードが間違っています。" });
      }
    }

    return createJsonResponse({ status: "error", message: "不明なアクション: " + action });

  } catch (err) {
    return createJsonResponse({ status: "error", message: "マスターサーバーエラー: " + err.message });
  }
}

// ==========================================
// 次の企業IDを自動生成するロジック（CP-001, CP-002...）
// ==========================================
function generateNextCompanyId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "CP-001";

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (let i = 0; i < data.length; i++) {
    const val = String(data[i][0]).trim();
    if (val.startsWith("CP-")) {
      const num = parseInt(val.replace("CP-", ""), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  const nextNum = maxNum + 1;
  return "CP-" + ("000" + nextNum).slice(-3);
}

// ==========================================
// 新規テナント環境構築ロジック（フォルダ一式生成）
// ==========================================
function createNewTenant(name, email, password) {
  try {
    if (!name || !email || !password) {
      return createJsonResponse({ status: "error", message: "企業名、メールアドレス、パスワードは必須です。" });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tenantSheet = ss.getSheetByName(SHEET_TENANT);

    // 1. 企業IDと初期管理者IDの生成
    const newId = generateNextCompanyId(tenantSheet);
    const initialAdminId = newId + "-admin";

    // 2. Googleドライブのフォルダ構成を作成
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    const rootFolder = parentFolder.createFolder(name + "様_システム一式");
    const rootFolderId = rootFolder.getId();

    const meisaiFolder = rootFolder.createFolder("給与明細");
    const daichoFolder = rootFolder.createFolder("賃金台帳");
    const shukkinFolder = rootFolder.createFolder("出勤簿");
    const meiboFolder = rootFolder.createFolder("労働者名簿");

    // 3. テンプレートシートをコピーしてルートフォルダに配置
    const templateFile = DriveApp.getFileById(TEMPLATE_SHEET_ID);
    const newSheet = templateFile.makeCopy(name + "様_データベース", rootFolder);
    const newSheetId = newSheet.getId();
    
    // 4. マスターDBへ書き込み (CSVの列構成に完全一致)
    tenantSheet.appendRow([
      newId,                // 企業ID
      name,                 // 企業名
      initialAdminId,       // 初期管理者ID
      password,             // 初期パスワード
      email,                // 管理者メール
      newSheetId,           // 企業用DB(SS)_ID
      rootFolderId,         // ルートフォルダ_ID
      meisaiFolder.getId(), // 明細フォルダ_ID
      daichoFolder.getId(), // 台帳フォルダ_ID
      shukkinFolder.getId(),// 出勤簿フォルダ_ID
      meiboFolder.getId(),  // 名簿フォルダ_ID
      new Date()            // 登録日時
    ]);
    
    return createJsonResponse({ 
      status: "success", 
      message: `構築完了: [${newId}] ${name} の環境（フォルダ群・データベース）を生成しました。`,
    });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "環境構築エラー: " + err.message });
  }
}