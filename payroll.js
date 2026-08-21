// ==========================================
// payroll.js - 給与自動計算エンジン（骨組み）
// ------------------------------------------
// スコープ: 打刻データの月次集計 → 社会保険料/雇用保険料/源泉所得税の算出 →
//          「勤怠・給与データ」への反映まで。PDF明細出力・実行用フロントエンドUIは次フェーズ。
// シート名・列名の定数は api.js の先頭（SHEET_EMPLOYEES 等）で一元定義している。
// ==========================================

// ==========================================
// 1. 変数定義
// ==========================================
const PAYROLL_TIME_LIMIT_MS = 5 * 60 * 1000; // GAS6分制限に対する安全マージン。超えたら自主停止し継続トリガーへ引き継ぐ
const PAYROLL_RESUME_DELAY_SEC = 60;          // 継続トリガーが発火するまでの待機秒数
const MONTHLY_OVERTIME_THRESHOLD_H = 60;      // 時間外割増率が引き上がる月間残業時間の境界（労基法）
const NIGHT_START_HOUR = 22;                  // 深夜割増の開始時刻
const NIGHT_END_HOUR = 5;                     // 深夜割増の終了時刻
const DEPENDENTS_EXTRA_DEDUCTION = 1610;      // 扶養親族8人目以降、1人あたりの税額控除（国税庁・源泉徴収税額表の備考）
const PUBLIC_TRANSIT_TAX_FREE_LIMIT = 150000; // 公共交通機関利用者の通勤手当非課税限度額（月額、国税庁No.2582）

// ==========================================
// 2. APIハンドラ
// ==========================================
function handleRunPayrollBatch(payload) {
  try {
    const dbId = String(payload.dbId || "").trim();
    const companyId = String(payload.companyId || "").trim();
    const targetMonth = normalizeMonthKey(payload.targetMonth);
    if (!dbId || !targetMonth) return createJsonResponse({ status: "error", message: "dbIdとtargetMonth(YYYY-MM)を指定してください" });

    const queueSheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_PAYROLL_QUEUE);
    if (!queueSheet) return createJsonResponse({ status: "error", message: "給与計算バッチキューが見つかりません（先にセットアップを実行してください）" });

    if (findActiveQueueRow(queueSheet, dbId, targetMonth)) {
      return createJsonResponse({ status: "error", message: `対象年月[${targetMonth}]は既に計算中です` });
    }

    const db = SpreadsheetApp.openById(dbId);
    const batchSheet = db.getSheetByName(SHEET_BATCH_STATUS);
    if (!batchSheet) return createJsonResponse({ status: "error", message: "テナントDBに計算バッチ状態シートが見つかりません" });

    const batchId = `${targetMonth}-${new Date().getTime()}`;
    appendRowByHeaderMap(batchSheet, {
      "バッチID": batchId, "対象年月": targetMonth, "ステータス": "実行中",
      "開始日時": new Date(), "最終更新日時": new Date(),
      "処理済み社員数": 0, "対象社員総数": 0, "トリガーID": "", "エラー件数": 0
    });
    appendRowByHeaderMap(queueSheet, {
      "バッチID": batchId, "企業ID": companyId, "dbId": dbId, "対象年月": targetMonth,
      "ステータス": "実行中", "トリガーID": "", "最終更新日時": new Date()
    });

    runPayrollForMonth(dbId, companyId, targetMonth, batchId);

    return createJsonResponse({ status: "success", message: "給与計算バッチを開始しました。", batchId: batchId });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "給与計算バッチ起動エラー: " + err.message });
  }
}

function handleGetPayrollStatus(payload) {
  try {
    const dbId = String(payload.dbId || "").trim();
    const targetMonth = normalizeMonthKey(payload.targetMonth);
    if (!dbId || !targetMonth) return createJsonResponse({ status: "error", message: "dbIdとtargetMonth(YYYY-MM)を指定してください" });

    const db = SpreadsheetApp.openById(dbId);
    const batchSheet = db.getSheetByName(SHEET_BATCH_STATUS);
    if (!batchSheet) return createJsonResponse({ status: "error", message: "計算バッチ状態シートが見つかりません" });

    const map = getHeaderMap(batchSheet);
    const lastRow = batchSheet.getLastRow();
    if (lastRow <= 1) return createJsonResponse({ status: "success", data: null });

    const data = batchSheet.getRange(2, 1, lastRow - 1, batchSheet.getLastColumn()).getValues();
    let latest = null;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][map["対象年月"]]).trim() === targetMonth) latest = data[i];
    }
    if (!latest) return createJsonResponse({ status: "success", data: null });

    return createJsonResponse({
      status: "success",
      data: {
        batchId: latest[map["バッチID"]],
        batchStatus: latest[map["ステータス"]],
        processedCount: latest[map["処理済み社員数"]],
        totalCount: latest[map["対象社員総数"]],
        errorCount: latest[map["エラー件数"]],
        lastUpdated: latest[map["最終更新日時"]]
      }
    });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "進捗取得エラー: " + err.message });
  }
}

// ==========================================
// 3. バッチキュー管理（SAAS_MASTER_SS_ID側。時間主導トリガーは全テナント共通の単一GASプロジェクトから
//    発火するため、「どのテナントDBのどのバッチを再開すべきか」を横断的に把握するための調停シート）
// ==========================================
function findActiveQueueRow(queueSheet, dbId, targetMonth) {
  const map = getHeaderMap(queueSheet);
  const lastRow = queueSheet.getLastRow();
  if (lastRow <= 1) return null;
  const data = queueSheet.getRange(2, 1, lastRow - 1, queueSheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const status = String(data[i][map["ステータス"]]).trim();
    if (String(data[i][map["dbId"]]).trim() === dbId &&
        String(data[i][map["対象年月"]]).trim() === targetMonth &&
        (status === "実行中" || status === "一時停止(継続予定)")) {
      return { rowIndex: i + 2, row: data[i] };
    }
  }
  return null;
}

function updateQueueRowByBatchId(queueSheet, batchId, fields) {
  const map = getHeaderMap(queueSheet);
  const lastRow = queueSheet.getLastRow();
  if (lastRow <= 1) return;
  const data = queueSheet.getRange(2, 1, lastRow - 1, queueSheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][map["バッチID"]]).trim() === batchId) {
      setRowValuesByMap(queueSheet, i + 2, map, fields);
      return;
    }
  }
}

// ==========================================
// 4. バッチ制御（テナントDB側の実処理）
// ==========================================
function runPayrollForMonth(dbId, companyId, targetMonth, batchId) {
  const startTime = new Date().getTime();
  const db = SpreadsheetApp.openById(dbId);
  const batchSheet = db.getSheetByName(SHEET_BATCH_STATUS);
  const batchMap = getHeaderMap(batchSheet);
  const batchRowIndex = findBatchRowIndexById(batchSheet, batchMap, batchId);
  if (!batchRowIndex) return;

  const queueSheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_PAYROLL_QUEUE);

  try {
    aggregateAttendanceForMonth(db, targetMonth);

    const ctx = loadPayrollMasterContext(db, targetMonth);
    const payrollSheet = db.getSheetByName(SHEET_PAYROLL);
    const payrollMap = getHeaderMap(payrollSheet);
    const lastRow = payrollSheet.getLastRow();

    if (lastRow <= 1) {
      setRowValuesByMap(batchSheet, batchRowIndex, batchMap, { "ステータス": "完了", "最終更新日時": new Date(), "対象社員総数": 0 });
      if (queueSheet) updateQueueRowByBatchId(queueSheet, batchId, { "ステータス": "完了", "最終更新日時": new Date() });
      return;
    }

    const data = payrollSheet.getRange(2, 1, lastRow - 1, payrollSheet.getLastColumn()).getValues();

    let processedCount = 0;
    let errorCount = 0;
    let totalTarget = 0;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (String(row[payrollMap["対象年月"]]).trim() !== targetMonth) continue;
      totalTarget++;
      if (String(row[payrollMap["計算ステータス"]] || "").trim() === "完了") { processedCount++; continue; }

      if (new Date().getTime() - startTime > PAYROLL_TIME_LIMIT_MS) {
        const trigger = ScriptApp.newTrigger("resumePayrollBatch").timeBased().after(PAYROLL_RESUME_DELAY_SEC * 1000).create();
        const triggerId = trigger.getUniqueId();
        setRowValuesByMap(batchSheet, batchRowIndex, batchMap, {
          "ステータス": "一時停止(継続予定)", "処理済み社員数": processedCount, "対象社員総数": totalTarget,
          "最終更新日時": new Date(), "トリガーID": triggerId
        });
        if (queueSheet) updateQueueRowByBatchId(queueSheet, batchId, {
          "ステータス": "一時停止(継続予定)", "トリガーID": triggerId, "最終更新日時": new Date()
        });
        return;
      }

      try {
        const empRow = ctx.employeeByNo[String(row[payrollMap["社員No"]]).trim()];
        if (!empRow) throw new Error("社員マスタに該当社員が見つかりません");
        const result = processSingleEmployeePayroll(empRow, ctx.employeeMap, row, payrollMap, targetMonth, ctx);
        finalizePayrollRow(payrollSheet, i + 2, payrollMap, result);
      } catch (rowErr) {
        payrollSheet.getRange(i + 2, payrollMap["計算ステータス"] + 1).setValue("エラー");
        payrollSheet.getRange(i + 2, payrollMap["エラーメッセージ"] + 1).setValue(String(rowErr.message));
        errorCount++;
      }
      processedCount++;

      setRowValuesByMap(batchSheet, batchRowIndex, batchMap, {
        "処理済み社員数": processedCount, "対象社員総数": totalTarget, "エラー件数": errorCount, "最終更新日時": new Date()
      });
    }

    setRowValuesByMap(batchSheet, batchRowIndex, batchMap, {
      "ステータス": "完了", "処理済み社員数": processedCount, "対象社員総数": totalTarget, "エラー件数": errorCount, "最終更新日時": new Date()
    });
    if (queueSheet) updateQueueRowByBatchId(queueSheet, batchId, { "ステータス": "完了", "最終更新日時": new Date() });
  } catch (err) {
    setRowValuesByMap(batchSheet, batchRowIndex, batchMap, { "ステータス": "エラー", "最終更新日時": new Date() });
    if (queueSheet) updateQueueRowByBatchId(queueSheet, batchId, { "ステータス": "エラー", "最終更新日時": new Date() });
    throw err;
  }
}

// 時間主導トリガーから起動されるエントリポイント。PropertiesServiceは使わず、dbId・targetMonth・batchIdの
// 引き継ぎはマスターDB側「給与計算バッチキュー」のスプレッドシート行から復元する（過去のPropertiesService
// 乱用によるクラッシュを踏まえた設計）。「一時停止(継続予定)」の行を全件処理する（同時に複数テナントの
// バッチが一時停止していても、発火元トリガーを厳密に特定する必要がなく取りこぼしがない）。
function resumePayrollBatch() {
  const queueSheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_PAYROLL_QUEUE);
  if (!queueSheet) return;

  const map = getHeaderMap(queueSheet);
  const lastRow = queueSheet.getLastRow();
  if (lastRow > 1) {
    const data = queueSheet.getRange(2, 1, lastRow - 1, queueSheet.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][map["ステータス"]]).trim() !== "一時停止(継続予定)") continue;
      const dbId = String(data[i][map["dbId"]]).trim();
      const companyId = String(data[i][map["企業ID"]]).trim();
      const targetMonth = String(data[i][map["対象年月"]]).trim();
      const batchId = String(data[i][map["バッチID"]]).trim();
      runPayrollForMonth(dbId, companyId, targetMonth, batchId);
    }
  }
  cleanupOrphanTriggers_();
}

// resumePayrollBatchに紐づく時間主導トリガーのうち、キュー上で「一時停止(継続予定)」として
// 参照されなくなったものを削除する（runPayrollForMonthが新規トリガーを発行した場合は
// キュー上に新しいトリガーIDが記録されているため、その行は削除対象から除外される）。
function cleanupOrphanTriggers_() {
  const queueSheet = SpreadsheetApp.openById(SAAS_MASTER_SS_ID).getSheetByName(SHEET_PAYROLL_QUEUE);
  const activeTriggerIds = {};
  if (queueSheet) {
    const map = getHeaderMap(queueSheet);
    const lastRow = queueSheet.getLastRow();
    if (lastRow > 1) {
      queueSheet.getRange(2, 1, lastRow - 1, queueSheet.getLastColumn()).getValues().forEach(r => {
        if (String(r[map["ステータス"]]).trim() === "一時停止(継続予定)") {
          const tid = String(r[map["トリガーID"]] || "").trim();
          if (tid) activeTriggerIds[tid] = true;
        }
      });
    }
  }
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "resumePayrollBatch" && !activeTriggerIds[t.getUniqueId()]) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function findBatchRowIndexById(batchSheet, batchMap, batchId) {
  const lastRow = batchSheet.getLastRow();
  if (lastRow <= 1) return null;
  const data = batchSheet.getRange(2, 1, lastRow - 1, batchSheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][batchMap["バッチID"]]).trim() === batchId) return i + 2;
  }
  return null;
}

function setRowValuesByMap(sheet, rowIndex, map, fields) {
  Object.keys(fields).forEach(key => {
    if (map[key] !== undefined) sheet.getRange(rowIndex, map[key] + 1).setValue(fields[key]);
  });
}

// ==========================================
// 5. 勤怠集計
// ==========================================

// 「打刻データ」から対象年月分をまとめ、「勤怠・給与データ」に未計算行として下書きする。
// 既に同一対象年月・社員Noの行が存在する場合はスキップする（再実行時の重複作成防止）。
function aggregateAttendanceForMonth(db, targetMonth) {
  const empSheet = db.getSheetByName(SHEET_EMPLOYEES);
  const logSheet = db.getSheetByName(SHEET_ATTENDANCE_LOG);
  const payrollSheet = db.getSheetByName(SHEET_PAYROLL);
  if (!empSheet || !logSheet || !payrollSheet) throw new Error("社員マスタ・打刻データ・勤怠給与データのいずれかが見つかりません");

  const payrollMap = getHeaderMap(payrollSheet);
  const existingKeys = {};
  const payrollLastRow = payrollSheet.getLastRow();
  if (payrollLastRow > 1) {
    payrollSheet.getRange(2, 1, payrollLastRow - 1, payrollSheet.getLastColumn()).getValues().forEach(r => {
      existingKeys[`${String(r[payrollMap["対象年月"]]).trim()}|${String(r[payrollMap["社員No"]]).trim()}`] = true;
    });
  }

  const empMap = getHeaderMap(empSheet);
  const empLastRow = empSheet.getLastRow();
  if (empLastRow <= 1) return;
  const employees = empSheet.getRange(2, 1, empLastRow - 1, empSheet.getLastColumn()).getValues()
    .filter(r => String(r[empMap["ステータス"]]).trim() === "有効");

  const logMap = getHeaderMap(logSheet);
  const logLastRow = logSheet.getLastRow();
  const logs = logLastRow > 1 ? logSheet.getRange(2, 1, logLastRow - 1, logSheet.getLastColumn()).getValues() : [];

  const logsByEmp = {};
  logs.forEach(r => {
    const dateVal = toDateSafe(r[logMap["日付"]]);
    if (!dateVal || formatMonthKey(dateVal) !== targetMonth) return;
    const empNo = String(r[logMap["社員No"]]).trim();
    if (!logsByEmp[empNo]) logsByEmp[empNo] = [];
    const dt = toDateSafe(r[logMap["打刻日時"]]) || dateVal;
    logsByEmp[empNo].push({ dateTime: dt, date: formatDateKey(dateVal), type: String(r[logMap["打刻種別"]]).trim() });
  });

  employees.forEach(emp => {
    const empNo = String(emp[empMap["社員No"]]).trim();
    const key = `${targetMonth}|${empNo}`;
    if (existingKeys[key]) return;

    const empLogs = (logsByEmp[empNo] || []).sort((a, b) => a.dateTime - b.dateTime);
    const daily = pairClockEvents(empLogs);
    const summary = calcOvertimeNightHoliday(daily);

    appendRowByHeaderMap(payrollSheet, {
      "対象年月": targetMonth,
      "社員No": empNo,
      "氏名": emp[empMap["氏名"]],
      "出勤日数": summary.workDays,
      "休日出勤日数": summary.holidayWorkDays,
      "欠勤日数": 0, // 骨組みスコープ: 所定稼働日との突合による欠勤判定は次フェーズ
      "総実働H": round2(summary.totalHours),
      "時間外H(60H以下)": round2(summary.overtimeUnder60),
      "時間外H(60H超)": round2(summary.overtimeOver60),
      "深夜H": round2(summary.nightHours),
      "休日実働H": round2(summary.holidayHours),
      "計算ステータス": "未計算"
    });
  });
}

// 出勤/退勤打刻をペアリングし、日別の実働時間(ms)・深夜重複時間(ms)・曜日を積算する。
// 骨組みスコープのため、休憩時間の自動控除・日跨ぎ勤務・打刻漏れの補正は行わない（次フェーズ課題）。
function pairClockEvents(events) {
  const dailyMap = {};
  let openIn = null;
  events.forEach(e => {
    if (e.type === "出勤") {
      openIn = e;
    } else if (e.type === "退勤" && openIn) {
      const dateKey = openIn.date;
      if (!dailyMap[dateKey]) dailyMap[dateKey] = { workedMs: 0, nightMs: 0, dayOfWeek: openIn.dateTime.getDay() };
      dailyMap[dateKey].workedMs += (e.dateTime.getTime() - openIn.dateTime.getTime());
      dailyMap[dateKey].nightMs += calcNightOverlapMs(openIn.dateTime, e.dateTime);
      openIn = null;
    }
  });
  return dailyMap;
}

// 打刻時刻区間と深夜帯(22:00〜翌5:00)の重なり時間(ms)を、日単位で解析的に算出する。
function calcNightOverlapMs(start, end) {
  let overlapMs = 0;
  let cursor = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const endTime = end.getTime();
  while (cursor.getTime() < endTime) {
    const dayStartTime = cursor.getTime();
    const dayEndTime = dayStartTime + 24 * 3600 * 1000;
    const nightMorningEnd = dayStartTime + NIGHT_END_HOUR * 3600 * 1000;   // 当日0:00-5:00
    const nightEveningStart = dayStartTime + NIGHT_START_HOUR * 3600 * 1000; // 当日22:00-翌0:00
    overlapMs += overlapRange(start.getTime(), endTime, dayStartTime, nightMorningEnd);
    overlapMs += overlapRange(start.getTime(), endTime, nightEveningStart, dayEndTime);
    cursor = new Date(dayEndTime);
  }
  return overlapMs;
}

function overlapRange(s1, e1, s2, e2) {
  const s = Math.max(s1, s2);
  const e = Math.min(e1, e2);
  return Math.max(0, e - s);
}

// 日別実働時間から、法定労働時間(1日8H)超過分を時間外として積算し、月60H基準で分割する。
// 休日出勤の判定は「日曜日=法定休日」という簡易ルール（次フェーズで現場マスタ側の休日設定に対応）。
function calcOvertimeNightHoliday(daily) {
  const STANDARD_DAILY_HOURS = 8;
  let totalHours = 0, workDays = 0, holidayWorkDays = 0, holidayHours = 0, nightHours = 0;
  let overtimeAccum = 0, overtimeUnder60 = 0, overtimeOver60 = 0;

  Object.keys(daily).sort().forEach(dateKey => {
    const rec = daily[dateKey];
    const hours = rec.workedMs / (1000 * 60 * 60);
    const nightH = rec.nightMs / (1000 * 60 * 60);
    totalHours += hours;
    nightHours += nightH;
    workDays++;

    if (rec.dayOfWeek === 0) {
      holidayWorkDays++;
      holidayHours += hours;
    } else {
      const dailyOvertime = Math.max(0, hours - STANDARD_DAILY_HOURS);
      const remainingUnder60 = Math.max(0, MONTHLY_OVERTIME_THRESHOLD_H - overtimeAccum);
      const under60Part = Math.min(dailyOvertime, remainingUnder60);
      const over60Part = dailyOvertime - under60Part;
      overtimeUnder60 += under60Part;
      overtimeOver60 += over60Part;
      overtimeAccum += dailyOvertime;
    }
  });

  return { totalHours, workDays, holidayWorkDays, holidayHours, nightHours, overtimeUnder60, overtimeOver60 };
}

// ==========================================
// 6. マスタ参照
// ==========================================

// 給与計算1バッチ分に必要な全マスタを一括ロードする。ループ内でのシートI/Oを避けるため、
// 各マスタは対象年月時点で有効な行のみに絞り込んだ状態でメモリに保持する。
function loadPayrollMasterContext(db, targetMonth) {
  const rateSheet = db.getSheetByName(SHEET_RATE_MASTER);
  const gradeSheet = db.getSheetByName(SHEET_INSURANCE_GRADE);
  const taxSheet = db.getSheetByName(SHEET_WITHHOLDING_TAX);
  const commuteSheet = db.getSheetByName(SHEET_COMMUTE_LIMIT);
  const empSheet = db.getSheetByName(SHEET_EMPLOYEES);
  if (!rateSheet || !gradeSheet || !taxSheet || !commuteSheet || !empSheet) {
    throw new Error("給与計算に必要なマスタシートが不足しています（税率・保険料率/社会保険等級表/源泉徴収税額表/通勤手当非課税限度額/社員マスタ）");
  }

  const rateResult = getActiveMasterRow(rateSheet, "適用開始日", targetMonth);
  if (!rateResult) throw new Error("税率・保険料率マスタに対象年月時点で有効な行がありません");

  const gradeMap = getHeaderMap(gradeSheet);
  const gradeRows = getActiveMasterRows(gradeSheet, gradeMap, "適用開始日", targetMonth);

  const taxMap = getHeaderMap(taxSheet);
  const taxLastRow = taxSheet.getLastRow();
  const taxRows = taxLastRow > 1 ? taxSheet.getRange(2, 1, taxLastRow - 1, taxSheet.getLastColumn()).getValues() : [];

  const commuteMap = getHeaderMap(commuteSheet);
  const commuteRows = getActiveMasterRows(commuteSheet, commuteMap, "適用開始日", targetMonth);

  const empMap = getHeaderMap(empSheet);
  const empLastRow = empSheet.getLastRow();
  const employeeByNo = {};
  if (empLastRow > 1) {
    empSheet.getRange(2, 1, empLastRow - 1, empSheet.getLastColumn()).getValues().forEach(r => {
      employeeByNo[String(r[empMap["社員No"]]).trim()] = r;
    });
  }

  return {
    rateRow: rateResult.row, rateMap: rateResult.map,
    gradeRows: gradeRows, gradeMap: gradeMap,
    taxRows: taxRows, taxMap: taxMap,
    commuteRows: commuteRows, commuteMap: commuteMap,
    employeeByNo: employeeByNo, employeeMap: empMap
  };
}

// 「適用開始日」列を持つマスタから、対象年月以前で最も新しい適用開始日の1行を返す（年度改定対応）。
function getActiveMasterRow(sheet, dateColName, targetMonth) {
  const map = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const targetDate = monthKeyToDate(targetMonth);
  let best = null, bestDate = null;
  data.forEach(r => {
    const d = toDateSafe(r[map[dateColName]]);
    if (!d || d > targetDate) return;
    if (!bestDate || d > bestDate) { bestDate = d; best = r; }
  });
  return best ? { row: best, map: map } : null;
}

// 同上だが、対象年月時点で有効な「適用開始日の世代」に属する行を全部返す
// （社会保険等級表・通勤手当限度額のように複数行を検索対象にする場合）。
function getActiveMasterRows(sheet, map, dateColName, targetMonth) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const targetDate = monthKeyToDate(targetMonth);
  const candidates = data.filter(r => {
    const d = toDateSafe(r[map[dateColName]]);
    return d && d <= targetDate;
  });
  if (candidates.length === 0) return [];
  let latestDate = toDateSafe(candidates[0][map[dateColName]]);
  candidates.forEach(r => {
    const d = toDateSafe(r[map[dateColName]]);
    if (d > latestDate) latestDate = d;
  });
  return candidates.filter(r => toDateSafe(r[map[dateColName]]).getTime() === latestDate.getTime());
}

// ==========================================
// 7. 社会保険計算
// ==========================================
function calcAge(birthDate, asOfDate) {
  let age = asOfDate.getFullYear() - birthDate.getFullYear();
  const m = asOfDate.getMonth() - birthDate.getMonth();
  if (m < 0 || (m === 0 && asOfDate.getDate() < birthDate.getDate())) age--;
  return age;
}

// 介護保険第2号被保険者(40歳以上65歳未満)判定
function isKaigoHokenTarget(age) {
  return age >= 40 && age < 65;
}

// 標準報酬月額のベースとなる報酬月額から、社会保険等級表マスタの健保用行を検索する。
// 標準報酬月額の正式な算定(定時決定・随時改定)は次フェーズ課題。今回は当月の固定的賃金相当額を簡易的に用いる。
function findHealthInsuranceGrade(gradeRows, gradeMap, monthlyAmount) {
  if (gradeRows.length === 0) return null;
  for (let i = 0; i < gradeRows.length; i++) {
    const lower = Number(gradeRows[i][gradeMap["健保報酬月額下限"]]);
    const upper = Number(gradeRows[i][gradeMap["健保報酬月額上限"]]);
    const isLast = i === gradeRows.length - 1;
    if (monthlyAmount >= lower && (isLast || monthlyAmount < upper)) return gradeRows[i];
  }
  return gradeRows[gradeRows.length - 1];
}

// 厚生年金は健保と等級・報酬月額範囲が一部異なる（協会けんぽ表の等級欄の()内数字）ため、
// 「等級(厚年)」が設定されている行だけを対象に別途検索する。
function findPensionGrade(gradeRows, gradeMap, monthlyAmount) {
  const pensionRows = gradeRows.filter(r => String(r[gradeMap["等級(厚年)"]]).trim() !== "");
  if (pensionRows.length === 0) return null;
  for (let i = 0; i < pensionRows.length; i++) {
    const lower = Number(pensionRows[i][gradeMap["厚年報酬月額下限"]]);
    const upper = Number(pensionRows[i][gradeMap["厚年報酬月額上限"]]);
    const isLast = i === pensionRows.length - 1;
    if (monthlyAmount >= lower && (isLast || monthlyAmount < upper)) return pensionRows[i];
  }
  return pensionRows[pensionRows.length - 1];
}

function calcHealthInsurancePremium(gradeRow, gradeMap, isKaigoTarget) {
  if (!gradeRow) return 0;
  const col = isKaigoTarget ? "健保折半額(介護該当)" : "健保折半額(介護非該当)";
  return Number(gradeRow[gradeMap[col]]) || 0;
}

function calcPensionPremium(gradeRow, gradeMap) {
  if (!gradeRow) return 0;
  return Number(gradeRow[gradeMap["厚年折半額"]]) || 0;
}

function calcEmploymentInsurancePremium(grossWage, rateRow, rateMap) {
  const rate = Number(rateRow[rateMap["雇用保険料率(労働者負担)"]]) || 0;
  return Math.floor(grossWage * rate);
}

// ==========================================
// 8. 通勤手当非課税判定
// ==========================================
function calcCommutingAllowanceTaxFree(commuteMethod, distanceKm, monthlyAllowance, commuteRows, commuteMap) {
  if (monthlyAllowance <= 0) return { taxFree: 0, taxable: 0 };
  if (!commuteMethod || commuteMethod === "なし" || commuteMethod === "公共交通機関") {
    const taxFree = Math.min(monthlyAllowance, PUBLIC_TRANSIT_TAX_FREE_LIMIT);
    return { taxFree: taxFree, taxable: monthlyAllowance - taxFree };
  }
  const dist = Number(distanceKm) || 0;
  let limit = 0;
  for (let i = 0; i < commuteRows.length; i++) {
    const lower = Number(commuteRows[i][commuteMap["片道距離下限(km)"]]);
    const upper = Number(commuteRows[i][commuteMap["片道距離上限(km)"]]);
    const isLast = i === commuteRows.length - 1;
    if (dist >= lower && (isLast || dist < upper)) { limit = Number(commuteRows[i][commuteMap["非課税限度額(月額)"]]) || 0; break; }
  }
  const taxFree = Math.min(monthlyAllowance, limit);
  return { taxFree: taxFree, taxable: monthlyAllowance - taxFree };
}

// ==========================================
// 9. 源泉所得税（源泉徴収税額表・月額表）
// ==========================================
function lookupWithholdingTax(amount, dependents, isOtsuran, taxRows, taxMap) {
  const dependentsCapped = Math.min(Math.max(0, dependents), 7);
  const extraDependents = Math.max(0, dependents - 7);

  if (amount < 105000) {
    if (!isOtsuran) return 0;
    const row = taxRows.find(r => String(r[taxMap["区分"]]).trim() === "105000円未満");
    const rate = row ? Number(row[taxMap["乙欄税率(%)"]]) / 100 : 0;
    return Math.floor(amount * rate);
  }

  if (amount < 740000) {
    const row = taxRows.find(r => {
      if (String(r[taxMap["区分"]]).trim() !== "テーブル") return false;
      const lower = Number(r[taxMap["給与額下限"]]);
      const upper = Number(r[taxMap["給与額上限"]]);
      return amount >= lower && amount < upper;
    });
    if (!row) throw new Error(`源泉徴収税額表に該当行が見つかりません(金額:${amount})`);
    if (isOtsuran) return Number(row[taxMap["乙欄税額/基準額"]]) || 0;
    let tax = Number(row[taxMap[`甲欄${dependentsCapped}人`]]) || 0;
    if (extraDependents > 0) tax = Math.max(0, tax - extraDependents * DEPENDENTS_EXTRA_DEDUCTION);
    return tax;
  }

  const row = taxRows.find(r => {
    if (String(r[taxMap["区分"]]).trim() !== "速算") return false;
    const lower = Number(r[taxMap["給与額下限"]]);
    const upperRaw = r[taxMap["給与額上限"]];
    const upper = (upperRaw === "" || upperRaw === null || upperRaw === undefined) ? Infinity : Number(upperRaw);
    return amount >= lower && amount < upper;
  });
  if (!row) throw new Error(`源泉徴収税額表(速算区分)に該当行が見つかりません(金額:${amount})`);

  const lower = Number(row[taxMap["給与額下限"]]);
  if (isOtsuran) {
    const base = Number(row[taxMap["乙欄税額/基準額"]]) || 0;
    const rate = Number(row[taxMap["乙欄税率(%)"]]) / 100;
    return Math.floor(base + (amount - lower) * rate);
  }
  // 速算区分も甲欄0〜7人列を基準額として使う（扶養人数ごとに790,000円等の基準額自体が異なるため、
  // テーブル区分と同じ列を共用する設計。マスタ側は payrollMasterSetup.js の
  // WITHHOLDING_TAX_BRACKET_RAW を参照）。
  const base = Number(row[taxMap[`甲欄${dependentsCapped}人`]]) || 0;
  const rate = Number(row[taxMap["甲欄税率(%)"]]) / 100;
  let tax = Math.floor(base + (amount - lower) * rate);
  if (extraDependents > 0) tax = Math.max(0, tax - extraDependents * DEPENDENTS_EXTRA_DEDUCTION);
  return tax;
}

// ==========================================
// 10. 給与確定
// ==========================================
function processSingleEmployeePayroll(empRow, empMap, payrollRow, payrollMap, targetMonth, ctx) {
  const asOfDate = monthKeyToDate(targetMonth);
  const birthDate = toDateSafe(empRow[empMap["生年月日"]]);
  const age = birthDate ? calcAge(birthDate, asOfDate) : null;
  const isKaigoTarget = age !== null && isKaigoHokenTarget(age);

  const basePay = Number(empRow[empMap["基本給/日当/時給"]]) || 0;
  const salaryType = String(empRow[empMap["給与形態"]]).trim();
  const fixedAllowance = ["職長手当", "家族手当", "住宅手当", "運転手当"]
    .reduce((sum, col) => sum + (Number(empRow[empMap[col]]) || 0), 0);

  const workDays = Number(payrollRow[payrollMap["出勤日数"]]) || 0;
  const totalHours = Number(payrollRow[payrollMap["総実働H"]]) || 0;

  // 日給/時給者の基本給は出勤日数・実働時間から算出。月給者は固定額（骨組みスコープ、日割り欠勤控除は次フェーズ）。
  let basePayAmount = basePay;
  if (salaryType === "日給") basePayAmount = basePay * workDays;
  if (salaryType === "時給") basePayAmount = Math.floor(basePay * totalHours);

  const dailyStdHours = Number(ctx.rateRow[ctx.rateMap["1日の所定労働時間H"]]) || 8;
  const monthlyStdHours = Number(ctx.rateRow[ctx.rateMap["月間平均所定労働時間H"]]) || 160;
  const hourlyRate = salaryType === "月給" ? basePay / monthlyStdHours
    : (salaryType === "日給" ? basePay / dailyStdHours : basePay);

  const otRateUnder60 = Number(ctx.rateRow[ctx.rateMap["時間外割増率(60H以下)"]]) || 1.25;
  const otRateOver60 = Number(ctx.rateRow[ctx.rateMap["時間外割増率(60H超)"]]) || 1.5;
  const nightRate = Number(ctx.rateRow[ctx.rateMap["深夜割増率"]]) || 0.25;
  const holidayRate = Number(ctx.rateRow[ctx.rateMap["休日割増率"]]) || 0.35;

  const overtimeUnder60H = Number(payrollRow[payrollMap["時間外H(60H以下)"]]) || 0;
  const overtimeOver60H = Number(payrollRow[payrollMap["時間外H(60H超)"]]) || 0;
  const nightH = Number(payrollRow[payrollMap["深夜H"]]) || 0;
  const holidayH = Number(payrollRow[payrollMap["休日実働H"]]) || 0;

  // 時間外・深夜・休日が重複する時間帯の割増率加算（例: 時間外深夜=1.25+0.25）は骨組みスコープでは
  // 個別に単純加算する簡易実装とし、重複区分の厳密な切り分けは次フェーズで詳細化する。
  const overtimePay = Math.floor(hourlyRate * otRateUnder60 * overtimeUnder60H + hourlyRate * otRateOver60 * overtimeOver60H);
  const nightPay = Math.floor(hourlyRate * nightRate * nightH);
  const holidayPay = Math.floor(hourlyRate * (1 + holidayRate) * holidayH);

  const commuteMethod = String(empRow[empMap["通勤手段"]]).trim();
  const commuteDistance = Number(empRow[empMap["通勤距離(km)"]]) || 0;
  const commuteAllowanceInput = Number(empRow[empMap["通勤手当(月額)"]]) || 0;
  const commute = calcCommutingAllowanceTaxFree(commuteMethod, commuteDistance, commuteAllowanceInput, ctx.commuteRows, ctx.commuteMap);

  const grossPay = basePayAmount + fixedAllowance + overtimePay + nightPay + holidayPay + commute.taxFree + commute.taxable;
  const taxableGross = basePayAmount + fixedAllowance + overtimePay + nightPay + holidayPay + commute.taxable;

  const isSyahoTarget = String(empRow[empMap["社会保険加入フラグ"]]).trim() === "加入";
  let healthPremium = 0, kaigoPremium = 0, pensionPremium = 0, gradeInfo = "";
  if (isSyahoTarget) {
    const standardMonthlyBase = basePayAmount + fixedAllowance; // 簡易算定。定時決定/随時改定は次フェーズ
    const healthGrade = findHealthInsuranceGrade(ctx.gradeRows, ctx.gradeMap, standardMonthlyBase);
    const pensionGrade = findPensionGrade(ctx.gradeRows, ctx.gradeMap, standardMonthlyBase);
    const healthTotal = calcHealthInsurancePremium(healthGrade, ctx.gradeMap, isKaigoTarget);
    healthPremium = isKaigoTarget ? 0 : healthTotal;
    kaigoPremium = isKaigoTarget ? healthTotal : 0;
    pensionPremium = calcPensionPremium(pensionGrade, ctx.gradeMap);
    gradeInfo = healthGrade ? String(healthGrade[ctx.gradeMap["等級(健保)"]]) : "";
  }

  const koyouFlag = String(empRow[empMap["雇用保険適用"]]).trim();
  const isKoyouTarget = koyouFlag === "有効" || koyouFlag === "加入";
  const employmentPremium = isKoyouTarget ? calcEmploymentInsurancePremium(grossPay, ctx.rateRow, ctx.rateMap) : 0;

  const insuranceTotal = healthPremium + kaigoPremium + pensionPremium + employmentPremium;
  const afterInsurance = Math.max(0, taxableGross - insuranceTotal);

  const dependents = Number(empRow[empMap["扶養人数"]]) || 0;
  const isOtsuran = String(empRow[empMap["甲欄/乙欄区分"]]).trim() === "乙欄";
  const incomeTax = lookupWithholdingTax(afterInsurance, dependents, isOtsuran, ctx.taxRows, ctx.taxMap);

  const deductionTotal = insuranceTotal + incomeTax;
  const netPay = grossPay - deductionTotal;

  return {
    basePay: basePayAmount, fixedAllowance, commuteTaxFree: commute.taxFree, commuteTaxable: commute.taxable,
    overtimePay, nightPay, holidayPay, taxableGross, grossPay,
    grade: gradeInfo, healthPremium, kaigoPremium, pensionPremium, employmentPremium,
    afterInsurance, incomeTax, deductionTotal, netPay
  };
}

function finalizePayrollRow(sheet, rowIndex, map, r) {
  setRowValuesByMap(sheet, rowIndex, map, {
    "基本給": r.basePay,
    "固定手当計": r.fixedAllowance,
    "通勤手当(非課税分)": r.commuteTaxFree,
    "通勤手当(課税分)": r.commuteTaxable,
    "時間外手当": r.overtimePay,
    "深夜手当": r.nightPay,
    "休日手当": r.holidayPay,
    "遅早欠控除": 0, // 骨組みスコープ: 欠勤・遅早控除の実額算出は次フェーズ
    "課税支給額計": r.taxableGross,
    "総支給額": r.grossPay,
    "適用等級": r.grade,
    "健康保険料": r.healthPremium,
    "介護保険料": r.kaigoPremium,
    "厚生年金保険料": r.pensionPremium,
    "雇用保険料": r.employmentPremium,
    "社保控除後給与額": r.afterInsurance,
    "源泉所得税": r.incomeTax,
    "控除額合計": r.deductionTotal,
    "差引支給額": r.netPay,
    "計算ステータス": "完了",
    "確定日時": new Date(),
    "エラーメッセージ": ""
  });
}

// ==========================================
// 11. 日付・共通ユーティリティ
// ==========================================
function normalizeMonthKey(value) {
  const s = String(value || "").trim();
  const m = s.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (!m) return "";
  return `${m[1]}-${String(m[2]).padStart(2, "0")}`;
}

function monthKeyToDate(monthKey) {
  const parts = monthKey.split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, 1);
}

function formatMonthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function formatDateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function toDateSafe(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function round2(n) {
  return Math.round(n * 100) / 100;
}
