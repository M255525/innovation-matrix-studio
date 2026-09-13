/**
 * 創意打版台 — 序號授權後端（比照 traffic-rank-estimator / amazon-listing-generator 的骨架）。
 * 部署為 Web App 後，把取得的網址填進前端 index.html 的 LICENSE_CHECK_URL。
 */

const VALID_AMOUNT = 12; // 月
const COL_SERIAL = "序號";
const COL_START = "開始日期";
const COL_END = "結束日期";
// 序號資料不在第一個工作表時，把分頁名稱填在這裡；留空則自動用第一個工作表
const SHEET_NAME = "";

function doPost(e) {
  let result;
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const serial = String(payload.serial || "").trim();
    result = serial ? checkOrActivate(serial) : { valid: false, reason: "missing_serial" };
  } catch (err) {
    result = { valid: false, reason: "server_error", message: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    message: "授權伺服器運作中。請用 POST 傳送 JSON body，例如 {\"serial\":\"your-serial-here\"}"
  })).setMimeType(ContentService.MimeType.JSON);
}

function getLicenseSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return (SHEET_NAME && ss.getSheetByName(SHEET_NAME)) || ss.getSheets()[0];
}

function checkOrActivate(serial) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getLicenseSheet_();
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { valid: false, reason: "serial_not_found" };

    const header = values[0];
    const colSerial = header.indexOf(COL_SERIAL);
    const colStart = header.indexOf(COL_START);
    const colEnd = header.indexOf(COL_END);
    if (colSerial < 0 || colStart < 0 || colEnd < 0) {
      return { valid: false, reason: "server_error", message: "表頭找不到「" + COL_SERIAL + "」「" + COL_START + "」「" + COL_END + "」欄位" };
    }

    let rowIdx = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][colSerial]).trim() === serial) { rowIdx = i; break; }
    }
    if (rowIdx === -1) return { valid: false, reason: "serial_not_found" };

    const sheetRow = rowIdx + 1;
    let startVal = values[rowIdx][colStart];
    let endVal = values[rowIdx][colEnd];
    const now = new Date();

    if (!startVal) {
      startVal = now;
      sheet.getRange(sheetRow, colStart + 1).setValue(startVal);
    }
    if (!endVal) {
      endVal = new Date(startVal);
      endVal.setMonth(endVal.getMonth() + VALID_AMOUNT);
      sheet.getRange(sheetRow, colEnd + 1).setValue(endVal);
    }

    const endDate = new Date(endVal);
    endDate.setHours(23, 59, 59, 999);
    const valid = now.getTime() <= endDate.getTime();

    return {
      valid: valid,
      reason: valid ? "ok" : "expired",
      activatedAt: new Date(startVal).toISOString(),
      expiresAt: new Date(endVal).toISOString()
    };
  } finally {
    lock.releaseLock();
  }
}
