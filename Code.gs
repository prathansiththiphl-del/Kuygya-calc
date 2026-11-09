const SHEET_NAME = "Survey";
const MAX_PAGE_SIZE = 100;

const HEADERS = [
  "วันที่",
  "ชื่อจริง",
  "ชั้น",
  "เลขที่",
  "วิชาที่ชอบ",
  "วิชาที่ไม่ชอบ",
  "เหตุผล"
];

/** Health check for the deployed Apps Script web app. */
function doGet() {
  return response({
    success: true,
    message: "Survey API is running"
  });
}

/** Accepts form-encoded requests and JSON request bodies. */
function doPost(e) {
  try {
    const data = getRequestData(e);
    const action = String(data.action || "");

    if (action === "submit") return submitSurvey(data);
    if (action === "adminLogin") return adminLogin(data);
    if (action === "adminStats") return adminStats(data);

    return response({ success: false, message: "ไม่พบคำสั่ง" });
  } catch (error) {
    console.error(error);
    return response({
      success: false,
      message: error && error.message ? error.message : "เกิดข้อผิดพลาด"
    });
  }
}

function getRequestData(e) {
  if (!e) return {};
  if (e.postData && e.postData.contents) {
    const contentType = String(e.postData.type || "").toLowerCase();
    if (contentType.indexOf("application/json") !== -1) {
      const parsed = JSON.parse(e.postData.contents);
      if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
        throw new Error("รูปแบบข้อมูลไม่ถูกต้อง");
      }
      return parsed;
    }
  }
  return e.parameter || {};
}

// เตรียมฐานข้อมูล — รันฟังก์ชันนี้หนึ่งครั้งจาก Apps Script Editor
function setupSurvey() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SURVEY_ID");
  let ss;

  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create("Kuygya Survey Database");
    props.setProperty("SURVEY_ID", ss.getId());
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

  Logger.log("Database URL: " + ss.getUrl());
  return ss.getUrl();
}

function getSurveySheet() {
  const id = PropertiesService.getScriptProperties().getProperty("SURVEY_ID");
  if (!id) throw new Error("กรุณารัน setupSurvey ก่อน");

  const sheet = SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("ไม่พบชีต Survey");
  return sheet;
}

function submitSurvey(data) {
  const fields = {
    name: { label: "ชื่อจริง", max: 100 },
    classroom: { label: "ชั้น", max: 30 },
    number: { label: "เลขที่", max: 20 },
    favorite: { label: "วิชาที่ชอบ", max: 100 },
    unfavorite: { label: "วิชาที่ไม่ชอบ", max: 100 },
    reason: { label: "เหตุผล", max: 500 }
  };
  const values = {};

  for (const key in fields) {
    const value = String(data[key] == null ? "" : data[key]).trim();
    if (!value) {
      return response({ success: false, message: "กรุณากรอกข้อมูลให้ครบ" });
    }
    if (value.length > fields[key].max) {
      return response({
        success: false,
        message: fields[key].label + "ยาวเกินกำหนด (สูงสุด " + fields[key].max + " ตัวอักษร)"
      });
    }
    values[key] = safeSheetText(value);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSurveySheet().appendRow([
      new Date(),
      values.name,
      values.classroom,
      values.number,
      values.favorite,
      values.unfavorite,
      values.reason
    ]);
  } finally {
    lock.releaseLock();
  }

  return response({ success: true, message: "บันทึกข้อมูลสำเร็จ" });
}

// Prevent user input from being interpreted as a spreadsheet formula.
function safeSheetText(value) {
  return /^[\\s]*[=+@-]/.test(value) ? "'" + value : value;
}

function adminLogin(data) {
  const savedPassword = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  if (!savedPassword) {
    return response({ success: false, message: "ยังไม่ได้ตั้งรหัสแอดมิน" });
  }
  if (String(data.password || "") !== savedPassword) {
    return response({ success: false, message: "รหัสผ่านไม่ถูกต้อง" });
  }

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put("ADMIN_" + token, "authorized", 1800);
  return response({ success: true, token: token });
}

// ส่งข้อมูลเป็นหน้า จำกัดไม่เกิน 100 รายการต่อคำขอ
function adminStats(data) {
  const token = String(data.token || "");
  const authorized = CacheService.getScriptCache().get("ADMIN_" + token);
  if (authorized !== "authorized") {
    return response({ success: false, message: "ไม่มีสิทธิ์เข้าถึง" });
  }

  const sheet = getSurveySheet();
  const total = Math.max(0, sheet.getLastRow() - 1);
  const requestedLimit = parseInt(data.limit, 10);
  const requestedOffset = parseInt(data.offset, 10);
  const limit = Math.min(MAX_PAGE_SIZE, Math.max(1, isNaN(requestedLimit) ? 50 : requestedLimit));
  const offset = Math.max(0, isNaN(requestedOffset) ? 0 : requestedOffset);
  const count = Math.max(0, Math.min(limit, total - offset));
  const values = count ? sheet.getRange(offset + 2, 1, count, HEADERS.length).getValues() : [];

  const records = values.map(row => ({
    date: row[0],
    name: row[1],
    classroom: row[2],
    number: row[3],
    favorite: row[4],
    unfavorite: row[5],
    reason: row[6]
  }));

  return response({
    success: true,
    total: total,
    offset: offset,
    limit: limit,
    records: records
  });
}

function response(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
