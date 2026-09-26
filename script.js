/**
 * Google Apps Script backend for the dessert shop order form.
 *
 * SETUP:
 * 1. Create a Google Sheet. (แถวหัวตารางจะถูกสร้างให้อัตโนมัติในชีตรายวัน)
 *    คอลัมน์: Timestamp | ชื่อลูกค้า | เบอร์โทร | ชื่อ Facebook | ที่อยู่ | รายการ | วิธีชำระเงิน | ค่าธรรมเนียมปลายทาง | ยอดรวม | หมายเหตุ | สลิปโอนเงิน
 * 2. In that sheet: Extensions > Apps Script.
 * 3. Delete any starter code and paste this whole file in.
 * 4. (Optional) Create a Google Drive folder to keep slip images in, open it,
 *    copy the folder ID from its URL, and paste it into SLIP_FOLDER_ID below.
 *    If you leave it blank, slips are saved to a folder named "สลิปออเดอร์ขนม"
 *    that this script creates automatically the first time it runs.
 * 5. Click Deploy > New deployment > select type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Click Deploy, authorize the permissions Google asks for.
 * 7. Copy the "Web app URL" it gives you.
 * 8. Paste that URL into the SCRIPT_URL constant near the top of index.html.
 *
 * ชีตรายวัน: ออเดอร์แต่ละวันจะถูกเขียนลงแท็บชื่อตามวันที่ (เช่น 2026-09-26)
 * ถ้ายังไม่มีแท็บวันนั้น สคริปต์จะสร้างใหม่และใส่หัวตารางให้อัตโนมัติ
 */

const SLIP_FOLDER_ID = "1AygB9R__9ACj0dRUGO-aVZLYV4q-PE7e"; // โฟลเดอร์ https://drive.google.com/drive/folders/1AygB9R__9ACj0dRUGO-aVZLYV4q-PE7e
const SLIP_FOLDER_NAME = "สลิปออเดอร์ขนม";

const SHEET_HEADERS = [
  "Timestamp",
  "ชื่อลูกค้า",
  "เบอร์โทร",
  "ชื่อ Facebook",
  "ที่อยู่",
  "รายการ",
  "วิธีชำระเงิน",
  "ค่าธรรมเนียมปลายทาง",
  "ยอดรวม",
  "หมายเหตุ",
  "สลิปโอนเงิน"
];

/** ชื่อแท็บตามวันที่ไทย เช่น 2026-09-26 */
function todaySheetName() {
  return Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
}

/** หาแท็บของวันนี้ ถ้ายังไม่มีให้สร้างใหม่ + ใส่หัวตาราง */
function getDailySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = todaySheetName();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(SHEET_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight("bold");
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight("bold");
  }

  return sheet;
}

function getSlipFolder() {
  if (SLIP_FOLDER_ID) {
    return DriveApp.getFolderById(SLIP_FOLDER_ID);
  }
  const existing = DriveApp.getFoldersByName(SLIP_FOLDER_NAME);
  if (existing.hasNext()) {
    return existing.next();
  }
  return DriveApp.createFolder(SLIP_FOLDER_NAME);
}

function saveSlip(base64, mimeType, fileName, orderTimestamp) {
  if (!base64) return "";
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      mimeType || "image/jpeg",
      (orderTimestamp || Date.now()) + "_" + (fileName || "slip.jpg")
    );
    const file = getSlipFolder().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    return "อัปโหลดสลิปไม่สำเร็จ: " + err.message;
  }
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: "Invalid payload" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getDailySheet();
  const slipUrl = saveSlip(data.slipBase64, data.slipMimeType, data.slipFileName, data.timestamp);

  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.name || "",
    data.phone || "",
    data.fbName || "",
    data.address || "",
    data.items || "",
    data.paymentMethod || "",
    data.codFee || 0,
    data.total || 0,
    data.note || "",
    slipUrl
  ]);

  return ContentService.createTextOutput(
    JSON.stringify({ status: "ok", sheet: sheet.getName() })
  ).setMimeType(ContentService.MimeType.JSON);
}

// Optional: lets you open the Web App URL directly in a browser to check it's alive.
function doGet(e) {
  return ContentService.createTextOutput("Order form backend is running.");
}
