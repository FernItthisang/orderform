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

const SLIP_FOLDER_ID = "1AygB9R__9ACj0dRUGO-aVZLYV4q-PE7e";
const SLIP_FOLDER_NAME = "สลิป";

const SPREADSHEET_ID = "1btwZ20ygLfLmpCcCeYIzD6PWWiaZ-I0ln5gWiNoqDWM";

const SHEET_HEADERS = [
  "Timestamp",
  "รหัสอ้างอิง",
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

/* =========================
   วันที่ประเทศไทย
========================= */

function todaySheetName() {
  return Utilities.formatDate(
    new Date(),
    "Asia/Bangkok",
    "yyyy-MM-dd"
  );
}


/* =========================
   Google Sheet
========================= */

function getDailySheet() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const name = todaySheetName();

  let sheet = ss.getSheetByName(name);


  // ถ้ายังไม่มี Sheet ของวันนี้ → สร้าง
  if (!sheet) {

    sheet = ss.insertSheet(name);

    sheet.appendRow(SHEET_HEADERS);

    sheet.setFrozenRows(1);

    sheet
      .getRange(1, 1, 1, SHEET_HEADERS.length)
      .setFontWeight("bold");

  }

  // ถ้ามี Sheet แต่ยังว่าง
  else if (sheet.getLastRow() === 0) {

    sheet.appendRow(SHEET_HEADERS);

    sheet.setFrozenRows(1);

    sheet
      .getRange(1, 1, 1, SHEET_HEADERS.length)
      .setFontWeight("bold");
  }


  return sheet;
}


/* =========================
   Google Drive Slip Folder
========================= */

function getSlipFolder() {

  if (SLIP_FOLDER_ID) {
    return DriveApp.getFolderById(SLIP_FOLDER_ID);
  }


  const existing =
    DriveApp.getFoldersByName(SLIP_FOLDER_NAME);


  if (existing.hasNext()) {
    return existing.next();
  }


  return DriveApp.createFolder(SLIP_FOLDER_NAME);
}


/* =========================
   Save Slip
========================= */

function saveSlip(base64, mimeType, fileName, orderId) {

  if (!base64) {
    return "";
  }

  try {

    const bytes = Utilities.base64Decode(base64);

    const blob = Utilities.newBlob(
      bytes,
      mimeType || "image/jpeg",
      (orderId || Date.now()) + "_" + (fileName || "slip.jpg")
    );

    const folder = DriveApp.getFolderById(SLIP_FOLDER_ID);

    const file = folder.createFile(blob);

    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    return file.getUrl();

  } catch (error) {

    console.error("SLIP ERROR:", error);

    return "อัปโหลดสลิปไม่สำเร็จ: "
    + err.message;

  }
}


/* =========================
   RECEIVE ORDER
========================= */

function generateOrderId(sheet) {

  const datePart = Utilities.formatDate(
    new Date(),
    "Asia/Bangkok",
    "yyMMdd"
  );

  // แถว 1 = header
  // ดังนั้น order แรกจะเป็น 0001
  const orderNumber = Math.max(1, sheet.getLastRow());

  const runningNumber =
    String(orderNumber).padStart(4, "0");

  return "JB" + datePart + "-" + runningNumber;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getDailySheet();

    // พยายามเก็บสลิป แต่ถ้าพัง ห้ามทำให้ออเดอร์หาย
    let slipUrl = "";
    let slipStatus = "";

    if (data.slipBase64) {
      try {
        slipUrl = saveSlip(
          data.slipBase64,
          data.slipMimeType,
          data.slipFileName,
          data.orderId
        );

        if (slipUrl) {
          slipStatus = `=HYPERLINK("${slipUrl}","ดูสลิป")`;
        } else {
          slipStatus = "อัปโหลดสลิปไม่สำเร็จ";
        }

      } catch (slipError) {
        console.error("SLIP ERROR:", slipError);
        slipStatus = "อัปโหลดสลิปไม่สำเร็จ";
      }
    }

    // สำคัญ: บันทึก Order ไม่ว่าสลิปจะสำเร็จหรือไม่
    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.orderId || "",
      data.name || "",
      data.phone || "",
      data.fbName || "",
      data.address || "",
      data.items || "",
      data.paymentMethod || "",
      Number(data.codFee) || 0,
      Number(data.total) || 0,
      data.note || "",
      slipStatus
    ]);

    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "ok",
        orderId: data.orderId || "",
        slipUrl: slipUrl
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {

    console.error("ORDER ERROR:", error);

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/* =========================
   TEST WEB APP
========================= */

function doGet() {

  return ContentService
  .createTextOutput(
    JSON.stringify({
      status: "ok",
      orderId: orderId,
      sheet: sheet.getName()
    })
  )
  .setMimeType(ContentService.MimeType.JSON);
}