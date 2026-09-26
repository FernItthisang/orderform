/**
 * Google Apps Script backend for the dessert shop order form.
 *
 * SETUP:
 * 1. Create a Google Sheet. (แถวหัวตารางจะถูกสร้างให้อัตโนมัติในชีตรายวัน)
 *    คอลัมน์: Timestamp | รหัสอ้างอิง | ชื่อลูกค้า | เบอร์โทร | ชื่อ Facebook | ที่อยู่ | รายการ | วิธีชำระเงิน | ค่าธรรมเนียมปลายทาง | ยอดรวม | หมายเหตุ | สลิปโอนเงิน
 * 2. ในชีตนั้น: Extensions > Apps Script
 * 3. ลบโค้ดเริ่มต้นแล้ววางไฟล์นี้ทั้งหมดแทน
 * 4. (Optional) สร้างโฟลเดอร์ Google Drive ไว้เก็บรูปสลิป เปิดโฟลเดอร์นั้น
 *    คัดลอก folder ID จาก URL แล้วใส่ใน SLIP_FOLDER_ID ด้านล่าง
 * 5. Deploy > New deployment > เลือกประเภท "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. กด Deploy แล้วอนุญาตสิทธิ์ที่ Google ขอ
 * 7. คัดลอก "Web app URL" ที่ได้
 * 8. นำ URL นั้นไปใส่ใน SCRIPT_URL ในไฟล์ index.html
 *
 * ชีตรายวัน: ออเดอร์แต่ละวันจะถูกเขียนลงแท็บชื่อตามวันที่ (เช่น 2026-09-26)
 * ถ้ายังไม่มีแท็บวันนั้น สคริปต์จะสร้างใหม่และใส่หัวตารางให้อัตโนมัติ
 *
 * หมายเหตุการแก้ไขจากไฟล์เดิม:
 * - เดิมมีฟังก์ชัน doPost ประกาศซ้ำ 2 อัน ใน Apps Script ตัวที่อยู่ล่างสุด
 *   จะทับตัวบนเสมอ ทำให้โค้ดตัวบน (ที่ทำงานถูกต้อง) ไม่เคยถูกเรียกใช้จริง
 *   ไฟล์นี้รวมเหลือ doPost เดียว
 * - saveSlip() คืนค่าเป็น string ของ URL ตรง ๆ (ไม่ใช่ object ที่มี .url)
 *   จุดที่เรียกใช้จึงต้องเช็คค่าตรง ๆ ไม่ใช่ .url
 * - เอาตัวแปร slipUrl ที่ไม่เคยถูกประกาศ (ทำให้เกิด error ทุกครั้ง) ออก
 */

const SLIP_FOLDER_ID = "1AygB9R__9ACj0dRUGO-aVZLYV4q-PE7e";
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

const SLIP_COLUMN_INDEX = SHEET_HEADERS.indexOf("สลิปโอนเงิน") + 1; // = 12

/* ==================================================
   DATE
================================================== */

function todaySheetName() {
  return Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
}

/* ==================================================
   GOOGLE SHEET
================================================== */

function getDailySheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = todaySheetName();
  let sheet = ss.getSheetByName(sheetName);

  const needsHeader = !sheet || sheet.getLastRow() === 0;

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (needsHeader) {
    sheet.appendRow(SHEET_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight("bold");
  }

  return sheet;
}

/* ==================================================
   ORDER ID
================================================== */

function generateOrderId(sheet) {
  const datePart = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyMMdd");

  // Row 1 = Header, ถ้าไม่มี order เลย getLastRow() = 1 → order แรก = 0001
  const orderNumber = Math.max(1, sheet.getLastRow());
  const runningNumber = String(orderNumber).padStart(4, "0");

  return "JB" + datePart + "-" + runningNumber;
}

/* ==================================================
   SAVE SLIP (คืนค่าเป็น URL string หรือ "" ถ้าไม่มีไฟล์)
================================================== */

function saveSlip(base64, mimeType, fileName, orderId) {
  if (!base64) {
    return "";
  }

  const bytes = Utilities.base64Decode(base64);

  const blob = Utilities.newBlob(
    bytes,
    mimeType || "image/jpeg",
    (orderId || Date.now()) + "_" + (fileName || "slip.jpg")
  );

  const folder = DriveApp.getFolderById(SLIP_FOLDER_ID);
  const file = folder.createFile(blob);

  // ทำให้ไฟล์เปิดดูได้จากลิงก์ (จำเป็นถ้าจะใช้ =IMAGE() หรือให้คนอื่นกดลิงก์ดูได้)
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

/* ==================================================
   RECEIVE ORDER FROM WEBSITE
================================================== */

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("ไม่พบข้อมูลจากเว็บไซต์");
    }

    const data = JSON.parse(e.postData.contents);
    const sheet = getDailySheet();
    const orderId = data.orderId || generateOrderId(sheet);

    /* ------------------------------
       Upload Slip
    ------------------------------ */

    let slipUrl = "";
    let slipCellValue = "";

    if (data.slipBase64) {
      try {
        slipUrl = saveSlip(
          data.slipBase64,
          data.slipMimeType,
          data.slipFileName,
          orderId
        );

        if (!slipUrl) {
          slipCellValue = "อัปโหลดสลิปไม่สำเร็จ";
        }
      } catch (slipError) {
        console.error("SLIP ERROR:", slipError);
        slipUrl = "";
        slipCellValue = "สลิปอยู่ใน Drive แต่สร้างลิงก์ไม่สำเร็จ";
      }
    }

    /* ------------------------------
       บันทึก Order (ยังไม่ใส่ค่าลิงก์สลิป ใส่ทีหลังด้วย RichText)
    ------------------------------ */

    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      orderId,
      data.name || "",
      data.phone || "",
      data.fbName || "",
      data.address || "",
      data.items || "",
      data.paymentMethod || "",
      Number(data.codFee) || 0,
      Number(data.total) || 0,
      data.note || "",
      slipCellValue // ว่าง หรือข้อความ error ถ้าอัปโหลดไม่สำเร็จ
    ]);

    const newRow = sheet.getLastRow();

    if (slipUrl) {
      const richText = SpreadsheetApp
        .newRichTextValue()
        .setText("ดูสลิป")
        .setLinkUrl(slipUrl)
        .build();

      sheet.getRange(newRow, SLIP_COLUMN_INDEX).setRichTextValue(richText);
    }

    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "ok",
        orderId: orderId,
        sheet: sheet.getName(),
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

/* ==================================================
   TEST WEB APP
   เปิด /exec ใน Browser ถ้าขึ้น status ok = Backend ทำงาน
================================================== */

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: "ok",
      message: "JBS Bakery backend is running"
    }))
    .setMimeType(ContentService.MimeType.JSON);
}