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

/**
 * JBS Bakery - Google Apps Script Backend
 *
 * Website -> Apps Script -> Google Sheets + Google Drive
 */

const SLIP_FOLDER_ID = "1AygB9R__9ACj0dRUGO-aVZLYV4q-PE7e";

const SPREADSHEET_ID =
  "1btwZ20ygLfLmpCcCeYIzD6PWWiaZ-I0ln5gWiNoqDWM";


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


/* ==================================================
   DATE
================================================== */

function todaySheetName() {

  return Utilities.formatDate(
    new Date(),
    "Asia/Bangkok",
    "yyyy-MM-dd"
  );

}


/* ==================================================
   GOOGLE SHEET
================================================== */

function getDailySheet() {

  const ss =
    SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheetName =
    todaySheetName();

  let sheet =
    ss.getSheetByName(sheetName);


  // ถ้ายังไม่มี Sheet วันนี้
  if (!sheet) {

    sheet =
      ss.insertSheet(sheetName);

    sheet.appendRow(SHEET_HEADERS);

    sheet.setFrozenRows(1);

    sheet
      .getRange(
        1,
        1,
        1,
        SHEET_HEADERS.length
      )
      .setFontWeight("bold");
  }


  // ถ้ามี Sheet แต่ไม่มีข้อมูลเลย
  else if (sheet.getLastRow() === 0) {

    sheet.appendRow(SHEET_HEADERS);

    sheet.setFrozenRows(1);

    sheet
      .getRange(
        1,
        1,
        1,
        SHEET_HEADERS.length
      )
      .setFontWeight("bold");
  }


  return sheet;

}


/* ==================================================
   ORDER ID
================================================== */

function generateOrderId(sheet) {

  const datePart =
    Utilities.formatDate(
      new Date(),
      "Asia/Bangkok",
      "yyMMdd"
    );

  /*
    Row 1 = Header

    ถ้าไม่มี order:
    getLastRow() = 1
    Order แรก = 0001

    Order ต่อไป:
    0002, 0003...
  */

  const orderNumber =
    Math.max(1, sheet.getLastRow());


  const runningNumber =
    String(orderNumber).padStart(4, "0");


  return (
    "JB" +
    datePart +
    "-" +
    runningNumber
  );

}


/* ==================================================
   GOOGLE DRIVE
================================================== */

function getSlipFolder() {

  return DriveApp
    .getFolderById(SLIP_FOLDER_ID);

}


/* ==================================================
   SAVE SLIP
================================================== */

function saveSlip(
  base64,
  mimeType,
  fileName,
  orderId
) {

  if (!base64) {
    return "";
  }


  try {

    const bytes =
      Utilities.base64Decode(base64);


    const safeFileName =
      orderId +
      "_" +
      (fileName || "slip.jpg");


    const blob =
      Utilities.newBlob(
        bytes,
        mimeType || "image/jpeg",
        safeFileName
      );


    const folder =
      getSlipFolder();


    const file =
      folder.createFile(blob);


    /*
      ทำให้กด link จาก Sheet แล้วดูสลิปได้
    */

    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );


    return file.getUrl();


  } catch (error) {

    console.error(
      "SAVE SLIP ERROR:",
      error
    );

    return "";

  }

}


/* ==================================================
   RECEIVE ORDER FROM WEBSITE
================================================== */

function doPost(e) {

  try {

    /* ------------------------------
       1. อ่านข้อมูลจากเว็บไซต์
    ------------------------------ */

    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      throw new Error(
        "ไม่พบข้อมูลจากเว็บไซต์"
      );

    }


    const data =
      JSON.parse(
        e.postData.contents
      );


    /* ------------------------------
       2. เปิด Sheet ของวันนี้
    ------------------------------ */

    const sheet =
      getDailySheet();


    /* ------------------------------
       3. Order ID

       ถ้า frontend ส่ง orderId มา
       ใช้อันนั้น

       ถ้าไม่ได้ส่งมา
       Apps Script สร้างให้
    ------------------------------ */

    const orderId =
      data.orderId ||
      generateOrderId(sheet);


    /* ------------------------------
       4. Upload Slip
    ------------------------------ */

    let slipUrl = "";

    let slipCell = "";


    if (data.slipBase64) {

      slipUrl =
        saveSlip(
          data.slipBase64,
          data.slipMimeType,
          data.slipFileName,
          orderId
        );


      if (slipUrl) {

        slipCell =
          '=HYPERLINK("' +
          slipUrl +
          '","ดูสลิป")';

      } else {

        slipCell =
          "อัปโหลดสลิปไม่สำเร็จ";

      }

    }


    /* ------------------------------
       5. บันทึก Order
    ------------------------------ */

    sheet.appendRow([

      data.timestamp ||
        new Date().toISOString(),

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

      slipCell

    ]);


    SpreadsheetApp.flush();


    /* ------------------------------
       6. Response
    ------------------------------ */

    return ContentService
      .createTextOutput(

        JSON.stringify({

          status: "ok",

          orderId: orderId,

          sheet: sheet.getName(),

          slipUrl: slipUrl

        })

      )
      .setMimeType(
        ContentService.MimeType.JSON
      );


  } catch (error) {

    console.error(
      "ORDER ERROR:",
      error
    );


    return ContentService
      .createTextOutput(

        JSON.stringify({

          status: "error",

          message:
            error.message

        })

      )
      .setMimeType(
        ContentService.MimeType.JSON
      );

  }

}


/* ==================================================
   TEST WEB APP

   เปิด /exec ใน Browser
   ถ้าขึ้น status ok = Backend ทำงาน
================================================== */

function doGet() {

  return ContentService
    .createTextOutput(

      JSON.stringify({

        status: "ok",

        message:
          "JBS Bakery backend is running"

      })

    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}