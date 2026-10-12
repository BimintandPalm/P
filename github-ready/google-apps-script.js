/**
 * =====================================================================
 * Bugfinn Pattern Manager - Google Apps Script
 * เชื่อมต่อ Google Sheet เพื่อเก็บและซิงค์ข้อมูล Pattern ทั้งหมด
 * =====================================================================
 * 
 * วิธีติดตั้ง (ทำเพียงครั้งเดียว):
 * 1. เปิด Google Sheet ที่คุณต้องการใช้เก็บข้อมูล (หรือสร้างใหม่ที่ https://sheets.new)
 * 2. ไปที่เมนู "ส่วนขยาย" (Extensions) > "Apps Script"
 * 3. ลบโค้ดเดิมทั้งหมดในไฟล์ Code.gs แล้ววางโค้ดชุดนี้ลงไปทั้งหมด
 * 4. กดบันทึก (ไอคอนแผ่นดิสก์ หรือ Ctrl + S)
 * 5. กดปุ่มสีน้ำเงิน "ทำให้ใช้งานได้" (Deploy) มุมขวาบน > เลือก "การทำให้ใช้งานได้รายการใหม่" (New deployment)
 * 6. กดที่ไอคอนรูปฟันเฟืองด้านซ้าย > เลือกประเภท "เว็บแอป" (Web app)
 *    - คำอธิบาย: Bugfinn Pattern API
 *    - ทำงานในฐานะ (Execute as): ตัวฉัน (Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง (Who has access): ทุกคน (Anyone)  <--- *** สำคัญมาก! ต้องเลือก Anyone ***
 * 7. กดปุ่ม "ทำให้ใช้งานได้" (Deploy) 
 *    - (หากมีหน้าต่างถามสิทธิ์ ให้กด "ให้สิทธิ์เข้าถึง" > เลือกบัญชี Google > กด Advanced > ไปที่... (ไม่ปลอดภัย) > ยอมรับ)
 * 8. คัดลอก "URL เว็บแอป" (Web App URL ที่ลงท้ายด้วย /exec)
 * 9. นำ URL นั้นมาวางในเมนู "📊 ตั้งค่า Google Sheet" บนหน้าเว็บ Bugfinn Pattern Manager
 */

const SHEET_NAME = "Patterns";

// ฟังก์ชันเข้าถึง Sheet หากยังไม่มีให้สร้างและตั้งหัวตารางอัตโนมัติ
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
  }

  // ถ้ายังไม่มีข้อมูล หรือไม่มีหัวตาราง ให้สร้างหัวตาราง
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "ชื่อ Pattern", "หมวดหมู่", "ข้อความ", "รายการโปรด", "อัปเดตล่าสุด"]);
    const headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#ead5da");
    headerRange.setFontColor("#49363b");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120); // ID
    sheet.setColumnWidth(2, 220); // ชื่อ
    sheet.setColumnWidth(3, 160); // หมวดหมู่
    sheet.setColumnWidth(4, 380); // ข้อความ
    sheet.setColumnWidth(5, 100); // รายการโปรด
    sheet.setColumnWidth(6, 170); // อัปเดตล่าสุด
  }
  return sheet;
}

// Helper ส่งข้อมูลกลับเป็น JSON
function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ดึงข้อมูล Pattern ทั้งหมด (GET)
 */
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const rows = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      return jsonOutput({ status: "success", count: 0, data: [] });
    }

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const id = String(row[0] || "").trim();
      const name = String(row[1] || "").trim();
      if (!id && !name) continue; // ข้ามบรรทัดว่าง

      data.push({
        id: id || ("p_" + i),
        name: name,
        category: String(row[2] || "📌 อื่น ๆ"),
        content: String(row[3] || ""),
        fav: String(row[4]).toLowerCase() === "true" || row[4] === true,
        updated_at: row[5] ? String(row[5]) : ""
      });
    }

    return jsonOutput({
      status: "success",
      count: data.length,
      data: data
    });
  } catch (err) {
    return jsonOutput({ status: "error", message: err.toString() });
  }
}

/**
 * บันทึก, แก้ไข, ลบ หรืออัปโหลดทั้งหมด (POST)
 */
function doPost(e) {
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        payload = e.parameter || {};
      }
    } else if (e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action || "add";
    const sheet = getOrCreateSheet();
    const nowStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

    // 1. เพิ่ม Pattern ใหม่
    if (action === "add") {
      const item = payload.item || payload;
      const id = String(item.id || ("p_" + new Date().getTime()));
      const name = String(item.name || "").trim();
      const category = String(item.category || "📌 อื่น ๆ").trim();
      const content = String(item.content || "");
      const fav = item.fav ? true : false;

      // ตรวจสอบว่ามี ID หรือชื่อซ้ำหรือไม่ ถ้าซ้ำให้อัปเดต
      const rows = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === id || (name && String(rows[i][1]) === name)) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow > 0) {
        sheet.getRange(foundRow, 1, 1, 6).setValues([[id, name, category, content, fav, nowStr]]);
        return jsonOutput({ status: "success", action: "updated", id: id });
      } else {
        sheet.appendRow([id, name, category, content, fav, nowStr]);
        return jsonOutput({ status: "success", action: "added", id: id });
      }
    }

    // 2. แก้ไข Pattern เดิม
    if (action === "update" || action === "edit") {
      const item = payload.item || payload;
      const id = String(item.id || "");
      const name = String(item.name || "").trim();
      const rows = sheet.getDataRange().getValues();
      let foundRow = -1;

      for (let i = 1; i < rows.length; i++) {
        if ((id && String(rows[i][0]) === id) || (!id && name && String(rows[i][1]) === name)) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow > 0) {
        const curFav = rows[foundRow - 1][4];
        const newFav = item.fav !== undefined ? item.fav : curFav;
        sheet.getRange(foundRow, 1, 1, 6).setValues([[
          id || rows[foundRow - 1][0],
          name || rows[foundRow - 1][1],
          item.category !== undefined ? item.category : rows[foundRow - 1][2],
          item.content !== undefined ? item.content : rows[foundRow - 1][3],
          newFav,
          nowStr
        ]]);
        return jsonOutput({ status: "success", action: "updated", id: id });
      } else {
        // หากไม่พบ ให้สร้างใหม่
        sheet.appendRow([id || ("p_" + new Date().getTime()), name, item.category || "📌 อื่น ๆ", item.content || "", !!item.fav, nowStr]);
        return jsonOutput({ status: "success", action: "added_as_new" });
      }
    }

    // 3. ลบ Pattern
    if (action === "delete") {
      const id = String(payload.id || (payload.item && payload.item.id) || "");
      const name = String(payload.name || (payload.item && payload.item.name) || "");
      const rows = sheet.getDataRange().getValues();
      let foundRow = -1;

      for (let i = 1; i < rows.length; i++) {
        if ((id && String(rows[i][0]) === id) || (name && String(rows[i][1]) === name)) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow > 0) {
        sheet.deleteRow(foundRow);
        return jsonOutput({ status: "success", action: "deleted", id: id });
      } else {
        return jsonOutput({ status: "warning", message: "Pattern not found to delete" });
      }
    }

    // 4. บันทึก / อัปโหลดข้อมูลทั้งหมด (Batch Sync / Upload All)
    if (action === "batch_save" || action === "sync_all") {
      const items = payload.items || [];
      if (!Array.isArray(items)) {
        return jsonOutput({ status: "error", message: "items must be an array" });
      }

      // ลบแถวข้อมูลเดิมทั้งหมด (เก็บแถว 1 หัวตารางไว้)
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }

      if (items.length > 0) {
        const rowsToAdd = items.map((x, idx) => [
          String(x.id || ("p_" + (new Date().getTime() + idx))),
          String(x.name || "").trim(),
          String(x.category || "📌 อื่น ๆ").trim(),
          String(x.content || ""),
          x.fav ? true : false,
          nowStr
        ]);
        sheet.getRange(2, 1, rowsToAdd.length, 6).setValues(rowsToAdd);
      }

      return jsonOutput({ status: "success", action: "batch_saved", count: items.length });
    }

    return jsonOutput({ status: "error", message: "Unknown action: " + action });
  } catch (err) {
    return jsonOutput({ status: "error", message: err.toString() });
  }
}
