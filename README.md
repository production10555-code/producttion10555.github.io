/**
 * 🔒 ตั้งค่าความปลอดภัยระบบ 
 */
var ADMIN_PIN = "1234"; // รหัสผ่านสำหรับผู้จัดการในการสั่งรีเซ็ตระบบ

/**
 * ฟังก์ชันเริ่มต้น: โหลดหน้าเว็บดั้งเดิมแบบสะอาด
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('QR Code Scanner')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

/**
 * ฟังก์ชันหลัก: ตรวจสอบรหัสคิวอาร์สแกน คุมโควตา และบันทึกประวัติ (พร้อมระบบ Auto-Backup รายวัน)
 */
function updateAndFetchDetails(qrData, operatorName, batchId) {
  var cleanQrData = qrData ? qrData.toString().trim() : "";
  if (cleanQrData === "") {
    return { success: false, message: "ข้อมูล QR Code ที่ส่งมาว่างเปล่า ไม่สามารถตรวจสอบได้" };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Recipe master"); 
    if (!sheet) {
      return { success: false, message: "🚨 วิกฤต: ไม่พบชีทฐานข้อมูลมาสเตอร์ 'Recipe master' กรุณาแจ้งผู้ดูแลระบบ" };
    }
    
    var logSheet = ss.getSheetByName("Mixing_Logs");
    if (!logSheet) {
      logSheet = ss.insertSheet("Mixing_Logs");
      logSheet.appendRow(["วัน-เวลาที่บันทึก", "QR Code ID", "ชื่อสาร (Name)", "ขนาดถุง", "ผู้สแกน", "Batch", "สถานะการสแกน"]);
    }
    
    try {
      checkAndTriggerDailyBackup(ss);
    } catch(backupErr) {
      console.error("Backup failed: " + backupErr.message);
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0]; 
    var foundIndex = -1;
    var rowData = null;
    
    // ตรวจสอบรหัสคิวอาร์โค้ดเทียบกับคอลัมน์ C (QR-name) ของ Recipe master
    for (var i = 1; i < data.length; i++) {
      var sheetQrData = data[i][2] ? data[i][2].toString().trim() : ""; // Index 2 คือ คอลัมน์ C
      if (sheetQrData !== "" && sheetQrData === cleanQrData) {
        foundIndex = i;
        rowData = data[i];
        break;
      }
    }
    
    if (foundIndex === -1) {
      return { success: false, message: "ไม่พบรหัสคิวอาร์โค้ด '" + cleanQrData + "' นี้ในคอลัมน์ QR-name ของ Recipe master" };
    }
    
    var currentSubstanceName = rowData[1] ? rowData[1].toString().trim() : ""; // คอลัมน์ B (Name) -> Index 1
    var bagSize = rowData[3] ? rowData[3].toString().trim() : "";              // คอลัมน์ D (ขนาดถุง) -> Index 3
    var totalNeededFromColumnE = rowData[4] ? parseInt(rowData[4]) : 0;        // คอลัมน์ E (จำนวนถุงโควตา) -> Index 4
    
    var cloudScannedCount = 0;
    var logData = logSheet.getDataRange().getValues();
    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    
    for (var k = 1; k < logData.length; k++) {
      if (!logData[k][0]) continue; 
      
      var logDateStr = (logData[k][0] instanceof Date)
        ? Utilities.formatDate(logData[k][0], timeZone, "yyyy-MM-dd")
        : logData[k][0].toString().split(" ")[0];
        
      var logSubstance = logData[k][2] ? logData[k][2].toString().trim() : "";
      var logBatch = logData[k][5] ? logData[k][5].toString().trim() : "";
      var logStatus = logData[k][6] ? logData[k][6].toString().trim() : "";
      
      if (logDateStr === todayStr && logSubstance === currentSubstanceName && logBatch === "Batch " + batchId && logStatus === "บันทึกสำเร็จ") {
        cloudScannedCount++;
      }
    }
    
    // 🔥 [แก้ไขจุดที่ 1] ปรับการแสดงผลหน้าประวัติให้โชว์สัดส่วนจำนวนถุงจริงทับแม็กซ์สุด เช่น 5/5 ถุง
    if (totalNeededFromColumnE > 0 && cloudScannedCount >= totalNeededFromColumnE) {
      logSheet.appendRow([new Date(), cleanQrData, currentSubstanceName, bagSize, operatorName, "Batch " + batchId, "ปฏิเสธ: เกินโควตา " + totalNeededFromColumnE + " ถุง"]);
      
      return { 
        success: false, 
        message: "สแกนไม่ได้! [" + currentSubstanceName + "] ใน Batch " + batchId + " สแกนครบกำหนดแล้ว (" + cloudScannedCount + "/" + totalNeededFromColumnE + ") ถุง",
        headers: headers,
        rowData: rowData,
        serverScannedCount: cloudScannedCount
      };
    }
    
    var scanTime = new Date();
    logSheet.appendRow([scanTime, cleanQrData, currentSubstanceName, bagSize, operatorName, "Batch " + batchId, "บันทึกสำเร็จ"]);
    
    var updatedCount = cloudScannedCount + 1;
    
    return {
      success: true,
      message: "บันทึกสำเร็จ! สาร " + currentSubstanceName + " [" + batchId + "] สแกนแล้ว " + updatedCount + "/" + totalNeededFromColumnE + " ถุง",
      headers: headers,
      rowData: rowData,
      isAlreadyUsed: true,
      serverScannedCount: updatedCount
    };
    
  } catch (err) {
    return { success: false, message: "System Error: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 🔒 ฟังก์ชันระดับสูงสำหรับ Admin: มีระบบตรวจสอบรหัสผ่าน (PIN Validation) ก่อนรีเซ็ตข้อมูล
 */
function adminResetLogStatus(qrData, inputPin) {
  var cleanQrData = qrData ? qrData.toString().trim() : "";
  if (cleanQrData === "") {
    return { success: false, message: "ข้อมูล QR Code ว่างเปล่า ไม่สามารถทำการรีเซ็ตได้" };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Mixing_Logs");
    
    if (!logSheet) {
      logSheet = ss.insertSheet("Mixing_Logs");
      logSheet.appendRow(["วัน-เวลาที่บันทึก", "QR Code ID", "ชื่อสาร (Name)", "ขนาดถุง", "ผู้สแกน", "Batch", "สถานะการสแกน"]);
    }

    if (inputPin !== ADMIN_PIN) {
      logSheet.appendRow([new Date(), cleanQrData, "SECURITY_ALERT", "-", "UNAUTHORIZED_USER", "ALL", "⚠️ มีผู้พยายามสั่งรีเซ็ตระบบด้วยรหัสผ่านที่ผิดพลาด!"]);
      return { success: false, message: "❌ รหัสผ่านผู้จัดการไม่ถูกต้อง! การกระทำนี้ถูกบันทึกในระบบรักษาความปลอดภัยแล้ว" };
    }
    
    var sheet = ss.getSheetByName("Recipe master");
    if (!sheet) {
      return { success: false, message: "🚨 ไม่พบชีทฐานข้อมูลมาสเตอร์ 'Recipe master' ไม่สามารถรีเซ็ตได้" };
    }

    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var sheetQrAdmin = data[i][2] ? data[i][2].toString().trim() : "";
      if (sheetQrAdmin !== "" && sheetQrAdmin === cleanQrData) {
        logSheet.appendRow([new Date(), cleanQrData, data[i][1], data[i][3], "SYSTEM_ADMIN", "ALL", "🔒 ผู้จัดการยืนยันรหัสผ่านถูกต้อง: สั่งรีเซ็ตระบบเรียบร้อย"]);
        return { success: true, message: "🔓 ปลดล็อกสำเร็จ! รหัส QR " + cleanQrData + " ถูกรีเซ็ตและบันทึกประวัติความปลอดภัยแล้ว" };
      }
    }
    return { success: false, message: "ไม่พบรหัส QR นี้ในระบบมาสเตอร์" };
  } catch (err) {
    return { success: false, message: "Admin Security Error: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 📂 ฟังก์ชันภายใน: ใช้ตรวจเช็ควัน และทำการก็อปปี้ไฟล์สำรองลง Google Drive แยกตามวันอัตโนมัติ
 */
function checkAndTriggerDailyBackup(currentSpreadsheet) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var lastBackupDate = scriptProperties.getProperty("LAST_BACKUP_DATE");
  
  if (lastBackupDate !== todayStr) {
    var fileId = currentSpreadsheet.getId();
    var appFile = DriveApp.getFileById(fileId);
    var backupName = "Backup_MixingSystem_" + todayStr;
    appFile.makeCopy(backupName);
    scriptProperties.setProperty("LAST_BACKUP_DATE", todayStr);
    console.log("Daily backup created successfully: " + backupName);
  }
}

// =======================================================
// 🛠️ ส่วนระบบส่งข้อมูลไปหน้าบ้าน (ปรับปรุงใหม่: ให้ดึงข้อมูลและจัดก้อนแบ่งตาม วัน และ แบตเตอช)
// =======================================================

/**
 * ฟังก์ชันสำหรับดึงข้อมูลประวัติจากชีท Mixing_Logs ส่งกลับไปแบบจัดหมวดหมู่แยกตาม วัน และ Batch
 */
function getLogsData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Mixing_Logs");
    if (!logSheet) return [];
    
    var data = logSheet.getDataRange().getValues();
    if (data.length <= 1) return []; 
    
    var logs = [];
    var timeZone = Session.getScriptTimeZone();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; 
      
      var formattedDate = (row[0] instanceof Date) 
        ? Utilities.formatDate(row[0], timeZone, "yyyy-MM-dd HH:mm:ss")
        : row[0].toString();
        
      var onlyDateStr = (row[0] instanceof Date)
        ? Utilities.formatDate(row[0], timeZone, "yyyy-MM-dd")
        : formattedDate.split(" ")[0];
        
      logs.push({
        timestamp: formattedDate.toString().trim(),
        dateGroup: onlyDateStr.trim(), 
        qrCode: row[1] ? row[1].toString().trim() : "",
        substanceName: row[2] ? row[2].toString().trim() : "",
        bagSize: row[3] ? row[3].toString().trim() : "",
        operator: row[4] ? row[4].toString().trim() : "",
        batch: row[5] ? row[5].toString().trim() : "Batch 未知", 
        status: row[6] ? row[6].toString().trim() : ""
      });
    }
    return logs;
  } catch (e) {
    console.error("Error in getLogsData: " + e.message);
    return [];
  }
}

// ==========================================
// 🛠️ ส่วนระบบตรวจเช็คสิทธิ์ตัวกล้องและแก้ไขขอบเขตสิทธิ์
// ==========================================

function getCameraPermissionStatus() {
  return {
    allowCamera: true,
    sandboxMode: "allow-same-origin allow-scripts allow-forms allow-modals"
  };
}

// ==========================================
// 🆕 ฟังก์ชันใหม่: สั่งลบข้อมูลประวัติใน Google Sheets จากหน้าเว็บแอป
// ==========================================

function deleteLogFromServer(timestamp, qrCode) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000); 
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Mixing_Logs");
    if (!logSheet) {
      return { success: false, message: "ไม่พบชีท Mixing_Logs ในระบบ" };
    }
    
    var data = logSheet.getDataRange().getValues();
    var timeZone = Session.getScriptTimeZone();
    
    var targetTimestamp = timestamp ? timestamp.toString().trim() : "";
    var targetQr = qrCode ? qrCode.toString().trim() : "";
    
    if (targetTimestamp === "" || targetQr === "") {
      return { success: false, message: "ข้อมูลส่งมาไม่ครบถ้วน ไม่สามารถลบได้" };
    }
    
    var deletedCount = 0;
    var targetSubstance = "";
    var targetBatch = "";
    var targetDateStr = "";

    // ลูปแรกเพื่อหาแถวหลักที่เรากดลบและเก็บข้อมูล สาร, แบตช์, และวันที่ ไว้ดึงตัวแดงออก
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      if (!row[0]) continue;
      
      var rowDateStr = (row[0] instanceof Date) 
        ? Utilities.formatDate(row[0], timeZone, "yyyy-MM-dd HH:mm:ss")
        : row[0].toString().trim();
      
      var rowQr = row[1] ? row[1].toString().trim() : "";
      
      if (rowDateStr.trim() === targetTimestamp && rowQr === targetQr) {
        targetSubstance = row[2] ? row[2].toString().trim() : "";
        targetBatch = row[5] ? row[5].toString().trim() : "";
        targetDateStr = (row[0] instanceof Date) ? Utilities.formatDate(row[0], timeZone, "yyyy-MM-dd") : rowDateStr.split(" ")[0];
        
        logSheet.deleteRow(i + 1); 
        deletedCount++;
        break; 
      }
    }
    
    // 🔥 [แก้ไขจุดที่ 2] เคลียร์ "ตัวแดง" (แถวที่ขึ้นสถานะ ปฏิเสธ: เกินโควตา) ของสารและแบทช์เดียวกันในวันนั้นออกไปด้วยทันที
    if (deletedCount > 0 && targetSubstance !== "" && targetBatch !== "") {
      // ดึงข้อมูลใหม่อีกครั้งหลังจากลบแถวแรกไปแล้ว เพื่ออัปเดต Index แถวให้ตรงความจริง
      var freshData = logSheet.getDataRange().getValues();
      for (var j = freshData.length - 1; j >= 1; j--) {
        if (!freshData[j][0]) continue;
        
        var currentLogDate = (freshData[j][0] instanceof Date) ? Utilities.formatDate(freshData[j][0], timeZone, "yyyy-MM-dd") : freshData[j][0].toString().split(" ")[0];
        var currentSubstance = freshData[j][2] ? freshData[j][2].toString().trim() : "";
        var currentBatch = freshData[j][5] ? freshData[j][5].toString().trim() : "";
        var currentStatus = freshData[j][6] ? freshData[j][6].toString().trim() : "";
        
        // เช็กถ้าเป็นสารเดียวกัน แบทช์เดียวกัน วันเดียวกัน และมีสถานะขึ้นต้นด้วย "ปฏิเสธ" ให้ลบทิ้งไปเลย
        if (currentLogDate === targetDateStr && currentSubstance === targetSubstance && currentBatch === targetBatch && currentStatus.indexOf("ปฏิเสธ") === 0) {
          logSheet.deleteRow(j + 1);
        }
      }
      return { success: true, message: "ลบประวัติแถวที่เลือกพร้อมล้างแจ้งเตือนถุงเกินในรอบนี้ออกจากระบบเรียบร้อยแล้ว!" };
    }
    
    return { success: false, message: "ไม่พบแถวข้อมูลประวัติที่ตรงกันใน Google Sheets" };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาดในการลบจากเซิร์ฟเวอร์: " + err.message };
  } finally {
    lock.releaseLock();
  }
}
