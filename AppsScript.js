// AppsScript.js - FINAL FIXED VERSION

const SHEET_NAME = 'Sheet1';
const FOLDER_ID = ''; 

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone(); // Ikuti zona waktu Spreadsheet
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const data = JSON.parse(e.postData.contents);
    
    const type = data.type; 
    // Normalisasi Nama: Huruf kecil, trim spasi, hapus spasi ganda
    const nameInput = data.name.toString().toLowerCase().replace(/\s+/g, ' ').trim();
    const timestamp = new Date(data.timestamp);
    const dateStr = Utilities.formatDate(timestamp, tz, "yyyy-MM-dd");
    const timeStr = Utilities.formatDate(timestamp, tz, "HH:mm:ss");
    const locationName = data.address || "Lokasi tidak diketahui";
    const coords = `${data.lat}, ${data.lng}`;
    
    // --- PROSES FOTO ---
    let photoUrl = "";
    if (data.photo && data.photo.includes("base64,")) {
      const base64Data = data.photo.split(",")[1];
      const decoded = Utilities.base64Decode(base64Data);
      const blob = Utilities.newBlob(decoded, "image/jpeg", `Absen_${nameInput}_${dateStr}_${type}.jpg`);
      
      let folder;
      if (FOLDER_ID) {
        folder = DriveApp.getFolderById(FOLDER_ID);
      } else {
        folder = DriveApp.getFileById(ss.getId()).getParents().next(); 
      }
      
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Tanggal", "Nama", "Jam Masuk", "Jam Pulang", "Lokasi Masuk", "Lokasi Pulang", "Foto Masuk", "Foto Pulang", "Durasi Bekerja"]);
      sheet.setRowHeight(1, 30);
      sheet.getRange("A1:I1").setBackground("#e11d48").setFontColor("white").setFontWeight("bold");
    }
    
    const rowHeight = 80;

    if (type === 'Masuk') {
      const newRow = [dateStr, data.name, timeStr, "", locationName + " (" + coords + ")", "", photoUrl ? '=IMAGE("' + photoUrl + '")' : "No Photo", "", ""];
      sheet.appendRow(newRow);
      sheet.setRowHeight(sheet.getLastRow(), rowHeight);
      return ContentService.createTextOutput("Success Masuk").setMimeType(ContentService.MimeType.TEXT);
    } 
    else if (type === 'Pulang') {
      const values = sheet.getDataRange().getValues();
      let rowIdx = -1;
      
      for (let i = values.length - 1; i >= 1; i--) {
        let rowDateStr = "";
        const rowDateRaw = values[i][0];
        if (rowDateRaw instanceof Date) {
          rowDateStr = Utilities.formatDate(rowDateRaw, tz, "yyyy-MM-dd");
        } else if (rowDateRaw) {
          rowDateStr = rowDateRaw.toString().substring(0, 10);
        }
        
        const rowName = values[i][1] ? values[i][1].toString().toLowerCase().replace(/\s+/g, ' ').trim() : "";
        const rowTimeOut = values[i][3];

        if (rowDateStr === dateStr && rowName === nameInput && (rowTimeOut === "" || rowTimeOut === null || rowTimeOut === undefined)) {
          rowIdx = i + 1;
          break;
        }
      }
      
      if (rowIdx !== -1) {
        const timeInStr = values[rowIdx-1][2]; 
        let duration = "0 jam 0 menit";
        
        if (timeInStr) {
          try {
            const timeIn = new Date(dateStr + "T" + timeInStr);
            const timeOut = new Date(dateStr + "T" + timeStr);
            if (!isNaN(timeIn.getTime()) && !isNaN(timeOut.getTime())) {
              const diffMs = timeOut - timeIn;
              const hrs = Math.floor(diffMs / 3600000);
              const mins = Math.floor((diffMs % 3600000) / 60000);
              duration = `${hrs} jam ${mins} menit`;
            }
          } catch (e) {}
        }
        
        sheet.getRange(rowIdx, 4).setValue(timeStr); 
        sheet.getRange(rowIdx, 6).setValue(locationName + " (" + coords + ")"); 
        if (photoUrl) {
          sheet.getRange(rowIdx, 8).setValue('=IMAGE("' + photoUrl + '")'); 
        }
        sheet.getRange(rowIdx, 9).setValue(duration); 
        
        return ContentService.createTextOutput("Success Pulang").setMimeType(ContentService.MimeType.TEXT);
      } else {
        const newRow = [dateStr, data.name, "", timeStr, "", locationName + " (" + coords + ")", "", photoUrl ? '=IMAGE("' + photoUrl + '")' : "", ""];
        sheet.appendRow(newRow);
        sheet.setRowHeight(sheet.getLastRow(), rowHeight);
        return ContentService.createTextOutput("New Row Added").setMimeType(ContentService.MimeType.TEXT);
      }
    }
  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet() {
  return ContentService.createTextOutput("Script Aktif: Versi Perbaikan Final").setMimeType(ContentService.MimeType.TEXT);
}
