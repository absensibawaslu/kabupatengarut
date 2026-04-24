// AppsScript.js - FINAL FIXED VERSION

const SHEET_NAME = 'Sheet1';
const FOLDER_ID = ''; 

function calculateDuration(timeInStr, timeOutStr) {
  if (!timeInStr || !timeOutStr) return "0 jam 0 menit";
  const inMatch = timeInStr.toString().match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  const outMatch = timeOutStr.toString().match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  
  if (inMatch && outMatch) {
    const inMs = (parseInt(inMatch[1], 10) * 3600 + parseInt(inMatch[2], 10) * 60 + (parseInt(inMatch[3] || 0, 10))) * 1000;
    const outMs = (parseInt(outMatch[1], 10) * 3600 + parseInt(outMatch[2], 10) * 60 + (parseInt(outMatch[3] || 0, 10))) * 1000;
    
    let diffMs = outMs - inMs;
    if (diffMs < 0) diffMs += 24 * 3600 * 1000;
    
    const hrs = Math.floor(diffMs / 3600000);
    const mins = Math.floor((diffMs % 3600000) / 60000);
    return `${hrs} jam ${mins} menit`;
  }
  return "0 jam 0 menit";
}

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

    const lastRowIndex = sheet.getLastRow();
    let isNewDate = false;

    if (lastRowIndex === 0) {
      isNewDate = true;
    } else {
      const lastRowDateRaw = sheet.getRange(lastRowIndex, 1).getValue();
      let lastRowDateStr = "";
      if (lastRowDateRaw instanceof Date) {
        lastRowDateStr = Utilities.formatDate(lastRowDateRaw, tz, "yyyy-MM-dd");
      } else if (lastRowDateRaw) {
        lastRowDateStr = lastRowDateRaw.toString().substring(0, 10);
      }
      
      if (lastRowDateStr !== dateStr && lastRowDateStr.toLowerCase() !== "tanggal") {
        isNewDate = true;
      }
    }

    if (isNewDate) {
      sheet.appendRow(["Tanggal", "Nama", "Jam Masuk", "Jam Pulang", "Lokasi Masuk", "Lokasi Pulang", "Foto Masuk", "Foto Pulang", "Durasi Bekerja"]);
      const newHeaderRow = sheet.getLastRow();
      sheet.setRowHeight(newHeaderRow, 30);
      sheet.getRange(newHeaderRow, 1, 1, 9).setBackground("#e11d48").setFontColor("white").setFontWeight("bold");
    }
    
    const rowHeight = 80;

    const values = sheet.getDataRange().getValues();
    const displayValues = sheet.getDataRange().getDisplayValues();
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

      if (rowDateStr === dateStr && rowName === nameInput) {
        rowIdx = i + 1;
        break;
      }
    }

    if (type === 'Masuk') {
      if (rowIdx !== -1) {
        sheet.getRange(rowIdx, 3).setValue(timeStr);
        sheet.getRange(rowIdx, 5).setValue(locationName + " (" + coords + ")");
        if (photoUrl) {
          sheet.getRange(rowIdx, 7).setValue('=IMAGE("' + photoUrl + '")');
        } else {
          sheet.getRange(rowIdx, 7).setValue("No Photo");
        }
        
        const timeOutStr = displayValues[rowIdx - 1][3];
        if (timeOutStr) {
          const duration = calculateDuration(timeStr, timeOutStr);
          sheet.getRange(rowIdx, 9).setValue(duration);
        }
        return ContentService.createTextOutput("Success Update Masuk").setMimeType(ContentService.MimeType.TEXT);
      } else {
        const newRow = [dateStr, data.name, timeStr, "", locationName + " (" + coords + ")", "", photoUrl ? '=IMAGE("' + photoUrl + '")' : "No Photo", "", ""];
        sheet.appendRow(newRow);
        sheet.setRowHeight(sheet.getLastRow(), rowHeight);
        return ContentService.createTextOutput("Success Masuk").setMimeType(ContentService.MimeType.TEXT);
      }
    } 
    else if (type === 'Pulang') {
      if (rowIdx !== -1) {
        const timeInStr = displayValues[rowIdx - 1][2];
        const duration = calculateDuration(timeInStr, timeStr);
        
        sheet.getRange(rowIdx, 4).setValue(timeStr);
        sheet.getRange(rowIdx, 6).setValue(locationName + " (" + coords + ")");
        if (photoUrl) {
          sheet.getRange(rowIdx, 8).setValue('=IMAGE("' + photoUrl + '")');
        } else {
          sheet.getRange(rowIdx, 8).setValue("No Photo");
        }
        sheet.getRange(rowIdx, 9).setValue(duration);
        
        return ContentService.createTextOutput("Success Pulang").setMimeType(ContentService.MimeType.TEXT);
      } else {
        const newRow = [dateStr, data.name, "", timeStr, "", locationName + " (" + coords + ")", "", photoUrl ? '=IMAGE("' + photoUrl + '")' : "No Photo", "0 jam 0 menit"];
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
