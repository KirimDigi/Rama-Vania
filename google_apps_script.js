// CONFIGURATION: Set your Spreadsheet ID here
const SPREADSHEET_ID = "1E3GbSP-H0b1viEBxDxT3IJA8vCDLX1Pb62FU0oum5Bo";
const SHEET_NAME = "Sheet1";

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] || row[2]) { // Ensure name or wish is present
        rows.push({
          timestamp: row[0],
          nama: row[1],
          ucapan: row[2],
          kehadiran: row[3],
          jumlah: row[4]
        });
      }
    }
    
    // Reverse to show the latest wish at the top
    rows.reverse();
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    let name = "";
    let message = "";
    let attendance = "";
    let guestCount = "";
    
    // Parse JSON body if present
    if (e.postData && e.postData.contents) {
      try {
        const body = JSON.parse(e.postData.contents);
        name = body.nama || body.name || "";
        message = body.ucapan || body.message || "";
        attendance = body.kehadiran || body.attendance || "";
        guestCount = body.jumlah || body.guestCount || "0";
      } catch (err) {
        // Fallback to URL-encoded parameters if parsing fails
      }
    }
    
    // Fallback to parameters
    if (!name) {
      name = e.parameter.nama || "";
      message = e.parameter.ucapan || "";
      attendance = e.parameter.kehadiran || "";
      guestCount = e.parameter.jumlah || "0";
    }
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    // Create headers if the sheet is completely empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["timestamp", "nama tamu", "ucapan", "konfirmasi kehadiran", "jumlah tamu"]);
    }
    
    const timestamp = new Date();
    sheet.appendRow([timestamp, name, message, attendance, guestCount]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
