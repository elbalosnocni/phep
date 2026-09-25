function doGet(e) {
  return doPost(e);
}
function doPost(e) {
  try {
    // 1. Kiem tra du lieu dau vao
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Khong co du lieu gui den"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 2. Giai ma du lieu JSON tu VBA gui sang
    var data = JSON.parse(e.postData.contents);
    
    // 3. Lay ID Spreadsheet truyen dong tu VBA
    var spreadsheetId = data.spreadsheet_id; 
    var sheetName = data.sheet_name || "DSCNV"; 
    
    if (!spreadsheetId) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Thieu thong tin spreadsheet_id"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. Mo Google Sheet nhan du lieu
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Khong tim thay sheet co ten: " + sheetName}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 5. Doc du lieu mang cac dong
    var allRows = data.values; 
    
    if (Array.isArray(allRows) && allRows.length > 0) {
      var lastRow = sheet.getLastRow();
      var startRow = lastRow + 1;
      
      var numRows = allRows.length;          // So luong dong du lieu (Ví dụ: 932)
      var numCols = allRows[0].length;       // DA SUA: Lay chinh xac so luong cot thuc te (16 cot)
      
      // Ghi hang loat xuong Google Sheet voi kich thuoc o chuan khop
      sheet.getRange(startRow, 1, numRows, numCols).setValues(allRows);
      
      return ContentService.createTextOutput(JSON.stringify({"status": "success", "message": "Da ghi thanh cong " + numRows + " dong vao Sheet!"}))
                           .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Du lieu 'values' trong hoac sai dinh dang"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    // Neu co loi can tro tren Google Sheet se duoc bao ve chính xác cho VBA qua day
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
