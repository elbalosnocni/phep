function doGet(e) {
  return doPost(e);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Khong co du lieu gui den"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = JSON.parse(e.postData.contents);
    var spreadsheetId = data.spreadsheet_id; 
    var sheetName = data.sheet_name || "DSCNV"; 
    
    if (!spreadsheetId) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Thieu thong tin spreadsheet_id"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Khong tim thay sheet co ten: " + sheetName}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    var allRows = data.values; 
    
    if (Array.isArray(allRows) && allRows.length > 0) {
      var lastRow = sheet.getLastRow();
      
      var existingData = [];
      var cccdMap = {}; 
      
      // Quét tìm CCCD cũ bắt đầu từ dòng số 2 (bỏ qua dòng tiêu đề 1)
      if (lastRow > 1) {
        // Lấy dữ liệu cột Số Căn Cước trên Google Sheet (Vị trí cột số 6 - Cột F trong danh sách ghi xuống)
        existingData = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
        for (var r = 0; r < existingData.length; r++) {
          var cccdKey = existingData[r][0]; 
          if (cccdKey) {
            cccdKey = String(cccdKey).replace(/'/g, "").trim();
            cccdMap[cccdKey] = r + 2; // Lưu số dòng thực tế trên Google Sheets
          }
        }
      }
      
      var countInsert = 0;
      var countUpdate = 0;
      
      for (var i = 0; i < allRows.length; i++) {
        var rowData = allRows[i];
        
        // Cột H trong chuỗi gửi từ VBA nằm ở vị trí số 5 (tính từ 0)
        var rawCCCD = String(rowData[5]).replace(/'/g, "").trim();
        
        // ĐÃ SỬA: Ép kiểu ngày chuẩn xác theo vị trí mảng JSON gửi sang (Index 1 là cột D, Index 2 là cột E)
        rowData[1] = parseDateString(rowData[1]); // Ngày sinh Nam
        rowData[2] = parseDateString(rowData[2]); // Ngày sinh Nữ
        
        if (cccdMap.hasOwnProperty(rawCCCD)) {
          // TRÙNG CĂN CƯỚC -> CẬP NHẬT LẠI THÔNG TIN DÒNG CŨ
          var targetRow = cccdMap[rawCCCD];
          sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
          countUpdate++;
        } else {
          // KHÔNG TRÙNG -> CHÈN DÒNG MỚI XUỐNG DƯỚI CÙNG
          sheet.appendRow(rowData);
          countInsert++;
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        "status": "success", 
        "message": "Da xu ly xong! Them moi: " + countInsert + " dong, Cap nhat trung: " + countUpdate + " dong."
      })).setMimeType(ContentService.MimeType.JSON);
      
    } else {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Du lieu 'values' trong hoac sai dinh dang"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function parseDateString(dateStr) {
  if (!dateStr || String(dateStr).trim() === "") return "";
  var cleanStr = String(dateStr).trim();
  
  var parts = cleanStr.split("/");
  if (parts.length === 3) {
    var day = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1; 
    var year = parseInt(parts[2], 10);
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
      return new Date(year, month, day);
    }
  }
  return cleanStr; 
}
