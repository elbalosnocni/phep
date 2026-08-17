/**
 * GOOGLE APPS SCRIPT - ANNUAL LEAVE
 *
 * Receives aggregated data from Excel and stores it in Google Sheets.
 *
 * Columns:
 * A EmployeeName
 * B PaidLeave       = SUM column Y
 * C Left2025        = column Z
 * D Left2026        = column AA
 * E RemainingLeave  = column AD (Gross)
 * F Factories
 * G UpdatedAt
 */

const SHEET_NAME = 'LeaveData';

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);

  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  sh.clear();

  sh.getRange(1, 1, 1, 7).setValues([[
    'EmployeeName',
    'PaidLeave',
    'Left2025',
    'Left2026',
    'RemainingLeave',
    'Factories',
    'UpdatedAt'
  ]]);

  sh.setFrozenRows(1);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ok:false, error:'Empty request'});
    }

    const body = JSON.parse(e.postData.contents);
    const employees = body.employees || [];
    const updatedAt = body.updatedAt || new Date().toISOString();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) {
      setupSheet();
      sh = ss.getSheetByName(SHEET_NAME);
    }

    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
    }

    if (employees.length > 0) {
      const values = employees.map(x => [
        x.employeeName || '',
        x.paidLeave ?? '',
        x.left2025 ?? '',
        x.left2026 ?? '',
        x.remainingLeave ?? '',
        x.factories || '',
        updatedAt
      ]);

      sh.getRange(2, 1, values.length, 7).setValues(values);
    }

    return jsonResponse({
      ok:true,
      rows:employees.length,
      updatedAt:updatedAt
    });

  } catch (err) {
    return jsonResponse({ok:false, error:String(err)});
  }
}

function doGet(e) {
  try {
    const employeeName =
      ((e && e.parameter && e.parameter.name) || '').trim();

    if (!employeeName) {
      return jsonResponse({
        ok:true,
        service:'Annual Leave Lookup'
      });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) {
      return jsonResponse({
        ok:false,
        error:'LeaveData sheet not found'
      });
    }

    const values = sh.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {

      if (
        String(values[i][0]).trim().toLowerCase() ===
        employeeName.toLowerCase()
      ) {
        return jsonResponse({
          ok:true,
          employee:{
            employeeName: values[i][0],
            paidLeave: values[i][1],
            left2025: values[i][2],
            left2026: values[i][3],
            remainingLeave: values[i][4],
            factories: values[i][5],
            updatedAt: values[i][6]
          }
        });
      }
    }

    return jsonResponse({
      ok:false,
      error:'Employee not found'
    });

  } catch(err) {
    return jsonResponse({
      ok:false,
      error:String(err)
    });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
