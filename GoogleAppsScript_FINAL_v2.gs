/**
 * GOOGLE APPS SCRIPT - ANNUAL LEAVE
 *
 * Search supports:
 *   - Vietnamese name WITH accents
 *   - Vietnamese name WITHOUT accents
 *   - Upper/lower case
 *
 * Examples:
 *   "BÙI THỊ TRÚC PHƯƠNG"
 *   "BUI THI TRUC PHUONG"
 * both return the same employee.
 *
 * Sheet:
 *   2026    -> XƯỞNG BÁNH
 *   2026 PL -> XƯỞNG IN
 *
 * Data columns in LeaveData:
 *   A EmployeeName
 *   B PaidLeave
 *   C Left2025
 *   D Left2026
 *   E RemainingLeave
 *   F Factories
 *   G UpdatedAt
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
    const updatedAt = body.updatedAt || '';

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
        round3(x.paidLeave),
        round3(x.left2025),
        round3(x.left2026),
        round3(x.remainingLeave),
        normalizeFactory(x.factories),
        updatedAt
      ]);

      sh.getRange(2, 1, values.length, 7).setValues(values);
      sh.getRange(2, 2, values.length, 4).setNumberFormat('0.000');
    }

    return jsonResponse({
      ok:true,
      rows:employees.length,
      updatedAt:updatedAt
    });

  } catch (err) {
    return jsonResponse({
      ok:false,
      error:String(err)
    });
  }
}

function doGet(e) {
  try {
    const name = ((e && e.parameter && e.parameter.name) || '').trim();

    if (!name) {
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
    const query = normalizeVietnamese(name);

    // Exact normalized match first.
    for (let i = 1; i < values.length; i++) {
      const dbName = String(values[i][0] || '').trim();

      if (normalizeVietnamese(dbName) === query) {
        return buildEmployeeResponse(values[i]);
      }
    }

    // Fallback: allow the query to be contained in the full normalized name.
    // This helps when a user types only part of the name.
    for (let i = 1; i < values.length; i++) {
      const dbName = String(values[i][0] || '').trim();
      const normalizedDbName = normalizeVietnamese(dbName);

      if (query && normalizedDbName.includes(query)) {
        return buildEmployeeResponse(values[i]);
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

function buildEmployeeResponse(row) {
  return jsonResponse({
    ok:true,
    employee:{
      employeeName: row[0],
      paidLeave: round3(row[1]),
      left2025: round3(row[2]),
      left2026: round3(row[3]),
      remainingLeave: round3(row[4]),
      factories: normalizeFactory(row[5]),
      updatedAt: row[6]
    }
  });
}

/**
 * Remove Vietnamese accents and normalize case/spaces.
 * Uses Unicode NFD plus Vietnamese-specific đ/Đ handling.
 */
function normalizeVietnamese(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function normalizeFactory(value) {
  const s = String(value || '').trim();

  if (s === '2026' || s === 'XƯỞNG BÁNH') {
    return 'Bánh';
  }

  if (s === '2026 PL' || s === 'XƯỞNG IN') {
    return 'In';
  }

  if (
    s.includes('2026') &&
    s.includes('2026 PL')
  ) {
    return 'Bánh + In';
  }

  if (
    s.toUpperCase().includes('BÁNH') &&
    s.toUpperCase().includes('IN')
  ) {
    return 'Bánh + In';
  }

  return s
    .replace(/XƯỞNG/gi, '')
    .replace(/\b2026\s+PL\b/gi, 'In')
    .replace(/\b2026\b/gi, 'Bánh')
    .trim();
}

function round3(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }

  const n = Number(value);

  if (isNaN(n)) {
    return '';
  }

  return Math.round((n + Number.EPSILON) * 1000) / 1000;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
