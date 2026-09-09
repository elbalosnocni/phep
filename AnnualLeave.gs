/**
 * ============================================================
 * GOOGLE APPS SCRIPT - ANNUAL LEAVE LOOKUP v4
 * ============================================================
 *
 * LeaveData columns:
 * A EmployeeName
 * B PaidLeave
 * C Left2025
 * D Left2026
 * E RemainingLeave
 * F Factories
 * G Department
 * H Section
 * I CitizenID
 * J UpdatedAt
 * K:V Jan-Dec
 *
 * Source:
 * - Npn2023.xlsb: sheets 2026 and 2026 PL
 * - DSCNV-2023.xlsb: sheet DSCNV
 *   H  = Citizen ID
 *   AE = Department (col 31)
 *   AF = Section (col 32)
 *
 * Timezone: Asia/Ho_Chi_Minh
 * ============================================================
 */

const SHEET_NAME = 'LeaveData';
const TIMEZONE = 'Asia/Ho_Chi_Minh';

const MONTHS = [
  'jan','feb','mar','apr','may','jun',
  'jul','aug','sep','oct','nov','dec'
];

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TIMEZONE);

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  sh.clear();

  sh.getRange(1, 1, 1, 22).setValues([[
    'EmployeeName',
    'PaidLeave',
    'Left2025',
    'Left2026',
    'RemainingLeave',
    'Factories',
    'Department',
    'Section',
    'CitizenID',
    'UpdatedAt',
    'Jan','Feb','Mar','Apr','May','Jun',
    'Jul','Aug','Sep','Oct','Nov','Dec'
  ]]);

  sh.setFrozenRows(1);
  sh.getRange('B:E').setNumberFormat('0.000');
  sh.getRange('I:I').setNumberFormat('@');
  sh.getRange('J:J').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange('K:V').setNumberFormat('0.000');
  sh.autoResizeColumns(1, 22);
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
    ss.setSpreadsheetTimeZone(TIMEZONE);

    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      setupSheet();
      sheet = ss.getSheetByName(SHEET_NAME);
    }

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 22).clearContent();
    }

    const updatedDate = parseVietnamDateTime(updatedAt);

    if (employees.length > 0) {
      const values = employees.map(function(x) {
        return [
          x.employeeName || '',
          round3(x.paidLeave),
          round3(x.left2025),
          round3(x.left2026),
          round3(x.remainingLeave),
          normalizeFactory(x.factories),
          String(x.department || '').trim(),
          String(x.section || '').trim(),
          normalizeCitizenId(x.citizenId),
          updatedDate,
          round3(x.jan), round3(x.feb), round3(x.mar), round3(x.apr),
          round3(x.may), round3(x.jun), round3(x.jul), round3(x.aug),
          round3(x.sep), round3(x.oct), round3(x.nov), round3(x.dec)
        ];
      });

      sheet.getRange(2, 1, values.length, 22).setValues(values);
      sheet.getRange(2, 2, values.length, 4).setNumberFormat('0.000');
      sheet.getRange(2, 9, values.length, 1).setNumberFormat('@');
      sheet.getRange(2, 10, values.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
      sheet.getRange(2, 11, values.length, 12).setNumberFormat('0.000');
    }

    return jsonResponse({
      ok: true,
      rows: employees.length,
      updatedAt: formatVietnamDateTime(updatedDate)
    });

  } catch (err) {
    return jsonResponse({ok:false, error:String(err)});
  }
}

function doGet(e) {
  try {
    const rawCitizenId =
      ((e && e.parameter && (e.parameter.cccd || e.parameter.citizenId)) || '').trim();

    if (!rawCitizenId) {
      return jsonResponse({
        ok:true,
        service:'Annual Leave Lookup',
        version:'4.0',
        lookup:'citizenId'
      });
    }

    const query = normalizeCitizenId(rawCitizenId);
    if (!/^\d{12}$/.test(query)) {
      return jsonResponse({
        ok:false,
        error:'Số căn cước phải gồm đúng 12 chữ số.'
      });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) {
      return jsonResponse({ok:false, error:'LeaveData sheet not found'});
    }

    const values = sh.getDataRange().getValues();

    // I = CitizenID -> index 8
    for (let i = 1; i < values.length; i++) {
      const dbCitizenId = normalizeCitizenId(values[i][8]);
      if (dbCitizenId === query) {
        return buildEmployeeResponse(values[i]);
      }
    }

    return jsonResponse({
      ok:false,
      error:'Không tìm thấy nhân viên với số căn cước này.'
    });

  } catch (err) {
    return jsonResponse({ok:false, error:String(err)});
  }
}

function buildEmployeeResponse(row) {
  const employee = {
    employeeName: row[0],
    paidLeave: round3(row[1]),
    left2025: round3(row[2]),
    left2026: round3(row[3]),
    remainingLeave: round3(row[4]),
    factories: normalizeFactory(row[5]),
    department: String(row[6] || '').trim(),
    section: String(row[7] || '').trim(),
    citizenId: normalizeCitizenId(row[8]),
    updatedAt: formatVietnamDateTime(row[9]),
    monthlyLeave: {}
  };

  MONTHS.forEach(function(month, index) {
    employee.monthlyLeave[month] = round3(row[10 + index]);
  });

  return jsonResponse({ok:true, employee:employee});
}

function parseVietnamDateTime(value) {
  if (value === null || value === undefined || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (!isNaN(value.getTime())) return value;
  }

  const s = String(value).trim();
  const match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/);

  if (!match) throw new Error('Invalid UpdatedAt: ' + s);

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);

  if (month < 1 || month > 12 || day < 1 || day > 31 ||
      hour < 0 || hour > 23 || minute < 0 || minute > 59 ||
      second < 0 || second > 59) {
    throw new Error('Invalid UpdatedAt value: ' + s);
  }

  return new Date(year, month - 1, day, hour, minute, second);
}

function formatVietnamDateTime(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' &&
      !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TIMEZONE, 'HH:mm:ss dd/MM/yyyy');
  }

  const s = String(value).trim();
  const match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/);

  if (match) {
    return match[4].padStart(2, '0') + ':' + match[5] + ':' + match[6] + ' ' +
           match[1].padStart(2, '0') + '/' + match[2].padStart(2, '0') + '/' + match[3];
  }

  return s;
}

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

function normalizeCitizenId(value) {
  if (value === null || value === undefined || value === '') return '';

  let s = String(value).trim();
  s = s.replace(/[^\d]/g, '');

  // CCCD is 12 digits. Keep leading zeroes.
  if (s.length < 12 && /^\d+$/.test(s)) {
    s = s.padStart(12, '0');
  }

  return s;
}

function normalizeFactory(value) {
  const s = String(value || '').trim();

  if (s === '2026' || s === 'XƯỞNG BÁNH') return 'Bánh';
  if (s === '2026 PL' || s === 'XƯỞNG IN') return 'In';

  if (s.toUpperCase().includes('BÁNH') && s.toUpperCase().includes('IN')) {
    return 'Bánh + In';
  }

  return s
    .replace(/XƯỞNG/gi, '')
    .replace(/\b2026\s+PL\b/gi, 'In')
    .replace(/\b2026\b/gi, 'Bánh')
    .trim();
}

function round3(value) {
  if (value === '' || value === null || value === undefined) return '';

  const n = Number(value);
  if (Number.isNaN(n)) return '';

  return Math.round((n + Number.EPSILON) * 1000) / 1000;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
