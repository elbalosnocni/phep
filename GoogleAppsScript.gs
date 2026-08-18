/**
 * GOOGLE APPS SCRIPT - ANNUAL LEAVE
 *
 * VERSION 2.0
 * Added monthly leave Jan-Dec.
 *
 * LeaveData:
 *
 * A EmployeeName
 * B PaidLeave
 * C Left2025
 * D Left2026
 * E RemainingLeave
 * F Factories
 * G UpdatedAt
 * H Jan
 * I Feb
 * J Mar
 * K Apr
 * L May
 * M Jun
 * N Jul
 * O Aug
 * P Sep
 * Q Oct
 * R Nov
 * S Dec
 */

const SHEET_NAME = 'LeaveData';

const MONTHS = [
  'jan',
  'feb',
  'mar',
  'apr',
  'may',
  'jun',
  'jul',
  'aug',
  'sep',
  'oct',
  'nov',
  'dec'
];

/**
 * Create / reset LeaveData
 */
function setupSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName(SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }

  sh.clear();

  sh.getRange(1, 1, 1, 19).setValues([[
    'EmployeeName',
    'PaidLeave',
    'Left2025',
    'Left2026',
    'RemainingLeave',
    'Factories',
    'UpdatedAt',
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec'
  ]]);

  sh.setFrozenRows(1);

}

/**
 * POST from VBA
 */
function doPost(e) {

  try {

    if (!e ||
        !e.postData ||
        !e.postData.contents) {

      return jsonResponse({
        ok: false,
        error: 'Empty request'
      });

    }

    const body =
      JSON.parse(e.postData.contents);

    const employees =
      body.employees || [];

    const updatedAt =
      body.updatedAt || '';

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    let sh =
      ss.getSheetByName(SHEET_NAME);

    if (!sh) {

      setupSheet();

      sh =
        ss.getSheetByName(SHEET_NAME);
    }

    /*
     * Clear old data.
     * Keep header.
     */
    if (sh.getLastRow() > 1) {

      sh.getRange(
        2,
        1,
        sh.getLastRow() - 1,
        19
      ).clearContent();

    }

    /*
     * Build rows
     */
    if (employees.length > 0) {

      const values =
        employees.map(x => [

          x.employeeName || '',

          round3(x.paidLeave),

          round3(x.left2025),

          round3(x.left2026),

          round3(x.remainingLeave),

          normalizeFactory(x.factories),

          updatedAt,

          round3(x.jan),

          round3(x.feb),

          round3(x.mar),

          round3(x.apr),

          round3(x.may),

          round3(x.jun),

          round3(x.jul),

          round3(x.aug),

          round3(x.sep),

          round3(x.oct),

          round3(x.nov),

          round3(x.dec)

        ]);

      sh.getRange(
        2,
        1,
        values.length,
        19
      ).setValues(values);

      /*
       * Number format
       *
       * B:E  = summary
       * H:S  = monthly
       */
      sh.getRange(
        2,
        2,
        values.length,
        4
      ).setNumberFormat('0.000');

      sh.getRange(
        2,
        8,
        values.length,
        12
      ).setNumberFormat('0.000');

    }

    return jsonResponse({

      ok: true,

      rows: employees.length,

      updatedAt: updatedAt

    });

  } catch (err) {

    return jsonResponse({

      ok: false,

      error: String(err)

    });

  }

}

/**
 * GET employee
 */
function doGet(e) {

  try {

    const name =
      ((e &&
        e.parameter &&
        e.parameter.name) || '')
        .trim();

    /*
     * Health check
     */
    if (!name) {

      return jsonResponse({

        ok: true,

        service: 'Annual Leave Lookup',

        version: '2.0'

      });

    }

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const sh =
      ss.getSheetByName(SHEET_NAME);

    if (!sh) {

      return jsonResponse({

        ok: false,

        error: 'LeaveData sheet not found'

      });

    }

    const values =
      sh.getDataRange().getValues();

    const query =
      normalizeVietnamese(name);

    /*
     * Exact normalized match
     */
    for (
      let i = 1;
      i < values.length;
      i++
    ) {

      const dbName =
        String(values[i][0] || '').trim();

      if (
        normalizeVietnamese(dbName)
        === query
      ) {

        return buildEmployeeResponse(
          values[i]
        );

      }

    }

    /*
     * Partial match
     */
    for (
      let i = 1;
      i < values.length;
      i++
    ) {

      const dbName =
        String(values[i][0] || '').trim();

      const normalizedDbName =
        normalizeVietnamese(dbName);

      if (
        query &&
        normalizedDbName.includes(query)
      ) {

        return buildEmployeeResponse(
          values[i]
        );

      }

    }

    return jsonResponse({

      ok: false,

      error: 'Employee not found'

    });

  } catch (err) {

    return jsonResponse({

      ok: false,

      error: String(err)

    });

  }

}

/**
 * Build employee response
 *
 * Sheet:
 * A = 0
 * B = 1
 * ...
 * G = 6
 * H:S = 7:18
 */
function buildEmployeeResponse(row) {

  const employee = {

    employeeName: row[0],

    paidLeave: round3(row[1]),

    left2025: round3(row[2]),

    left2026: round3(row[3]),

    remainingLeave: round3(row[4]),

    factories: normalizeFactory(row[5]),

    updatedAt: row[6],

    monthlyLeave: {}

  };

  /*
   * H:S
   */
  MONTHS.forEach(function(month, index) {

    employee.monthlyLeave[month] =
      round3(row[7 + index]);

  });

  return jsonResponse({

    ok: true,

    employee: employee

  });

}

/**
 * Remove Vietnamese accents
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

/**
 * Normalize factory
 */
function normalizeFactory(value) {

  const s =
    String(value || '').trim();

  if (
    s === '2026' ||
    s === 'XƯỞNG BÁNH'
  ) {

    return 'Bánh';

  }

  if (
    s === '2026 PL' ||
    s === 'XƯỞNG IN'
  ) {

    return 'In';

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

/**
 * Round 3 decimals
 */
function round3(value) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {

    return '';

  }

  const n =
    Number(value);

  if (Number.isNaN(n)) {

    return '';

  }

  return Math.round(
    (n + Number.EPSILON) * 1000
  ) / 1000;

}

/**
 * JSON response
 */
function jsonResponse(obj) {

  return ContentService

    .createTextOutput(
      JSON.stringify(obj)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}
