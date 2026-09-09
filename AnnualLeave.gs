/**
 * ============================================================
 * GOOGLE APPS SCRIPT - ANNUAL LEAVE
 * ============================================================
 *
 * Sheet: LeaveData
 *
 * A  EmployeeName
 * B  PaidLeave
 * C  Left2025
 * D  Left2026
 * E  RemainingLeave
 * F  Factories
 * G  Department
 * H  Section
 * I  CitizenID
 * J  UpdatedAt
 * K  Jan
 * I  Feb
 * J  Mar
 * K  Apr
 * L  May
 * M  Jun
 * N  Jul
 * O  Aug
 * P  Sep
 * Q  Oct
 * R  Nov
 * V  Dec
 *
 * Timezone:
 * Asia/Ho_Chi_Minh
 * ============================================================
 */

const SHEET_NAME = 'LeaveData';
const TIMEZONE = 'Asia/Ho_Chi_Minh';

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
 * ============================================================
 * SETUP SHEET
 * ============================================================
 */
function setupSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Luôn dùng giờ Việt Nam
  ss.setSpreadsheetTimeZone(TIMEZONE);

  let sh = ss.getSheetByName(SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }

  sh.clear();

  // Header A:V
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

  // Format số
  sh.getRange('B:E').setNumberFormat('0.000');
  sh.getRange('K:V').setNumberFormat('0.000');
  sh.getRange('I:I').setNumberFormat('@');

  // ==========================================================
  // QUAN TRỌNG:
  // Cột G phải là DATE thật, KHÔNG phải TEXT
  // ==========================================================
  sh.getRange('G:G')
    .setNumberFormat('dd/MM/yyyy HH:mm:ss');
}


/**
 * ============================================================
 * POST
 * VBA -> Google Apps Script -> Google Sheet
 * ============================================================
 */
function doPost(e) {

  try {

    // --------------------------------------------------------
    // Kiểm tra request
    // --------------------------------------------------------
    if (!e || !e.postData || !e.postData.contents) {

      return jsonResponse({
        ok: false,
        error: 'Empty request'
      });
    }


    // --------------------------------------------------------
    // Parse JSON
    // --------------------------------------------------------
    const body = JSON.parse(e.postData.contents);

    const employees = body.employees || [];

    const updatedAt = body.updatedAt || '';


    // --------------------------------------------------------
    // Spreadsheet
    // --------------------------------------------------------
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    ss.setSpreadsheetTimeZone(TIMEZONE);

    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {

      setupSheet();

      sheet = ss.getSheetByName(SHEET_NAME);
    }


    // --------------------------------------------------------
    // Xóa dữ liệu cũ
    // --------------------------------------------------------
    if (sheet.getLastRow() > 1) {

      sheet
        .getRange(
          2,
          1,
          sheet.getLastRow() - 1,
          19
        )
        .clearContent();
    }


    // --------------------------------------------------------
    // UPDATED AT
    //
    // VBA gửi:
    //
    // dd/mm/yyyy HH:mm:ss
    //
    // Ví dụ:
    // 04/09/2026 12:00:24
    //
    // Không dùng:
    // new Date("04/09/2026 12:00:24")
    //
    // vì JavaScript có thể hiểu sai dd/mm.
    // --------------------------------------------------------
    const updatedDate = parseVietnamDateTime(updatedAt);


    // --------------------------------------------------------
    // Tạo data
    // --------------------------------------------------------
    if (employees.length > 0) {

      const values = employees.map(function (x) {

        return [
          // A
          x.employeeName || '',
          // B:E
          round3(x.paidLeave),
          round3(x.left2025),
          round3(x.left2026),
          round3(x.remainingLeave),
          // F
          normalizeFactory(x.factories),
          // G:I
          x.department || '',
          x.section || '',
          String(x.citizenId || ''),
          // J
          updatedDate,
          // K:V
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
        ];

      });


      // ------------------------------------------------------
      // Ghi dữ liệu
      // ------------------------------------------------------
      sheet
        .getRange(
          2,
          1,
          values.length,
          22
        )
        .setValues(values);


      // ------------------------------------------------------
      // Format B:E
      // ------------------------------------------------------
      sheet
        .getRange(
          2,
          2,
          values.length,
          4
        )
        .setNumberFormat('0.000');


      // ------------------------------------------------------
      // Format H:S
      // ------------------------------------------------------
      sheet
        .getRange(
          2,
          8,
          values.length,
          12
        )
        .setNumberFormat('0.000');


      // ------------------------------------------------------
      // FORMAT G - UPDATEDAT
      //
      // Đây là phần quan trọng nhất.
      // ------------------------------------------------------
      sheet
        .getRange(
          2,
          7,
          values.length,
          1
        )
        .setNumberFormat('dd/MM/yyyy HH:mm:ss');

    }


    // --------------------------------------------------------
    // Response
    // --------------------------------------------------------
    return jsonResponse({

      ok: true,

      rows: employees.length,

      updatedAt: formatVietnamDateTime(updatedDate)

    });


  } catch (err) {

    return jsonResponse({

      ok: false,

      error: String(err)

    });

  }

}


/**
 * ============================================================
 * GET
 * Tra cứu nhân viên
 * ============================================================
 */
function doGet(e) {

  try {

    const name =
      ((e &&
        e.parameter &&
        e.parameter.name) || '')
        .trim();


    // --------------------------------------------------------
    // Health check
    // --------------------------------------------------------
    if (!name) {

      return jsonResponse({

        ok: true,

        service: 'Annual Leave Lookup',

        version: '3.0'

      });

    }


    // --------------------------------------------------------
    // Spreadsheet
    // --------------------------------------------------------
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


    // --------------------------------------------------------
    // Get data
    // --------------------------------------------------------
    const values =
      sh.getDataRange().getValues();


    const query =
      normalizeVietnamese(name);


    // --------------------------------------------------------
    // Exact match
    // --------------------------------------------------------
    for (
      let i = 1;
      i < values.length;
      i++
    ) {

      const dbName =
        String(values[i][0] || '')
          .trim();


      if (
        normalizeVietnamese(dbName)
        === query
      ) {

        return buildEmployeeResponse(
          values[i]
        );

      }

    }


    // --------------------------------------------------------
    // Partial match
    // --------------------------------------------------------
    for (
      let i = 1;
      i < values.length;
      i++
    ) {

      const dbName =
        String(values[i][0] || '')
          .trim();


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


    // --------------------------------------------------------
    // Not found
    // --------------------------------------------------------
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
 * ============================================================
 * BUILD EMPLOYEE RESPONSE
 * ============================================================
 */
function buildEmployeeResponse(row) {

  const employee = {

    employeeName: row[0],

    department: row[6] || '',
    section: row[7] || '',
    citizenId: String(row[8] || ''),

    paidLeave: round3(row[1]),

    left2025: round3(row[2]),

    left2026: round3(row[3]),

    remainingLeave: round3(row[4]),

    factories: normalizeFactory(row[5]),

    // G = UpdatedAt
    updatedAt: formatVietnamDateTime(row[9]),

    monthlyLeave: {}

  };


  // Jan-Dec
  MONTHS.forEach(function (month, index) {

    employee.monthlyLeave[month] =
      round3(row[10 + index]);

  });


  return jsonResponse({

    ok: true,

    employee: employee

  });

}


/**
 * ============================================================
 * PARSE NGÀY GIỜ VIỆT NAM
 * ============================================================
 *
 * Input từ VBA:
 *
 * 04/09/2026 12:00:24
 *
 * Output:
 *
 * JavaScript Date object
 *
 * ============================================================
 */
function parseVietnamDateTime(value) {

  // Không có dữ liệu
  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {

    return '';

  }


  // ----------------------------------------------------------
  // Nếu đã là Date
  // ----------------------------------------------------------
  if (
    Object.prototype.toString.call(value)
      === '[object Date]'
  ) {

    if (!isNaN(value.getTime())) {

      return value;

    }

  }


  const s =
    String(value).trim();


  // ----------------------------------------------------------
  // VBA:
  //
  // dd/mm/yyyy HH:mm:ss
  //
  // Ví dụ:
  // 04/09/2026 12:00:24
  // ----------------------------------------------------------
  const match =
    s.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/
    );


  if (!match) {

    throw new Error(
      'Invalid UpdatedAt: ' + s
    );

  }


  const day =
    Number(match[1]);

  const month =
    Number(match[2]);

  const year =
    Number(match[3]);

  const hour =
    Number(match[4]);

  const minute =
    Number(match[5]);

  const second =
    Number(match[6]);


  // ----------------------------------------------------------
  // Validate
  // ----------------------------------------------------------
  if (
    month < 1 ||
    month > 12 ||

    day < 1 ||
    day > 31 ||

    hour < 0 ||
    hour > 23 ||

    minute < 0 ||
    minute > 59 ||

    second < 0 ||
    second > 59
  ) {

    throw new Error(
      'Invalid UpdatedAt value: ' + s
    );

  }


  // ----------------------------------------------------------
  // Date thật
  //
  // Không parse bằng new Date(string)
  // ----------------------------------------------------------
  return new Date(
    year,
    month - 1,
    day,
    hour,
    minute,
    second
  );

}


/**
 * ============================================================
 * FORMAT DATE
 * ============================================================
 *
 * Output:
 *
 * HH:mm:ss dd/MM/yyyy
 *
 * Ví dụ:
 *
 * 12:00:24 04/09/2026
 * ============================================================
 */
function formatVietnamDateTime(value) {

  if (!value) {

    return '';

  }


  // Nếu là Date thật
  if (
    Object.prototype.toString.call(value)
      === '[object Date]' &&
    !isNaN(value.getTime())
  ) {

    return Utilities.formatDate(
      value,
      TIMEZONE,
      'HH:mm:ss dd/MM/yyyy'
    );

  }


  // Nếu chẳng may là string
  const s =
    String(value).trim();


  // dd/MM/yyyy HH:mm:ss
  const match =
    s.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/
    );


  if (match) {

    return (
      match[4].padStart(2, '0') +
      ':' +
      match[5] +
      ':' +
      match[6] +
      ' ' +
      match[1].padStart(2, '0') +
      '/' +
      match[2].padStart(2, '0') +
      '/' +
      match[3]
    );

  }


  return s;

}


/**
 * ============================================================
 * NORMALIZE VIETNAMESE
 * ============================================================
 */
function normalizeVietnamese(value) {

  return String(value || '')

    .normalize('NFD')

    .replace(
      /[\u0300-\u036f]/g,
      ''
    )

    .replace(/đ/g, 'd')

    .replace(/Đ/g, 'D')

    .replace(/\s+/g, ' ')

    .trim()

    .toLowerCase();

}


/**
 * ============================================================
 * NORMALIZE FACTORY
 * ============================================================
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

    .replace(
      /XƯỞNG/gi,
      ''
    )

    .replace(
      /\b2026\s+PL\b/gi,
      'In'
    )

    .replace(
      /\b2026\b/gi,
      'Bánh'
    )

    .trim();

}


/**
 * ============================================================
 * ROUND 3
 * ============================================================
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
 * ============================================================
 * JSON RESPONSE
 * ============================================================
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