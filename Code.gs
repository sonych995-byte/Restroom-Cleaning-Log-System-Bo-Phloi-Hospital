const SHEET_NAME = 'Sheet1';
const MASTER_SHEET_NAME = 'Sheet2';
const SALT_KEY = 'Hosp!tal_Salt_99';

const ADMIN_EMAIL_HASH = 'af394ecb78108d31c2da7665f9c29904eb34dbd23d252607a9390e355974ea0c';
const ADMIN_PASS_HASH = 'f5e78d17d5d6040341cd4f6ed4b972516eb879cc3ebb54efb8fa1338023cb461';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบตรวจสอบความสะอาดห้องน้ำ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function hashWithSalt(text) {
  const signature = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    text + SALT_KEY,
    Utilities.Charset.UTF_8
  );

  return signature.map(function(byte) {
    return ('0' + ((byte < 0) ? 256 + byte : byte).toString(16)).slice(-2);
  }).join('');
}

function verifyLogin(email, password) {
  if (!email && !password) {
    return { success: true, role: 'guest' };
  }

  const hashedEmail = hashWithSalt(email);
  const hashedPassword = hashWithSalt(password);

  if (hashedEmail === ADMIN_EMAIL_HASH && hashedPassword === ADMIN_PASS_HASH) {
    return { success: true, role: 'admin' };
  }

  return {
    success: false,
    error: 'อีเมลหรือรหัสผ่านไม่ถูกต้อง'
  };
}

/*
 * Sheet1 โครงสร้างเดิม:
 * A Timestamp
 * B Location
 * C Worker
 * D Inspector (คงไว้เพื่อไม่ให้ข้อมูลเก่าเลื่อนคอลัมน์ แต่ระบบใหม่จะบันทึกเป็นค่าว่าง)
 * E Floor_Action
 * F Floor_State
 * ...
 * V Tissue_State
 * W Status
 *
 * ผู้ปฏิบัติงานจะส่งเฉพาะ *_Action
 * ส่วน *_State จะเว้นว่างจนกว่า Admin จะกด "ตรวจสอบ"
 */
function saveRecord(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('ไม่พบ Sheet1');
  }

  if (!data || !data.location || !data.worker) {
    throw new Error('ข้อมูลผู้ปฏิบัติงานหรือสถานที่ไม่ครบถ้วน');
  }

  const timestamp = new Date();

    const rowData = [
    timestamp,
    data.location,
    data.worker,
    'ฐิติชญาน์ กมลปิตุลารัตน์', // Inspector: ผู้ตรวจสอบประจำระบบ
    data.floor_action,
    '', // Floor_State
    data.wall_action,
    '', // Wall_State
    data.toilet_action,
    '', // Toilet_State
    data.bidet_action,
    '', // Bidet_State
    data.sink_action,
    '', // Sink_State
    data.mirror_action,
    '', // Mirror_State
    data.trash_action,
    '', // Trash_State
    data.soap_action,
    '', // Soap_State
    data.tissue_action,
    '', // Tissue_State
    'รอตรวจสอบ'
  ];

  sheet.appendRow(rowData);
  return true;
}

function getRecords() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('ไม่พบ Sheet1');
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return [];
  }

  const headers = data[0];

  const records = data.slice(1).map((row, index) => {
    const obj = { rowId: index + 2 };

    headers.forEach((header, i) => {
      let cellValue = row[i];

      if (cellValue instanceof Date) {
        cellValue = cellValue.toISOString();
      }

      if (cellValue === undefined || cellValue === null || cellValue === '') {
        cellValue = null;
      }

      obj[header] = cellValue;
    });

    return obj;
  });

  return records.reverse();
}

function updateStatus(rowId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('ไม่พบ Sheet1');
  }

  if (!rowId || rowId < 2) {
    throw new Error('Row ID ไม่ถูกต้อง');
  }

  sheet.getRange(Number(rowId), 23).setValue(newStatus);
  return true;
}

/*
 * บันทึกผล "สะอาด/เพียงพอ" จากหน้า รอตรวจสอบ
 *
 * states:
 * {
 *   floor: true/false,
 *   wall: true/false,
 *   ...
 * }
 */
function updateReviewStates(rowId, states) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('ไม่พบ Sheet1');
  }

  if (!rowId || rowId < 2) {
    throw new Error('Row ID ไม่ถูกต้อง');
  }

  if (!states) {
    throw new Error('ไม่พบผลการตรวจสอบ');
  }

  const columns = {
    floor: 6,
    wall: 8,
    toilet: 10,
    bidet: 12,
    sink: 14,
    mirror: 16,
    trash: 18,
    soap: 20,
    tissue: 22
  };

  Object.keys(columns).forEach(function(key) {
    if (typeof states[key] !== 'boolean') {
      throw new Error('ผลตรวจสอบของ "' + key + '" ไม่ครบถ้วน');
    }

    sheet.getRange(Number(rowId), columns[key]).setValue(states[key]);
  });

  return true;
}

/*
 * Sheet2 ใหม่:
 * A = ชื่อผู้ปฏิบัติงาน
 * B = สถานที่รับผิดชอบ
 *
 * ตัวอย่าง:
 * A2 = สมชาย
 * B2 = "สถานที่ 1" "สถานที่ 2" "สถานที่ 3"
 *
 * ระบบจะลบเครื่องหมาย " หน้า/หลังชื่อสถานที่
 * และแยกหลายสถานที่จากข้อความที่ครอบด้วย "..."
 */
function parseLocations(value) {
  if (value === null || value === undefined || value === '') {
    return [];
  }

  const text = String(value).trim();
  const locations = [];

  const quotedMatches = text.match(/"([^"]+)"/g);

  if (quotedMatches && quotedMatches.length > 0) {
    quotedMatches.forEach(function(item) {
      const cleaned = item
        .replace(/^"+|"+$/g, '')
        .trim();

      if (cleaned) {
        locations.push(cleaned);
      }
    });
  } else {
    const cleaned = text
      .replace(/^["“”]+|["“”]+$/g, '')
      .trim();

    if (cleaned) {
      locations.push(cleaned);
    }
  }

  return [...new Set(locations)];
}

function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);

  if (!sheet) {
    throw new Error('ไม่พบ ' + MASTER_SHEET_NAME + ' กรุณาใช้ Sheet2 เป็นตารางรายชื่อและสถานที่');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 1) {
    return {
      workers: [],
      workerLocations: {}
    };
  }

  const values = sheet.getRange(1, 1, lastRow, 2).getValues();

  const workers = [];
  const workerLocations = {};

  values.forEach(function(row, index) {
    const worker = String(row[0] ?? '').trim();
    const locationCell = row[1];

    // ถ้าแถวแรกเป็นหัวตาราง ให้ข้าม
    if (index === 0 && worker === 'รายชื่อ') {
      return;
    }

    if (!worker) {
      return;
    }

    if (!workerLocations[worker]) {
      workerLocations[worker] = [];
      workers.push(worker);
    }

    parseLocations(locationCell).forEach(function(location) {
      if (!workerLocations[worker].includes(location)) {
        workerLocations[worker].push(location);
      }
    });
  });

  return {
    workers: workers,
    workerLocations: workerLocations
  };
}
