/**
 * รับข้อมูลใบสมัครอบรมช่างโซลาร์เซลล์ แล้วบันทึกลง Google Sheet
 * วิธีติดตั้ง: ดู README.md
 */

const SHEET_NAME = 'ผู้สมัคร';
const NOTIFY_EMAIL = '';   // ใส่อีเมลถ้าต้องการให้แจ้งเตือนทุกครั้งที่มีคนสมัคร เช่น 'you@gmail.com'
const ID_PREFIX = 'SOL';

// [ชื่อช่องในฟอร์ม, หัวคอลัมน์ในชีต]
const FIELDS = [
  ['timestamp',   'วันเวลาที่สมัคร'],
  ['appId',       'เลขที่ใบสมัคร'],
  ['prefix',      'คำนำหน้า'],
  ['firstName',   'ชื่อ'],
  ['lastName',    'นามสกุล'],
  ['age',         'อายุ'],
  ['province',    'จังหวัด'],
  ['phone',       'เบอร์โทร'],
  ['lineId',      'LINE ID'],
  ['email',       'อีเมล'],
  ['education',   'วุฒิการศึกษา'],
  ['major',       'สาขา'],
  ['occupation',  'อาชีพ'],
  ['experience',  'ประสบการณ์ไฟฟ้า'],
  ['certificate', 'หนังสือรับรอง/มาตรฐานฝีมือ'],
  ['courses',     'หลักสูตรที่สนใจ'],
  ['schedule',    'ช่วงเวลาสะดวก'],
  ['goal',        'เป้าหมาย'],
  ['source',      'รู้จักจาก'],
  ['note',        'หมายเหตุ'],
  ['consent',     'ความยินยอม PDPA'],
  ['status',      'สถานะ (เจ้าหน้าที่กรอก)'],
];

const REQUIRED = ['prefix', 'firstName', 'lastName', 'age', 'province', 'phone', 'education',
  'occupation', 'experience', 'certificate', 'courses', 'schedule', 'goal', 'consent'];

// เปิด URL ของเว็บแอป (/exec) = แสดงหน้าฟอร์มจากไฟล์ index.html ในโปรเจกต์นี้
// ต่อท้าย ?ping=1 เพื่อเช็กว่าระบบทำงาน
function doGet(e) {
  if (e && e.parameter && e.parameter.ping) return json_({ ok: true, message: 'ระบบรับใบสมัครทำงานปกติ' });
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ใบสมัครอบรมช่างโซลาร์เซลล์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// หน้าฟอร์มที่เปิดจาก Apps Script เรียกฟังก์ชันนี้ผ่าน google.script.run
function submitApplication(payload) {
  return saveApplication_(payload || {});
}

// หน้าฟอร์มที่วางไว้ที่อื่น (เช่น GitHub Pages) ส่งข้อมูลมาทาง POST
function doPost(e) {
  return json_(saveApplication_((e && e.parameters) || {}));
}

function saveApplication_(p) {
  const get = k => [].concat(p[k] || []).map(v => String(v).trim()).filter(Boolean).join(', ').slice(0, 1000);

  // ช่องดักบอท: คนจริงมองไม่เห็นช่องนี้
  if (get('website')) return { ok: true, id: '-' };

  const missing = REQUIRED.filter(k => !get(k));
  if (missing.length) return { ok: false, error: 'missing: ' + missing.join(',') };
  if (!/^0\d{8,9}$/.test(get('phone'))) return { ok: false, error: 'phone' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    const seq = sheet.getLastRow(); // แถว 1 คือหัวตาราง
    const id = ID_PREFIX + '-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMM') + '-' + String(seq).padStart(4, '0');

    const row = FIELDS.map(([k]) => {
      if (k === 'timestamp') return new Date();
      if (k === 'appId') return id;
      if (k === 'status') return 'รอติดต่อ';
      const v = get(k);
      // ใส่ ' นำหน้าให้เป็นข้อความ: เก็บเลข 0 หน้าเบอร์โทร และกันสูตร (=, +, -, @) ที่ผู้กรอกพิมพ์มา
      return v ? "'" + v : '';
    });
    sheet.appendRow(row);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL,
        'ผู้สมัครอบรมโซลาร์ใหม่: ' + get('firstName') + ' ' + get('lastName') + ' (' + id + ')',
        FIELDS.slice(2, -1).map(([k, label]) => label + ': ' + (get(k) || '-')).join('\n') +
        '\n\nเปิดชีต: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
    }
    return { ok: true, id };
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(FIELDS.map(f => f[1]));
    sheet.getRange(1, 1, 1, FIELDS.length).setFontWeight('bold').setBackground('#1b3a5c').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    const statusCol = FIELDS.findIndex(f => f[0] === 'status') + 1;
    sheet.getRange(2, statusCol, 999, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['รอติดต่อ', 'ติดต่อแล้ว', 'ยืนยันเข้าอบรม', 'ยกเลิก'], true).build());
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
