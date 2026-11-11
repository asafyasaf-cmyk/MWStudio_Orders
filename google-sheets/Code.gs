/**
 * MW Studio – order intake for Google Sheets.
 * Paste into the sheet: Extensions → Apps Script → Code.gs, then Deploy → Web app.
 * The website POSTs each order as JSON; this script appends one row per order.
 */
const SHEET_NAME = 'הזמנות';
const STATUSES = ['התקבלה', 'בטיפול', 'בייצור', 'מוכנה', 'נמסרה', 'בוטלה'];
// Optional: an address that gets an email for each new order. Empty = no email.
const NOTIFY_EMAIL = '';

const COLUMNS = [
  ['id', 'מספר הזמנה'], ['created', 'תאריך'], ['status', 'סטטוס'], ['name', 'שם'], ['phone', 'טלפון'],
  ['product', 'מוצר'], ['material', 'חומר'], ['qty', 'כמות'], ['unit', 'מחיר ליח׳'], ['total', 'סה״כ'],
  ['color', 'צבע'], ['note', 'הערות'], ['desc', 'בקשת הלקוח'], ['text', 'טקסט מלא']
];
const COL = k => COLUMNS.findIndex(c => c[0] === k) + 1;

/** Run once from the editor (▶ setup) to create and format the sheet. */
function setup() {
  const sh = sheet_();
  const n = COLUMNS.length;
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#F4F6F8');
  sh.getRange(2, COL('phone'), sh.getMaxRows() - 1, 1).setNumberFormat('@');
  sh.getRange(2, COL('created'), sh.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(2, COL('unit'), sh.getMaxRows() - 1, 2).setNumberFormat('₪#,##0');
  const status = sh.getRange(2, COL('status'), sh.getMaxRows() - 1, 1);
  status.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  const colors = { 'התקבלה': '#FBE7D6', 'בטיפול': '#FFF4C2', 'בייצור': '#D6E8FB', 'מוכנה': '#D3F2E0', 'נמסרה': '#E6E6E6', 'בוטלה': '#F3D6D6' };
  sh.setConditionalFormatRules(Object.keys(colors).map(s =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s]).setRanges([status]).build()));
  sh.setColumnWidth(COL('text'), 320);
  return 'ok';
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNS.map(c => c[1]));
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  if (!d || d.website) return json_({ ok: true }); // honeypot filled = bot; pretend success
  if (!d.id || !d.name || !d.phone) return json_({ ok: false, error: 'missing_fields' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, COL('id'), sh.getLastRow() - 1, 1).getValues().flat() : [];
    if (ids.indexOf(String(d.id)) >= 0) return json_({ ok: true, id: d.id, duplicate: true });
    sh.appendRow(COLUMNS.map(([k]) => k === 'created' ? new Date() : k === 'status' ? STATUSES[0] : clean_(d[k], k)));
    const row = sh.getLastRow();
    sh.getRange(row, COL('phone')).setNumberFormat('@').setValue(clean_(d.phone, 'phone'));
    sh.getRange(row, COL('status')).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
    if (NOTIFY_EMAIL) MailApp.sendEmail(NOTIFY_EMAIL, `הזמנה חדשה ${clean_(d.id)} – ${clean_(d.name)}`, String(clean_(d.text)));
    return json_({ ok: true, id: d.id, row: row });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return json_({ ok: true, service: 'MW Studio orders' }); }

// Numbers stay numbers; text is trimmed, length-capped and protected from formula injection.
function clean_(v, k) {
  if (v == null) return '';
  if (typeof v === 'number' && (k === 'qty' || k === 'unit' || k === 'total')) return isFinite(v) ? v : '';
  let s = String(v).slice(0, k === 'text' ? 5000 : 500).trim();
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
