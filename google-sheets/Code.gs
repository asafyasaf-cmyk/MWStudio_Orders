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
  ['color', 'צבע'], ['note', 'הערות'], ['desc', 'בקשת הלקוח'], ['text', 'טקסט מלא'], ['stepUrl', 'קובץ STEP']
];
// STEP files of profile orders are saved in this Google Drive folder (created on first use).
const STEP_FOLDER = 'MW Studio – קובצי STEP';
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
  } else if (sh.getLastColumn() < COLUMNS.length) { // a sheet made by an older version: add the new headers
    sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => c[1])]);
  }
  return sh;
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  if (d && d.action === 'design') return json_(aiDesign_(d));
  if (!d || d.website) return json_({ ok: true }); // honeypot filled = bot; pretend success
  if (!d.id || !d.name || !d.phone) return json_({ ok: false, error: 'missing_fields' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, COL('id'), sh.getLastRow() - 1, 1).getValues().flat() : [];
    if (ids.indexOf(String(d.id)) >= 0) return json_({ ok: true, id: d.id, duplicate: true });
    d.stepUrl = '';
    if (d.step && String(d.step).indexOf('ISO-10303-21') === 0 && String(d.step).length < 20000000) {
      try { d.stepUrl = saveStep_(String(d.step), String(d.stepName || d.id + '.step')); } catch (err) { d.stepUrl = 'שגיאה בשמירה: ' + err; }
    }
    delete d.step;
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

function saveStep_(text, name) {
  const it = DriveApp.getFoldersByName(STEP_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(STEP_FOLDER);
  const safe = name.replace(/[^\w.\-א-ת]+/g, '_').slice(0, 120);
  return folder.createFile(safe.endsWith('.step') ? safe : safe + '.step', text, 'application/step').getUrl();
}

function doGet() { return json_({ ok: true, service: 'MW Studio orders' }); }

// Numbers stay numbers; text is trimmed, length-capped and protected from formula injection.
function clean_(v, k) {
  if (v == null) return '';
  if (k === 'stepUrl') return String(v).slice(0, 500);
  if (typeof v === 'number' && (k === 'qty' || k === 'unit' || k === 'total')) return isFinite(v) ? v : '';
  let s = String(v).slice(0, k === 'text' ? 5000 : 500).trim();
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- AI: read the customer's text, sketch or photo and return a part design ----------
 * The Anthropic API key is stored in Project Settings → Script properties as ANTHROPIC_API_KEY,
 * never in this file or on the website. AI_DAILY_LIMIT caps how many requests a day the site can make.
 */
const AI_MODEL = 'claude-opus-5-5';
const AI_DAILY_LIMIT = 150;
const FAMILIES = ['shs', 'rhs', 'chs', 'angle', 'flat', 'round', 'upn', 'ipe', 'hea', 'heb'];

const AI_SYSTEM = `You are the ordering assistant of an Israeli laser-cutting, sheet-metal bending and steel-profile shop. Customers write in Hebrew, often attach a photo of a hand sketch drawn in pen, and cannot read technical drawings. Turn what they send into a part design that the website draws in 3D and prices.

Part types ("template"):
- profile: a standard steel section cut to length, with optional holes and notches. Fill "profile": family (shs square tube, rhs rectangular tube, chs round tube, angle equal L-angle, flat flat bar, round round bar, upn U channel, ipe / hea / heb beams), size = an id from the catalog list in the message (pick the closest real size), length = one piece in mm, ops = holes and notches.
- plate (W length, H width, r corner radius, hn holes, hd hole diameter, off hole centre to edge, center central hole diameter), disc (D, Di inner hole, hn holes on bolt circle, hd, pcd), angle (bent sheet L: A, B, L, hn, hd), channel (bent sheet U: W, H, L, hn, hd), zed (a, h, c, L, hn, hd), tray (L, W, H, hn, hd). For these put the numbers in "params" as {k, v} pairs and set "profile" to null. For "profile" leave "params" empty.

Operations on a profile (every field is required; fill the ones that do not apply with these defaults: c 0, d 9, n 1, s 100, thru true, shape "rect", wdt 20, span 0, ang 90, depth 10):
- hole: x = distance in mm from the start of the piece to the first hole centre; face = top | bottom | front | back (holes seen in a top view are "top", holes seen in a side / front view are "front"); c = offset across the face from its centre line; d = diameter; n holes spaced s apart along the length; thru = drilled through both walls / the whole section.
- notch, shape "rect": a rectangular cut from the face, x = where the notch starts, wdt = its length along the piece, span = its width across the face (0 = the whole face), c = offset of that width from the face centre line, depth = how deep it goes in.
- notch, shape "v": a V cut for bending the profile, x = centre of the V, ang = opening angle (90 for a right-angle bend), depth usually the section height minus one wall thickness.

Reading sketches: the left end of the main view is the start of the piece (x = 0). Turn chained dimensions into positions measured from the start. Numbers are millimetres unless the sketch says cm or m. "Ø" or "קוטר" means diameter; a bolt Mx needs a hole of x+1 mm. If a dimension is unreadable or missing, use the most plausible value, say in the reply what you assumed, and ask one question about the most important missing value. Never invent holes or notches that are not drawn or described.

Materials: steel (ברזל שחור), galv (מגולוון), ss (נירוסטה), al (אלומיניום). ral is a 4-digit RAL code, or "none" for no paint; keep the customer's current values for anything they did not mention. qty is the number of pieces. thickness applies to sheet parts only (mm).

Every message starts with the part currently on screen. When the customer describes a new part or a change, return the full updated design; when they only ask a question, return design null.
"reply": plain Hebrew, at most 25 words, no markdown, no prices or delivery dates. Name what you built and any assumption.
"suggestions": 2-3 Hebrew quick replies of 1-3 words each.`;

const OP_SCHEMA = { type: 'object', additionalProperties: false,
  required: ['kind', 'face', 'x', 'c', 'd', 'n', 's', 'thru', 'shape', 'wdt', 'span', 'ang', 'depth'],
  properties: { kind: { type: 'string', enum: ['hole', 'notch'] }, face: { type: 'string', enum: ['top', 'bottom', 'front', 'back'] },
    x: { type: 'number' }, c: { type: 'number' }, d: { type: 'number' }, n: { type: 'integer' }, s: { type: 'number' }, thru: { type: 'boolean' },
    shape: { type: 'string', enum: ['rect', 'v'] }, wdt: { type: 'number' }, span: { type: 'number' }, ang: { type: 'number' }, depth: { type: 'number' } } };
const DESIGN_SCHEMA = { type: 'object', additionalProperties: false,
  required: ['template', 'params', 'profile', 'material', 'thickness', 'qty', 'ral'],
  properties: {
    template: { type: 'string', enum: ['profile', 'plate', 'disc', 'angle', 'channel', 'zed', 'tray'] },
    params: { type: 'array', items: { type: 'object', additionalProperties: false, required: ['k', 'v'], properties: { k: { type: 'string' }, v: { type: 'number' } } } },
    profile: { anyOf: [{ type: 'null' }, { type: 'object', additionalProperties: false, required: ['family', 'size', 'length', 'ops'],
      properties: { family: { type: 'string', enum: FAMILIES }, size: { type: 'string' }, length: { type: 'number' }, ops: { type: 'array', items: OP_SCHEMA } } }] },
    material: { type: 'string', enum: ['steel', 'galv', 'ss', 'al'] }, thickness: { type: 'number' }, qty: { type: 'integer' }, ral: { type: 'string' } } };
const REPLY_SCHEMA = { type: 'object', additionalProperties: false, required: ['reply', 'design', 'suggestions'],
  properties: { reply: { type: 'string' }, design: { anyOf: [{ type: 'null' }, DESIGN_SCHEMA] }, suggestions: { type: 'array', items: { type: 'string' } } } };

function aiDesign_(d) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'ai_not_configured' };
  if (!countToday_()) return { ok: false, error: 'rate_limited' };

  const content = [];
  if (d.image) {
    const img = String(d.image);
    if (img.length > 6000000) return { ok: false, error: 'image_too_large' };
    const types = ['image/jpeg', 'image/png', 'image/webp', 'image/gif'];
    content.push({ type: 'image', source: { type: 'base64', media_type: types.indexOf(d.mediaType) >= 0 ? d.mediaType : 'image/jpeg', data: img } });
  }
  content.push({ type: 'text', text:
    `[החלק שעל המסך עכשיו: ${String(d.current || '').slice(0, 4000)}]\n` +
    `[קטלוג פרופילים (family: size ids): ${String(d.catalog || '').slice(0, 8000)}]\n` +
    (d.image ? '[צורף שרטוט או תמונה של הלקוח. קראו ממנו את הצורה, המידות, החורים והחירוצים.]\n' : '') +
    String(d.text || '(ראו את השרטוט המצורף)').slice(0, 2000) });

  // Earlier turns as plain text, alternating user / assistant, ending with an assistant turn.
  let history = (Array.isArray(d.history) ? d.history : []).slice(-10)
    .filter(m => m && (m.role === 'user' || m.role === 'assistant') && typeof m.content === 'string')
    .map(m => ({ role: m.role, content: m.content.slice(0, 4000) }));
  while (history.length && history[0].role !== 'user') history.shift();
  history = history.filter((m, i) => i === 0 || m.role !== history[i - 1].role);
  if (history.length && history[history.length - 1].role !== 'assistant') history.pop();

  const body = {
    model: AI_MODEL,
    max_tokens: 8000,
    fallbacks: 'default',
    system: AI_SYSTEM,
    output_config: { effort: 'medium', format: { type: 'json_schema', schema: REPLY_SCHEMA } },
    messages: history.concat([{ role: 'user', content: content }])
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body),
    headers: { 'x-api-key': key, 'anthropic-version': '2026-10-21', 'anthropic-beta': 'server-side-fallback-2026-07-01' }
  });
  const code = res.getResponseCode();
  let j = {};
  try { j = JSON.parse(res.getContentText() || '{}'); } catch (err) { /* keep empty */ }
  if (code === 429 || code === 529) return { ok: false, error: 'rate_limited' };
  if (code >= 400) return { ok: false, error: 'api_error', status: code, detail: String((j.error && j.error.message) || '').slice(0, 300) };
  if (j.stop_reason === 'refusal') return { ok: false, error: 'refused' };
  if (j.stop_reason === 'max_tokens') return { ok: false, error: 'invalid_json' };
  const text = (j.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  try { return { ok: true, result: JSON.parse(text) }; } catch (err) { return { ok: false, error: 'invalid_json' }; }
}

// Counts today's AI requests; false once AI_DAILY_LIMIT is reached.
function countToday_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const props = PropertiesService.getScriptProperties();
    const k = 'ai_count_' + Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyyMMdd');
    const n = Number(props.getProperty(k) || 0);
    if (n >= AI_DAILY_LIMIT) return false;
    props.setProperty(k, String(n + 1));
    return true;
  } finally { lock.releaseLock(); }
}
