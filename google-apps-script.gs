// SRCAM waitlist endpoint — hardened for public signups.
// Paste this into Extensions > Apps Script inside the destination Google Sheet,
// then Deploy > Manage deployments > Edit > New version > Deploy.
const SPREADSHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = 'Waitlist';
const HEADERS = ['Submitted at', 'Full name', 'Email', 'Phone', 'Country', 'Would pay', 'Client timestamp'];
const DEDUPE_SECONDS = 60 * 60 * 24; // one accepted submission per email every 24 hours
const MIN_FORM_AGE_MS = 1200;
const MAX_FORM_AGE_MS = 2 * 60 * 60 * 1000;

function json_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function cleanText_(value, maxLength) {
  if (typeof value !== 'string') return '';
  const cleaned = value.replace(/[\u0000-\u001F\u007F]/g, '').trim().replace(/\s+/g, ' ');
  if (cleaned.length > maxLength) throw new Error('field_too_long');
  return cleaned;
}

function safeCell_(value) {
  // Stops spreadsheet formula injection from values beginning =, +, -, or @.
  const text = String(value || '');
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function sha256_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return bytes.map(byte => ('0' + (byte & 0xff).toString(16)).slice(-2)).join('');
}

function getWaitlistSheet_() {
  const spreadsheet = SPREADSHEET_ID === 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE'
    ? SpreadsheetApp.getActiveSpreadsheet()
    : SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function validate_(data) {
  if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('invalid_payload');
  if (cleanText_(data.company, 100)) throw new Error('bot_detected'); // hidden honeypot

  const startedAt = Number(data.startedAt);
  const age = Date.now() - startedAt;
  if (!Number.isFinite(startedAt) || age < MIN_FORM_AGE_MS || age > MAX_FORM_AGE_MS) throw new Error('invalid_form_timing');

  const name = cleanText_(data.name, 80);
  const email = cleanText_(data.email, 254).toLowerCase();
  const phone = cleanText_(data.phone, 32);
  const country = cleanText_(data.country, 100);
  const wouldPay = cleanText_(data.wouldPay, 3);
  const clientTimestamp = cleanText_(data.timestamp, 40);

  if (name.length < 2) throw new Error('invalid_name');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error('invalid_email');
  if (phone && !/^\+[1-9]\d{6,14}$/.test(phone)) throw new Error('invalid_phone');
  if (wouldPay !== 'Yes' && wouldPay !== 'No') throw new Error('invalid_choice');

  return { name, email, phone, country, wouldPay, clientTimestamp };
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || typeof e.postData.contents !== 'string') throw new Error('missing_body');
    const data = validate_(JSON.parse(e.postData.contents));
    if (!lock.tryLock(5000)) return json_({ ok: false, error: 'busy' });

    const cache = CacheService.getScriptCache();
    const dedupeKey = 'signup:' + sha256_(data.email);
    if (cache.get(dedupeKey)) return json_({ ok: true, duplicate: true });

    getWaitlistSheet_().appendRow([
      new Date(),
      safeCell_(data.name),
      safeCell_(data.email),
      safeCell_(data.phone),
      safeCell_(data.country),
      data.wouldPay,
      safeCell_(data.clientTimestamp)
    ]);
    SpreadsheetApp.flush();
    cache.put(dedupeKey, '1', DEDUPE_SECONDS);
    return json_({ ok: true });
  } catch (error) {
    // Log the real error privately; never expose Sheet/project details to a public endpoint.
    console.warn(String(error));
    return json_({ ok: false, error: 'invalid_request' });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function doGet() {
  // Do not expose Sheet data or diagnostics from the public URL.
  return json_({ ok: false, error: 'method_not_allowed' });
}
