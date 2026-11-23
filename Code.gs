/**
 * Birthday RSVP backend
 * 1. Create a Google Sheet, then Extensions > Apps Script.
 * 2. Paste this file in, change HOST_PIN below, and Save.
 * 3. Deploy > New deployment > Web app
 *      Execute as: Me    Who has access: Anyone
 * 4. Copy the Web App URL into PARTY.scriptUrl in index.html.
 */
const HOST_PIN = 'change-me';   // you'll type this to see the guest list
const SHEET_NAME = 'RSVPs';
const HEADERS = ['Updated', 'Family', 'Response', 'Kids', 'Adults', 'Guest names'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const name = clean(d.name, 80);
    if (!name) return json({ ok: false, error: 'name' });
    const response = ['yes', 'no', 'maybe'].includes(d.response) ? d.response : 'maybe';
    const row = [new Date(), name, response, num(d.kids), num(d.adults), clean(d.guests, 900)];

    const sh = sheet();
    const names = sh.getRange(1, 2, sh.getLastRow(), 1).getValues().flat();
    const key = name.toLowerCase();
    const idx = names.findIndex((n, i) => i > 0 && String(n).trim().toLowerCase() === key);
    if (idx > 0) sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);   // update existing RSVP
    else sh.appendRow(row);
    return json({ ok: true, updated: idx > 0 });
  } catch (err) {
    return json({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (!e.parameter.pin || e.parameter.pin !== HOST_PIN) return json({ ok: false, error: 'pin' });
  const values = sheet().getDataRange().getValues().slice(1);
  const rows = values.map(r => ({ name: r[1], response: r[2], kids: r[3], adults: r[4], guests: r[5] }));
  return json({ ok: true, rows });
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  } else {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);   // keep headers current
  }
  return sh;
}

function clean(v, max) {
  let s = String(v || '').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // stop spreadsheet formula injection
  return s;
}
function num(v) { const n = parseInt(v, 10); return isNaN(n) ? 0 : Math.max(0, Math.min(10, n)); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
