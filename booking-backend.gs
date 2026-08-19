/**
 * ZW Site — Booking backend
 * ---------------------------
 * Paste this whole file into: Google Sheets → Extensions → Apps Script
 * (replace whatever boilerplate is already there), then:
 *
 * 1. Change SECRET below to something only you know (any random string).
 * 2. Deploy → New deployment → type: Web app
 *      Execute as:      Me
 *      Who has access:  Anyone
 * 3. Authorize when Google prompts you.
 * 4. Copy the "Web app URL" (ends in /exec).
 * 5. In your admin panel's Bookings tab, paste that URL + your SECRET.
 *
 * A "Bookings" sheet gets created automatically the first time it's used —
 * you never need to create it by hand.
 */

const SECRET = 'CHANGE_ME_TO_YOUR_OWN_SECRET';
const BOOKINGS_SHEET = 'Bookings';

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  // Admin panel updating a status — requires the secret.
  if (data.action === 'updateStatus') {
    if (data.secret !== SECRET) return jsonOut({ ok: false, error: 'Unauthorized' });
    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const idCol = headers.indexOf('id');
    const statusCol = headers.indexOf('status');
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]) === String(data.id)) {
        sheet.getRange(i + 1, statusCol + 1).setValue(data.status);
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: 'Not found' });
  }

  // Public form submission — no secret required (same trust model as any contact form).
  const sheet = getSheet();
  const id = Utilities.getUuid();
  const createdAt = new Date().toISOString();
  sheet.appendRow([id, createdAt, data.name || '', data.date || '', data.time || '',
    data.editingType || '', data.contacts || '', 'new']);
  return jsonOut({ ok: true, id: id });
}

function doGet(e) {
  if (e.parameter.secret !== SECRET) return jsonOut({ ok: false, error: 'Unauthorized' });
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const items = rows.slice(1)
    .filter(r => r.some(cell => cell !== ''))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
  return jsonOut({ ok: true, items: items });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BOOKINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BOOKINGS_SHEET);
    sheet.appendRow(['id', 'createdAt', 'name', 'date', 'time', 'editingType', 'contacts', 'status']);
  }
  return sheet;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
