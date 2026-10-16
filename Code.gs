/**
 * Bike Survey — Google Sheets receiver
 *
 * Paste this into your Google Sheet via Extensions → Apps Script,
 * run setup() once, then Deploy → New deployment → Web app
 * (Execute as: Me, Who has access: Anyone).
 */

const SHEET_NAME = 'Responses';

// Private key the dashboard must send to read the data. Keep it out of the public repo.
const DASHBOARD_KEY = '41i4LtdxDorJuwpAyV2BMJft';

// [key sent by the app, column heading in the sheet]
const QUESTION_COLUMNS = [
  ['quickRelease', 'Quick-release axles'],
  ['frontLocked',  'Front wheel locked'],
  ['missingWheel', 'Missing a wheel'],
];

const HEADERS = ['ID', 'Saved at', 'Collector']
  .concat(QUESTION_COLUMNS.map(function (q) { return q[1]; }))
  .concat(['Latitude', 'Longitude', 'Accuracy (m)', 'Map link', 'Received at']);

/** Run once from the editor: creates the Responses tab and asks for permission. */
function setup() {
  const sheet = getSheet_();
  Logger.log('Ready: "' + sheet.getName() + '" tab with ' + dataRows_(sheet) + ' entries.');
}

/**
 * Opening the web app URL in a browser shows a "live" message.
 * The dashboard calls ?action=data&key=… to read every entry.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const sheet = getSheet_();
  if (p.action === 'data') {
    if (p.key !== DASHBOARD_KEY) return json_({ ok: false, error: 'Wrong dashboard key' });
    return json_({ ok: true, entries: readEntries_(sheet), serverTime: new Date().toISOString() });
  }
  return json_({ ok: true, message: 'Bike Survey endpoint is live', rows: dataRows_(sheet) });
}

function readEntries_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(String);
  const col = function (name) { return head.indexOf(name); };
  const c = {
    id: col('ID'), savedAt: col('Saved at'), collector: col('Collector'),
    lat: col('Latitude'), lng: col('Longitude'), acc: col('Accuracy (m)'),
  };
  const qCols = QUESTION_COLUMNS.map(function (q) { return [q[0], col(q[1])]; });
  return values.slice(1).filter(function (r) { return r[c.id] !== ''; }).map(function (r) {
    const out = {
      id: String(r[c.id]),
      savedAt: r[c.savedAt] instanceof Date ? r[c.savedAt].toISOString() : String(r[c.savedAt]),
      collector: String(r[c.collector] || ''),
      lat: typeof r[c.lat] === 'number' ? r[c.lat] : null,
      lng: typeof r[c.lng] === 'number' ? r[c.lng] : null,
      accuracy: typeof r[c.acc] === 'number' ? r[c.acc] : null,
    };
    qCols.forEach(function (q) { out[q[0]] = q[1] >= 0 ? String(r[q[1]]) : ''; });
    return out;
  });
}

/** The app sends {op:'add', record}, {op:'delete', id}, or {op:'batch', ops:[…]} for a backlog. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // several phones can upload at once
    const op = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    if (op.op === 'add') addRecord_(sheet, op.record);
    else if (op.op === 'delete') deleteRecord_(sheet, op.id);
    else if (op.op === 'batch') applyBatch_(sheet, op.ops || []);
    else throw new Error('Unknown operation');
    return json_({ ok: true, rows: dataRows_(sheet) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function addRecord_(sheet, r) {
  if (!r || !r.id) throw new Error('Missing record id');
  if (findRow_(sheet, r.id)) return; // already saved by an earlier retry — skip the duplicate
  sheet.appendRow(buildRow_(r));
}

/** Writes a whole backlog in one go: one ID lookup, one write. */
function applyBatch_(sheet, ops) {
  const last = sheet.getLastRow();
  const seen = {};
  if (last > 1) sheet.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { seen[String(r[0])] = true; });
  const rows = [];
  ops.forEach(function (o) {
    if (o.op === 'add' && o.record && o.record.id && !seen[o.record.id]) {
      rows.push(buildRow_(o.record));
      seen[o.record.id] = true;
    }
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  ops.forEach(function (o) { if (o.op === 'delete') deleteRecord_(sheet, o.id); });
}

function buildRow_(r) {
  const answers = r.answers || {};
  const hasLoc = typeof r.lat === 'number' && typeof r.lng === 'number';
  const row = [r.id, r.savedAt ? new Date(r.savedAt) : new Date(), r.collector || '']
    .concat(QUESTION_COLUMNS.map(function (q) { return answers[q[0]] || ''; }))
    .concat([
      hasLoc ? r.lat : '',
      hasLoc ? r.lng : '',
      r.accuracy != null ? r.accuracy : '',
      hasLoc ? 'https://maps.google.com/?q=' + r.lat + ',' + r.lng : '',
      new Date(),
    ]);
  return row;
}

function deleteRecord_(sheet, id) {
  const row = findRow_(sheet, id);
  if (row) sheet.deleteRow(row);
}

function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const hit = sheet.getRange(2, 1, last - 1, 1)
    .createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.hideColumns(1); // ID column: only used to stop duplicate uploads
  }
  return sheet;
}

function dataRows_(sheet) {
  return Math.max(sheet.getLastRow() - 1, 0);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
