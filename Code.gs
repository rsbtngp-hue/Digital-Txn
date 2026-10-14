/**
 * Digital-Txn Google Sheets backend
 * Bind this Apps Script project to the spreadsheet created for Digital-Txn.
 * Then run setupSpreadsheet() once from the Apps Script editor and authorize it.
 */

const APP_NAME = 'Digital-Txn';
const DATA_SHEET = 'Digital Txn Data';
const META_SHEET = 'Sync Status';
const HEADERS = [
  'txn_date', 'office_id', 'office_name', 'office_type', 'division',
  'sub_division', 'cash', 'pos', 'qr', 'total_digital', 'total_txns',
  'pct', 'updated_at'
];

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'health');
  if (action === 'health' || action === 'status') {
    return output_({
      ok: true,
      app: APP_NAME,
      spreadsheet: SpreadsheetApp.getActiveSpreadsheet().getName(),
      action: action,
      serverTime: new Date().toISOString()
    }, e);
  }
  if (action === 'getRecords' || action === 'getData') {
    const records = readRecords_();
    return output_({
      ok: true,
      app: APP_NAME,
      records: records,
      count: records.length,
      serverTime: new Date().toISOString()
    }, e);
  }
  return output_({ok:false, error:'Unknown action: ' + action}, e);
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return output_({ok:false, error:'Request body must be valid JSON: ' + err.message}, e);
  }

  const action = String(payload.action || '');
  if (action === 'setupSpreadsheet') {
    setupSpreadsheet();
    return output_({ok:true, message:'Spreadsheet setup completed'}, e);
  }
  if (action === 'appendRecords') {
    const records = Array.isArray(payload.records) ? payload.records : [];
    const result = upsertRecords_(records);
    return output_({ok:true, action:action, received:records.length, ...result}, e);
  }
  return output_({ok:false, error:'Unknown action: ' + action}, e);
}

/** Run this once from the Apps Script editor after creating/binding the sheet. */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open Apps Script from the target Google Spreadsheet (Extensions > Apps Script).');

  let data = ss.getSheetByName(DATA_SHEET);
  if (!data) data = ss.insertSheet(DATA_SHEET);
  if (data.getLastRow() === 0) {
    data.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    const existing = data.getRange(1, 1, 1, Math.max(data.getLastColumn(), HEADERS.length))
      .getDisplayValues()[0];
    HEADERS.forEach((h, i) => {
      if (existing[i] !== h) data.getRange(1, i + 1).setValue(h);
    });
  }
  data.setFrozenRows(1);
  data.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  data.autoResizeColumns(1, HEADERS.length);

  let meta = ss.getSheetByName(META_SHEET);
  if (!meta) meta = ss.insertSheet(META_SHEET);
  if (meta.getLastRow() === 0) {
    meta.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    meta.getRange(2, 1, 3, 2).setValues([
      ['App', APP_NAME],
      ['Created At', new Date().toISOString()],
      ['Last Successful Sync', 'Not synced yet']
    ]);
    meta.setFrozenRows(1);
    meta.getRange(1, 1, 1, 2).setFontWeight('bold');
    meta.autoResizeColumns(1, 2);
  }
  return {spreadsheetId:ss.getId(), spreadsheetName:ss.getName()};
}

function upsertRecords_(records) {
  if (!records.length) return {insertedOrUpdated:0, skipped:0};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) {
    setupSpreadsheet();
    sheet = ss.getSheetByName(DATA_SHEET);
  }
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues()
    : [];
  const rowByKey = new Map();
  existing.forEach((r, i) => rowByKey.set(String(r[0]).trim() + '|' + String(r[1]).trim(), i + 2));

  let insertedOrUpdated = 0;
  let skipped = 0;
  const appendRows = [];
  const updateRows = [];

  records.forEach(r => {
    const date = normalizeDate_(r.txn_date || r['txn-date'] || r.date);
    const officeId = String(r.office_id || r['office-id'] || '').trim();
    if (!date || !officeId) { skipped++; return; }
    const row = [
      date, officeId, value_(r.office_name || r['office-name']),
      value_(r.office_type || r['office-type']), value_(r.division),
      value_(r.sub_division || r['sub-division']), number_(r.cash),
      number_(r.pos), number_(r.qr), number_(r.total_digital),
      number_(r.total_txns), value_(r.pct), new Date().toISOString()
    ];
    const key = date + '|' + officeId;
    const rowNo = rowByKey.get(key);
    if (rowNo && rowNo <= sheet.getLastRow()) updateRows.push({rowNo:rowNo, values:row});
    else if (rowNo) appendRows[rowNo - sheet.getLastRow() - 1] = row;
    else {
      const newRowNo = sheet.getLastRow() + appendRows.length + 1;
      appendRows.push(row);
      rowByKey.set(key, newRowNo);
    }
  });

  updateRows.forEach(item => sheet.getRange(item.rowNo, 1, 1, HEADERS.length).setValues([item.values]));
  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, HEADERS.length).setValues(appendRows);
  }

  const meta = ss.getSheetByName(META_SHEET);
  if (meta) {
    const row = Math.max(meta.getLastRow() + 1, 4);
    meta.getRange(4, 1, 1, 2).setValues([['Last Successful Sync', new Date().toISOString()]]);
  }
  return {insertedOrUpdated:updateRows.length + appendRows.length, updated:updateRows.length, inserted:appendRows.length, skipped:skipped};
}

function readRecords_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(1, 1, sheet.getLastRow(), HEADERS.length).getValues();
  const headers = values.shift().map(String);
  return values.map(row => {
    const out = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      out[h] = v;
    });
    return out;
  });
}

function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const dmy = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (dmy) return dmy[3] + '-' + ('0' + dmy[2]).slice(-2) + '-' + ('0' + dmy[1]).slice(-2);
  const d = new Date(s);
  return isNaN(d) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function value_(v) { return v === null || v === undefined ? '' : v; }
function number_(v) { const n = Number(v); return v === '' || v === null || v === undefined || !isFinite(n) ? 0 : n; }

function output_(obj, e) {
  const json = JSON.stringify(obj);
  const callback = e && e.parameter && e.parameter.callback;
  if (callback && /^[A-Za-z_$][\w.$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
