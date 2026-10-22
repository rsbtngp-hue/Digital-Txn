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
  if (action === 'replaceDates') {
    const dates = Array.isArray(payload.dates) ? payload.dates.map(normalizeDate_).filter(Boolean) : [];
    const records = Array.isArray(payload.records) ? payload.records : [];
    if (!dates.length) return output_({ok:false, error:'replaceDates requires one or more valid dates'}, e);
    const result = replaceDates_(dates, records);
    return output_({ok:true, action:action, dates:dates, received:records.length, ...result}, e);
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
  if (!records.length) return {inserted:0, updated:0, skipped:0};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DATA_SHEET);
    if (!sheet) { setupSpreadsheet(); sheet = ss.getSheetByName(DATA_SHEET); }
    const lastRow = sheet.getLastRow();
    const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues() : [];
    const rowByKey = new Map();
    existing.forEach((r, i) => {
      const key = normalizeDate_(r[0]) + '|' + String(r[1]).trim();
      if (r[0] && r[1]) rowByKey.set(key, i + 2);
    });

    let skipped = 0, inserted = 0, updated = 0;
    const appendRows = [];
    const updates = [];
    const now = new Date().toISOString();
    records.forEach(r => {
      const date = normalizeDate_(r.txn_date || r['txn-date'] || r.date);
      const officeId = String(r.office_id || r['office-id'] || '').trim();
      if (!date || !officeId) { skipped++; return; }
      const row = recordToRow_(r, date, officeId, now);
      const key = date + '|' + officeId;
      if (rowByKey.has(key)) {
        const rowNo = rowByKey.get(key);
        if (rowNo === -1) {
          const index = appendRows.findIndex(x => x[0] === date && String(x[1]) === officeId);
          if (index >= 0) appendRows[index] = row;
        } else {
          updates.push({rowNo:rowNo, values:row});
          updated++;
        }
      } else {
        appendRows.push(row);
        rowByKey.set(key, -1);
        inserted++;
      }
    });

    updates.sort((a,b) => a.rowNo - b.rowNo);
    let i = 0;
    while (i < updates.length) {
      let j = i + 1;
      while (j < updates.length && updates[j].rowNo === updates[j-1].rowNo + 1) j++;
      sheet.getRange(updates[i].rowNo, 1, j-i, HEADERS.length).setValues(updates.slice(i,j).map(x=>x.values));
      i = j;
    }
    if (appendRows.length) sheet.getRange(sheet.getLastRow()+1, 1, appendRows.length, HEADERS.length).setValues(appendRows);
    markSuccessfulSync_();
    return {inserted:inserted, updated:updated, skipped:skipped};
  } finally { lock.releaseLock(); }
}

function replaceDates_(dates, records) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DATA_SHEET);
    if (!sheet) { setupSpreadsheet(); sheet = ss.getSheetByName(DATA_SHEET); }
    const dateSet = new Set(dates);
    const lastRow = sheet.getLastRow();
    const existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];
    const keepRows = existingRows.filter(row => !dateSet.has(normalizeDate_(row[0])));
    const newRows = [];
    const seen = new Set();
    let skipped = 0;
    const now = new Date().toISOString();
    records.forEach(r => {
      const date = normalizeDate_(r.txn_date || r['txn-date'] || r.date);
      const officeId = String(r.office_id || r['office-id'] || '').trim();
      if (!date || !officeId || !dateSet.has(date)) { skipped++; return; }
      const key = date + '|' + officeId;
      if (seen.has(key)) return;
      seen.add(key);
      newRows.push(recordToRow_(r, date, officeId, now));
    });
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
    const finalRows = keepRows.concat(newRows);
    if (finalRows.length) sheet.getRange(2, 1, finalRows.length, HEADERS.length).setValues(finalRows);
    markSuccessfulSync_();
    return {replacedDates:dates.length, removed:existingRows.length-keepRows.length, inserted:newRows.length, skipped:skipped};
  } finally { lock.releaseLock(); }
}

function recordToRow_(r, date, officeId, now) {
  return [
    date, officeId, value_(r.office_name || r['office-name']),
    value_(r.office_type || r['office-type']), value_(r.division),
    value_(r.sub_division || r['sub-division']), number_(r.cash),
    number_(r.pos), number_(r.qr), number_(r.total_digital),
    number_(r.total_txns), value_(r.pct), now
  ];
}

function markSuccessfulSync_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meta = ss.getSheetByName(META_SHEET);
  if (meta) meta.getRange(4, 1, 1, 2).setValues([['Last Successful Sync', new Date().toISOString()]]);
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
