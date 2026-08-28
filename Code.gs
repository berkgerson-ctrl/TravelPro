/**
 * All In One Travel Pro — Google Apps Script backend
 * -----------------------------------------------------
 * Deploy: Extensions > Apps Script > paste this file > Deploy > New deployment
 *         Type: Web app | Execute as: Me | Who has access: Anyone
 *         Copy the Web App URL into the app's Settings screen.
 */

const SHEET_TRIPS = 'Seyahatler';
const SHEET_PLANS = 'Planlar';
const SHEET_EXPENSES = 'Harcamalar';
const SHEET_TODOS = 'Yapilacaklar';

const HEADERS = {
  [SHEET_TRIPS]: ['SeyahatID', 'SeyahatAdi', 'Lokasyonlar', 'BaslangicTarihi', 'BitisTarihi', 'ToplamButce', 'ParaBirimi'],
  [SHEET_PLANS]: ['PlanID', 'SeyahatID', 'Gun', 'Saat', 'Baslik', 'Kategori', 'KonumLink', 'HarcamaID', 'Notlar', 'Sira'],
  [SHEET_EXPENSES]: ['HarcamaID', 'SeyahatID', 'PlanID', 'Tarih', 'Kategori', 'Tutar', 'ParaBirimi', 'Aciklama', 'Notlar', 'TutarOrijinal', 'ParaBirimiOrijinal'],
  [SHEET_TODOS]: ['TodoID', 'SeyahatID', 'Metin', 'Tamamlandi'],
};
// Columns that must always be stored as plain text so Sheets doesn't
// auto-convert them into Date/Time serial values (e.g. "09:00" -> time,
// or "2026-09-04" -> a Date that round-trips wrong once JSON-serialized).
const TEXT_COLUMNS = {
  [SHEET_TRIPS]: ['BaslangicTarihi', 'BitisTarihi'],
  [SHEET_PLANS]: ['Saat'],
  [SHEET_EXPENSES]: ['Tarih'],
};

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _applyTextFormats(sheet, sheetName) {
  const cols = TEXT_COLUMNS[sheetName];
  if (!cols) return;
  const headers = HEADERS[sheetName];
  cols.forEach((colName) => {
    const idx = headers.indexOf(colName);
    if (idx === -1) return;
    sheet.getRange(1, idx + 1, Math.max(sheet.getMaxRows(), 1000), 1).setNumberFormat('@');
  });
}

function _ensureSheet(name) {
  const ss = _ss();
  let sheet = ss.getSheetByName(name);
  const requiredHeaders = HEADERS[name];
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    _applyTextFormats(sheet, name);
    return { created: true, sheet };
  }
  // Repair: fill a blank header row, or append any headers that are missing
  // (without touching existing columns/data) so older sheets self-heal.
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const isBlank = existing.join('') === '';
  let repaired = false;
  if (isBlank) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    repaired = true;
  } else {
    const missing = requiredHeaders.filter((h) => existing.indexOf(h) === -1);
    if (missing.length) {
      sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
      repaired = true;
    }
  }
  _applyTextFormats(sheet, name);
  return { created: false, repaired, sheet };
}

function initDatabase() {
  const report = {};
  Object.keys(HEADERS).forEach((name) => {
    const result = _ensureSheet(name);
    report[name] = result.created ? 'created' : (result.repaired ? 'repaired' : 'ok');
  });
  return { success: true, report: report };
}

function _sheetToObjects(name) {
  const sheet = _ss().getSheetByName(name);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter((row) => row.join('') !== '')
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function getAllData() {
  return {
    success: true,
    trips: _sheetToObjects(SHEET_TRIPS),
    plans: _sheetToObjects(SHEET_PLANS),
    expenses: _sheetToObjects(SHEET_EXPENSES),
    todos: _sheetToObjects(SHEET_TODOS),
  };
}

function _headerRow(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

function _findRowById(sheet, idColIndex, id) {
  const values = sheet.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idColIndex]) === String(id)) return r + 1; // 1-indexed row
  }
  return -1;
}

function _upsert(sheetName, idKey, row) {
  const { sheet } = _ensureSheet(sheetName);
  const headers = _headerRow(sheet); // actual on-sheet header order (handles legacy/repaired sheets)
  const idColIndex = headers.indexOf(idKey);
  if (!row[idKey]) row[idKey] = Utilities.getUuid();
  const rowArray = headers.map((h) => (row[h] !== undefined ? row[h] : ''));
  const existingRow = _findRowById(sheet, idColIndex, row[idKey]);
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, headers.length).setValues([rowArray]);
  } else {
    sheet.appendRow(rowArray);
  }
  return { success: true, id: row[idKey] };
}

function _delete(sheetName, idKey, id) {
  const { sheet } = _ensureSheet(sheetName);
  const headers = _headerRow(sheet);
  const idColIndex = headers.indexOf(idKey);
  const existingRow = _findRowById(sheet, idColIndex, id);
  if (existingRow > 0) sheet.deleteRow(existingRow);
  return { success: true };
}

function deleteTripCascade(tripId) {
  _delete(SHEET_TRIPS, 'SeyahatID', tripId);
  [SHEET_PLANS, SHEET_EXPENSES, SHEET_TODOS].forEach((name) => {
    const { sheet } = _ensureSheet(name);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const tripColIndex = headers.indexOf('SeyahatID');
    for (let r = values.length - 1; r >= 1; r--) {
      if (String(values[r][tripColIndex]) === String(tripId)) sheet.deleteRow(r + 1);
    }
  });
  return { success: true };
}

function doGet(e) {
  const action = e.parameter.action || 'getData';
  try {
    if (action === 'init') return _json(initDatabase());
    if (action === 'getData') return _json(getAllData());
    return _json({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return _json({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    switch (action) {
      case 'init':
        result = initDatabase(); break;
      case 'saveTrip':
        result = _upsert(SHEET_TRIPS, 'SeyahatID', body.data); break;
      case 'deleteTrip':
        result = deleteTripCascade(body.data.SeyahatID); break;
      case 'saveActivity':
        result = _upsert(SHEET_PLANS, 'PlanID', body.data); break;
      case 'deleteActivity':
        result = _delete(SHEET_PLANS, 'PlanID', body.data.PlanID); break;
      case 'saveExpense':
        result = _upsert(SHEET_EXPENSES, 'HarcamaID', body.data); break;
      case 'deleteExpense':
        result = _delete(SHEET_EXPENSES, 'HarcamaID', body.data.HarcamaID); break;
      case 'saveTodo':
        result = _upsert(SHEET_TODOS, 'TodoID', body.data); break;
      case 'deleteTodo':
        result = _delete(SHEET_TODOS, 'TodoID', body.data.TodoID); break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
    return _json(result);
  } catch (err) {
    return _json({ success: false, error: err.message });
  }
}
