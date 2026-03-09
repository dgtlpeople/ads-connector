function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeader_(sheetName, headers) {
  const sh = getSheet_(sheetName);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
}

function clearDataKeepHeader_(sheetName) {
  const sh = getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}

function appendRows_(sheetName, rows) {
  if (!rows || !rows.length) return;
  const sh = getSheet_(sheetName);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function formatDate_(d) {
  if (!d) return '';
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function toNumber_(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function log_(message, detail) {
  ensureHeader_(SHEETS.LOG, HEADERS.LOG);
  appendRows_(SHEETS.LOG, [[new Date(), message, detail || '']]);
}

function readObjects_(sheetName) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];

  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function normalizeId_(value) {
  return String(value || '').replace(/-/g, '').trim();
}

function statusOf_(value) {
  if (!value) return 'MISSING';
  const v = String(value).trim().toUpperCase();
  if (v === '1' || v === 'PENDING' || v === 'TEMP' || v === 'TO_FILL') {
    return 'PLACEHOLDER';
  }
  return 'OK';
}