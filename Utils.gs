function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeader_(sheetName, headers) {
  const sh = getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  const current = lastRow > 0 ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0] : [];
  const mismatch = headers.some(function (h, i) {
    return String(current[i] || '') !== h;
  }) || current.length < headers.length;

  if (lastRow === 0 || mismatch) {
    sh.clear();
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

function readObjects_(sheetName) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  return values.slice(1).map(function (row) {
    const out = {};
    headers.forEach(function (h, i) {
      out[h] = row[i];
    });
    return out;
  });
}

function formatDate_(value) {
  if (!value) return '';
  return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function toNumber_(value) {
  if (value === '' || value === null || value === undefined) return 0;
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}

function normalizeId_(value) {
  return String(value || '').trim();
}

function normalizePlatform_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeEntityLevel_(value) {
  return String(value || '').trim().toLowerCase();
}

function entityKey_(platform, accountId, entityLevel, entityId) {
  return [
    normalizePlatform_(platform),
    normalizeId_(accountId),
    normalizeEntityLevel_(entityLevel),
    normalizeId_(entityId)
  ].join('|');
}

function isConfigured_(value) {
  return String(value || '').trim() !== '';
}

function log_(message, detail) {
  try {
    ensureHeader_(SHEETS.LOG, HEADERS.LOG);
    appendRows_(SHEETS.LOG, [[new Date(), message, detail || '']]);
  } catch (e) {
    console.error('LOG_FAILED', message, detail || '', e.message);
  }
}

function withErrorLogging_(message, fn) {
  try {
    return fn();
  } catch (e) {
    log_(message, e && e.stack ? e.stack : String(e));
    throw e;
  }
}

function sortCampaignsEnabled_() {
  const sh = getSheet_(SHEETS.CAMPAIGNS_ENABLED);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 2 || lastCol <= 0) return;

  sh.getRange(2, 1, lastRow - 1, lastCol).sort([
    { column: 1, ascending: true }, // platform
    { column: 5, ascending: true }  // entity_name
  ]);
}

function ensureReachCacheSampleRow_() {
  ensureHeader_(SHEETS.REACH_CACHE, HEADERS.REACH_CACHE);
  const sh = getSheet_(SHEETS.REACH_CACHE);
  if (sh.getLastRow() > 1) return;

  sh.getRange(2, 1, 1, HEADERS.REACH_CACHE.length).setValues([[
    'google',
    '',
    'campaign',
    '1234567890',
    'Sample Campaign',
    250000,
    new Date()
  ]]);
}
