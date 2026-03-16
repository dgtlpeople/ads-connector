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

function getYesterdayDate_() {
  const date = new Date();
  date.setHours(0, 0, 0, 0);
  date.setDate(date.getDate() - 1);
  return date;
}

function getYesterdayDateKey_() {
  return formatDate_(getYesterdayDate_());
}

function toNumber_(value) {
  if (value === '' || value === null || value === undefined) return 0;
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}

function sanitizePlanGoal_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const normalized = typeof value === 'string'
    ? value.replace(/,/g, '').trim()
    : value;
  if (normalized === '') return '';

  const n = Number(normalized);
  if (isNaN(n)) return '';

  // PLAN uses 1 as a default placeholder, not a meaningful goal.
  return n === 1 ? '' : n;
}

function hasUsablePlanGoal_(value) {
  return sanitizePlanGoal_(value) !== '';
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

function logGoogleChange_(entry) {
  ensureHeader_(SHEETS.GOOGLE_CHANGES_LOG, HEADERS.GOOGLE_CHANGES_LOG);
  appendRows_(SHEETS.GOOGLE_CHANGES_LOG, [[
    new Date(),
    entry.action || '',
    entry.entity_level || '',
    entry.entity_id || '',
    entry.resource_name || '',
    entry.status || '',
    entry.request_payload || '',
    entry.response_or_error || ''
  ]]);
}

function enqueueVideoAction_(entry) {
  ensureHeader_(SHEETS.VIDEO_ACTION_QUEUE, HEADERS.VIDEO_ACTION_QUEUE);
  const existingId = findExistingQueuedVideoAction_(entry.campaign_id, entry.action);
  if (existingId) {
    return existingId;
  }

  const actionId = Utilities.getUuid();
  appendRows_(SHEETS.VIDEO_ACTION_QUEUE, [[
    new Date(),
    actionId,
    'QUEUED',
    'google',
    entry.campaign_id || '',
    entry.campaign_name || '',
    entry.action || '',
    entry.requested_by || Session.getActiveUser().getEmail() || '',
    entry.detail || '',
    0,
    '',
    '',
    ''
  ]]);
  return actionId;
}

function findExistingQueuedVideoAction_(campaignId, action) {
  const rows = readObjects_(SHEETS.VIDEO_ACTION_QUEUE);
  const targetCampaignId = normalizeId_(campaignId).replace(/-/g, '');
  const targetAction = String(action || '').trim();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const status = String(r.status || '').toUpperCase();
    if (status !== 'QUEUED' && status !== 'PROCESSING') continue;
    if (normalizeId_(r.campaign_id).replace(/-/g, '') !== targetCampaignId) continue;
    if (String(r.action || '').trim() !== targetAction) continue;
    return String(r.action_id || '');
  }
  return '';
}

function generateVideoQueueAdsScript_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const script = [
    "function main() {",
    "  var MAX_ATTEMPTS = 5;",
    "  var SPREADSHEET_URL = '" + ss.getUrl() + "';",
    "  var SHEET_NAME = 'VIDEO_ACTION_QUEUE';",
    "  var sh = SpreadsheetApp.openByUrl(SPREADSHEET_URL).getSheetByName(SHEET_NAME);",
    "  if (!sh) return;",
    "  var values = sh.getDataRange().getValues();",
    "  if (values.length < 2) return;",
    "  var headers = values[0];",
    "  var idx = {};",
    "  for (var h = 0; h < headers.length; h++) idx[String(headers[h])] = h;",
    "  function col(name) { return (idx[name] || 0) + 1; }",
    "  for (var i = 1; i < values.length; i++) {",
    "    var row = values[i];",
    "    var status = String(row[idx.status] || '').toUpperCase();",
    "    var attempts = Number(row[idx.attempts] || 0);",
    "    if (status !== 'QUEUED' && status !== 'ERROR') continue;",
    "    if (attempts >= MAX_ATTEMPTS) continue;",
    "    var campaignId = String(row[idx.campaign_id] || '');",
    "    var action = String(row[idx.action] || '');",
    "    sh.getRange(i + 1, col('status')).setValue('PROCESSING');",
    "    sh.getRange(i + 1, col('attempts')).setValue(attempts + 1);",
    "    sh.getRange(i + 1, col('processed_at')).setValue(new Date());",
    "    sh.getRange(i + 1, col('last_error')).setValue('');",
    "    SpreadsheetApp.flush();",
    "    try {",
    "      var it = AdsApp.videoCampaigns().withIds([campaignId]).get();",
    "      if (!it.hasNext()) throw new Error('Campaign not found');",
    "      var c = it.next();",
    "      var caps = c.getFrequencyCaps();",
    "      var eventType = 'IMPRESSION';",
    "      var timeUnit = 'MONTH';",
    "      var currentCap = caps.getFrequencyCapFor(eventType, timeUnit);",
    "      var currentLimit = currentCap ? Number(currentCap.getLimit()) : 2;",
    "      var nextLimit = action === 'Increase frequency cap'",
    "        ? Math.min(100, currentLimit + 1)",
    "        : Math.max(1, currentLimit - 1);",
    "      caps.removeFrequencyCapFor(eventType, timeUnit);",
    "      caps.newFrequencyCapBuilder()",
    "        .withEventType(eventType)",
    "        .withTimeUnit(timeUnit)",
    "        .withLimit(nextLimit)",
    "        .build();",
    "      sh.getRange(i + 1, col('status')).setValue('DONE');",
    "      sh.getRange(i + 1, col('processed_at')).setValue(new Date());",
    "      sh.getRange(i + 1, col('result')).setValue('Applied: ' + currentLimit + ' -> ' + nextLimit + ' / ' + timeUnit);",
    "      sh.getRange(i + 1, col('last_error')).setValue('');",
    "    } catch (e) {",
    "      sh.getRange(i + 1, col('status')).setValue('ERROR');",
    "      sh.getRange(i + 1, col('processed_at')).setValue(new Date());",
    "      sh.getRange(i + 1, col('last_error')).setValue(String(e));",
    "      sh.getRange(i + 1, col('result')).setValue('');",
    "    }",
    "  }",
    "}"
  ].join('\n');

  const sh = getSheet_('ADS_SCRIPT_TEMPLATE');
  sh.clear();
  sh.getRange(1, 1).setValue(script);
  sh.getRange(1, 1).setWrap(true);
  sh.getRange(2, 1).setValue('Copy only cell A1 into Google Ads Script.');
  sh.setColumnWidth(1, 1200);
  sh.setRowHeight(1, 900);
}
