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
  const normalized = typeof value === 'string'
    ? value.replace(/,/g, '').trim()
    : value;
  const n = Number(normalized);
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
  ensureHeader_(SHEETS.CAMPAIGNS_ENABLED, HEADERS.CAMPAIGNS_ENABLED);
  const rows = readObjects_(SHEETS.CAMPAIGNS_ENABLED);
  if (!rows.length) return;

  rows.sort(function (a, b) {
    const aEnded = String(a.plan_status || '') === 'ENDED_NOT_REQUIRED_IN_PLAN';
    const bEnded = String(b.plan_status || '') === 'ENDED_NOT_REQUIRED_IN_PLAN';
    if (aEnded !== bEnded) return aEnded ? 1 : -1;

    const p = normalizePlatform_(a.platform).localeCompare(normalizePlatform_(b.platform));
    if (p !== 0) return p;

    return normalizeId_(a.entity_name).localeCompare(normalizeId_(b.entity_name));
  });

  clearDataKeepHeader_(SHEETS.CAMPAIGNS_ENABLED);
  appendRows_(SHEETS.CAMPAIGNS_ENABLED, rows.map(mapCampaignEnabledRow_));
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

function parseCompositeEntityName_(entityName) {
  const raw = String(entityName || '');
  const split = raw.split(' | ');
  if (split.length < 2) {
    return { campaign_name: raw, child_name: raw };
  }
  return {
    campaign_name: split[0].trim(),
    child_name: split.slice(1).join(' | ').trim()
  };
}

function getCachedReachEntry_(platform, accountId, entityLevel, entityId, options) {
  ensureHeader_(SHEETS.REACH_CACHE, HEADERS.REACH_CACHE);
  const now = new Date().getTime();
  const todayKey = formatDate_(new Date());
  const rows = readObjects_(SHEETS.REACH_CACHE);
  const targetPlatform = normalizePlatform_(platform);
  const targetAccountId = normalizeId_(accountId);
  const targetLevel = normalizeEntityLevel_(entityLevel);
  const targetEntityId = normalizeId_(entityId);
  const opts = options || {};
  const allowExpired = !!opts.allowExpired;
  let freshestValid = null;
  let freshestAny = null;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (normalizePlatform_(r.platform) !== targetPlatform) continue;
    const rowAccountId = normalizeId_(r.account_id);
    if (rowAccountId !== targetAccountId) {
      // Backward compatibility: old Google cache rows may have blank account_id.
      if (!(targetPlatform === 'google' && rowAccountId === '')) continue;
    }
    if (normalizeEntityLevel_(r.entity_level) !== targetLevel) continue;
    if (normalizeId_(r.entity_id) !== targetEntityId) continue;

    const cachedAt = r.cached_at ? new Date(r.cached_at).getTime() : 0;
    if (!cachedAt || isNaN(cachedAt)) continue;
    const reach = toNumber_(r.reach);

    const entry = {
      cachedAt: cachedAt,
      reach: reach,
      ageHours: (now - cachedAt) / 3600000,
      isToday: formatDate_(new Date(cachedAt)) === todayKey
    };
    entry.isExpired = entry.ageHours > REACH_CACHE_TTL_HOURS;

    if (!freshestAny || cachedAt > freshestAny.cachedAt) {
      freshestAny = entry;
    }

    if (!entry.isExpired && (!freshestValid || cachedAt > freshestValid.cachedAt)) {
      freshestValid = entry;
    }
  }

  if (freshestValid) return freshestValid;
  if (allowExpired && freshestAny) return freshestAny;
  return null;
}

function getCachedReach_(platform, accountId, entityLevel, entityId, options) {
  const entry = getCachedReachEntry_(platform, accountId, entityLevel, entityId, options);
  return entry ? entry.reach : null;
}

function setCachedReach_(platform, accountId, entityLevel, entityId, entityName, reach) {
  ensureHeader_(SHEETS.REACH_CACHE, HEADERS.REACH_CACHE);
  const sh = getSheet_(SHEETS.REACH_CACHE);
  const data = sh.getDataRange().getValues();
  const rowData = [
    normalizePlatform_(platform),
    normalizeId_(accountId),
    normalizeEntityLevel_(entityLevel),
    normalizeId_(entityId),
    normalizeId_(entityName),
    toNumber_(reach),
    new Date()
  ];

  if (data.length < 2) {
    sh.getRange(2, 1, 1, rowData.length).setValues([rowData]);
    return;
  }

  const headers = data[0];
  const idx = {
    platform: headers.indexOf('platform'),
    account_id: headers.indexOf('account_id'),
    entity_level: headers.indexOf('entity_level'),
    entity_id: headers.indexOf('entity_id')
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowPlatform = normalizePlatform_(row[idx.platform]);
    const rowAccountId = normalizeId_(row[idx.account_id]);
    const accountMatches = rowAccountId === rowData[1] || (rowData[0] === 'google' && rowAccountId === '');
    if (
      rowPlatform === rowData[0] &&
      accountMatches &&
      normalizeEntityLevel_(row[idx.entity_level]) === rowData[2] &&
      normalizeId_(row[idx.entity_id]) === rowData[3]
    ) {
      sh.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
      return;
    }
  }

  sh.getRange(sh.getLastRow() + 1, 1, 1, rowData.length).setValues([rowData]);
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

function buildPlanEntityLookup_() {
  ensureHeader_(SHEETS.PLAN, HEADERS.PLAN);
  const planRows = readObjects_(SHEETS.PLAN);
  const lookup = {};

  planRows.forEach(function (r) {
    const platform = normalizePlatform_(r.platform);
    const accountId = normalizeId_(r.account_id);
    const level = normalizeEntityLevel_(r.entity_level);
    const entityId = normalizeId_(r.entity_id);
    if (!platform || !level || !entityId) return;
    lookup[entityKey_(platform, accountId, level, entityId)] = true;
  });

  return lookup;
}

function isEntityInPlan_(lookup, row) {
  const platform = normalizePlatform_(row.platform);
  const accountId = normalizeId_(row.account_id);
  const level = normalizeEntityLevel_(row.entity_level);
  const entityId = normalizeId_(row.entity_id);
  if (!platform || !level || !entityId) return false;

  if (lookup[entityKey_(platform, accountId, level, entityId)]) return true;

  // Backward compatibility with historical blank account_id in PLAN.
  if (lookup[entityKey_(platform, '', level, entityId)]) return true;

  return false;
}

function isLiveByEndDate_(endDateValue) {
  const end = formatDate_(endDateValue);
  if (!end) return false;
  const today = formatDate_(new Date());
  return end >= today;
}

function updateCampaignsEnabledPlanStatusForPlatform_(platform) {
  const targetPlatform = normalizePlatform_(platform);
  ensureHeader_(SHEETS.CAMPAIGNS_ENABLED, HEADERS.CAMPAIGNS_ENABLED);
  ensureHeader_(SHEETS.PLAN, HEADERS.PLAN);
  const rows = readObjects_(SHEETS.CAMPAIGNS_ENABLED);
  if (!rows.length) return;

  const lookup = buildPlanEntityLookup_();
  let liveNotInPlanCount = 0;
  const planAppendRows = [];

  rows.forEach(function (r) {
    if (normalizePlatform_(r.platform) !== targetPlatform) return;

    if (!isLiveByEndDate_(r.end_date)) {
      r.plan_status = 'ENDED_NOT_REQUIRED_IN_PLAN';
      return;
    }

    if (isEntityInPlan_(lookup, r)) {
      r.plan_status = 'LIVE_IN_PLAN';
    } else {
      r.plan_status = 'LIVE_NOT_IN_PLAN';
      liveNotInPlanCount += 1;

      const planKey = entityKey_(r.platform, r.account_id, r.entity_level, r.entity_id);
      if (!lookup[planKey]) {
        planAppendRows.push([
          r.platform || '',
          r.account_id || '',
          r.entity_level || '',
          r.entity_id || '',
          r.entity_name || '',
          1,
          1
        ]);
        lookup[planKey] = true;
      }
    }
  });

  clearDataKeepHeader_(SHEETS.CAMPAIGNS_ENABLED);
  appendRows_(SHEETS.CAMPAIGNS_ENABLED, rows.map(mapCampaignEnabledRow_));
  sortCampaignsEnabled_();
  if (planAppendRows.length) {
    appendRows_(SHEETS.PLAN, planAppendRows);
  }
  log_('PLAN status updated', 'platform=' + targetPlatform + '; live_not_in_plan=' + liveNotInPlanCount);
}
