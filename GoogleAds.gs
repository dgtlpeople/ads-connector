var GOOGLE_SKIP_UNIQUE_USERS_FOR_RUN_ = false;

function getGoogleAdsConfig_() {
  const props = getScriptProps_();
  const cfg = {
    developerToken: props.getProperty('GOOGLE_ADS_DEVELOPER_TOKEN'),
    customerId: normalizeId_(props.getProperty('GOOGLE_ADS_CUSTOMER_ID')).replace(/-/g, ''),
    loginCustomerId: normalizeId_(props.getProperty('GOOGLE_ADS_LOGIN_CUSTOMER_ID')).replace(/-/g, ''),
    clientId: props.getProperty('GOOGLE_OAUTH_CLIENT_ID'),
    clientSecret: props.getProperty('GOOGLE_OAUTH_CLIENT_SECRET'),
    refreshToken: props.getProperty('GOOGLE_ADS_REFRESH_TOKEN')
  };

  if (!isConfigured_(cfg.developerToken) || !isConfigured_(cfg.customerId) || !isConfigured_(cfg.loginCustomerId) || !isConfigured_(cfg.clientId) || !isConfigured_(cfg.clientSecret) || !isConfigured_(cfg.refreshToken)) {
    throw new Error('Missing required Google Ads Script Properties.');
  }

  return cfg;
}

function getGoogleAccessToken_() {
  const cfg = getGoogleAdsConfig_();

  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      client_id: cfg.clientId,
      client_secret: cfg.clientSecret,
      refresh_token: cfg.refreshToken,
      grant_type: 'refresh_token'
    },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    throw new Error('Failed to refresh Google OAuth token: ' + res.getContentText());
  }

  return JSON.parse(res.getContentText()).access_token;
}

function googleAdsSearchStream_(query) {
  const cfg = getGoogleAdsConfig_();
  const token = getGoogleAccessToken_();
  const url = 'https://googleads.googleapis.com/v20/customers/' + cfg.customerId + '/googleAds:searchStream';

  const headers = {
    Authorization: 'Bearer ' + token,
    'developer-token': cfg.developerToken,
    'Content-Type': 'application/json'
  };

  if (isConfigured_(cfg.loginCustomerId)) {
    headers['login-customer-id'] = cfg.loginCustomerId;
  }

  const maxAttempts = 4;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: headers,
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code >= 200 && code < 300) {
      return JSON.parse(body);
    }

    const retryable = isGoogleRetryableError_(code, body);
    if (!retryable || attempt === maxAttempts) {
      throw new Error('Google Ads API error ' + code + ': ' + body);
    }

    Utilities.sleep(Math.pow(2, attempt) * 1000);
  }
}

function loadGoogleEntities() {
  withErrorLogging_('loadGoogleEntities failed', function () {
    ensureHeader_(SHEETS.CAMPAIGNS_ENABLED, HEADERS.CAMPAIGNS_ENABLED);
    const keep = readObjects_(SHEETS.CAMPAIGNS_ENABLED).filter(function (r) {
      return normalizePlatform_(r.platform) !== 'google';
    });

    clearDataKeepHeader_(SHEETS.CAMPAIGNS_ENABLED);

    const query = [
      'SELECT',
      '  campaign.id,',
      '  campaign.name,',
      '  campaign.start_date,',
      '  campaign.end_date,',
      '  campaign.status,',
      '  campaign.advertising_channel_type',
      'FROM campaign',
      'WHERE campaign.status = ENABLED',
      'ORDER BY campaign.name'
    ].join('\n');

    const out = [];

    keep.forEach(function (r) {
      out.push(mapCampaignEnabledRow_(r));
    });

    const chunks = googleAdsSearchStream_(query);
    chunks.forEach(function (chunk) {
      (chunk.results || []).forEach(function (row) {
        const c = row.campaign || {};
        out.push([
          'google',
          '',
          'campaign',
          String(c.id || ''),
          c.name || '',
          String(c.id || ''),
          c.name || '',
          '',
          '',
          c.startDate || '',
          c.endDate || '',
          c.status || '',
          c.advertisingChannelType || ''
        ]);
      });
    });

    if (out.length) appendRows_(SHEETS.CAMPAIGNS_ENABLED, out);
    sortCampaignsEnabled_();
  });
}

function fetchGoogleEntityMetrics_(entity) {
  if (normalizeEntityLevel_(entity.entity_level) !== 'campaign') {
    throw new Error('Google Ads currently supports campaign level only.');
  }

  const campaignId = normalizeId_(entity.entity_id).replace(/-/g, '');
  if (!/^\d+$/.test(campaignId)) {
    throw new Error('Invalid Google campaign ID: ' + entity.entity_id);
  }

  const fields = [
    'campaign.id',
    'campaign.name',
    'campaign.start_date',
    'campaign.end_date',
    'campaign.status',
    'campaign.advertising_channel_type',
    'metrics.impressions',
    'metrics.average_cpm',
    'metrics.video_quartile_p25_rate',
    'metrics.video_quartile_p50_rate',
    'metrics.video_quartile_p75_rate',
    'metrics.video_quartile_p100_rate'
  ];
  if (!GOOGLE_SKIP_UNIQUE_USERS_FOR_RUN_) {
    fields.push('metrics.unique_users');
  }

  const baseQuery = [
    'SELECT',
    '  ' + fields.join(',\n  '),
    'FROM campaign',
    'WHERE campaign.id = ' + campaignId
  ].join('\n');

  let selected = null;
  let uniqueUsersFetchSucceeded = false;
  try {
    selected = flattenGoogleMetricRow_(googleAdsSearchStream_(baseQuery));
    uniqueUsersFetchSucceeded = !GOOGLE_SKIP_UNIQUE_USERS_FOR_RUN_ && fields.indexOf('metrics.unique_users') !== -1;
  } catch (e) {
    const attemptedUniqueUsers = fields.indexOf('metrics.unique_users') !== -1;
    if (!attemptedUniqueUsers) throw e;

    if (!GOOGLE_SKIP_UNIQUE_USERS_FOR_RUN_ && isGoogleBandwidthQuotaError_(e.message || '')) {
      GOOGLE_SKIP_UNIQUE_USERS_FOR_RUN_ = true;
      log_('Google unique_users disabled for run', e.message);
    } else {
      log_('Google unique_users query failed', e.message);
    }

    const fallbackQuery = baseQuery.replace(',\n  metrics.unique_users', '');
    selected = flattenGoogleMetricRow_(googleAdsSearchStream_(fallbackQuery));
    uniqueUsersFetchSucceeded = false;
  }

  if (!selected) {
    return null;
  }

  let reach = selected.uniqueUsers === '' || selected.uniqueUsers === undefined
    ? 0
    : toNumber_(selected.uniqueUsers);
  const impressions = toNumber_(selected.impressions);

  if (uniqueUsersFetchSucceeded && selected.uniqueUsers !== undefined && selected.uniqueUsers !== '') {
    upsertGoogleReachCache_(campaignId, selected.name || normalizeId_(entity.entity_name), reach);
  } else {
    const cachedReach = getCachedGoogleReach_(campaignId);
    if (cachedReach !== null) {
      reach = cachedReach;
    }
  }

  return {
    platform: 'google',
    account_id: '',
    entity_level: 'campaign',
    entity_id: String(selected.id || campaignId),
    entity_name: selected.name || normalizeId_(entity.entity_name),
    campaign_id: String(selected.id || campaignId),
    campaign_name: selected.name || normalizeId_(entity.entity_name),
    adset_id: '',
    adset_name: '',
    start_date: selected.startDate || '',
    end_date: selected.endDate || '',
    impressions: impressions,
    reach: reach,
    frequency: reach > 0 ? impressions / reach : 0,
    cpm: toNumber_(selected.averageCpm) / 1000000,
    video_p25: toNumber_(selected.p25),
    video_p50: toNumber_(selected.p50),
    video_p75: toNumber_(selected.p75),
    video_p100: toNumber_(selected.p100),
    status: selected.status || '',
    channel_type: selected.channelType || ''
  };
}

function isGoogleRetryableError_(code, body) {
  if (code === 429 || code === 500 || code === 502 || code === 503 || code === 504) return true;
  const text = String(body || '').toLowerCase();
  return (
    text.indexOf('resource_exhausted') !== -1 ||
    text.indexOf('rate exceeded') !== -1 ||
    text.indexOf('quota exceeded') !== -1 ||
    text.indexOf('temporarily unavailable') !== -1
  );
}

function isGoogleBandwidthQuotaError_(text) {
  const t = String(text || '').toLowerCase();
  return t.indexOf('bandwidth quota exceeded') !== -1;
}

function flattenGoogleMetricRow_(chunks) {
  let out = null;
  chunks.forEach(function (chunk) {
    (chunk.results || []).forEach(function (row) {
      const c = row.campaign || {};
      const m = row.metrics || {};
      out = {
        id: c.id,
        name: c.name,
        startDate: c.startDate,
        endDate: c.endDate,
        status: c.status,
        channelType: c.advertisingChannelType,
        impressions: m.impressions,
        uniqueUsers: m.uniqueUsers,
        averageCpm: m.averageCpm,
        p25: m.videoQuartileP25Rate,
        p50: m.videoQuartileP50Rate,
        p75: m.videoQuartileP75Rate,
        p100: m.videoQuartileP100Rate
      };
    });
  });
  return out;
}

function mapCampaignEnabledRow_(row) {
  return [
    row.platform || '',
    row.account_id || '',
    row.entity_level || '',
    row.entity_id || '',
    row.entity_name || '',
    row.campaign_id || '',
    row.campaign_name || '',
    row.adset_id || '',
    row.adset_name || '',
    row.start_date || '',
    row.end_date || '',
    row.status || '',
    row.channel_type || ''
  ];
}

function getCachedGoogleReach_(campaignId) {
  ensureHeader_(SHEETS.REACH_CACHE, HEADERS.REACH_CACHE);
  const cacheRows = readObjects_(SHEETS.REACH_CACHE);
  for (let i = 0; i < cacheRows.length; i++) {
    const r = cacheRows[i];
    if (
      normalizePlatform_(r.platform) === 'google' &&
      normalizeEntityLevel_(r.entity_level) === 'campaign' &&
      normalizeId_(r.entity_id).replace(/-/g, '') === normalizeId_(campaignId).replace(/-/g, '')
    ) {
      return toNumber_(r.reach);
    }
  }
  return null;
}

function upsertGoogleReachCache_(campaignId, entityName, reach) {
  ensureHeader_(SHEETS.REACH_CACHE, HEADERS.REACH_CACHE);
  const sh = getSheet_(SHEETS.REACH_CACHE);
  const lastRow = sh.getLastRow();
  const safeCampaignId = normalizeId_(campaignId).replace(/-/g, '');
  const rowData = ['google', '', 'campaign', safeCampaignId, entityName || '', toNumber_(reach), new Date()];

  if (lastRow <= 1) {
    sh.getRange(2, 1, 1, rowData.length).setValues([rowData]);
    return;
  }

  const values = sh.getRange(2, 1, lastRow - 1, HEADERS.REACH_CACHE.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const platform = normalizePlatform_(values[i][0]);
    const level = normalizeEntityLevel_(values[i][2]);
    const entityId = normalizeId_(values[i][3]).replace(/-/g, '');
    if (platform === 'google' && level === 'campaign' && entityId === safeCampaignId) {
      sh.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
      return;
    }
  }

  sh.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
}
