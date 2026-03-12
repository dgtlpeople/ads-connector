function getTikTokConfig_() {
  const props = getScriptProps_();
  const cfg = {
    accessToken: normalizeId_(props.getProperty('TIKTOK_ACCESS_TOKEN')),
    advertiserId: normalizeId_(props.getProperty('TIKTOK_ADVERTISER_ID')),
    refreshToken: normalizeId_(props.getProperty('TIKTOK_REFRESH_TOKEN')),
    appId: normalizeId_(props.getProperty('TIKTOK_APP_ID')),
    secret: normalizeId_(props.getProperty('TIKTOK_SECRET'))
  };

  if (!isConfigured_(cfg.accessToken) || !isConfigured_(cfg.advertiserId)) {
    throw new Error('Missing TIKTOK_ACCESS_TOKEN or TIKTOK_ADVERTISER_ID in Script Properties.');
  }

  return cfg;
}

function tiktokHasRefreshConfig_(cfg) {
  return isConfigured_(cfg.refreshToken) && isConfigured_(cfg.appId) && isConfigured_(cfg.secret);
}

function isTikTokAuthError_(responseCode, body) {
  const text = String(body || '').toLowerCase();
  if (responseCode === 401 || responseCode === 403) return true;
  return text.indexOf('access token') !== -1 || text.indexOf('invalid token') !== -1 || text.indexOf('auth') !== -1;
}

function buildQueryString_(params) {
  return Object.keys(params || {})
    .filter(function (k) {
      return params[k] !== undefined && params[k] !== null && params[k] !== '';
    })
    .map(function (k) {
      const value = params[k];
      const serialized = (typeof value === 'object') ? JSON.stringify(value) : String(value);
      return encodeURIComponent(k) + '=' + encodeURIComponent(serialized);
    })
    .join('&');
}

function tiktokApiRequestRaw_(path, method, payload, accessToken, options) {
  const requestOptions = options || {};
  const baseUrl = String(requestOptions.baseUrl || 'https://business-api.tiktok.com/open_api/v1.3/').replace(/\/+$/, '') + '/';
  const normalizedMethod = String(method || 'get').toLowerCase();
  let url = baseUrl + path.replace(/^\//, '');

  const fetchOptions = {
    method: normalizedMethod,
    muteHttpExceptions: true,
    headers: {
      'Access-Token': accessToken,
      'Content-Type': 'application/json'
    }
  };

  if (normalizedMethod === 'get') {
    const qs = buildQueryString_(payload || {});
    if (qs) url += '?' + qs;
  } else if (payload) {
    fetchOptions.payload = JSON.stringify(payload);
  }

  const res = UrlFetchApp.fetch(url, fetchOptions);
  const code = res.getResponseCode();
  const bodyText = res.getContentText();
  let body = {};

  try {
    body = JSON.parse(bodyText);
  } catch (e) {
    body = { raw: bodyText };
  }

  return {
    code: code,
    body: body,
    bodyText: bodyText,
    requestUrl: url,
    requestMethod: normalizedMethod
  };
}

function tiktokApiRequestWithAuthRetry_(path, method, payload, options) {
  const cfg = getTikTokConfig_();
  let response = tiktokApiRequestRaw_(path, method, payload, cfg.accessToken, options);

  const apiCode = response.body && response.body.code !== undefined ? toNumber_(response.body.code) : 0;
  const failed = response.code < 200 || response.code >= 300 || apiCode !== 0;

  if (!failed) {
    return response.body;
  }

  const authError = isTikTokAuthError_(response.code, response.bodyText) || isTikTokAuthError_(response.code, response.body && response.body.message);
  if (!authError) {
    throw new Error('TikTok API error ' + response.code + ' [' + response.requestMethod.toUpperCase() + ' ' + response.requestUrl + ']: ' + response.bodyText);
  }

  if (!tiktokHasRefreshConfig_(cfg)) {
    log_('TikTok auth failed', 'Refresh is not configured. Set TIKTOK_REFRESH_TOKEN, TIKTOK_APP_ID, TIKTOK_SECRET.');
    throw new Error('TikTok auth failed and refresh is not configured.');
  }

  const refreshedToken = refreshTikTokAccessToken_();
  response = tiktokApiRequestRaw_(path, method, payload, refreshedToken, options);
  const retryCode = response.body && response.body.code !== undefined ? toNumber_(response.body.code) : 0;
  const retryFailed = response.code < 200 || response.code >= 300 || retryCode !== 0;

  if (retryFailed) {
    throw new Error('TikTok API error after token refresh ' + response.code + ' [' + response.requestMethod.toUpperCase() + ' ' + response.requestUrl + ']: ' + response.bodyText);
  }

  return response.body;
}

function fetchTikTokReportBodyWithFallback_(payload) {
  const attempts = [
    { method: 'get', path: '/report/integrated/get/', baseUrl: 'https://business-api.tiktok.com/open_api/v1.3/' },
    { method: 'get', path: '/report/integrated/get', baseUrl: 'https://business-api.tiktok.com/open_api/v1.3/' },
    { method: 'post', path: '/report/integrated/get/', baseUrl: 'https://business-api.tiktok.com/open_api/v1.3/' },
    { method: 'post', path: '/report/integrated/get', baseUrl: 'https://business-api.tiktok.com/open_api/v1.3/' },
    { method: 'get', path: '/report/integrated/get/', baseUrl: 'https://ads.tiktok.com/open_api/v1.3/' },
    { method: 'post', path: '/report/integrated/get/', baseUrl: 'https://ads.tiktok.com/open_api/v1.3/' }
  ];

  let lastError = null;
  for (let i = 0; i < attempts.length; i++) {
    const attempt = attempts[i];
    try {
      return tiktokApiRequestWithAuthRetry_(attempt.path, attempt.method, payload, { baseUrl: attempt.baseUrl });
    } catch (e) {
      lastError = e;
      if (String(e && e.message || '').indexOf('TikTok API error 405') === -1) {
        throw e;
      }
    }
  }

  throw lastError || new Error('TikTok report request failed for all fallback attempts.');
}

function normalizeTikTokDate_(value) {
  if (value === null || value === undefined || value === '') return '';

  if (typeof value === 'number') {
    const ms = value > 1000000000000 ? value : value * 1000;
    return formatDate_(new Date(ms));
  }

  const asNumber = Number(value);
  if (!isNaN(asNumber) && String(value).trim() !== '') {
    const ms = asNumber > 1000000000000 ? asNumber : asNumber * 1000;
    return formatDate_(new Date(ms));
  }

  return formatDate_(value);
}

function mapTikTokEntityRow_(adgroup) {
  const advertiserId = normalizeId_(adgroup.advertiser_id);
  const campaignName = normalizeId_(adgroup.campaign_name);
  const adgroupName = normalizeId_(adgroup.adgroup_name);

  return [
    'tiktok',
    advertiserId,
    'adgroup',
    normalizeId_(adgroup.adgroup_id),
    campaignName + ' | ' + adgroupName,
    normalizeId_(adgroup.campaign_id),
    campaignName,
    normalizeId_(adgroup.adgroup_id),
    adgroupName,
    normalizeTikTokDate_(adgroup.schedule_start_time),
    normalizeTikTokDate_(adgroup.schedule_end_time),
    normalizeId_(adgroup.operation_status),
    normalizeId_(adgroup.optimization_goal),
    ''
  ];
}

function listTikTokAdgroups_() {
  const cfg = getTikTokConfig_();
  const out = [];
  let page = 1;
  const pageSize = 100;

  while (true) {
    const body = tiktokApiRequestWithAuthRetry_('/adgroup/get/', 'get', {
      advertiser_id: cfg.advertiserId,
      page: page,
      page_size: pageSize
    });

    const data = body.data || {};
    const list = data.list || [];
    list.forEach(function (item) {
      out.push(item);
    });

    const pageInfo = data.page_info || {};
    const totalPage = toNumber_(pageInfo.total_page);
    if (!list.length || (totalPage > 0 && page >= totalPage)) {
      break;
    }

    if (totalPage === 0 && list.length < pageSize) {
      break;
    }

    page += 1;
  }

  return out;
}

function loadTikTokEntities() {
  withErrorLogging_('loadTikTokEntities failed', function () {
    ensureHeader_(SHEETS.CAMPAIGNS_ENABLED, HEADERS.CAMPAIGNS_ENABLED);
    const keep = readObjects_(SHEETS.CAMPAIGNS_ENABLED).filter(function (r) {
      return normalizePlatform_(r.platform) !== 'tiktok';
    });

    clearDataKeepHeader_(SHEETS.CAMPAIGNS_ENABLED);

    const out = keep.map(function (r) {
      return mapCampaignEnabledRow_(r);
    });

    const adgroups = listTikTokAdgroups_();
    adgroups.forEach(function (adgroup) {
      if (String(adgroup.operation_status || '').toUpperCase() !== 'ENABLE') return;
      out.push(mapTikTokEntityRow_(adgroup));
    });

    if (out.length) appendRows_(SHEETS.CAMPAIGNS_ENABLED, out);
    sortCampaignsEnabled_();
    updateCampaignsEnabledPlanStatusForPlatform_('tiktok');
  });
}

function getTikTokMetricMap_() {
  return {
    impressions: ['impressions', 'show_cnt'],
    reach: ['reach', 'unique_users', 'unique_viewers'],
    cpm: ['cpm', 'cost_per_1000_impression'],
    video_p25_count: ['video_watched_25pct', 'video_play_actions_pct_25', 'video_views_p25'],
    video_p50_count: ['video_watched_50pct', 'video_play_actions_pct_50', 'video_views_p50'],
    video_p75_count: ['video_watched_75pct', 'video_play_actions_pct_75', 'video_views_p75'],
    video_p100_count: ['video_watched_100pct', 'video_play_actions_pct_100', 'video_views_p100']
  };
}

function readTikTokMetricValue_(container, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const key = candidates[i];
    if (container[key] !== undefined && container[key] !== null && container[key] !== '') {
      return toNumber_(container[key]);
    }
  }
  return 0;
}

function findTikTokAdgroupById_(advertiserId, adgroupId) {
  const body = tiktokApiRequestWithAuthRetry_('/adgroup/get/', 'get', {
    advertiser_id: advertiserId,
    filtering: JSON.stringify({ adgroup_ids: [String(adgroupId)] }),
    page: 1,
    page_size: 1
  });

  const list = (body.data && body.data.list) ? body.data.list : [];
  return list.length ? list[0] : null;
}

function fetchTikTokReportRow_(advertiserId, adgroupId, startDate, endDate) {
  const metricMap = getTikTokMetricMap_();
  const metrics = [
    metricMap.impressions[0],
    metricMap.reach[0],
    metricMap.cpm[0],
    metricMap.video_p25_count[0],
    metricMap.video_p50_count[0],
    metricMap.video_p75_count[0],
    metricMap.video_p100_count[0]
  ];

  const payload = {
    advertiser_id: advertiserId,
    data_level: 'AUCTION_ADGROUP',
    dimensions: ['adgroup_id'],
    metrics: metrics,
    start_date: startDate,
    end_date: endDate,
    page: 1,
    page_size: 1000,
    filters: [{ field_name: 'adgroup_ids', filter_type: 'IN', filter_value: [String(adgroupId)] }]
  };

  const body = fetchTikTokReportBodyWithFallback_(payload);
  const list = body.data && body.data.list ? body.data.list : [];

  for (let i = 0; i < list.length; i++) {
    const row = list[i] || {};
    const dimensions = row.dimensions || {};
    if (normalizeId_(dimensions.adgroup_id) === normalizeId_(adgroupId) || normalizeId_(row.adgroup_id) === normalizeId_(adgroupId)) {
      return row;
    }
  }

  return null;
}

function fetchTikTokEntityMetrics_(entityId, accountId, maybeEntity) {
  let entity = maybeEntity || {};
  if (typeof entityId === 'object' && entityId !== null) {
    entity = entityId;
    entityId = entity.entity_id;
    accountId = entity.account_id;
  }

  const cfg = getTikTokConfig_();
  const advertiserId = normalizeId_(accountId) || cfg.advertiserId;
  const adgroupId = normalizeId_(entityId || entity.entity_id);
  if (!adgroupId) {
    throw new Error('Missing TikTok adgroup_id');
  }

  const adgroup = findTikTokAdgroupById_(advertiserId, adgroupId) || {};
  const campaignName = normalizeId_(adgroup.campaign_name) || normalizeId_(entity.campaign_name) || parseCompositeEntityName_(entity.entity_name).campaign_name;
  const adgroupName = normalizeId_(adgroup.adgroup_name) || normalizeId_(entity.adset_name) || parseCompositeEntityName_(entity.entity_name).child_name;
  const startDate = normalizeTikTokDate_(adgroup.schedule_start_time) || formatDate_(new Date());

  let endDate = normalizeTikTokDate_(adgroup.schedule_end_time) || formatDate_(new Date());
  const today = formatDate_(new Date());
  if (endDate > today) endDate = today;

  const reportRow = fetchTikTokReportRow_(advertiserId, adgroupId, startDate, endDate) || {};
  const metricMap = getTikTokMetricMap_();
  const sourceMetrics = reportRow.metrics || reportRow;

  const impressions = readTikTokMetricValue_(sourceMetrics, metricMap.impressions);
  const reach = readTikTokMetricValue_(sourceMetrics, metricMap.reach);
  const p25Count = readTikTokMetricValue_(sourceMetrics, metricMap.video_p25_count);
  const p50Count = readTikTokMetricValue_(sourceMetrics, metricMap.video_p50_count);
  const p75Count = readTikTokMetricValue_(sourceMetrics, metricMap.video_p75_count);
  const p100Count = readTikTokMetricValue_(sourceMetrics, metricMap.video_p100_count);

  return {
    platform: 'tiktok',
    account_id: advertiserId,
    entity_level: 'adgroup',
    entity_id: adgroupId,
    entity_name: campaignName + ' | ' + adgroupName,
    campaign_id: normalizeId_(adgroup.campaign_id) || normalizeId_(entity.campaign_id),
    campaign_name: campaignName,
    adset_id: adgroupId,
    adset_name: adgroupName,
    start_date: startDate,
    end_date: endDate,
    impressions: impressions,
    reach: reach,
    frequency: reach > 0 ? impressions / reach : 0,
    cpm: readTikTokMetricValue_(sourceMetrics, metricMap.cpm),
    video_p25: impressions > 0 ? p25Count / impressions : 0,
    video_p50: impressions > 0 ? p50Count / impressions : 0,
    video_p75: impressions > 0 ? p75Count / impressions : 0,
    video_p100: impressions > 0 ? p100Count / impressions : 0,
    status: normalizeId_(adgroup.operation_status) || 'UNKNOWN',
    channel_type: normalizeId_(adgroup.optimization_goal) || normalizeId_(entity.channel_type)
  };
}

function refreshTikTokAccessToken_() {
  const cfg = getTikTokConfig_();
  if (!tiktokHasRefreshConfig_(cfg)) {
    throw new Error('TikTok refresh is not configured.');
  }

  const url = 'https://business-api.tiktok.com/open_api/v1.3/oauth2/refresh_token/';
  const payload = {
    app_id: cfg.appId,
    secret: cfg.secret,
    grant_type: 'refresh_token',
    refresh_token: cfg.refreshToken
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    throw new Error('TikTok token refresh failed ' + res.getResponseCode() + ': ' + res.getContentText());
  }

  const body = JSON.parse(res.getContentText());
  if (toNumber_(body.code) !== 0 || !body.data || !body.data.access_token) {
    throw new Error('TikTok token refresh returned invalid response: ' + res.getContentText());
  }

  const props = getScriptProps_();
  props.setProperty('TIKTOK_ACCESS_TOKEN', String(body.data.access_token));
  if (body.data.refresh_token) {
    props.setProperty('TIKTOK_REFRESH_TOKEN', String(body.data.refresh_token));
  }

  log_('TikTok token refreshed', 'Updated TIKTOK_ACCESS_TOKEN in Script Properties.');
  return String(body.data.access_token);
}
