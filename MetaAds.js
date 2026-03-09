function getMetaConfig_() {
  const props = getScriptProps();
  const accessToken = props.getProperty('META_ACCESS_TOKEN');
  const adAccountIdsRaw = props.getProperty('META_AD_ACCOUNT_IDS');

  if (statusOf_(accessToken) !== 'OK' || statusOf_(adAccountIdsRaw) !== 'OK') {
    throw new Error('Missing META_ACCESS_TOKEN or META_AD_ACCOUNT_IDS in Script Properties.');
  }

  const adAccountIds = String(adAccountIdsRaw)
    .split(',')
    .map(x => x.trim())
    .filter(Boolean);

  return {
    accessToken: accessToken,
    adAccountIds: adAccountIds
  };
}

function metaApiGet_(path, params) {
  const cfg = getMetaConfig_();
  const baseUrl = 'https://graph.facebook.com/v21.0/' + path;

  const query = Object.assign({}, params || {}, {
    access_token: cfg.accessToken
  });

  const parts = [];
  Object.keys(query).forEach(key => {
    const value = query[key];
    if (value === undefined || value === null || value === '') return;
    parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(value)));
  });

  const url = baseUrl + '?' + parts.join('&');

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('Meta API error ' + code + ': ' + body);
  }

  return JSON.parse(body);
}

function metaApiGetAllPages_(path, params) {
  const cfg = getMetaConfig_();
  let nextUrl = 'https://graph.facebook.com/v21.0/' + path;
  const query = Object.assign({}, params || {}, {
    access_token: cfg.accessToken
  });

  const parts = [];
  Object.keys(query).forEach(key => {
    const value = query[key];
    if (value === undefined || value === null || value === '') return;
    parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(value)));
  });

  nextUrl += '?' + parts.join('&');

  const all = [];

  while (nextUrl) {
    const res = UrlFetchApp.fetch(nextUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code < 200 || code >= 300) {
      throw new Error('Meta API error ' + code + ': ' + body);
    }

    const json = JSON.parse(body);
    (json.data || []).forEach(item => all.push(item));

    nextUrl = json.paging && json.paging.next ? json.paging.next : null;
  }

  return all;
}

function toMetaDateOnly_(value) {
  if (!value) return '';
  return formatDate_(value);
}

function loadMetaEnabledCampaigns() {
  ensureHeader_(SHEETS.CAMPAIGNS_ENABLED, HEADERS.CAMPAIGNS_ENABLED);

  const sh = getSheet_(SHEETS.CAMPAIGNS_ENABLED);
  const existing = readObjects_(SHEETS.CAMPAIGNS_ENABLED).filter(
    r => String(r.platform).toLowerCase() !== 'meta'
  );

  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.CAMPAIGNS_ENABLED.length).setValues([HEADERS.CAMPAIGNS_ENABLED]);
  sh.setFrozenRows(1);

  const cfg = getMetaConfig_();
  const output = [];

  existing.forEach(r => {
    output.push([
      r.platform,
      r.account_id,
      r.campaign_id,
      r.campaign_name,
      r.start_date,
      r.end_date,
      r.status,
      r.channel_type
    ]);
  });

  cfg.adAccountIds.forEach(accountId => {
    const campaigns = metaApiGetAllPages_(accountId + '/campaigns', {
      fields: 'id,name,status,effective_status,objective,start_time,stop_time',
      effective_status: '["ACTIVE"]',
      limit: 200
    });

    campaigns.forEach(c => {
      output.push([
        'meta',
        accountId,
        String(c.id || ''),
        c.name || '',
        toMetaDateOnly_(c.start_time),
        toMetaDateOnly_(c.stop_time),
        c.effective_status || c.status || '',
        c.objective || ''
      ]);
    });
  });

  if (output.length) {
    appendRows_(SHEETS.CAMPAIGNS_ENABLED, output);
  }

  SpreadsheetApp.getUi().alert('Loaded Meta enabled campaigns.');
}

function getMetaCampaignObject_(campaignId) {
  const campaign = metaApiGet_(String(campaignId), {
    fields: 'id,name,status,effective_status,objective,start_time,stop_time'
  });

  return {
    platform: 'meta',
    account_id: '',
    campaign_id: String(campaign.id || ''),
    campaign_name: campaign.name || '',
    start_date: toMetaDateOnly_(campaign.start_time),
    end_date: toMetaDateOnly_(campaign.stop_time),
    status: campaign.effective_status || campaign.status || '',
    channel_type: campaign.objective || '',
    impressions: 0,
    average_cpm: '',
    video_quartile_p25_rate: '',
    video_quartile_p50_rate: '',
    video_quartile_p75_rate: '',
    video_quartile_p100_rate: '',
    reach_or_unique_users: ''
  };
}

function getMetaInsightsForCampaign_(campaignId, startDate, endDate) {
  const timeRange = JSON.stringify({
    since: startDate,
    until: endDate
  });

  const data = metaApiGet_(String(campaignId) + '/insights', {
    fields: [
      'campaign_id',
      'campaign_name',
      'impressions',
      'reach',
      'cpm',
      'video_p25_watched_actions',
      'video_p50_watched_actions',
      'video_p75_watched_actions',
      'video_p100_watched_actions'
    ].join(','),
    time_range: timeRange,
    level: 'campaign',
    limit: 50
  });

  if (!data.data || !data.data.length) {
    return null;
  }

  const row = data.data[0];

  function actionValue_(arr) {
    if (!Array.isArray(arr) || !arr.length) return 0;
    const first = arr[0];
    return toNumber_(first.value);
  }

  const impressions = toNumber_(row.impressions);
  const p25 = actionValue_(row.video_p25_watched_actions);
  const p50 = actionValue_(row.video_p50_watched_actions);
  const p75 = actionValue_(row.video_p75_watched_actions);
  const p100 = actionValue_(row.video_p100_watched_actions);

  return {
    impressions: impressions,
    reach_or_unique_users: toNumber_(row.reach),
    average_cpm: toNumber_(row.cpm),
    video_quartile_p25_rate: impressions > 0 ? p25 / impressions : '',
    video_quartile_p50_rate: impressions > 0 ? p50 / impressions : '',
    video_quartile_p75_rate: impressions > 0 ? p75 / impressions : '',
    video_quartile_p100_rate: impressions > 0 ? p100 / impressions : ''
  };
}

function fetchMetaCampaignMetrics_(campaignId, accountId) {
  const base = getMetaCampaignObject_(campaignId);
  if (!base || !base.campaign_id) return null;

  base.account_id = accountId || '';

  const today = new Date();
  const todayStr = formatDate_(today);

  let startDate = base.start_date || todayStr;
  let endDate = base.end_date || todayStr;

  if (endDate > todayStr) {
    endDate = todayStr;
  }

  if (startDate > endDate) {
    startDate = endDate;
  }

  const insights = getMetaInsightsForCampaign_(campaignId, startDate, endDate);
  if (!insights) {
    return base;
  }

  base.impressions = insights.impressions;
  base.reach_or_unique_users = insights.reach_or_unique_users;
  base.average_cpm = insights.average_cpm;
  base.video_quartile_p25_rate = insights.video_quartile_p25_rate;
  base.video_quartile_p50_rate = insights.video_quartile_p50_rate;
  base.video_quartile_p75_rate = insights.video_quartile_p75_rate;
  base.video_quartile_p100_rate = insights.video_quartile_p100_rate;

  return base;
}