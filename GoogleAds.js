function getGoogleAccessToken_() {
  const props = getScriptProps();
  const clientId = props.getProperty('GOOGLE_OAUTH_CLIENT_ID');
  const clientSecret = props.getProperty('GOOGLE_OAUTH_CLIENT_SECRET');
  const refreshToken = props.getProperty('GOOGLE_ADS_REFRESH_TOKEN');

  if (
    statusOf_(clientId) !== 'OK' ||
    statusOf_(clientSecret) !== 'OK' ||
    statusOf_(refreshToken) !== 'OK'
  ) {
    throw new Error('Missing OAuth credentials or refresh token in Script Properties.');
  }

  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      client_id: clientId,
      client_secret: clientSecret,
      refresh_token: refreshToken,
      grant_type: 'refresh_token'
    },
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('Failed to refresh Google access token: ' + body);
  }

  return JSON.parse(body).access_token;
}

function googleAdsSearchStream_(query) {
  const props = getScriptProps();
  const accessToken = getGoogleAccessToken_();
  const developerToken = props.getProperty('GOOGLE_ADS_DEVELOPER_TOKEN');
  const customerId = normalizeId_(props.getProperty('GOOGLE_ADS_CUSTOMER_ID'));
  const loginCustomerId = normalizeId_(props.getProperty('GOOGLE_ADS_LOGIN_CUSTOMER_ID'));

  if (statusOf_(developerToken) !== 'OK' || !customerId) {
    throw new Error('Missing Google Ads developer token or customer ID.');
  }

  const url = `https://googleads.googleapis.com/v20/customers/${customerId}/googleAds:searchStream`;

  const headers = {
    Authorization: `Bearer ${accessToken}`,
    'developer-token': developerToken,
    'Content-Type': 'application/json'
  };

  if (loginCustomerId) {
    headers['login-customer-id'] = loginCustomerId;
  }

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: headers,
    payload: JSON.stringify({ query: query }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`Google Ads API error ${code}: ${body}`);
  }

  return JSON.parse(body);
}

function loadEnabledCampaigns() {
  ensureHeader_(SHEETS.CAMPAIGNS_ENABLED, HEADERS.CAMPAIGNS_ENABLED);

  const sh = getSheet_(SHEETS.CAMPAIGNS_ENABLED);
  const existing = readObjects_(SHEETS.CAMPAIGNS_ENABLED).filter(r => String(r.platform).toLowerCase() !== 'google');

  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.CAMPAIGNS_ENABLED.length).setValues([HEADERS.CAMPAIGNS_ENABLED]);
  sh.setFrozenRows(1);

  const query = `
    SELECT
      campaign.id,
      campaign.name,
      campaign.start_date,
      campaign.end_date,
      campaign.status,
      campaign.advertising_channel_type
    FROM campaign
    WHERE campaign.status = ENABLED
    ORDER BY campaign.name
  `;

  const chunks = googleAdsSearchStream_(query);
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

  chunks.forEach(chunk => {
    (chunk.results || []).forEach(r => {
      output.push([
        'google',
        '',
        String(r.campaign?.id || ''),
        r.campaign?.name || '',
        r.campaign?.startDate || '',
        r.campaign?.endDate || '',
        r.campaign?.status || '',
        r.campaign?.advertisingChannelType || ''
      ]);
    });
  });

  if (output.length) {
    appendRows_(SHEETS.CAMPAIGNS_ENABLED, output);
  }

  SpreadsheetApp.getUi().alert('Loaded Google enabled campaigns.');
}

function fetchBaseCampaignMetrics_(campaignId) {
  const safeCampaignId = normalizeId_(campaignId);

  if (!/^\d+$/.test(safeCampaignId)) {
    throw new Error(`Invalid campaign_id: ${campaignId}`);
  }

  const query = `
    SELECT
      campaign.id,
      campaign.name,
      campaign.status,
      campaign.start_date,
      campaign.end_date,
      campaign.advertising_channel_type,
      metrics.impressions,
      metrics.average_cpm,
      metrics.video_quartile_p25_rate,
      metrics.video_quartile_p50_rate,
      metrics.video_quartile_p75_rate,
      metrics.video_quartile_p100_rate
    FROM campaign
    WHERE campaign.id = ${safeCampaignId}
      AND campaign.status = ENABLED
  `;

  const chunks = googleAdsSearchStream_(query);
  let row = null;

  chunks.forEach(chunk => {
    (chunk.results || []).forEach(r => {
      row = {
        platform: 'google',
        account_id: '',
        campaign_id: String(r.campaign?.id || ''),
        campaign_name: r.campaign?.name || '',
        start_date: r.campaign?.startDate || '',
        end_date: r.campaign?.endDate || '',
        impressions: toNumber_(r.metrics?.impressions),
        average_cpm: toNumber_(r.metrics?.averageCpm) / 1000000,
        video_quartile_p25_rate: toNumber_(r.metrics?.videoQuartileP25Rate),
        video_quartile_p50_rate: toNumber_(r.metrics?.videoQuartileP50Rate),
        video_quartile_p75_rate: toNumber_(r.metrics?.videoQuartileP75Rate),
        video_quartile_p100_rate: toNumber_(r.metrics?.videoQuartileP100Rate),
        status: r.campaign?.status || '',
        channel_type: r.campaign?.advertisingChannelType || '',
        reach_or_unique_users: ''
      };
    });
  });

  return row;
}

function isUniqueUsersEligibleChannel_(channelType) {
  const t = String(channelType || '').toUpperCase();
  return ['DISPLAY', 'VIDEO', 'DISCOVERY', 'APP'].indexOf(t) !== -1;
}

function getUniqueUsersWindow_(startDateStr, endDateStr) {
  const today = new Date();
  const start = new Date(startDateStr);
  const end = new Date(endDateStr);

  if (isNaN(start.getTime()) || isNaN(end.getTime())) return null;

  const effectiveEnd = end < today ? end : today;
  const minStart = new Date(effectiveEnd);
  minStart.setDate(minStart.getDate() - 91);

  const effectiveStart = start > minStart ? start : minStart;
  if (effectiveStart > effectiveEnd) return null;

  return {
    start: formatDate_(effectiveStart),
    end: formatDate_(effectiveEnd)
  };
}

function fetchUniqueUsers_(campaignId, startDate, endDate) {
  const safeCampaignId = normalizeId_(campaignId);

  const query = `
    SELECT
      campaign.id,
      metrics.unique_users
    FROM campaign
    WHERE campaign.id = ${safeCampaignId}
      AND campaign.status = ENABLED
      AND segments.date BETWEEN '${startDate}' AND '${endDate}'
  `;

  const chunks = googleAdsSearchStream_(query);
  let uniqueUsers = '';

  chunks.forEach(chunk => {
    (chunk.results || []).forEach(r => {
      uniqueUsers = toNumber_(r.metrics?.uniqueUsers);
    });
  });

  return uniqueUsers;
}

function fetchGoogleCampaignMetrics_(campaignId) {
  const base = fetchBaseCampaignMetrics_(campaignId);
  if (!base) return null;

  if (!isUniqueUsersEligibleChannel_(base.channel_type)) {
    return base;
  }

  const window = getUniqueUsersWindow_(base.start_date, base.end_date);
  if (!window) return base;

  try {
    base.reach_or_unique_users = fetchUniqueUsers_(campaignId, window.start, window.end);
  } catch (e) {
    log_(
      'unique_users fallback',
      `campaign_id=${campaignId}; channel=${base.channel_type}; window=${window.start}..${window.end}; reason=${e.message}`
    );
    base.reach_or_unique_users = '';
  }

  return base;
}

function refreshRawAllFromPlan_(platform) {
  ensureHeader_(SHEETS.RAW_ALL, HEADERS.RAW_ALL);

  const sh = getSheet_(SHEETS.RAW_ALL);
  const existing = readObjects_(SHEETS.RAW_ALL).filter(
    r => String(r.platform).toLowerCase() !== String(platform).toLowerCase()
  );

  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.RAW_ALL.length).setValues([HEADERS.RAW_ALL]);
  sh.setFrozenRows(1);

  const output = [];

  existing.forEach(r => {
    output.push([
      r.platform,
      r.account_id,
      r.campaign_id,
      r.campaign_name,
      r.start_date,
      r.end_date,
      r.goal_impressions,
      r.goal_reach,
      r.impressions,
      r.reach_or_unique_users,
      r.frequency,
      r.average_cpm,
      r.video_quartile_p25_rate,
      r.video_quartile_p50_rate,
      r.video_quartile_p75_rate,
      r.video_quartile_p100_rate,
      r.status,
      r.channel_type
    ]);
  });

  const planRows = readObjects_(SHEETS.PLAN).filter(r => {
    const rowPlatform = String(r.platform || '').toLowerCase().trim();
    return rowPlatform === String(platform).toLowerCase();
  });

  if (!planRows.length) {
    log_('RAW_ALL', `No PLAN rows found for platform=${platform}`);
    if (output.length) appendRows_(SHEETS.RAW_ALL, output);
    return;
  }

  planRows.forEach(p => {
    const campaignId = normalizeId_(p.campaign_id);
    const campaignName = p.campaign_name || '';
    const accountId = p.account_id || '';
    const goalImpressions = toNumber_(p.goal_impressions);
    const goalReach = toNumber_(p.goal_reach);

    if (!campaignId) {
      log_('Skipped PLAN row', `Missing campaign_id for ${campaignName}`);
      return;
    }

    if (platform === 'google' && !/^\d+$/.test(campaignId)) {
      log_('Skipped PLAN row', `Invalid Google campaign_id=${campaignId} for ${campaignName}`);
      return;
    }

    try {
      let metrics = null;

      if (platform === 'google') {
        metrics = fetchGoogleCampaignMetrics_(campaignId);
      } else if (platform === 'meta') {
        metrics = fetchMetaCampaignMetrics_(campaignId, accountId);
      }

      if (!metrics) {
        output.push([
          platform,
          accountId,
          campaignId,
          campaignName,
          '',
          '',
          goalImpressions,
          goalReach,
          0,
          '',
          '',
          '',
          '',
          '',
          '',
          '',
          'NO_DATA',
          ''
        ]);
        return;
      }

      const reachValue =
        metrics.reach_or_unique_users === '' ? '' : toNumber_(metrics.reach_or_unique_users);
      const frequency =
        reachValue !== '' && reachValue > 0 ? toNumber_(metrics.impressions) / reachValue : '';

      output.push([
        platform,
        metrics.account_id || accountId,
        metrics.campaign_id || campaignId,
        metrics.campaign_name || campaignName,
        metrics.start_date,
        metrics.end_date,
        goalImpressions,
        goalReach,
        metrics.impressions,
        reachValue,
        frequency,
        metrics.average_cpm,
        metrics.video_quartile_p25_rate,
        metrics.video_quartile_p50_rate,
        metrics.video_quartile_p75_rate,
        metrics.video_quartile_p100_rate,
        metrics.status || 'ENABLED',
        metrics.channel_type || ''
      ]);
    } catch (e) {
      log_('Campaign fetch failed', `platform=${platform}; account_id=${accountId}; campaign_id=${campaignId}; ${e.message}`);
      output.push([
        platform,
        accountId,
        campaignId,
        campaignName,
        '',
        '',
        goalImpressions,
        goalReach,
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        'ERROR',
        ''
      ]);
    }
  });

  if (output.length) {
    appendRows_(SHEETS.RAW_ALL, output);
  }

  log_('RAW_ALL', `Loaded ${output.length} rows for platform=${platform}`);
}