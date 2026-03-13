function getMetaConfig_() {
  const props = getScriptProps_();
  const accessToken = props.getProperty('META_ACCESS_TOKEN');
  const rawAccountIds = props.getProperty('META_AD_ACCOUNT_IDS');

  if (!isConfigured_(accessToken) || !isConfigured_(rawAccountIds)) {
    throw new Error('Missing META_ACCESS_TOKEN or META_AD_ACCOUNT_IDS in Script Properties.');
  }

  const adAccountIds = String(rawAccountIds)
    .split(',')
    .map(function (x) {
      return x.trim();
    })
    .filter(function (x) {
      return x;
    })
    .map(function (x) {
      return x.indexOf('act_') === 0 ? x : 'act_' + x;
    });

  return {
    accessToken: accessToken,
    adAccountIds: adAccountIds
  };
}

function metaApiGet_(path, params) {
  const cfg = getMetaConfig_();
  const query = Object.assign({}, params || {}, { access_token: cfg.accessToken });
  const qs = Object.keys(query)
    .filter(function (k) {
      return query[k] !== undefined && query[k] !== null && query[k] !== '';
    })
    .map(function (k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(String(query[k]));
    })
    .join('&');

  const url = 'https://graph.facebook.com/v21.0/' + path + (qs ? '?' + qs : '');
  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });

  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    throw new Error('Meta API error ' + res.getResponseCode() + ': ' + res.getContentText());
  }

  return JSON.parse(res.getContentText());
}

function metaApiGetAllPages_(path, params) {
  const cfg = getMetaConfig_();
  const merged = Object.assign({}, params || {}, { access_token: cfg.accessToken });
  const qs = Object.keys(merged)
    .filter(function (k) {
      return merged[k] !== undefined && merged[k] !== null && merged[k] !== '';
    })
    .map(function (k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(String(merged[k]));
    })
    .join('&');

  let url = 'https://graph.facebook.com/v21.0/' + path + (qs ? '?' + qs : '');
  const out = [];

  while (url) {
    const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
      throw new Error('Meta API error ' + res.getResponseCode() + ': ' + res.getContentText());
    }

    const body = JSON.parse(res.getContentText());
    (body.data || []).forEach(function (item) {
      out.push(item);
    });
    url = body.paging && body.paging.next ? body.paging.next : '';
  }

  return out;
}

function loadMetaEntities() {
  withErrorLogging_('loadMetaEntities failed', function () {
    ensureHeader_(SHEETS.CAMPAIGNS_ENABLED, HEADERS.CAMPAIGNS_ENABLED);
    const keep = readObjects_(SHEETS.CAMPAIGNS_ENABLED).filter(function (r) {
      return normalizePlatform_(r.platform) !== 'meta';
    });

    clearDataKeepHeader_(SHEETS.CAMPAIGNS_ENABLED);

    const out = [];
    keep.forEach(function (r) {
      out.push(mapCampaignEnabledRow_(r));
    });

    const cfg = getMetaConfig_();
    cfg.adAccountIds.forEach(function (accountId) {
      const adsets = metaApiGetAllPages_(accountId + '/adsets', {
        fields: 'id,name,campaign_id,start_time,end_time,effective_status,status,optimization_goal,campaign{id,name,objective}',
        filtering: JSON.stringify([{ field: 'effective_status', operator: 'IN', value: ['ACTIVE'] }]),
        limit: 200
      });

      adsets.forEach(function (adset) {
        const campaign = adset.campaign || {};
        const campaignName = campaign.name || '';
        const adsetName = adset.name || '';

        out.push([
          'meta',
          accountId,
          'adset',
          String(adset.id || ''),
          campaignName + ' | ' + adsetName,
          String(campaign.id || adset.campaign_id || ''),
          campaignName,
          String(adset.id || ''),
          adsetName,
          formatDate_(adset.start_time),
          formatDate_(adset.end_time),
          adset.effective_status || adset.status || '',
          campaign.objective || adset.optimization_goal || ''
        ]);
      });
    });

    if (out.length) appendRows_(SHEETS.CAMPAIGNS_ENABLED, out);
    sortCampaignsEnabled_();
  });
}

function fetchMetaEntityMetrics_(entity) {
  if (normalizeEntityLevel_(entity.entity_level) !== 'adset') {
    throw new Error('Meta Ads currently supports ad set level only.');
  }

  const adsetId = normalizeId_(entity.entity_id);
  if (!adsetId) {
    throw new Error('Missing Meta ad set ID');
  }

  const adset = metaApiGet_(adsetId, {
    fields: 'id,name,campaign_id,start_time,end_time,effective_status,status,optimization_goal,campaign{id,name,objective}'
  });

  const yesterday = getYesterdayDateKey_();
  let since = formatDate_(adset.start_time) || yesterday;
  let until = formatDate_(adset.end_time) || yesterday;
  if (until > yesterday) until = yesterday;
  if (since > until) since = until;

  const insights = metaApiGet_(adsetId + '/insights', {
    level: 'adset',
    fields: [
      'impressions',
      'reach',
      'cpm',
      'video_p25_watched_actions',
      'video_p50_watched_actions',
      'video_p75_watched_actions',
      'video_p100_watched_actions'
    ].join(','),
    time_range: JSON.stringify({ since: since, until: until }),
    limit: 1
  });

  const row = insights.data && insights.data.length ? insights.data[0] : {};
  const impressions = toNumber_(row.impressions);
  const reach = toNumber_(row.reach);

  function actionValue_(arr) {
    if (!Array.isArray(arr) || !arr.length) return 0;
    return toNumber_(arr[0].value);
  }

  const p25Count = actionValue_(row.video_p25_watched_actions);
  const p50Count = actionValue_(row.video_p50_watched_actions);
  const p75Count = actionValue_(row.video_p75_watched_actions);
  const p100Count = actionValue_(row.video_p100_watched_actions);

  const campaign = adset.campaign || {};
  const campaignName = campaign.name || normalizeId_(entity.campaign_name);
  const adsetName = adset.name || normalizeId_(entity.adset_name);

  return {
    platform: 'meta',
    account_id: normalizeId_(entity.account_id),
    entity_level: 'adset',
    entity_id: String(adset.id || adsetId),
    entity_name: campaignName + ' | ' + adsetName,
    campaign_id: String(campaign.id || adset.campaign_id || normalizeId_(entity.campaign_id)),
    campaign_name: campaignName,
    adset_id: String(adset.id || adsetId),
    adset_name: adsetName,
    start_date: formatDate_(adset.start_time),
    end_date: formatDate_(adset.end_time),
    impressions: impressions,
    reach: reach,
    frequency: reach > 0 ? impressions / reach : 0,
    cpm: toNumber_(row.cpm),
    video_p25: impressions > 0 ? p25Count / impressions : 0,
    video_p50: impressions > 0 ? p50Count / impressions : 0,
    video_p75: impressions > 0 ? p75Count / impressions : 0,
    video_p100: impressions > 0 ? p100Count / impressions : 0,
    status: adset.effective_status || adset.status || '',
    channel_type: campaign.objective || adset.optimization_goal || ''
  };
}
