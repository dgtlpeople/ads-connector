function refreshLast30AdHierarchyAll_() {
  ensureHeader_(SHEETS.AD_HIERARCHY_LAST_30, HEADERS.AD_HIERARCHY_LAST_30);
  clearDataKeepHeader_(SHEETS.AD_HIERARCHY_LAST_30);

  const dateRange = getLastNDaysDateRange_(30);
  const rows = [];
  const succeeded = [];
  const failed = [];
  const platforms = ['google', 'meta', 'tiktok'];

  platforms.forEach(function (platform) {
    try {
      let platformRows = [];
      if (platform === 'google') {
        platformRows = fetchGoogleLast30AdHierarchyRows_(dateRange);
      } else if (platform === 'meta') {
        platformRows = fetchMetaLast30AdHierarchyRows_(dateRange);
      } else if (platform === 'tiktok') {
        platformRows = fetchTikTokLast30AdHierarchyRows_(dateRange);
      }

      platformRows.forEach(function (row) {
        rows.push(mapLast30AdHierarchyRow_(row));
      });
      succeeded.push(platform + ' (' + platformRows.length + ')');
    } catch (e) {
      const reason = e && e.message ? e.message : String(e);
      failed.push(platform + ': ' + reason);
      log_('Last 30d hierarchy export failed', 'platform=' + platform + '; reason=' + reason);
    }
  });

  if (rows.length) {
    appendRows_(SHEETS.AD_HIERARCHY_LAST_30, rows);
  }

  log_(
    'Last 30d hierarchy export',
    'date_start=' + dateRange.start + '; date_end=' + dateRange.end +
    '; rows=' + rows.length + '; success=' + succeeded.join(', ') +
    '; failed=' + failed.join(' | ')
  );

  return {
    date_start: dateRange.start,
    date_end: dateRange.end,
    rows: rows.length,
    succeeded: succeeded,
    failed: failed
  };
}

function mapLast30AdHierarchyRow_(row) {
  return [
    row.platform || '',
    row.account_id || '',
    row.date_start || '',
    row.date_end || '',
    row.entity_level || 'ad',
    row.campaign_id || '',
    row.campaign_name || '',
    row.campaign_status || '',
    row.child_level || '',
    row.child_id || '',
    row.child_name || '',
    row.child_status || '',
    row.ad_id || '',
    row.ad_name || '',
    row.ad_status || '',
    row.channel_type || '',
    toNumber_(row.impressions),
    row.reach === '' ? '' : toNumber_(row.reach),
    toNumber_(row.cost)
  ];
}
