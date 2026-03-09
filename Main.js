function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ads Connector')
    .addItem('Setup sheets', 'setupSheets')
    .addItem('Reset PLAN header', 'resetPlanHeader')
    .addItem('Check config', 'checkAllConfigs')
    .addSeparator()
    .addItem('Load Google enabled campaigns', 'loadEnabledCampaigns')
    .addItem('Load Meta enabled campaigns', 'loadMetaEnabledCampaigns')
    .addSeparator()
    .addItem('Refresh RAW_ALL from PLAN (Google)', 'refreshGoogleFromPlan')
    .addItem('Refresh RAW_ALL from PLAN (Meta)', 'refreshMetaFromPlan')
    .addItem('Build SUMMARY', 'buildSummary')
    .addItem('Run full pipeline', 'runFullPipeline')
    .addToUi();
}

function setupSheets() {
  ensureHeader_(SHEETS.CAMPAIGNS_ENABLED, HEADERS.CAMPAIGNS_ENABLED);
  ensureHeader_(SHEETS.PLAN, HEADERS.PLAN);
  ensureHeader_(SHEETS.RAW_ALL, HEADERS.RAW_ALL);
  ensureHeader_(SHEETS.SUMMARY, HEADERS.SUMMARY);
  ensureHeader_(SHEETS.LOG, HEADERS.LOG);

  SpreadsheetApp.getUi().alert('Sheets created successfully.');
}

function resetPlanHeader() {
  const sh = getSheet_(SHEETS.PLAN);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.PLAN.length).setValues([HEADERS.PLAN]);
  sh.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('PLAN reset. Paste only: platform, account_id, campaign_id, campaign_name, goal_impressions, goal_reach');
}

function checkAllConfigs() {
  const props = getScriptProps();

  const keys = [
    'GOOGLE_ADS_DEVELOPER_TOKEN',
    'GOOGLE_ADS_CUSTOMER_ID',
    'GOOGLE_ADS_LOGIN_CUSTOMER_ID',
    'GOOGLE_OAUTH_CLIENT_ID',
    'GOOGLE_OAUTH_CLIENT_SECRET',
    'GOOGLE_ADS_REFRESH_TOKEN',
    'META_ACCESS_TOKEN',
    'META_AD_ACCOUNT_IDS'
  ];

  const result = keys.map(k => [k, statusOf_(props.getProperty(k))]);

  const sh = getSheet_('CONFIG_CHECK');
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['key', 'status']]);
  sh.getRange(2, 1, result.length, 2).setValues(result);

  SpreadsheetApp.getUi().alert('Config check written to CONFIG_CHECK.');
}

function refreshGoogleFromPlan() {
  refreshRawAllFromPlan_('google');
  SpreadsheetApp.getUi().alert('RAW_ALL updated for Google.');
}

function refreshMetaFromPlan() {
  refreshRawAllFromPlan_('meta');
  SpreadsheetApp.getUi().alert('RAW_ALL updated for Meta.');
}

function runFullPipeline() {
  refreshRawAllFromPlan_('google');
  refreshRawAllFromPlan_('meta');
  buildSummary();
  SpreadsheetApp.getUi().alert('Pipeline finished.');
}