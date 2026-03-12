function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ads Connector')
    .addItem('Setup sheets', 'setupSheets')
    .addItem('Reset PLAN header', 'resetPlanHeader')
    .addItem('Check config', 'checkConfig')
    .addSeparator()
    .addItem('Load Google entities', 'loadGoogleEntities')
    .addItem('Load Meta entities', 'loadMetaEntities')
    .addSeparator()
    .addItem('Refresh RAW_ALL from PLAN (Google)', 'refreshRawAllFromPlanGoogle')
    .addItem('Refresh RAW_ALL from PLAN (Meta)', 'refreshRawAllFromPlanMeta')
    .addItem('Build SUMMARY', 'buildSummary')
    .addItem('Build DASHBOARD', 'buildDashboard')
    .addItem('Generate VIDEO Ads Script', 'generateVideoAdsScript')
    .addItem('Setup Dashboard Action Trigger', 'setupDashboardActionTrigger')
    .addItem('Run full pipeline', 'runFullPipeline')
    .addToUi();
}

function onEdit(e) {
  if (!e || e.authMode !== ScriptApp.AuthMode.FULL) return;
  handleDashboardActionEdit_(e);
}

function dashboardActionOnEdit(e) {
  handleDashboardActionEdit_(e);
}

function setupSheets() {
  withErrorLogging_('setupSheets failed', function () {
    ensureHeader_(SHEETS.CAMPAIGNS_ENABLED, HEADERS.CAMPAIGNS_ENABLED);
    ensureHeader_(SHEETS.PLAN, HEADERS.PLAN);
    ensureHeader_(SHEETS.RAW_ALL, HEADERS.RAW_ALL);
    ensureHeader_(SHEETS.SUMMARY, HEADERS.SUMMARY);
    ensureHeader_(SHEETS.LOG, HEADERS.LOG);
    ensureHeader_(SHEETS.GOOGLE_CHANGES_LOG, HEADERS.GOOGLE_CHANGES_LOG);
    ensureHeader_(SHEETS.VIDEO_ACTION_QUEUE, HEADERS.VIDEO_ACTION_QUEUE);
    ensureHeader_(SHEETS.REACH_CACHE, HEADERS.REACH_CACHE);
    ensureReachCacheSampleRow_();
    SpreadsheetApp.getUi().alert('Sheets created/validated.');
  });
}

function resetPlanHeader() {
  withErrorLogging_('resetPlanHeader failed', function () {
    const sh = getSheet_(SHEETS.PLAN);
    sh.clear();
    sh.getRange(1, 1, 1, HEADERS.PLAN.length).setValues([HEADERS.PLAN]);
    sh.setFrozenRows(1);
    SpreadsheetApp.getUi().alert('PLAN header reset.');
  });
}

function checkConfig() {
  withErrorLogging_('checkConfig failed', function () {
    ensureHeader_(SHEETS.LOG, HEADERS.LOG);
    const sh = getSheet_('CONFIG_CHECK');
    sh.clear();
    sh.getRange(1, 1, 1, 2).setValues([['key', 'status']]);

    const props = getScriptProps_();
    const rows = SCRIPT_PROPERTY_KEYS.map(function (k) {
      return [k, isConfigured_(props.getProperty(k)) ? 'OK' : 'MISSING'];
    });

    if (rows.length) {
      sh.getRange(2, 1, rows.length, 2).setValues(rows);
    }

    SpreadsheetApp.getUi().alert('Config check written to CONFIG_CHECK.');
  });
}

function refreshRawAllFromPlanGoogle() {
  withErrorLogging_('refreshRawAllFromPlanGoogle failed', function () {
    refreshRawAllFromPlan_('google');
    SpreadsheetApp.getUi().alert('RAW_ALL refreshed for Google.');
  });
}

function refreshRawAllFromPlanMeta() {
  withErrorLogging_('refreshRawAllFromPlanMeta failed', function () {
    refreshRawAllFromPlan_('meta');
    SpreadsheetApp.getUi().alert('RAW_ALL refreshed for Meta.');
  });
}

function runFullPipeline() {
  withErrorLogging_('runFullPipeline failed', function () {
    loadGoogleEntities();
    loadMetaEntities();
    refreshRawAllFromPlan_('google');
    refreshRawAllFromPlan_('meta');
    buildSummary();
    buildDashboard();
    SpreadsheetApp.getUi().alert('Full pipeline complete.');
  });
}

function setupDashboardActionTrigger() {
  withErrorLogging_('setupDashboardActionTrigger failed', function () {
    const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const existing = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getHandlerFunction() === 'dashboardActionOnEdit';
    });

    if (!existing) {
      ScriptApp.newTrigger('dashboardActionOnEdit')
        .forSpreadsheet(ssId)
        .onEdit()
        .create();
    }

    SpreadsheetApp.getUi().alert('Dashboard action trigger is configured.');
  });
}

function generateVideoAdsScript() {
  withErrorLogging_('generateVideoAdsScript failed', function () {
    generateVideoQueueAdsScript_();
    SpreadsheetApp.getUi().alert('VIDEO Ads Script template generated in ADS_SCRIPT_TEMPLATE sheet.');
  });
}
