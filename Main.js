function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ads Connector')
    .addItem('Setup sheets', 'setupSheets')
    .addItem('Reset PLAN header', 'resetPlanHeader')
    .addItem('Check config', 'checkConfig')
    .addSeparator()
    .addItem('Load Google entities', 'loadGoogleEntities')
    .addItem('Load Meta entities', 'loadMetaEntities')
    .addItem('Load TikTok entities', 'loadTikTokEntities')
    .addSeparator()
    .addItem('Refresh RAW_ALL from PLAN (Google)', 'refreshRawAllFromPlanGoogle')
    .addItem('Refresh RAW_ALL from PLAN (Meta)', 'refreshRawAllFromPlanMeta')
    .addItem('Refresh RAW_ALL from PLAN (TikTok)', 'refreshRawAllFromPlanTikTok')
    .addItem('Refresh RAW_ALL from PLAN (All platforms)', 'refreshRawAllFromPlanAll')
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
    ensureHeader_(SHEETS.REACH_CACHE, HEADERS.REACH_CACHE);
    ensureHeader_(SHEETS.GOOGLE_CHANGES_LOG, HEADERS.GOOGLE_CHANGES_LOG);
    ensureHeader_(SHEETS.VIDEO_ACTION_QUEUE, HEADERS.VIDEO_ACTION_QUEUE);
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
    sh.getRange(1, 1, 1, 3).setValues([['key', 'required', 'status']]);

    const props = getScriptProps_();
    const rows = [];

    SCRIPT_PROPERTY_KEYS_REQUIRED.forEach(function (k) {
      rows.push([k, 'yes', isConfigured_(props.getProperty(k)) ? 'OK' : 'MISSING']);
    });

    SCRIPT_PROPERTY_KEYS_OPTIONAL.forEach(function (k) {
      rows.push([k, 'no', isConfigured_(props.getProperty(k)) ? 'OK' : 'OPTIONAL_MISSING']);
    });

    if (rows.length) {
      sh.getRange(2, 1, rows.length, 3).setValues(rows);
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

function refreshRawAllFromPlanTikTok() {
  withErrorLogging_('refreshRawAllFromPlanTikTok failed', function () {
    refreshRawAllFromPlan_('tiktok');
    SpreadsheetApp.getUi().alert('RAW_ALL refreshed for TikTok.');
  });
}

function refreshRawAllFromPlanAll() {
  withErrorLogging_('refreshRawAllFromPlanAll failed', function () {
    const platforms = ['google', 'meta', 'tiktok'];
    const succeeded = [];
    const failed = [];

    platforms.forEach(function (platform) {
      try {
        refreshRawAllFromPlan_(platform);
        succeeded.push(platform);
      } catch (e) {
        const reason = e && e.message ? e.message : String(e);
        failed.push(platform + ': ' + reason);
        log_('RAW_ALL refresh failed', 'platform=' + platform + '; reason=' + reason);
      }
    });

    if (failed.length) {
      SpreadsheetApp.getUi().alert(
        'RAW_ALL refresh finished with errors.\nSuccess: ' + (succeeded.length ? succeeded.join(', ') : 'none') +
        '\nFailed:\n- ' + failed.join('\n- ')
      );
      return;
    }

    SpreadsheetApp.getUi().alert('RAW_ALL refreshed for all platforms: ' + succeeded.join(', ') + '.');
  });
}

function runFullPipeline() {
  withErrorLogging_('runFullPipeline failed', function () {
    loadGoogleEntities();
    loadMetaEntities();
    loadTikTokEntities();
    refreshRawAllFromPlan_('google');
    refreshRawAllFromPlan_('meta');
    refreshRawAllFromPlan_('tiktok');
    buildSummary();
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
