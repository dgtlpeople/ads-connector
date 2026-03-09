function buildDashboard() {
  withErrorLogging_('buildDashboard failed', function () {
    ensureHeader_(SHEETS.SUMMARY, HEADERS.SUMMARY);

    const sh = getSheet_(SHEETS.DASHBOARD);
    sh.clear();
    sh.clearConditionalFormatRules();
    sh.getCharts().forEach(function (chart) {
      sh.removeChart(chart);
    });
    sh.setHiddenGridlines(true);

    const rows = readObjects_(SHEETS.SUMMARY);

    const totals = rows.reduce(function (acc, r) {
      acc.goalImp += toNumber_(r.goal_impressions);
      acc.goalReach += toNumber_(r.goal_reach);
      acc.imp += toNumber_(r.impressions);
      acc.reach += toNumber_(r.reach);
      acc.expImp += toNumber_(r.expected_impressions_to_date);
      acc.expReach += toNumber_(r.expected_reach_to_date);
      return acc;
    }, { goalImp: 0, goalReach: 0, imp: 0, reach: 0, expImp: 0, expReach: 0 });

    const impDelivery = totals.goalImp > 0 ? totals.imp / totals.goalImp : 0;
    const reachDelivery = totals.goalReach > 0 ? totals.reach / totals.goalReach : 0;
    const impPace = totals.expImp > 0 ? totals.imp / totals.expImp : 0;
    const reachPace = totals.expReach > 0 ? totals.reach / totals.expReach : 0;

    sh.getRange('A1:H2').merge();
    sh.getRange('A1').setValue('Ads Connector Dashboard');
    sh.getRange('A1').setFontSize(22).setFontWeight('bold').setFontColor('#FFFFFF');
    sh.getRange('A1:H2').setBackground('#0E7490');

    sh.getRange('A3:H3').setBackground('#CCFBF1');
    sh.getRange('A3').setValue('Auto-generated from SUMMARY. Refresh via: Ads Connector > Build DASHBOARD');
    sh.getRange('A3').setFontColor('#134E4A').setFontWeight('bold');

    writeKpiCard_(sh, 'A5:B7', 'Impressions Delivery', impDelivery, 'percent', '#0EA5E9');
    writeKpiCard_(sh, 'C5:D7', 'Reach Delivery', reachDelivery, 'percent', '#14B8A6');
    writeKpiCard_(sh, 'E5:F7', 'Impressions Pace', impPace, 'percent', '#F59E0B');
    writeKpiCard_(sh, 'G5:H7', 'Reach Pace', reachPace, 'percent', '#F97316');

    sh.getRange('A9:I9').setValues([['Entity', 'Platform', 'Campaign ID', 'Imp Pace', 'Reach Pace', 'Imp Delivery', 'Action', 'Run', 'Last run']]);
    sh.getRange('A9:I9').setBackground('#E2E8F0').setFontWeight('bold').setFontColor('#1E293B');

    const ranked = rows
      .map(function (r) {
        const impP = toNumber_(r.impression_pace_pct);
        const reachP = toNumber_(r.reach_pace_pct);
        return {
          entity: r.entity_name || r.entity_id || '',
          platform: r.platform || '',
          campaignId: String(r.campaign_id || ''),
          impPace: impP,
          reachPace: reachP,
          impDelivery: toNumber_(r.impression_delivery_pct),
          action: r.action || '',
          volatility: Math.abs(impP - 1) + Math.abs(reachP - 1)
        };
      })
      .sort(function (a, b) {
        return b.volatility - a.volatility;
      })
      .slice(0, 12);

    if (ranked.length) {
      const values = ranked.map(function (r) {
        return [r.entity, r.platform, r.campaignId, r.impPace, r.reachPace, r.impDelivery, r.action, false, ''];
      });
      sh.getRange(10, 1, values.length, 9).setValues(values);
      sh.getRange(10, 4, values.length, 3).setNumberFormat('0.00%');
      sh.getRange(10, 1, values.length, 9).setBackground('#F8FAFC');
      sh.getRange(10, 8, values.length, 1).insertCheckboxes();
    }

    const dataLastRow = Math.max(10, 9 + ranked.length);

    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0.9)
        .setBackground('#FEE2E2')
        .setRanges([sh.getRange('D10:E' + dataLastRow)])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(0.95, 1.1)
        .setBackground('#DCFCE7')
        .setRanges([sh.getRange('D10:E' + dataLastRow)])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(1.5)
        .setBackground('#FEF3C7')
        .setRanges([sh.getRange('D10:E' + dataLastRow)])
        .build()
    ];
    sh.setConditionalFormatRules(rules);

    sh.setColumnWidths(1, 1, 330);
    sh.setColumnWidths(2, 1, 90);
    sh.setColumnWidth(3, 120);
    sh.setColumnWidths(4, 3, 120);
    sh.setColumnWidth(7, 180);
    sh.setColumnWidth(8, 60);
    sh.setColumnWidth(9, 220);
    sh.setFrozenRows(9);
  });
}

function writeKpiCard_(sheet, a1, title, value, type, color) {
  const range = sheet.getRange(a1);
  range.merge();
  range.setBackground(color);
  range.setFontColor('#FFFFFF');
  range.setVerticalAlignment('middle');

  const label = type === 'percent'
    ? title + '\n' + Utilities.formatString('%.1f%%', value * 100)
    : title + '\n' + value;

  range.setValue(label);
  range.setFontWeight('bold');
  range.setFontSize(14);
  range.setWrap(true);
}

function handleDashboardActionEdit_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEETS.DASHBOARD) return;
  if (e.range.getRow() < 10 || e.range.getColumn() !== 8) return;
  if (String(e.value).toUpperCase() !== 'TRUE') return;

  const row = e.range.getRow();
  const rowValues = sh.getRange(row, 1, 1, 9).getValues()[0];
  const platform = normalizePlatform_(rowValues[1]);
  const campaignId = normalizeId_(rowValues[2]).replace(/-/g, '');
  const action = String(rowValues[6] || '');

  let result = '';
  try {
    if (platform !== 'google') {
      result = 'Skipped: auto-run enabled only for Google';
    } else if (!campaignId) {
      result = 'Skipped: missing campaign ID';
    } else {
      const impressionPace = toNumber_(rowValues[3]);
      result = executeGoogleDashboardAction_(campaignId, action, impressionPace);
    }
  } catch (err) {
    result = 'ERROR: ' + summarizeDashboardError_(err);
    log_('Dashboard run failed', 'row=' + row + '; campaign_id=' + campaignId + '; ' + err.message);
  }

  sh.getRange(row, 9).setValue(formatDate_(new Date()) + ' ' + result);
  sh.getRange(row, 8).setValue(false);
}

function summarizeDashboardError_(err) {
  const raw = String((err && err.message) || err || '');
  const shortGoogle = extractGoogleErrorSummary_(raw);
  if (shortGoogle) return shortGoogle;
  const singleLine = raw.replace(/\s+/g, ' ').trim();
  return singleLine.length > 180 ? singleLine.slice(0, 177) + '...' : singleLine;
}

function extractGoogleErrorSummary_(rawMessage) {
  const marker = 'Google campaigns mutate failed';
  if (rawMessage.indexOf(marker) === -1 && rawMessage.indexOf('Google campaignBudgets mutate failed') === -1) {
    return '';
  }

  try {
    const jsonStart = rawMessage.indexOf('{');
    if (jsonStart === -1) return '';
    const parsed = JSON.parse(rawMessage.slice(jsonStart));
    const err = parsed && parsed.error ? parsed.error : {};
    const details = Array.isArray(err.details) ? err.details : [];
    const firstDetail = details.length ? details[0] : {};
    const firstError = firstDetail && Array.isArray(firstDetail.errors) && firstDetail.errors.length
      ? firstDetail.errors[0]
      : {};
    const code = err.code ? String(err.code) : '';
    const status = err.status ? String(err.status) : '';
    const message = firstError.message || err.message || 'Google Ads mutate failed';
    const trigger = firstError.trigger && firstError.trigger.stringValue
      ? ' (trigger: ' + firstError.trigger.stringValue + ')'
      : '';
    return [code, status, message + trigger].filter(Boolean).join(' | ');
  } catch (_e) {
    return '';
  }
}
