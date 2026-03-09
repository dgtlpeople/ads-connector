function buildDashboard() {
  withErrorLogging_('buildDashboard failed', function () {
    ensureHeader_(SHEETS.SUMMARY, HEADERS.SUMMARY);

    const sh = getSheet_(SHEETS.DASHBOARD);
    sh.clear();
    sh.clearConditionalFormatRules();
    sh.clearCharts();
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

    sh.getRange('A9:F9').setValues([['Entity', 'Platform', 'Imp Pace', 'Reach Pace', 'Imp Delivery', 'Action']]);
    sh.getRange('A9:F9').setBackground('#E2E8F0').setFontWeight('bold').setFontColor('#1E293B');

    const ranked = rows
      .map(function (r) {
        const impP = toNumber_(r.impression_pace_pct);
        const reachP = toNumber_(r.reach_pace_pct);
        return {
          entity: r.entity_name || r.entity_id || '',
          platform: r.platform || '',
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
        return [r.entity, r.platform, r.impPace, r.reachPace, r.impDelivery, r.action];
      });
      sh.getRange(10, 1, values.length, 6).setValues(values);
      sh.getRange(10, 3, values.length, 3).setNumberFormat('0.00%');
      sh.getRange(10, 1, values.length, 6).setBackground('#F8FAFC');
    }

    const dataLastRow = Math.max(10, 9 + ranked.length);

    const chart1 = sh.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange('A9:C' + dataLastRow))
      .setPosition(9, 8, 0, 0)
      .setOption('title', 'Impression Pace by Entity')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { format: 'percent', viewWindow: { min: 0 } })
      .build();
    sh.insertChart(chart1);

    const chart2 = sh.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange('A9:E' + dataLastRow))
      .setPosition(26, 1, 0, 0)
      .setOption('title', 'Impression vs Reach Performance')
      .setOption('legend', { position: 'top' })
      .setOption('vAxis', { format: 'percent' })
      .build();
    sh.insertChart(chart2);

    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0.9)
        .setBackground('#FEE2E2')
        .setRanges([sh.getRange('C10:D' + dataLastRow)])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(0.95, 1.1)
        .setBackground('#DCFCE7')
        .setRanges([sh.getRange('C10:D' + dataLastRow)])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(1.5)
        .setBackground('#FEF3C7')
        .setRanges([sh.getRange('C10:D' + dataLastRow)])
        .build()
    ];
    sh.setConditionalFormatRules(rules);

    sh.setColumnWidths(1, 1, 330);
    sh.setColumnWidths(2, 1, 90);
    sh.setColumnWidths(3, 3, 120);
    sh.setColumnWidth(6, 180);
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
