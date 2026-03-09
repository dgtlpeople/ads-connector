function getActionRecommendation_(impressionPacePct, reachPacePct, frequency, hasReachData) {
  const imp = toNumber_(impressionPacePct);
  const reach = hasReachData ? toNumber_(reachPacePct) : null;
  const freq = frequency === '' ? null : toNumber_(frequency);

  if (hasReachData && imp < 1 && reach >= 1) {
    if (freq !== null && freq < 2) return 'Increase frequency cap';
    return 'Increase delivery pressure';
  }

  if (hasReachData && imp > 1 && reach < 0.9) {
    if (freq !== null && freq > 4) return 'Reduce frequency cap';
    return 'Expand reach';
  }

  if (hasReachData && imp < 0.9 && reach < 0.9) return 'Increase budget';
  if (hasReachData && imp > 1.5 && reach > 1.5) return 'Decrease budget';
  if (imp > 1.5) return 'Decrease budget';
  if (imp > 1.2) return 'Decrease budget slightly';
  if (imp < 0.8) return 'Increase budget aggressively';
  if (imp < 0.95) return 'Increase budget slightly';

  if (hasReachData && imp >= 0.95 && imp <= 1.1 && reach >= 0.95 && reach <= 1.1) {
    return 'On track';
  }

  return 'Monitor';
}

function buildSummary() {
  ensureHeader_(SHEETS.SUMMARY, HEADERS.SUMMARY);
  clearDataKeepHeader_(SHEETS.SUMMARY);

  const rows = readObjects_(SHEETS.RAW_ALL);
  const today = new Date();
  const output = [];

  rows.forEach(r => {
    const start = new Date(r.start_date);
    const end = new Date(r.end_date);
    const goalImpressions = toNumber_(r.goal_impressions);
    const goalReach = toNumber_(r.goal_reach);
    const impressions = toNumber_(r.impressions);
    const reachOrUsers = r.reach_or_unique_users === '' ? '' : toNumber_(r.reach_or_unique_users);
    const frequency = r.frequency === '' ? '' : toNumber_(r.frequency);

    let daysTotal = 0;
    let daysElapsed = 0;

    if (!isNaN(start.getTime()) && !isNaN(end.getTime())) {
      daysTotal = Math.max(1, Math.floor((end - start) / 86400000) + 1);

      if (today < start) {
        daysElapsed = 0;
      } else if (today > end) {
        daysElapsed = daysTotal;
      } else {
        daysElapsed = Math.floor((today - start) / 86400000) + 1;
      }
    }

    const expectedImpressionsToDate = daysTotal > 0 ? goalImpressions * (daysElapsed / daysTotal) : 0;
    const expectedReachToDate = daysTotal > 0 ? goalReach * (daysElapsed / daysTotal) : 0;

    const impressionDeliveryPct = goalImpressions > 0 ? impressions / goalImpressions : 0;
    const reachDeliveryPct = goalReach > 0 && reachOrUsers !== '' ? reachOrUsers / goalReach : '';
    const impressionPacePct = expectedImpressionsToDate > 0 ? impressions / expectedImpressionsToDate : 0;
    const reachPacePct = expectedReachToDate > 0 && reachOrUsers !== '' ? reachOrUsers / expectedReachToDate : '';
    const hasReachData = reachPacePct !== '';

    const action = getActionRecommendation_(
      impressionPacePct,
      reachPacePct,
      frequency,
      hasReachData
    );

    output.push([
      r.platform,
      r.account_id,
      r.campaign_id,
      r.campaign_name,
      r.start_date,
      r.end_date,
      goalImpressions,
      goalReach,
      impressions,
      reachOrUsers,
      frequency,
      toNumber_(r.average_cpm),
      toNumber_(r.video_quartile_p25_rate),
      toNumber_(r.video_quartile_p50_rate),
      toNumber_(r.video_quartile_p75_rate),
      toNumber_(r.video_quartile_p100_rate),
      daysTotal,
      daysElapsed,
      expectedImpressionsToDate,
      expectedReachToDate,
      impressionDeliveryPct,
      reachDeliveryPct,
      impressionPacePct,
      reachPacePct,
      action,
      r.status,
      r.channel_type
    ]);
  });

  if (output.length) {
    appendRows_(SHEETS.SUMMARY, output);
  }

  formatSummary_();
  SpreadsheetApp.getUi().alert('SUMMARY built.');
}

function formatSummary_() {
  const sh = getSheet_(SHEETS.SUMMARY);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  sh.getRange(2, 11, lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, 12, lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, 13, lastRow - 1, 4).setNumberFormat('0.00%');
  sh.getRange(2, 19, lastRow - 1, 2).setNumberFormat('#,##0');
  sh.getRange(2, 21, lastRow - 1, 1).setNumberFormat('0.00%');
  sh.getRange(2, 22, lastRow - 1, 1).setNumberFormat('0.00%');
  sh.getRange(2, 23, lastRow - 1, 1).setNumberFormat('0.00%');
  sh.getRange(2, 24, lastRow - 1, 1).setNumberFormat('0.00%');
}