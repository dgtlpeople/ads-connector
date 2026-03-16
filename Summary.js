function buildSummary() {
  withErrorLogging_('buildSummary failed', function () {
    ensureHeader_(SHEETS.RAW_ALL, HEADERS.RAW_ALL);
    ensureHeader_(SHEETS.SUMMARY, HEADERS.SUMMARY);
    clearDataKeepHeader_(SHEETS.SUMMARY);

    const yesterday = getYesterdayDate_();
    const rows = readObjects_(SHEETS.RAW_ALL);
    const out = rows.map(function (r) {
      const start = r.start_date ? new Date(r.start_date) : null;
      const end = r.end_date ? new Date(r.end_date) : null;

      let daysTotal = 0;
      let daysElapsed = 0;

      if (start && end && !isNaN(start.getTime()) && !isNaN(end.getTime())) {
        daysTotal = Math.max(1, Math.floor((end.getTime() - start.getTime()) / 86400000) + 1);
        if (yesterday < start) {
          daysElapsed = 0;
        } else if (yesterday >= end) {
          daysElapsed = daysTotal;
        } else {
          daysElapsed = Math.floor((yesterday.getTime() - start.getTime()) / 86400000) + 1;
        }
      }

      const rawGoalImpressions = sanitizePlanGoal_(r.goal_impressions);
      const rawGoalReach = sanitizePlanGoal_(r.goal_reach);
      const hasGoalImpressions = rawGoalImpressions !== '';
      const hasGoalReach = rawGoalReach !== '';
      const goalImpressions = hasGoalImpressions ? toNumber_(rawGoalImpressions) : '';
      const goalReach = hasGoalReach ? toNumber_(rawGoalReach) : '';
      const impressions = toNumber_(r.impressions);
      const reach = toNumber_(r.reach);

      const expectedImpressionsToDate = hasGoalImpressions && daysTotal > 0
        ? goalImpressions * (daysElapsed / daysTotal)
        : '';
      const expectedReachToDate = hasGoalReach && daysTotal > 0
        ? goalReach * (daysElapsed / daysTotal)
        : '';

      const impressionDeliveryPct = hasGoalImpressions && goalImpressions > 0
        ? impressions / goalImpressions
        : '';
      const reachDeliveryPct = hasGoalReach && goalReach > 0
        ? reach / goalReach
        : '';
      const impressionPacePct = hasGoalImpressions && expectedImpressionsToDate > 0
        ? impressions / expectedImpressionsToDate
        : '';
      const reachPacePct = hasGoalReach && expectedReachToDate > 0
        ? reach / expectedReachToDate
        : '';

      const action = recommendAction_(
        impressionPacePct,
        reachPacePct,
        toNumber_(r.frequency),
        hasGoalImpressions,
        hasGoalReach
      );

      return [
        r.platform || '',
        r.account_id || '',
        r.entity_level || '',
        r.entity_id || '',
        r.entity_name || '',
        r.campaign_id || '',
        r.campaign_name || '',
        r.adset_id || '',
        r.adset_name || '',
        r.start_date || '',
        r.end_date || '',
        goalImpressions,
        goalReach,
        impressions,
        reach,
        toNumber_(r.frequency),
        toNumber_(r.cpm),
        toNumber_(r.video_p25),
        toNumber_(r.video_p50),
        toNumber_(r.video_p75),
        toNumber_(r.video_p100),
        daysTotal,
        daysElapsed,
        expectedImpressionsToDate,
        expectedReachToDate,
        impressionDeliveryPct,
        reachDeliveryPct,
        impressionPacePct,
        reachPacePct,
        action,
        r.status || '',
        r.channel_type || ''
      ];
    });

    if (out.length) appendRows_(SHEETS.SUMMARY, out);
    formatSummary_();
  });
}

function recommendAction_(impressionPacePct, reachPacePct, frequency, hasGoalImpressions, hasGoalReach) {
  const hasImpGoal = !!hasGoalImpressions;
  const hasReachGoal = !!hasGoalReach;

  if (!hasImpGoal && !hasReachGoal) return 'No goal';

  const imp = toNumber_(impressionPacePct);
  const reach = toNumber_(reachPacePct);
  const freq = toNumber_(frequency);

  if (hasImpGoal && !hasReachGoal) {
    if (imp < 0.9) return 'Increase budget';
    if (imp > 1.5) return 'Decrease budget';
    if (imp >= 0.95 && imp <= 1.1) return 'On track';
    return 'Monitor';
  }

  if (!hasImpGoal && hasReachGoal) {
    if (reach < 0.9) return 'Expand reach';
    if (reach >= 0.95 && reach <= 1.1) return 'On track';
    return 'Monitor';
  }

  if (imp < 0.9) return 'Increase budget';
  if (imp > 1.5) return 'Decrease budget';
  if (reach >= 1 && imp < 1) return 'Increase frequency cap';
  if (imp > 1 && reach < 0.9) return freq >= 3 ? 'Decrease frequency cap' : 'Expand reach';
  if (imp >= 0.95 && imp <= 1.1 && reach >= 0.95 && reach <= 1.1) return 'On track';
  return 'Monitor';
}

function formatSummary_() {
  const sh = getSheet_(SHEETS.SUMMARY);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  sh.getRange(2, 16, lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, 17, lastRow - 1, 1).setNumberFormat('0.00');
  sh.getRange(2, 18, lastRow - 1, 4).setNumberFormat('0.00%');
  sh.getRange(2, 24, lastRow - 1, 2).setNumberFormat('#,##0');
  sh.getRange(2, 26, lastRow - 1, 4).setNumberFormat('0.00%');
}
