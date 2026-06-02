function refreshRawAllFromPlan_(platform) {
  withErrorLogging_('refreshRawAllFromPlan_ failed for ' + platform, function () {
    ensureHeader_(SHEETS.PLAN, HEADERS.PLAN);
    ensureHeader_(SHEETS.RAW_ALL, HEADERS.RAW_ALL);

    const targetPlatform = normalizePlatform_(platform);
    const existingOtherPlatformRows = readObjects_(SHEETS.RAW_ALL).filter(function (r) {
      return normalizePlatform_(r.platform) !== targetPlatform;
    });

    clearDataKeepHeader_(SHEETS.RAW_ALL);

    const output = existingOtherPlatformRows.map(function (r) {
      return mapRawRow_(r);
    });

    const enabledRows = readObjects_(SHEETS.CAMPAIGNS_ENABLED).filter(function (r) {
      return normalizePlatform_(r.platform) === targetPlatform;
    });

    const enabledByKey = {};
    enabledRows.forEach(function (r) {
      enabledByKey[entityKey_(r.platform, r.account_id, r.entity_level, r.entity_id)] = r;
    });

    const planRows = readObjects_(SHEETS.PLAN).filter(function (r) {
      return normalizePlatform_(r.platform) === targetPlatform;
    });

    if (!planRows.length) {
      log_('RAW_ALL refresh', 'No PLAN rows for platform=' + targetPlatform);
      if (output.length) appendRows_(SHEETS.RAW_ALL, output);
      return;
    }

    planRows.forEach(function (planRow) {
      const goalImpressions = sanitizePlanGoal_(planRow.goal_impressions);
      const goalReach = sanitizePlanGoal_(planRow.goal_reach);

      try {
        const normalizedEntity = normalizePlanEntity_(planRow);
        const enabledMatch = findEnabledEntityMatch_(enabledByKey, targetPlatform, normalizedEntity);

        const mergedEntity = mergePlanAndEnabledEntity_(normalizedEntity, enabledMatch);
        mergedEntity.goal_reach = goalReach;
        mergedEntity.goal_impressions = goalImpressions;

        let metrics = null;
        if (targetPlatform === 'google') {
          metrics = fetchGoogleEntityMetrics_(mergedEntity);
        } else if (targetPlatform === 'meta') {
          metrics = fetchMetaEntityMetrics_(mergedEntity);
        } else if (targetPlatform === 'tiktok') {
          metrics = fetchTikTokEntityMetrics_(mergedEntity);
        } else {
          throw new Error('Unsupported platform: ' + targetPlatform);
        }

        if (!metrics) {
          throw new Error('No metrics returned for entity_id=' + mergedEntity.entity_id);
        }

        metrics.goal_impressions = goalImpressions;
        metrics.goal_reach = goalReach;

        const rawReach = metrics.reach === '' ? 0 : toNumber_(metrics.reach);
        metrics.frequency = rawReach > 0
          ? toNumber_(metrics.impressions) / rawReach
          : 0;

        output.push(mapRawRow_(metrics));
      } catch (e) {
        log_(
          'RAW entity refresh failed',
          'platform=' + targetPlatform + '; entity_id=' + normalizeId_(planRow.entity_id) + '; reason=' + e.message
        );

        const fallback = normalizePlanEntity_(planRow);
        output.push(mapRawRow_({
          platform: targetPlatform,
          account_id: fallback.account_id,
          entity_level: fallback.entity_level,
          entity_id: fallback.entity_id,
          entity_name: fallback.entity_name,
          campaign_id: fallback.campaign_id,
          campaign_name: fallback.campaign_name,
          adset_id: fallback.adset_id,
          adset_name: fallback.adset_name,
          start_date: '',
          end_date: '',
          goal_impressions: goalImpressions,
          goal_reach: goalReach,
          impressions: 0,
          reach: '',
          frequency: 0,
          cpm: 0,
          video_p25: 0,
          video_p50: 0,
          video_p75: 0,
          video_p100: 0,
          status: 'ERROR',
          channel_type: ''
        }));
      }
    });

    if (output.length) appendRows_(SHEETS.RAW_ALL, output);
    log_('RAW_ALL refresh', 'platform=' + targetPlatform + '; rows=' + output.length);
  });
}

function findEnabledEntityMatch_(enabledByKey, platform, normalizedEntity) {
  const accountId = normalizeId_(normalizedEntity.account_id);
  const entityLevel = normalizeEntityLevel_(normalizedEntity.entity_level);
  const entityId = normalizeId_(normalizedEntity.entity_id);
  const platformKey = normalizePlatform_(platform);

  const exact = enabledByKey[entityKey_(platformKey, accountId, entityLevel, entityId)];
  if (exact) return exact;

  // Legacy compatibility: older PLAN rows may have blank account_id.
  const blankAccount = enabledByKey[entityKey_(platformKey, '', entityLevel, entityId)];
  if (blankAccount) return blankAccount;

  // Fallback scan by normalized entity tuple when account differs.
  const keys = Object.keys(enabledByKey);
  for (let i = 0; i < keys.length; i++) {
    const r = enabledByKey[keys[i]];
    if (normalizePlatform_(r.platform) !== platformKey) continue;
    if (normalizeEntityLevel_(r.entity_level) !== entityLevel) continue;
    if (normalizeId_(r.entity_id) !== entityId) continue;
    return r;
  }

  return {};
}

function normalizePlanEntity_(row) {
  const platform = normalizePlatform_(row.platform);
  const defaultLevel = platform === 'google' ? 'campaign' : (platform === 'tiktok' ? 'adgroup' : 'adset');
  const entityLevel = normalizeEntityLevel_(row.entity_level || defaultLevel);
  const entityId = normalizeId_(row.entity_id);
  const entityName = normalizeId_(row.entity_name);
  const parsed = parseCompositeEntityName_(entityName);

  return {
    platform: platform,
    account_id: normalizeId_(row.account_id),
    entity_level: entityLevel,
    entity_id: entityId,
    entity_name: entityName,
    campaign_id: normalizeId_(row.campaign_id),
    campaign_name: normalizeId_(row.campaign_name) || parsed.campaign_name,
    adset_id: normalizeId_(row.adset_id),
    adset_name: normalizeId_(row.adset_name) || parsed.child_name
  };
}

function mergePlanAndEnabledEntity_(planEntity, enabledEntity) {
  const merged = {
    platform: planEntity.platform,
    account_id: planEntity.account_id || normalizeId_(enabledEntity.account_id),
    entity_level: planEntity.entity_level || normalizeEntityLevel_(enabledEntity.entity_level),
    entity_id: planEntity.entity_id || normalizeId_(enabledEntity.entity_id),
    entity_name: planEntity.entity_name || normalizeId_(enabledEntity.entity_name),
    campaign_id: planEntity.campaign_id || normalizeId_(enabledEntity.campaign_id),
    campaign_name: planEntity.campaign_name || normalizeId_(enabledEntity.campaign_name),
    adset_id: planEntity.adset_id || normalizeId_(enabledEntity.adset_id),
    adset_name: planEntity.adset_name || normalizeId_(enabledEntity.adset_name),
    start_date: normalizeId_(enabledEntity.start_date),
    end_date: normalizeId_(enabledEntity.end_date),
    status: normalizeId_(enabledEntity.status),
    channel_type: normalizeId_(enabledEntity.channel_type)
  };

  if (!merged.entity_id) {
    throw new Error('Missing entity_id in PLAN row');
  }

  if (!merged.entity_level) {
    throw new Error('Missing entity_level in PLAN row');
  }

  if (!merged.entity_name) {
    merged.entity_name = merged.entity_id;
  }

  if (!merged.campaign_id && merged.entity_level === 'campaign') {
    merged.campaign_id = merged.entity_id;
  }

  if (!merged.campaign_name) {
    merged.campaign_name = parseCompositeEntityName_(merged.entity_name).campaign_name;
  }

  if (!merged.adset_id && (merged.entity_level === 'adset' || merged.entity_level === 'adgroup')) {
    merged.adset_id = merged.entity_id;
  }

  if (!merged.adset_name && (merged.entity_level === 'adset' || merged.entity_level === 'adgroup')) {
    merged.adset_name = parseCompositeEntityName_(merged.entity_name).child_name;
  }

  return merged;
}

function mapRawRow_(row) {
  return [
    row.platform || '',
    row.account_id || '',
    row.entity_level || '',
    row.entity_id || '',
    row.entity_name || '',
    row.campaign_id || '',
    row.campaign_name || '',
    row.adset_id || '',
    row.adset_name || '',
    row.start_date || '',
    row.end_date || '',
    hasUsablePlanGoal_(row.goal_impressions) ? toNumber_(row.goal_impressions) : '',
    hasUsablePlanGoal_(row.goal_reach) ? toNumber_(row.goal_reach) : '',
    toNumber_(row.impressions),
    row.reach === '' ? '' : toNumber_(row.reach),
    toNumber_(row.frequency),
    toNumber_(row.cpm),
    toNumber_(row.video_p25),
    toNumber_(row.video_p50),
    toNumber_(row.video_p75),
    toNumber_(row.video_p100),
    row.status || '',
    row.channel_type || ''
  ];
}
