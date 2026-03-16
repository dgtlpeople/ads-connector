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
        const enabledMatch = enabledByKey[entityKey_(
          targetPlatform,
          normalizedEntity.account_id,
          normalizedEntity.entity_level,
          normalizedEntity.entity_id
        )] || {};

        const mergedEntity = mergePlanAndEnabledEntity_(normalizedEntity, enabledMatch);
        mergedEntity.goal_reach = goalReach;
        mergedEntity.goal_impressions = goalImpressions;

        let metrics = null;
        if (targetPlatform === 'google') {
          metrics = fetchGoogleEntityMetrics_(mergedEntity);
        } else if (targetPlatform === 'meta') {
          metrics = fetchMetaEntityMetrics_(mergedEntity);
        } else {
          throw new Error('Unsupported platform: ' + targetPlatform);
        }

        if (!metrics) {
          throw new Error('No metrics returned for entity_id=' + mergedEntity.entity_id);
        }

        metrics.goal_impressions = goalImpressions;
        metrics.goal_reach = goalReach;

        // Enforce required formula source for frequency.
        metrics.frequency = toNumber_(metrics.reach) > 0
          ? toNumber_(metrics.impressions) / toNumber_(metrics.reach)
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
          reach: 0,
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

function normalizePlanEntity_(row) {
  const platform = normalizePlatform_(row.platform);
  const entityLevel = normalizeEntityLevel_(row.entity_level || (platform === 'google' ? 'campaign' : 'adset'));
  const entityId = normalizeId_(row.entity_id);
  const entityName = normalizeId_(row.entity_name);
  const parsed = parseMetaEntityName_(entityName);

  return {
    platform: platform,
    account_id: normalizeId_(row.account_id),
    entity_level: entityLevel,
    entity_id: entityId,
    entity_name: entityName,
    campaign_id: entityLevel === 'campaign' ? entityId : normalizeId_(row.campaign_id),
    campaign_name: entityLevel === 'campaign'
      ? entityName
      : (normalizeId_(row.campaign_name) || parsed.campaign_name),
    adset_id: entityLevel === 'adset' ? entityId : normalizeId_(row.adset_id),
    adset_name: entityLevel === 'adset'
      ? (normalizeId_(row.adset_name) || parsed.adset_name)
      : normalizeId_(row.adset_name)
  };
}

function parseMetaEntityName_(entityName) {
  const raw = String(entityName || '');
  const split = raw.split(' | ');
  if (split.length < 2) {
    return { campaign_name: raw, adset_name: raw };
  }
  return {
    campaign_name: split[0].trim(),
    adset_name: split.slice(1).join(' | ').trim()
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
    end_date: normalizeId_(enabledEntity.end_date)
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
    toNumber_(row.reach),
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
