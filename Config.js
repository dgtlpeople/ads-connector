const SHEETS = {
  CAMPAIGNS_ENABLED: 'CAMPAIGNS_ENABLED',
  PLAN: 'PLAN',
  RAW_ALL: 'RAW_ALL',
  SUMMARY: 'SUMMARY',
  LOG: 'LOG',
};

const HEADERS = {
  CAMPAIGNS_ENABLED: [
    'platform',
    'account_id',
    'campaign_id',
    'campaign_name',
    'start_date',
    'end_date',
    'status',
    'channel_type'
  ],
  PLAN: [
    'platform',
    'account_id',
    'campaign_id',
    'campaign_name',
    'goal_impressions',
    'goal_reach'
  ],
  RAW_ALL: [
    'platform',
    'account_id',
    'campaign_id',
    'campaign_name',
    'start_date',
    'end_date',
    'goal_impressions',
    'goal_reach',
    'impressions',
    'reach_or_unique_users',
    'frequency',
    'average_cpm',
    'video_quartile_p25_rate',
    'video_quartile_p50_rate',
    'video_quartile_p75_rate',
    'video_quartile_p100_rate',
    'status',
    'channel_type'
  ],
  SUMMARY: [
    'platform',
    'account_id',
    'campaign_id',
    'campaign_name',
    'start_date',
    'end_date',
    'goal_impressions',
    'goal_reach',
    'impressions',
    'reach_or_unique_users',
    'frequency',
    'average_cpm',
    'video_quartile_p25_rate',
    'video_quartile_p50_rate',
    'video_quartile_p75_rate',
    'video_quartile_p100_rate',
    'days_total',
    'days_elapsed',
    'expected_impressions_to_date',
    'expected_reach_to_date',
    'impression_delivery_pct',
    'reach_delivery_pct',
    'impression_pace_pct',
    'reach_pace_pct',
    'action',
    'status',
    'channel_type'
  ],
  LOG: ['timestamp', 'message', 'detail']
};

function getScriptProps() {
  return PropertiesService.getScriptProperties();
}