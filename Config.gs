const SHEETS = {
  CAMPAIGNS_ENABLED: 'CAMPAIGNS_ENABLED',
  PLAN: 'PLAN',
  RAW_ALL: 'RAW_ALL',
  SUMMARY: 'SUMMARY',
  LOG: 'LOG'
};

const HEADERS = {
  CAMPAIGNS_ENABLED: [
    'platform',
    'account_id',
    'entity_level',
    'entity_id',
    'entity_name',
    'campaign_id',
    'campaign_name',
    'adset_id',
    'adset_name',
    'start_date',
    'end_date',
    'status',
    'channel_type'
  ],
  PLAN: [
    'platform',
    'account_id',
    'entity_level',
    'entity_id',
    'entity_name',
    'goal_impressions',
    'goal_reach'
  ],
  RAW_ALL: [
    'platform',
    'account_id',
    'entity_level',
    'entity_id',
    'entity_name',
    'campaign_id',
    'campaign_name',
    'adset_id',
    'adset_name',
    'start_date',
    'end_date',
    'goal_impressions',
    'goal_reach',
    'impressions',
    'reach',
    'frequency',
    'cpm',
    'video_p25',
    'video_p50',
    'video_p75',
    'video_p100',
    'status',
    'channel_type'
  ],
  SUMMARY: [
    'platform',
    'account_id',
    'entity_level',
    'entity_id',
    'entity_name',
    'campaign_id',
    'campaign_name',
    'adset_id',
    'adset_name',
    'start_date',
    'end_date',
    'goal_impressions',
    'goal_reach',
    'impressions',
    'reach',
    'frequency',
    'cpm',
    'video_p25',
    'video_p50',
    'video_p75',
    'video_p100',
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

const SCRIPT_PROPERTY_KEYS = [
  'GOOGLE_ADS_DEVELOPER_TOKEN',
  'GOOGLE_ADS_CUSTOMER_ID',
  'GOOGLE_ADS_LOGIN_CUSTOMER_ID',
  'GOOGLE_OAUTH_CLIENT_ID',
  'GOOGLE_OAUTH_CLIENT_SECRET',
  'GOOGLE_ADS_REFRESH_TOKEN',
  'META_ACCESS_TOKEN',
  'META_AD_ACCOUNT_IDS'
];

function getScriptProps_() {
  return PropertiesService.getScriptProperties();
}
