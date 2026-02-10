const PROPS = PropertiesService.getScriptProperties();

const CONFIG = {
  LINE_CHANNEL_ACCESS_TOKEN: PROPS.getProperty('LINE_CHANNEL_ACCESS_TOKEN'),
  LINE_CHANNEL_SECRET: PROPS.getProperty('LINE_CHANNEL_SECRET'),
  GEMINI_API_KEY: PROPS.getProperty('GEMINI_API_KEY'),
  SPREADSHEET_ID: PROPS.getProperty('SPREADSHEET_ID'),
  FOLDER_ID: PROPS.getProperty('FOLDER_ID')
};

const API_ENDPOINTS = {
  LINE_REPLY: 'https://api.line.me/v2/bot/message/reply',
  LINE_PUSH: 'https://api.line.me/v2/bot/message/push',
  LINE_CONTENT_BASE: 'https://api-data.line.me/v2/bot/message',
  GEMINI_GENERATE_CONTENT:
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' +
    CONFIG.GEMINI_API_KEY
};

const SETTINGS = {
  DEDUPE_TTL_SECONDS: 600,
  DRIVE_MAX_RETRY: 3,
  API_MAX_RETRY: 2,
  API_RETRY_BASE_MS: 1000,
  SUMMARY_CATEGORIES: ['食費', '日用品', '交際費', '交通費', '医療費', 'その他']
};

function assertRequiredConfig_() {
  const required = ['LINE_CHANNEL_ACCESS_TOKEN', 'GEMINI_API_KEY', 'SPREADSHEET_ID', 'FOLDER_ID'];
  const missing = required.filter(function(key) {
    return !CONFIG[key];
  });

  if (missing.length > 0) {
    throw new Error('Missing script properties: ' + missing.join(', '));
  }
}
