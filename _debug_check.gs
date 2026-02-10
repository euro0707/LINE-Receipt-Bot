function runSystemCheck() {
  const results = [];
  const log = (msg) => {
    console.log(msg);
    results.push(msg);
  };

  log('=== SYSTEM CHECK START ===');

  // 1. Check Properties
  try {
    const props = PropertiesService.getScriptProperties();
    const keys = ['LINE_CHANNEL_ACCESS_TOKEN', 'GEMINI_API_KEY', 'SPREADSHEET_ID', 'FOLDER_ID'];
    const missing = keys.filter(k => !props.getProperty(k));
    if (missing.length > 0) {
      log('❌ MISSING PROPERTIES: ' + missing.join(', '));
    } else {
      log('✅ Script Properties: OK');
    }
  } catch (e) {
    log('❌ Property Check Error: ' + e);
  }

  // 2. Check Spreadsheet
  if (CONFIG.SPREADSHEET_ID) {
    try {
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      log('✅ Spreadsheet: OK (' + ss.getName() + ')');
    } catch (e) {
      log('❌ Spreadsheet Error: ' + e);
    }
  }

  // 3. Check Drive Folder
  if (CONFIG.FOLDER_ID) {
    try {
      const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
      log('✅ Drive Folder: OK (' + folder.getName() + ')');
    } catch (e) {
      log('❌ Drive Folder Error: ' + e);
    }
  }

  // 4. Check Gemini API
  if (CONFIG.GEMINI_API_KEY) {
    try {
      log('Wait... testing Gemini API...');
      // Note: Using 1.5-flash for the test to match config
      const response = UrlFetchApp.fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + CONFIG.GEMINI_API_KEY,
        {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({
            contents: [{ parts: [{ text: "Hello" }] }]
          }),
          muteHttpExceptions: true
        }
      );
      const status = response.getResponseCode();
      if (status >= 200 && status < 300) {
        log('✅ Gemini API: OK (Status ' + status + ')');
      } else {
        log('❌ Gemini API Error: ' + status + ' ' + response.getContentText());
      }
    } catch (e) {
      log('❌ Gemini API Connection Error: ' + e);
    }
  }

  log('=== SYSTEM CHECK END ===');
  return results.join('\n');
}
