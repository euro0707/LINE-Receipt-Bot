/**
 * エラー診断ツール
 * GASエディタでこの関数を実行すると、最近のエラーを自動的に検索して報告します
 */

function diagnoseRecentErrors() {
  const results = [];
  const log = (msg) => {
    console.log(msg);
    results.push(msg);
  };

  log('=== エラー診断開始 ===');
  log('実行時刻: ' + new Date().toISOString());
  log('');

  // 1. Script Propertiesの確認
  log('【1. Script Properties確認】');
  try {
    const props = PropertiesService.getScriptProperties();
    const keys = ['LINE_CHANNEL_ACCESS_TOKEN', 'LINE_CHANNEL_SECRET', 'GEMINI_API_KEY', 'SPREADSHEET_ID', 'FOLDER_ID'];
    const missing = keys.filter(k => !props.getProperty(k));
    if (missing.length > 0) {
      log('❌ 未設定: ' + missing.join(', '));
    } else {
      log('✅ 全ての設定が存在します');
    }
  } catch (e) {
    log('❌ エラー: ' + e);
  }
  log('');

  // 2. 最近のdoPost実行を確認
  log('【2. 最近のWebhook呼び出し確認】');
  try {
    const cache = CacheService.getScriptCache();
    const lastWebhookTime = cache.get('last_webhook_time');
    if (lastWebhookTime) {
      log('✅ 最後のWebhook: ' + lastWebhookTime);
    } else {
      log('⚠️ 最近のWebhook呼び出しが記録されていません');
      log('   → LINEからメッセージを送っても反応がない場合、Webhook URLの設定を確認してください');
    }
  } catch (e) {
    log('⚠️ キャッシュ確認エラー: ' + e);
  }
  log('');

  // 3. 画像ジョブキューの確認
  log('【3. 画像処理キュー確認】');
  try {
    const queueLength = getImageJobQueueLength_();
    if (queueLength > 0) {
      log('⚠️ 処理待ちの画像: ' + queueLength + '件');
      log('   → processPendingImageJobs_ を手動実行してみてください');
    } else {
      log('✅ キューは空です');
    }
  } catch (e) {
    log('❌ エラー: ' + e);
  }
  log('');

  // 4. トリガーの確認
  log('【4. トリガー確認】');
  try {
    const triggers = ScriptApp.getProjectTriggers();
    log('登録済みトリガー数: ' + triggers.length);
    triggers.forEach(function(trigger) {
      log('  - ' + trigger.getHandlerFunction() + ' (' + trigger.getTriggerSource() + ')');
    });
    if (triggers.length === 0) {
      log('⚠️ トリガーが登録されていません（通常は問題ありません）');
    }
  } catch (e) {
    log('❌ エラー: ' + e);
  }
  log('');

  // 5. Webhook URLの表示
  log('【5. Webhook URL】');
  try {
    const scriptId = ScriptApp.getScriptId();
    const webappUrl = 'https://script.google.com/macros/s/' + scriptId + '/exec';
    log('あなたのWebhook URL:');
    log(webappUrl);
    log('');
    log('このURLをLINE Developersコンソールに設定してください');
  } catch (e) {
    log('❌ エラー: ' + e);
  }
  log('');

  // 6. テスト実行の提案
  log('【6. 次のステップ】');
  log('1. LINE Developersコンソールで上記のWebhook URLを設定');
  log('2. Webhookを「有効」にする');
  log('3. LINEで「ヘルプ」と送信してテスト');
  log('4. 画像を送信してテスト');
  log('');
  // 7. 最後のdoPostログを表示
  log('【7. 最後のdoPost実行ログ】');
  try {
    const cache = CacheService.getScriptCache();
    const lastLog = cache.get('last_dopost_log');
    if (lastLog) {
      log(lastLog);
    } else {
      log('⚠️ doPostのログが記録されていません');
      log('   → LINEからメッセージを送ってから再度実行してください');
    }
  } catch (e) {
    log('❌ エラー: ' + e);
  }
  log('');

  // 8. 最後の画像処理エラーを表示
  log('【8. 最後の画像処理エラー】');
  try {
    const cache = CacheService.getScriptCache();
    const lastError = cache.get('last_image_error');
    if (lastError) {
      log('❌ エラー発生:');
      log(lastError);
    } else {
      log('✅ 最近のエラーはありません');
    }
  } catch (e) {
    log('❌ エラー: ' + e);
  }
  log('');

  log('=== 診断完了 ===');
  
  return results.join('\n');
}

/**
 * LINE APIトークンの有効性をテストする
 * GASエディタで実行して、結果を確認してください
 */
function testLineApiToken() {
  console.log('=== LINE APIトークンテスト開始 ===');
  
  const token = CONFIG.LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) {
    console.log('❌ LINE_CHANNEL_ACCESS_TOKEN が設定されていません');
    return;
  }
  console.log('トークンの長さ: ' + token.length);
  console.log('トークンの先頭: ' + token.substring(0, 10) + '...');
  
  // 1. Bot Info APIでトークンを確認
  console.log('');
  console.log('【1. Bot Info API テスト】');
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const status = res.getResponseCode();
    const body = res.getContentText();
    console.log('ステータス: ' + status);
    console.log('レスポンス: ' + body);
    
    if (status >= 200 && status < 300) {
      console.log('✅ トークンは有効です');
      const info = JSON.parse(body);
      console.log('Bot名: ' + (info.displayName || '不明'));
    } else {
      console.log('❌ トークンが無効です！ステータス: ' + status);
      console.log('→ LINE Developersコンソールで新しいトークンを発行してください');
    }
  } catch (e) {
    console.log('❌ エラー: ' + e);
  }
  
  // 2. Reply APIエンドポイントの確認
  console.log('');
  console.log('【2. Reply APIエンドポイント確認】');
  console.log('LINE_REPLY URL: ' + API_ENDPOINTS.LINE_REPLY);
  console.log('LINE_PUSH URL: ' + API_ENDPOINTS.LINE_PUSH);
  
  // 3. Channel Secret確認
  console.log('');
  console.log('【3. Channel Secret確認】');
  const secret = CONFIG.LINE_CHANNEL_SECRET;
  if (secret) {
    console.log('✅ LINE_CHANNEL_SECRET: 設定済み（長さ: ' + secret.length + '）');
  } else {
    console.log('⚠️ LINE_CHANNEL_SECRET: 未設定（署名検証がスキップされます）');
  }
  
  console.log('');
  console.log('=== テスト完了 ===');
}

/**
 * 最近のエラーログを検索
 * 注意: Apps Scriptでは実行ログを直接取得できないため、
 * この関数は手動でログを確認する必要があることを通知します
 */
function checkRecentExecutionLogs() {
  const message = [
    '【実行ログの確認方法】',
    '',
    '1. GASエディタの左メニューで「実行数」をクリック',
    '2. 「関数」列で「doPost」を探す',
    '3. 最新の実行をクリック',
    '4. ログを確認',
    '',
    '【探すべきログ】',
    '✅ 正常: "=== doPost called at ..."',
    '❌ エラー: "doPost: ERROR - ..."',
    '',
    '【doPostが見つからない場合】',
    '→ Webhookが呼ばれていません',
    '→ LINE DevelopersでWebhook URLを確認してください'
  ].join('\n');
  
  console.log(message);
  return message;
}

/**
 * Webhook URLを取得して表示
 */
function showWebhookUrl() {
  const scriptId = ScriptApp.getScriptId();
  const webappUrl = 'https://script.google.com/macros/s/' + scriptId + '/exec';
  
  console.log('Webhook URL: ' + webappUrl);
  console.log('このURLをLINE Developersコンソールに設定してください');
  
  return webappUrl;
}

/**
 * 処理待ちの画像を手動で処理する
 * GASエディタから実行できるラッパー関数
 */
function processQueuedImages() {
  console.log('=== 処理待ち画像の処理を開始 ===');
  
  try {
    // キューの長さを確認
    const queueLength = getImageJobQueueLength_();
    console.log('処理待ちの画像: ' + queueLength + '件');
    
    if (queueLength === 0) {
      console.log('✅ 処理待ちの画像はありません');
      return '処理待ちの画像はありません';
    }
    
    // 処理を実行
    processPendingImageJobs_();
    
    console.log('=== 処理完了 ===');
    console.log('LINEを確認してください。解析結果が送信されているはずです。');
    
    return '処理完了。LINEを確認してください。';
  } catch (e) {
    console.error('エラー: ' + (e && e.stack ? e.stack : e));
    return 'エラーが発生しました: ' + e;
  }
}

/**
 * Gemini APIで利用可能なモデルを一覧表示
 */
function listAvailableGeminiModels() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    console.error('❌ GEMINI_API_KEY が設定されていません');
    return;
  }
  
  console.log('=== Gemini利用可能モデルを確認中 ===');
  console.log('');
  
  const url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + apiKey;
  
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const status = response.getResponseCode();
    
    if (status !== 200) {
      console.error('❌ ListModels API failed: status=' + status);
      console.error('Response: ' + response.getContentText());
      return;
    }
    
    const data = JSON.parse(response.getContentText());
    const models = data.models || [];
    
    console.log('総モデル数: ' + models.length + '個');
    console.log('');
    
    // Flash系モデルでgenerateContent対応のものをフィルタ
    const flashModels = models.filter(function(model) {
      const methods = model.supportedGenerationMethods || [];
      const supportsGenerate = methods.indexOf('generateContent') >= 0;
      const isFlash = model.name.indexOf('flash') >= 0 || model.name.indexOf('Flash') >= 0;
      return supportsGenerate && isFlash;
    });
    
    console.log('=== Flash系モデル（generateContent対応） ===');
    console.log('');
    
    if (flashModels.length === 0) {
      console.warn('⚠️ Flash系モデルが見つかりませんでした');
      console.log('全モデルを表示します:');
      models.forEach(function(model) {
        console.log('- ' + model.name);
        console.log('  methods: ' + (model.supportedGenerationMethods || []).join(', '));
      });
      return;
    }
    
    flashModels.forEach(function(model, index) {
      console.log('【' + (index + 1) + '】 ' + model.name);
      console.log('  displayName: ' + (model.displayName || 'N/A'));
      console.log('  methods: ' + (model.supportedGenerationMethods || []).join(', '));
      console.log('');
    });
    
    const recommended = flashModels[0].name;
    console.log('=== 推奨設定 ===');
    console.log('config.gs の GEMINI_GENERATE_CONTENT を以下に変更:');
    console.log('');
    console.log("'https://generativelanguage.googleapis.com/v1beta/" + recommended + ":generateContent?key=' + CONFIG.GEMINI_API_KEY");
    console.log('');
    
    return flashModels;
  } catch (error) {
    console.error('❌ ListModels error: ' + error);
    if (error.stack) {
      console.error(error.stack);
    }
  }
}

/**
 * Gemini APIキーを更新する (セキュリティのため、関数内にキーを直接書かない)
 * 引数にキーを渡して実行するか、GASエディタのプロパティ設定から直接更新してください
 */
function updateGeminiApiKey(newKey) {
  if (!newKey) {
    console.error('❌ キーが指定されていません。 updateGeminiApiKey("あなたのキー") のように実行してください。');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', newKey);
  console.log('✅ Gemini APIキーを更新しました: ' + newKey.slice(0, 8) + '...');
  return 'APIキーを更新しました。';
}
