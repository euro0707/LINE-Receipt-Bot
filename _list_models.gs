/**
 * Gemini APIで利用可能なモデルを一覧表示
 */
function listGeminiModels() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    console.error('GEMINI_API_KEY が設定されていません');
    return;
  }
  
  const url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + apiKey;
  
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const status = response.getResponseCode();
    
    if (status !== 200) {
      console.error('ListModels failed: status=' + status);
      console.error('Response: ' + response.getContentText());
      return;
    }
    
    const data = JSON.parse(response.getContentText());
    const models = data.models || [];
    
    console.log('=== 利用可能なGeminiモデル (v1beta) ===');
    console.log('総数: ' + models.length + '個');
    console.log('');
    
    // Vision対応モデル（generateContentサポート）をフィルタ
    const visionModels = models.filter(function(model) {
      const methods = model.supportedGenerationMethods || [];
      const supportsGenerate = methods.indexOf('generateContent') >= 0;
      const isFlash = model.name.indexOf('flash') >= 0;
      return supportsGenerate && isFlash;
    });
    
    console.log('=== Flash系モデル（generateContent対応） ===');
    visionModels.forEach(function(model) {
      console.log('✅ ' + model.name);
      console.log('   displayName: ' + (model.displayName || 'N/A'));
      console.log('   description: ' + (model.description || 'N/A'));
      console.log('   methods: ' + (model.supportedGenerationMethods || []).join(', '));
      console.log('');
    });
    
    if (visionModels.length > 0) {
      const recommended = visionModels[0].name;
      console.log('=== 推奨モデル ===');
      console.log(recommended);
      console.log('');
      console.log('config.gsで以下のように設定してください:');
      console.log("'https://generativelanguage.googleapis.com/v1beta/" + recommended + ":generateContent?key=' + CONFIG.GEMINI_API_KEY");
    }
    
    return visionModels;
  } catch (error) {
    console.error('ListModels error: ' + error);
    console.error(error.stack);
  }
}
