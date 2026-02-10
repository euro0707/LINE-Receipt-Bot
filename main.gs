const APP_VERSION = '2026-02-11-v20';
const IMAGE_JOB_QUEUE_KEY = 'image_job_queue_v1';
const IMAGE_JOB_BATCH_SIZE = 2;

function doGet() {
  return createTextResponse_('LINE receipt bot is running. ' + APP_VERSION);
}

/**
 * 処理待ちの画像を手動で処理する
 * GASエディタの関数リストから実行できます
 */
function processQueuedImages() {
  console.log('=== 処理待ち画像の処理を開始 ===');
  
  try {
    const queueLength = getImageJobQueueLength_();
    console.log('処理待ちの画像: ' + queueLength + '件');
    
    if (queueLength === 0) {
      console.log('✅ 処理待ちの画像はありません');
      return '処理待ちの画像はありません';
    }
    
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
 * 画像処理キューをクリアする
 * 古いメッセージIDが残っている場合に使用
 */
function clearImageQueue() {
  console.log('=== 画像処理キューのクリア ===');
  
  try {
    const beforeLength = getImageJobQueueLength_();
    console.log('クリア前のキュー: ' + beforeLength + '件');
    
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(IMAGE_JOB_QUEUE_KEY);
    
    const afterLength = getImageJobQueueLength_();
    console.log('クリア後のキュー: ' + afterLength + '件');
    console.log('✅ キューをクリアしました');
    
    return 'キューをクリアしました（' + beforeLength + '件削除）';
  } catch (e) {
    console.error('エラー: ' + (e && e.stack ? e.stack : e));
    return 'エラーが発生しました: ' + e;
  }
}

function doPost(e) {
  const debugLog = [];
  const dlog = function(msg) {
    console.log(msg);
    debugLog.push(new Date().toISOString().slice(11,19) + ' ' + msg);
  };
  
  dlog('=== doPost called === VERSION=' + APP_VERSION);
  
  // Webhook呼び出しを記録
  try {
    const cache = CacheService.getScriptCache();
    cache.put('last_webhook_time', new Date().toISOString(), 600);
  } catch (cacheError) {
    dlog('WARN: cache write failed: ' + cacheError);
  }
  
  try {
    assertRequiredConfig_();
    dlog('config OK');

    const requestBody = e && e.postData && e.postData.contents ? e.postData.contents : '';
    if (!requestBody) {
      dlog('empty body');
      saveDebugLog_(debugLog);
      return createTextResponse_('ok');
    }
    dlog('body len=' + requestBody.length);

    const signature = extractLineSignature_(e);
    dlog('signature=' + (signature ? 'found' : 'NOT FOUND'));
    
    if (!verifyLineSignature_(requestBody, signature)) {
      dlog('❌ SIGNATURE VERIFICATION FAILED');
      saveDebugLog_(debugLog);
      return createTextResponse_('ok');
    }
    dlog('signature verified OK');

    const payload = JSON.parse(requestBody);
    const events = payload.events || [];
    dlog('events count=' + events.length);
    
    events.forEach(function(event, i) {
      dlog('event[' + i + '] type=' + (event ? event.type : 'null') + ' msgType=' + (event && event.message ? event.message.type : 'N/A'));
      dlog('event[' + i + '] replyToken=' + (event && event.replyToken ? 'exists' : 'null'));
      dlog('event[' + i + '] source=' + JSON.stringify(event && event.source ? event.source : null));
      
      try {
        processWebhookEventSafely_(event);
        dlog('event[' + i + '] processed OK');
      } catch (evtErr) {
        dlog('event[' + i + '] ERROR: ' + evtErr);
      }
    });

    dlog('=== doPost completed ===');
    saveDebugLog_(debugLog);
    return createTextResponse_('ok');
  } catch (error) {
    dlog('❌ FATAL ERROR: ' + (error && error.stack ? error.stack : error));
    saveDebugLog_(debugLog);
    return createTextResponse_('ok');
  }
}

function saveDebugLog_(logLines) {
  try {
    const cache = CacheService.getScriptCache();
    cache.put('last_dopost_log', logLines.join('\n'), 600);
  } catch (e) {
    console.warn('saveDebugLog_ failed: ' + e);
  }
}

function processWebhookEventSafely_(event) {
  try {
    processWebhookEvent_(event);
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    const pushTarget = extractPushTargetId_(event);
    notifyUserByPushOrReply_(event, pushTarget, 'システムエラーが発生しました。時間をおいてお試しください。');
  }
}

function processWebhookEvent_(event) {
  console.log('processWebhookEvent_: start eventType=' + (event ? event.type : 'null'));
  
  if (!event || event.mode === 'standby') {
    console.log('processWebhookEvent_: skipped (standby or null)');
    return;
  }
  if (event.type !== 'message' || !event.message) {
    console.log('processWebhookEvent_: skipped (not message type)');
    return;
  }
  if (isDuplicateEvent_(event)) {
    console.log('processWebhookEvent_: skipped (duplicate)');
    return;
  }

  const replyToken = event.replyToken;
  if (!replyToken) {
    console.log('processWebhookEvent_: no replyToken');
    return;
  }

  const messageType = event.message.type;
  console.log('processWebhookEvent_: messageType=' + messageType);
  
  switch (messageType) {
    case 'image':
      console.log('processWebhookEvent_: calling handleImageMessage_');
      handleImageMessage_(event);
      return;
    case 'text':
      console.log('processWebhookEvent_: calling handleTextMessage_');
      handleTextMessage_(event);
      return;
    default:
      console.log('processWebhookEvent_: unknown message type, sending help');
      replyLineText_(replyToken, buildHelpMessage_());
  }
}

function handleImageMessage_(event) {
  const messageId = event.message.id;
  const pushTarget = extractPushTargetId_(event);
  console.log('handleImageMessage_: messageId=' + messageId + ' pushTarget=' + pushTarget);
  
  if (!messageId) {
    notifyUserByPushOrReply_(event, pushTarget, '画像IDが取得できませんでした。もう一度お試しください。');
    return;
  }

  // 受付メッセージを送信
  const acceptedMessage = '画像を受け取りました。解析中です...';
  const acceptedByReply = safeReplyLineText_(event.replyToken, acceptedMessage);
  if (acceptedByReply) {
    event.__replied = true;
  }

  // 直接画像を処理（キュー・トリガーを使わない）
  try {
    console.log('handleImageMessage_: fetching content');
    const content = fetchLineMessageContent_(messageId);
    console.log('handleImageMessage_: content fetched bytes=' + content.bytes.length);

    console.log('handleImageMessage_: saving to Drive');
    const savedFile = saveReceiptImageToDrive_(content.bytes, content.mimeType, new Date(), messageId);
    console.log('handleImageMessage_: image saved fileId=' + savedFile.id);

    console.log('handleImageMessage_: analyzing image');
    const analysis = analyzeReceiptImage_(content.bytes, content.mimeType);
    console.log('handleImageMessage_: analysis done');

    const hasItems = analysis.items && analysis.items.length > 0;
    if (!hasItems && !toNumber_(analysis.total)) {
      console.log('handleImageMessage_: no items detected');
      safePushLineText_(pushTarget, 'レシートを読み取れませんでした。別の画像でお試しください。');
      return;
    }

    console.log('handleImageMessage_: appending to sheet');
    appendReceiptToSheet_(analysis);
    console.log('handleImageMessage_: building summary');
    const summary = buildMonthlySummary_(new Date());
    console.log('handleImageMessage_: formatting reply');
    const message = formatReceiptReply_(analysis, summary, savedFile);
    console.log('handleImageMessage_: sending result');
    safePushLineText_(pushTarget, message);
    console.log('handleImageMessage_: completed successfully');
  } catch (error) {
    const errorDetails = (error && error.stack ? error.stack : String(error));
    console.error('handleImageMessage_: ERROR - ' + errorDetails);
    
    // エラー詳細をキャッシュに保存
    try {
      const cache = CacheService.getScriptCache();
      cache.put('last_image_error', errorDetails, 600);
    } catch (cacheErr) {
      console.warn('Failed to cache error: ' + cacheErr);
    }
    
    safePushLineText_(pushTarget, '画像の解析中にエラーが発生しました。時間をおいて再度お試しください。');
  }
}

function handleTextMessage_(event) {
  const text = (event.message.text || '').trim();
  if (!text) {
    replyLineText_(event.replyToken, buildHelpMessage_());
    return;
  }

  if (text === 'ヘルプ') {
    replyLineText_(event.replyToken, buildHelpMessage_());
    return;
  }

  if (text === '今月' || text === '先月') {
    const targetDate = text === '今月' ? new Date() : addMonths_(new Date(), -1);
    const summary = buildMonthlySummary_(targetDate);
    replyLineText_(event.replyToken, formatSummaryReply_(summary));
    return;
  }

  const nowDate = new Date();
  const analysis = parseSimpleExpenseText_(text, nowDate);
  if (!analysis || !analysis.item_name || !toNumber_(analysis.amount)) {
    replyLineText_(event.replyToken, '内容を解釈できませんでした。例: ランチ 1200円');
    return;
  }

  appendManualExpenseToSheet_(analysis);
  const summary = buildMonthlySummary_(new Date());
  replyLineText_(event.replyToken, formatTextRecordReply_(analysis, summary));
}

function buildHelpMessage_() {
  return [
    'version: ' + APP_VERSION,
    '',
    '使い方',
    '・レシート画像を送ると自動記録します',
    '・例: ランチ 1200円',
    '・「今月」: 今月の集計',
    '・「先月」: 先月の集計',
    '・「ヘルプ」: このメッセージ'
  ].join('\n');
}

function formatReceiptReply_(analysis, summary, savedFile) {
  const lines = [];
  lines.push('📝 レシート読み取り完了！');
  lines.push('');
  lines.push('🏪 ' + safeText_(analysis.store_name, '不明') + '（' + safeText_(analysis.receipt_type, 'その他') + '）');
  lines.push('📅 ' + formatDateDisplay_(analysis.date));

  const items = analysis.items || [];
  if (items.length > 0) {
    lines.push('━━━━━━━━━━');
    items.slice(0, 20).forEach(function(item) {
      const name = safeText_(item.name, '不明');
      const price = formatYen_(item.price);
      const category = safeText_(item.category, 'その他');
      lines.push('・' + name + ' ' + price + '【' + category + '】');
    });
    if (items.length > 20) {
      lines.push('... ほか ' + (items.length - 20) + '件');
    }
  }

  lines.push('━━━━━━━━━━');
  lines.push('💰 合計: ' + formatYen_(analysis.total));
  lines.push('');
  lines.push(formatSummaryBlock_(summary));
  if (savedFile && savedFile.url) {
    lines.push('');
    lines.push('画像保存: ' + savedFile.url);
  }
  lines.push('');
  lines.push('スプレッドシートに記録しました✅');

  return trimToLineMessageLimit_(lines.join('\n'));
}

function formatTextRecordReply_(analysis, summary) {
  const lines = [];
  lines.push('📝 記録しました！');
  lines.push('');
  lines.push('・' + safeText_(analysis.item_name, '項目') + ' ' + formatYen_(analysis.amount) + '【' + safeText_(analysis.category, 'その他') + '】');
  lines.push('📅 ' + formatDateDisplay_(analysis.date));
  lines.push('');
  lines.push(formatSummaryBlock_(summary));

  return trimToLineMessageLimit_(lines.join('\n'));
}

function formatSummaryReply_(summary) {
  return trimToLineMessageLimit_([
    '📊 ' + summary.monthLabel + ' の集計',
    '',
    '総合計: ' + formatYen_(summary.total),
    'カテゴリ別支出:',
    formatCategoryLines_(summary.categories)
  ].join('\n'));
}

function formatSummaryBlock_(summary) {
  return [
    '--- 今月の状況 ---',
    '総合計: ' + formatYen_(summary.total),
    'カテゴリ別支出:',
    formatCategoryLines_(summary.categories)
  ].join('\n');
}

function formatCategoryLines_(categories) {
  return SETTINGS.SUMMARY_CATEGORIES.map(function(category) {
    return '・' + category + ': ' + formatYen_(categories[category] || 0);
  }).join('\n');
}

function parseSimpleExpenseText_(text, nowDate) {
  const normalized = String(text || '')
    .replace(/\u3000/g, ' ')
    .replace(/[，]/g, ',')
    .replace(/[０-９]/g, function(ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0xfee0);
    })
    .trim();

  if (!normalized) {
    return null;
  }

  let match = normalized.match(/^(.+?)\s+([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+)(?:\s*円)?$/);
  if (!match) {
    match = normalized.match(/^(.+?)([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+)\s*円?$/);
  }
  if (!match) {
    return null;
  }

  const itemName = safeText_(match[1], '');
  const amount = toNumber_(match[2]);
  if (!itemName || !amount) {
    return null;
  }

  return {
    item_name: itemName,
    amount: amount,
    category: inferCategoryFromText_(itemName),
    store_name: null,
    date: formatDateOnly_(nowDate || new Date())
  };
}

function inferCategoryFromText_(text) {
  const source = String(text || '');
  if (/電車|バス|タクシー|交通|切符|IC|定期/.test(source)) {
    return '交通費';
  }
  if (/病院|クリニック|薬|ドラッグ|処方/.test(source)) {
    return '医療費';
  }
  if (/飲み|会食|プレゼント|祝い|交際/.test(source)) {
    return '交際費';
  }
  if (/洗剤|ティッシュ|トイレット|日用品|シャンプー/.test(source)) {
    return '日用品';
  }
  if (/食|ランチ|朝食|夕食|弁当|カフェ|コーヒー|スーパー|コンビニ/.test(source)) {
    return '食費';
  }
  return 'その他';
}
function extractPushTargetId_(event) {
  if (!event || !event.source) {
    return null;
  }
  return (
    event.source.userId ||
    event.source.groupId ||
    event.source.roomId ||
    null
  );
}

function safePushLineText_(to, messageText) {
  if (!to) {
    return false;
  }
  try {
    pushLineText_(to, messageText);
    return true;
  } catch (error) {
    console.error('Push fallback failed: ' + (error && error.stack ? error.stack : error));
    return false;
  }
}

function safeReplyLineText_(replyToken, messageText) {
  try {
    replyLineText_(replyToken, messageText);
    return true;
  } catch (error) {
    console.error('Reply failed: ' + (error && error.stack ? error.stack : error));
    return false;
  }
}

function notifyUserByPushOrReply_(event, pushTarget, messageText) {
  if (pushTarget && safePushLineText_(pushTarget, messageText)) {
    return true;
  }

  if (event && event.replyToken && event.mode !== 'standby' && !event.__replied) {
    if (safeReplyLineText_(event.replyToken, messageText)) {
      event.__replied = true;
      return true;
    }
  }

  return false;
}

function enqueueImageJob_(job) {
  return mutateImageJobQueue_(function(queue) {
    queue.push(job);
    return queue.length;
  });
}

function dequeueImageJobBatch_(limit) {
  const maxJobs = Math.max(1, Number(limit) || 1);
  return mutateImageJobQueue_(function(queue) {
    const jobs = queue.splice(0, maxJobs);
    return {
      jobs: jobs,
      remaining: queue.length
    };
  });
}

function getImageJobQueueLength_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(IMAGE_JOB_QUEUE_KEY);
  if (!raw) {
    return 0;
  }

  try {
    const queue = JSON.parse(raw);
    return Array.isArray(queue) ? queue.length : 0;
  } catch (error) {
    console.warn('getImageJobQueueLength_: invalid queue JSON. ' + error);
    return 0;
  }
}

function mutateImageJobQueue_(mutator) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(IMAGE_JOB_QUEUE_KEY);
    let queue = [];

    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        queue = Array.isArray(parsed) ? parsed : [];
      } catch (parseError) {
        console.warn('mutateImageJobQueue_: failed to parse queue; reset. ' + parseError);
      }
    }

    const result = mutator(queue);
    if (queue.length > 0) {
      props.setProperty(IMAGE_JOB_QUEUE_KEY, JSON.stringify(queue));
    } else {
      props.deleteProperty(IMAGE_JOB_QUEUE_KEY);
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

function ensureImageWorkerTrigger_() {
  const handler = 'processPendingImageJobs_';
  const exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === handler;
  });
  if (exists) {
    return false;
  }

  ScriptApp.newTrigger(handler).timeBased().after(10000).create();
  return true;
}

function processPendingImageJobs_() {
  assertRequiredConfig_();

  const batch = dequeueImageJobBatch_(IMAGE_JOB_BATCH_SIZE);
  const jobs = batch.jobs || [];
  if (!jobs.length) {
    console.log('processPendingImageJobs_: no queued jobs');
    return;
  }

  console.log('processPendingImageJobs_: start jobs=' + jobs.length + ' remaining=' + batch.remaining);
  jobs.forEach(function(job) {
    processSingleImageJobSafely_(job);
  });

  if (batch.remaining > 0 || getImageJobQueueLength_() > 0) {
    ensureImageWorkerTrigger_();
  }
}

function processSingleImageJobSafely_(job) {
  const messageId = job && job.messageId ? String(job.messageId) : '';
  const pushTarget = job && job.pushTarget ? String(job.pushTarget) : '';
  if (!messageId) {
    console.warn('processSingleImageJobSafely_: skipped invalid job');
    return;
  }

  try {
    console.log('processSingleImageJobSafely_: fetch content messageId=' + messageId);
    const content = fetchLineMessageContent_(messageId);
    console.log('processSingleImageJobSafely_: content fetched bytes=' + content.bytes.length + ' mimeType=' + content.mimeType);

    const savedFile = saveReceiptImageToDrive_(content.bytes, content.mimeType, new Date(), messageId);
    console.log('processSingleImageJobSafely_: image saved fileId=' + savedFile.id);

    const analysis = analyzeReceiptImage_(content.bytes, content.mimeType);
    const itemCount = analysis.items && analysis.items.length ? analysis.items.length : 0;
    console.log('processSingleImageJobSafely_: analysis done itemCount=' + itemCount + ' total=' + toNumber_(analysis.total));

    const hasItems = analysis.items && analysis.items.length > 0;
    if (!hasItems && !toNumber_(analysis.total)) {
      safePushLineText_(pushTarget, 'レシートを読み取れませんでした。別の画像でお試しください。');
      return;
    }

    appendReceiptToSheet_(analysis);
    const summary = buildMonthlySummary_(new Date());
    const message = formatReceiptReply_(analysis, summary, savedFile);
    safePushLineText_(pushTarget, message);
  } catch (error) {
    console.error('processSingleImageJobSafely_ failed messageId=' + messageId + ': ' + (error && error.stack ? error.stack : error));
    safePushLineText_(pushTarget, '画像の解析中にエラーが発生しました。時間をおいて再度お試しください。');
  }
}

function isDuplicateEvent_(event) {
  const keySource =
    (event.message && event.message.id) ||
    event.webhookEventId ||
    event.replyToken;

  if (!keySource) {
    return false;
  }

  const cacheKey = 'line:dedupe:' + keySource;
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get(cacheKey)) {
      return true;
    }
    cache.put(cacheKey, '1', SETTINGS.DEDUPE_TTL_SECONDS);
    return false;
  } catch (error) {
    console.warn('CacheService unavailable: ' + error);
    return false;
  }
}

function executeWithRetry_(fn, maxAttempts, baseDelayMs, label) {
  let lastError = null;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return fn();
    } catch (error) {
      lastError = error;
      if (attempt === maxAttempts) {
        break;
      }
      const delayMs = baseDelayMs * Math.pow(2, attempt - 1);
      console.warn((label || 'retry') + ' failed. attempt=' + attempt + ' delayMs=' + delayMs);
      Utilities.sleep(delayMs);
    }
  }
  throw lastError;
}

function createTextResponse_(text) {
  return HtmlService.createHtmlOutput(String(text || 'ok'));
}

function parseIsoDate_(value) {
  if (!value || typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const isoDate = new Date(trimmed + 'T00:00:00+09:00');
    return isNaN(isoDate.getTime()) ? null : isoDate;
  }

  const normalized = trimmed.replace(/\//g, '-');
  const parsed = new Date(normalized);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDateOnly_(dateObj) {
  return Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function formatDateDisplay_(dateValue) {
  const parsed = dateValue instanceof Date ? dateValue : parseIsoDate_(dateValue);
  if (!parsed) {
    return '不明';
  }
  return Utilities.formatDate(parsed, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function formatTimestamp_(dateObj) {
  return Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

function addMonths_(dateObj, months) {
  const d = new Date(dateObj.getFullYear(), dateObj.getMonth(), 1);
  d.setMonth(d.getMonth() + months);
  return d;
}

function toNumber_(value) {
  if (typeof value === 'number') {
    return isFinite(value) ? value : 0;
  }
  if (typeof value !== 'string') {
    return 0;
  }
  const normalized = value.replace(/[^0-9.-]/g, '');
  const n = Number(normalized);
  return isFinite(n) ? n : 0;
}

function formatYen_(value) {
  const amount = Math.round(toNumber_(value));
  return '¥' + amount.toLocaleString('ja-JP');
}

function safeText_(value, fallback) {
  if (value === null || value === undefined) {
    return fallback;
  }
  const text = String(value).trim();
  return text ? text : fallback;
}

function trimToLineMessageLimit_(text) {
  const max = 4900;
  if (text.length <= max) {
    return text;
  }
  return text.slice(0, max - 4) + '\n...';
}

