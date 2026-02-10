function extractLineSignature_(eventObj) {
  if (!eventObj) {
    return null;
  }

  const headers = eventObj.headers || eventObj.header || null;
  if (headers && typeof headers === 'object') {
    const key = Object.keys(headers).find(function(name) {
      return name && name.toLowerCase() === 'x-line-signature';
    });
    if (key) {
      return headers[key];
    }
  }

  const parameter = eventObj.parameter || {};
  return (
    parameter['x-line-signature'] ||
    parameter.x_line_signature ||
    parameter.signature ||
    null
  );
}

function verifyLineSignature_(requestBody, signatureHeader) {
  if (!CONFIG.LINE_CHANNEL_SECRET) {
    console.warn('LINE_CHANNEL_SECRET is not set. Signature verification skipped.');
    return true;
  }

  if (!signatureHeader) {
    console.warn('x-line-signature header is not available in event object. Verification skipped.');
    return true;
  }

  const digest = Utilities.computeHmacSha256Signature(requestBody, CONFIG.LINE_CHANNEL_SECRET);
  const expected = Utilities.base64Encode(digest);
  return constantTimeEquals_(expected, String(signatureHeader));
}

function constantTimeEquals_(a, b) {
  if (a.length !== b.length) {
    return false;
  }
  let diff = 0;
  for (let i = 0; i < a.length; i += 1) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

function replyLineText_(replyToken, messageText) {
  if (!replyToken) {
    return;
  }

  const payload = {
    replyToken: replyToken,
    messages: [
      {
        type: 'text',
        text: trimToLineMessageLimit_(messageText)
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    headers: {
      Authorization: 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  return executeWithRetry_(function() {
    const res = UrlFetchApp.fetch(API_ENDPOINTS.LINE_REPLY, options);
    const status = res.getResponseCode();
    if (status >= 200 && status < 300) {
      return res;
    }
    throw new Error('LINE reply failed: status=' + status + ' body=' + res.getContentText());
  }, SETTINGS.API_MAX_RETRY, SETTINGS.API_RETRY_BASE_MS, 'LINE reply');
}

function pushLineText_(to, messageText) {
  if (!to) {
    return;
  }

  const payload = {
    to: to,
    messages: [
      {
        type: 'text',
        text: trimToLineMessageLimit_(messageText)
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    headers: {
      Authorization: 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  return executeWithRetry_(function() {
    const res = UrlFetchApp.fetch(API_ENDPOINTS.LINE_PUSH, options);
    const status = res.getResponseCode();
    if (status >= 200 && status < 300) {
      return res;
    }
    throw new Error('LINE push failed: status=' + status + ' body=' + res.getContentText());
  }, SETTINGS.API_MAX_RETRY, SETTINGS.API_RETRY_BASE_MS, 'LINE push');
}

function fetchLineMessageContent_(messageId) {
  const url = API_ENDPOINTS.LINE_CONTENT_BASE + '/' + encodeURIComponent(messageId) + '/content';
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN
    },
    muteHttpExceptions: true
  };

  const response = executeWithRetry_(function() {
    const res = UrlFetchApp.fetch(url, options);
    const status = res.getResponseCode();
    if (status >= 200 && status < 300) {
      return res;
    }
    throw new Error('LINE content fetch failed: status=' + status + ' body=' + res.getContentText());
  }, SETTINGS.API_MAX_RETRY, SETTINGS.API_RETRY_BASE_MS, 'LINE content');

  const headers = response.getAllHeaders();
  const mimeType = findHeaderValue_(headers, 'content-type') || response.getBlob().getContentType() || 'image/jpeg';

  return {
    bytes: response.getContent(),
    mimeType: mimeType
  };
}

function findHeaderValue_(headers, targetName) {
  if (!headers || typeof headers !== 'object') {
    return null;
  }
  const key = Object.keys(headers).find(function(name) {
    return name && name.toLowerCase() === targetName.toLowerCase();
  });
  return key ? headers[key] : null;
}
