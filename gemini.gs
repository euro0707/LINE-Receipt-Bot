const RECEIPT_PROMPT = [
  'あなたはレシートの画像から情報を抽出するエキスパートです。',
  '以下の項目をJSON形式で出力してください。',
  '読み取れない項目はnullとしてください。',
  '',
  '- "store_name": 店舗名（自然な日本語表記に変換）',
  '- "date": 購入日 "YYYY-MM-DD" 形式',
  '- "receipt_type": レシート種別（"スーパーマーケット", "コンビニ", "ドラッグストア", "飲食店", "医療", "薬局", "交通", "その他"）',
  '- "items": 各品目の配列',
  '  - "name": 品名',
  '  - "quantity": 個数',
  '  - "price": 金額（税込）',
  '  - "category": カテゴリ（"食費", "日用品", "交際費", "交通費", "医療費", "その他"）',
  '- "total": 実際に支払った最終的な税込み合計金額（数値）',
  '- "payment_method": 支払い方法',
  '',
  '医療系レシート（クリニック、病院、薬局）の場合はitemsは空配列[]でよい。',
  'JSONのみを返し、それ以外のテキストは含めないでください。'
].join('\n');

const RECEIPT_TYPES = ['スーパーマーケット', 'コンビニ', 'ドラッグストア', '飲食店', '医療', '薬局', '交通', 'その他'];

function analyzeReceiptImage_(imageBytes, mimeType) {
  const request = {
    contents: [
      {
        role: 'user',
        parts: [
          { text: RECEIPT_PROMPT },
          {
            inline_data: {
              mime_type: mimeType || 'image/jpeg',
              data: Utilities.base64Encode(imageBytes)
            }
          }
        ]
      }
    ],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0.1
    }
  };

  const json = callGeminiAndParseJson_(request);
  return normalizeReceiptResult_(json);
}

function analyzeTextExpense_(userInput, nowDate) {
  const today = formatDateOnly_(nowDate);
  const prompt = [
    '現在の日付は「' + today + '」です。',
    'ユーザーが入力した以下のテキストから、家計簿に記録するための情報を抽出してください。',
    'ユーザー入力テキスト: "' + escapeForPrompt_(userInput) + '"',
    '',
    '以下の項目をJSON形式で出力してください。',
    '- "item_name": 品目名・内容',
    '- "amount": 金額（数値）',
    '- "category": カテゴリ（"食費", "日用品", "交際費", "交通費", "医療費", "その他"）',
    '- "store_name": 店舗名（わかれば。不明ならnull）',
    '- "date": 取引日 "YYYY-MM-DD" 形式。「昨日」「今日」等は現在日付を基準に解釈。日付情報がなければnull',
    'JSONのみを返してください。'
  ].join('\n');

  const request = {
    contents: [
      {
        role: 'user',
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0.1
    }
  };

  const json = callGeminiAndParseJson_(request);
  return normalizeTextResult_(json, nowDate);
}

function callGeminiAndParseJson_(requestBody) {
  const options = {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  const response = executeWithRetry_(function() {
    const res = UrlFetchApp.fetch(API_ENDPOINTS.GEMINI_GENERATE_CONTENT, options);
    const status = res.getResponseCode();
    if (status >= 200 && status < 300) {
      return res;
    }
    throw new Error('Gemini API failed: status=' + status + ' body=' + res.getContentText());
  }, SETTINGS.API_MAX_RETRY, SETTINGS.API_RETRY_BASE_MS, 'Gemini API');

  const responseJson = JSON.parse(response.getContentText());
  const generatedText = extractGeminiResponseText_(responseJson);
  return parseJsonWithFallback_(generatedText);
}

function extractGeminiResponseText_(responseJson) {
  const candidates = responseJson.candidates || [];
  if (!candidates.length) {
    throw new Error('Gemini response has no candidates: ' + JSON.stringify(responseJson));
  }

  const parts = (candidates[0].content && candidates[0].content.parts) || [];
  const text = parts
    .map(function(part) {
      return part.text || '';
    })
    .join('\n')
    .trim();

  if (!text) {
    throw new Error('Gemini response text is empty: ' + JSON.stringify(responseJson));
  }

  return text;
}

function parseJsonWithFallback_(rawText) {
  const cleaned = stripCodeFence_(String(rawText || '').trim());

  try {
    return JSON.parse(cleaned);
  } catch (firstError) {
    const start = cleaned.indexOf('{');
    const end = cleaned.lastIndexOf('}');
    if (start >= 0 && end > start) {
      const fragment = cleaned.slice(start, end + 1);
      return JSON.parse(fragment);
    }
    throw firstError;
  }
}

function stripCodeFence_(text) {
  return text
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/, '')
    .trim();
}

function normalizeReceiptResult_(json) {
  const items = Array.isArray(json.items)
    ? json.items
        .map(function(item) {
          return {
            name: safeText_(item && item.name, ''),
            quantity: toNumber_(item && item.quantity) || 1,
            price: toNumber_(item && item.price),
            category: sanitizeCategory_(item && item.category)
          };
        })
        .filter(function(item) {
          return item.name || item.price > 0;
        })
    : [];

  const computedTotal = items.reduce(function(sum, item) {
    return sum + toNumber_(item.price);
  }, 0);

  const total = toNumber_(json.total) || computedTotal;

  return {
    store_name: safeText_(json.store_name, null),
    date: normalizeDateString_(json.date) || formatDateOnly_(new Date()),
    receipt_type: sanitizeReceiptType_(json.receipt_type),
    items: items,
    total: total,
    payment_method: safeText_(json.payment_method, null)
  };
}

function normalizeTextResult_(json, nowDate) {
  const parsedDate = normalizeDateString_(json.date) || formatDateOnly_(nowDate);

  return {
    item_name: safeText_(json.item_name || json.name, ''),
    amount: toNumber_(json.amount),
    category: sanitizeCategory_(json.category),
    store_name: safeText_(json.store_name, null),
    date: parsedDate
  };
}

function normalizeDateString_(value) {
  if (!value) {
    return null;
  }

  const parsed = parseIsoDate_(String(value));
  return parsed ? formatDateOnly_(parsed) : null;
}

function sanitizeReceiptType_(value) {
  const text = safeText_(value, 'その他');
  return RECEIPT_TYPES.indexOf(text) >= 0 ? text : 'その他';
}

function sanitizeCategory_(value) {
  const text = safeText_(value, 'その他');
  return SETTINGS.SUMMARY_CATEGORIES.indexOf(text) >= 0 ? text : 'その他';
}

function escapeForPrompt_(text) {
  return String(text)
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"');
}
