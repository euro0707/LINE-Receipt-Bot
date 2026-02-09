function appendReceiptToSheet_(analysis) {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const purchaseDate = parseIsoDate_(analysis.date) || new Date();
  const sheet = getOrCreateMonthlySheet_(spreadsheet, purchaseDate);

  const recordedAt = formatTimestamp_(new Date());
  const purchaseDateText = formatDateOnly_(purchaseDate);

  const items = analysis.items && analysis.items.length > 0
    ? analysis.items
    : [
        {
          name: '(明細なし)',
          quantity: '',
          price: toNumber_(analysis.total),
          category: 'その他'
        }
      ];

  const rows = items.map(function(item) {
    const quantity = item.quantity === '' ? '' : toNumber_(item.quantity) || 1;
    return [
      recordedAt,
      safeText_(analysis.store_name, ''),
      purchaseDateText,
      safeText_(analysis.receipt_type, 'その他'),
      safeText_(item.name, '(不明)'),
      quantity,
      toNumber_(item.price),
      sanitizeCategory_(item.category),
      safeText_(analysis.payment_method, ''),
      toNumber_(analysis.total)
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
  return rows.length;
}

function appendManualExpenseToSheet_(analysis) {
  const targetDate = parseIsoDate_(analysis.date) || new Date();
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = getOrCreateMonthlySheet_(spreadsheet, targetDate);

  const row = [[
    formatTimestamp_(new Date()),
    safeText_(analysis.store_name, ''),
    formatDateOnly_(targetDate),
    '手動入力',
    safeText_(analysis.item_name, '(不明)'),
    1,
    toNumber_(analysis.amount),
    sanitizeCategory_(analysis.category),
    '',
    toNumber_(analysis.amount)
  ]];

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 10).setValues(row);
  return 1;
}

function buildMonthlySummary_(targetDate) {
  const date = targetDate || new Date();
  const monthLabel = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年MM月');
  const categories = createCategorySummary_();

  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(monthLabel);
  if (!sheet) {
    return {
      monthLabel: monthLabel,
      total: 0,
      categories: categories
    };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return {
      monthLabel: monthLabel,
      total: 0,
      categories: categories
    };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  let total = 0;

  values.forEach(function(row) {
    const amount = toNumber_(row[6]);
    if (amount <= 0) {
      return;
    }
    total += amount;

    const categoryText = safeText_(row[7], 'その他');
    const category = SETTINGS.SUMMARY_CATEGORIES.indexOf(categoryText) >= 0 ? categoryText : 'その他';
    categories[category] += amount;
  });

  return {
    monthLabel: monthLabel,
    total: total,
    categories: categories
  };
}

function getOrCreateMonthlySheet_(spreadsheet, targetDate) {
  const name = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy年MM月');
  let sheet = spreadsheet.getSheetByName(name);
  if (sheet) {
    return sheet;
  }

  sheet = spreadsheet.insertSheet(name);
  const headers = [[
    '記録日時',
    '店舗名',
    '購入日',
    'レシート種別',
    '品名',
    '個数',
    '金額',
    'カテゴリ',
    '支払い方法',
    '合計金額'
  ]];

  sheet.getRange(1, 1, 1, 10).setValues(headers);
  sheet.setFrozenRows(1);
  return sheet;
}

function createCategorySummary_() {
  const summary = {};
  SETTINGS.SUMMARY_CATEGORIES.forEach(function(category) {
    summary[category] = 0;
  });
  return summary;
}

function debugCheckSpreadsheet() {
  assertRequiredConfig_();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const message = [
    'Spreadsheet connected.',
    'id=' + CONFIG.SPREADSHEET_ID,
    'name=' + ss.getName(),
    'url=' + ss.getUrl()
  ].join('\n');
  Logger.log(message);
  return message;
}

function debugWriteTestRow() {
  assertRequiredConfig_();
  const now = new Date();
  appendManualExpenseToSheet_({
    store_name: 'debug',
    item_name: 'テスト入力',
    amount: 123,
    category: 'その他',
    date: formatDateOnly_(now)
  });
  const monthName = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月');
  const result = 'Wrote a test row to "' + monthName + '".';
  Logger.log(result);
  return result;
}
