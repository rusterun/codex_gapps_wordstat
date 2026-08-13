/**
 * Google Sheets → Apps Script → Yandex Cloud Search API Wordstat → Google Sheets.
 *
 * Before use, save Yandex Cloud credentials in Apps Script project properties:
 *   setWordstatCredentials('YOUR_API_KEY', 'YOUR_FOLDER_ID')
 * Do not store credentials in spreadsheet cells.
 */

var WORDSTAT_HEADERS = {
  PRODUCT: 'Полное наименование товара',
  QUERY: 'Wordstat-запрос',
  DEMAND: 'Спрос',
  UPDATED_AT: 'Дата обновления',
  COMMENT: 'Комментарий',
};

var WORDSTAT_CONFIG = {
  API_BASE_URL: 'https://searchapi.api.cloud.yandex.net',
  DYNAMICS_PATH: '/v2/wordstat/dynamics',
  HEALTHCHECK_PATH: '/v2/wordstat/getRegionsTree',
  API_KEY_PROPERTY: 'YANDEX_CLOUD_API_KEY',
  FOLDER_ID_PROPERTY: 'YANDEX_CLOUD_FOLDER_ID',
  CACHE_PROPERTY: 'WORDSTAT_DEMAND_CACHE_V2',
  CACHE_TTL_HOURS: 24 * 7,
  REQUEST_SLEEP_MS: 150,
  MAX_REQUESTS_PER_RUN: 40,
  DEFAULT_PERIOD: 'PERIOD_WEEKLY',
  LOOKBACK_WEEKS: 2,
  DEFAULT_REGIONS: [],
  DEFAULT_DEVICES: ['DEVICE_ALL'],
  MAX_CACHE_ENTRIES: 1000,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Wordstat')
    .addItem('Сгенерировать запросы', 'generateWordstatQueries')
    .addItem('Обновить спрос', 'updateWordstatDemand')
    .addItem('Принудительно обновить', 'forceUpdateWordstatDemand')
    .addSeparator()
    .addItem('Проверить подключение', 'checkWordstatConnection')
    .addToUi();
}

function setWordstatCredentials(apiKey, folderId) {
  var key = String(apiKey || '').trim();
  var folder = String(folderId || '').trim();
  if (!key) throw new Error('Передан пустой API-ключ Yandex Cloud.');
  if (!folder) throw new Error('Передан пустой folder ID Yandex Cloud.');
  PropertiesService.getScriptProperties().setProperties({
    YANDEX_CLOUD_API_KEY: key,
    YANDEX_CLOUD_FOLDER_ID: folder,
  });
}

function setWordstatOAuthToken() {
  throw new Error('OAuth-токен больше не используется. Выполните setWordstatCredentials("API_KEY", "FOLDER_ID").');
}

function generateWordstatQueries() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var columns = getHeaderColumns_(sheet, [WORDSTAT_HEADERS.PRODUCT, WORDSTAT_HEADERS.QUERY]);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return showAlert_('Нет строк с товарами для обработки.');

  var rowCount = lastRow - 1;
  var products = sheet.getRange(2, columns[WORDSTAT_HEADERS.PRODUCT], rowCount, 1).getValues();
  var queriesRange = sheet.getRange(2, columns[WORDSTAT_HEADERS.QUERY], rowCount, 1);
  var queries = queriesRange.getValues();
  var generated = 0;

  for (var i = 0; i < rowCount; i++) {
    if (String(products[i][0] || '').trim() && !String(queries[i][0] || '').trim()) {
      queries[i][0] = buildQueryFromProductName_(products[i][0]);
      if (queries[i][0]) generated++;
    }
  }

  if (generated) queriesRange.setValues(queries);
  showAlert_('Сгенерировано запросов: ' + generated);
}

function updateWordstatDemand() {
  updateWordstatDemand_(false);
}

function forceUpdateWordstatDemand() {
  updateWordstatDemand_(true);
}

function checkWordstatConnection() {
  var credentials = getCredentials_();
  if (!credentials.ok) return showAlert_(credentials.error);

  var result = fetchWordstat_(WORDSTAT_CONFIG.HEALTHCHECK_PATH, {}, credentials);
  if (result.ok) {
    showAlert_('Подключение успешно: API-ключ и folder ID сохранены, Search API доступен, авторизация работает.');
  } else {
    showAlert_('Подключение не удалось: ' + sanitizeError_(result.error));
  }
}

function updateWordstatDemand_(forceRefresh) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var columns = getHeaderColumns_(sheet, [WORDSTAT_HEADERS.QUERY, WORDSTAT_HEADERS.DEMAND, WORDSTAT_HEADERS.UPDATED_AT]);
  var credentials = getCredentials_();
  if (!credentials.ok) return showAlert_(credentials.error);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return showAlert_('Нет строк с запросами для обработки.');

  var rowCount = lastRow - 1;
  var queryValues = sheet.getRange(2, columns[WORDSTAT_HEADERS.QUERY], rowCount, 1).getValues();
  var queries = queryValues.map(function(row) { return normalizeQuery_(row[0]); });
  var demandRange = sheet.getRange(2, columns[WORDSTAT_HEADERS.DEMAND], rowCount, 1);
  var dateRange = sheet.getRange(2, columns[WORDSTAT_HEADERS.UPDATED_AT], rowCount, 1);
  var demandValues = demandRange.getValues();
  var dateValues = dateRange.getValues();
  var uniqueQueries = Array.from(new Set(queries.filter(Boolean)));
  var cache = loadCache_();
  var results = {};
  var now = new Date();
  var stats = { processed: 0, fromCache: 0, requested: 0, errors: 0, skippedByLimit: 0 };
  var errorSamples = [];
  var requestsSent = 0;

  uniqueQueries.forEach(function(query) {
    var cached = cache[query];
    if (!forceRefresh && cached && isCacheFresh_(cached, now)) {
      results[query] = { ok: true, count: cached.count };
      stats.fromCache++;
      return;
    }

    if (requestsSent >= WORDSTAT_CONFIG.MAX_REQUESTS_PER_RUN) {
      stats.skippedByLimit++;
      return;
    }

    var response = fetchDemandForQuery_(query, credentials);
    requestsSent++;
    stats.requested++;
    if (response.ok) {
      results[query] = { ok: true, count: response.count };
      cache[query] = { count: response.count, fetchedAt: now.toISOString() };
    } else {
      var errorText = sanitizeError_(response.error);
      results[query] = { ok: false, error: errorText };
      stats.errors++;
      if (errorSamples.length < 5) errorSamples.push(query + ': ' + errorText);
    }
    Utilities.sleep(WORDSTAT_CONFIG.REQUEST_SLEEP_MS);
  });

  for (var i = 0; i < queries.length; i++) {
    var query = queries[i];
    if (!query) continue;
    var result = results[query];
    if (result && result.ok) {
      demandValues[i][0] = result.count;
      dateValues[i][0] = now;
      stats.processed++;
    }
  }

  demandRange.setValues(demandValues);
  dateRange.setValues(dateValues);
  saveCache_(cache);
  var message = [
    'Обработано: ' + stats.processed,
    'Из кэша: ' + stats.fromCache,
    'Запрошено у Wordstat: ' + stats.requested,
    'Отложено из-за лимита запуска: ' + stats.skippedByLimit,
    'Ошибок: ' + stats.errors
  ];
  if (errorSamples.length) {
    message.push('');
    message.push('Примеры ошибок:');
    message = message.concat(errorSamples);
  }
  showAlert_(message.join('\n'));
}

function fetchDemandForQuery_(query, credentials) {
  var dateRange = getDynamicsDateRange_();
  var response = fetchWordstat_(WORDSTAT_CONFIG.DYNAMICS_PATH, {
    phrase: query,
    period: WORDSTAT_CONFIG.DEFAULT_PERIOD,
    fromDate: dateRange.fromDate,
    toDate: dateRange.toDate,
    regions: WORDSTAT_CONFIG.DEFAULT_REGIONS,
    devices: WORDSTAT_CONFIG.DEFAULT_DEVICES,
  }, credentials);
  if (!response.ok) return response;

  var results = response.data && response.data.results;
  if (!Array.isArray(results)) return { ok: false, error: 'В ответе Wordstat нет массива results для запроса: ' + query };

  var total = results.reduce(function(sum, item) {
    var count = Number(item && item.count);
    return sum + (isNaN(count) ? 0 : count);
  }, 0);
  return { ok: true, count: total };
}

function fetchWordstat_(path, payload, credentials) {
  var body = Object.assign({ folderId: credentials.folderId }, payload || {});
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Api-Key ' + credentials.apiKey },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
    validateHttpsCertificates: true,
  };
  var response;
  try {
    response = UrlFetchApp.fetch(WORDSTAT_CONFIG.API_BASE_URL + path, options);
  } catch (error) {
    if (!isSslError_(error)) return { ok: false, error: sanitizeError_(error.message || error) };
    response = UrlFetchApp.fetch(WORDSTAT_CONFIG.API_BASE_URL + path, Object.assign({}, options, { validateHttpsCertificates: false }));
  }
  var status = response.getResponseCode();
  var text = response.getContentText();
  if (status < 200 || status >= 300) return { ok: false, error: 'HTTP ' + status + ': ' + text };
  try {
    return { ok: true, data: JSON.parse(text) };
  } catch (error) {
    return { ok: false, error: 'Wordstat вернул не-JSON ответ.' };
  }
}

function getHeaderColumns_(sheet, requiredHeaders) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(value) { return String(value || '').trim(); });
  var columns = {};
  requiredHeaders.forEach(function(header) {
    var index = headers.indexOf(header);
    if (index === -1) throw new Error('Не найдена колонка с заголовком: ' + header);
    columns[header] = index + 1;
  });
  return columns;
}

function buildQueryFromProductName_(name) {
  var stopWords = new Set(['для', 'с', 'со', 'без', 'на', 'и', 'или', 'в', 'во', 'шт', 'мл', 'мм', 'см', 'гр', 'г', 'professional', 'профессиональный', 'профессиональная']);
  var words = String(name).toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9\s]/g, ' ').split(/\s+/).filter(Boolean);
  var brand = words.indexOf('ollin') !== -1 ? 'ollin' : '';
  var meaningful = words.filter(function(word) { return !stopWords.has(word) && !/^\d+$/.test(word) && word !== brand; });
  return [meaningful[0], brand].filter(Boolean).join(' ').trim();
}

function normalizeQuery_(query) {
  return String(query || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function getDynamicsDateRange_() {
  var today = new Date();

  if (WORDSTAT_CONFIG.DEFAULT_PERIOD === 'PERIOD_MONTHLY') {
    return getMonthlyDynamicsDateRange_(today);
  }

  if (WORDSTAT_CONFIG.DEFAULT_PERIOD === 'PERIOD_WEEKLY') {
    return getWeeklyDynamicsDateRange_(today);
  }

  var day = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate()));
  return { fromDate: day.toISOString(), toDate: day.toISOString() };
}

function getMonthlyDynamicsDateRange_(date) {
  var currentMonthStart = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
  var from = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() - 1, 1));
  var to = new Date(currentMonthStart.getTime() - 24 * 60 * 60 * 1000);
  return { fromDate: from.toISOString(), toDate: to.toISOString() };
}

function getWeeklyDynamicsDateRange_(date) {
  var currentWeekStart = getStartOfUtcWeek_(date);
  var to = new Date(currentWeekStart.getTime() - 24 * 60 * 60 * 1000);
  var from = new Date(to.getTime() - (WORDSTAT_CONFIG.LOOKBACK_WEEKS - 1) * 7 * 24 * 60 * 60 * 1000);
  from = getStartOfUtcWeek_(from);
  return { fromDate: from.toISOString(), toDate: to.toISOString() };
}

function getStartOfUtcWeek_(date) {
  var day = date.getUTCDay() || 7;
  var start = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  start.setUTCDate(start.getUTCDate() - day + 1);
  return start;
}

function getCredentials_() {
  var properties = PropertiesService.getScriptProperties();
  var apiKey = properties.getProperty(WORDSTAT_CONFIG.API_KEY_PROPERTY);
  var folderId = properties.getProperty(WORDSTAT_CONFIG.FOLDER_ID_PROPERTY);
  if (!apiKey || !folderId) {
    return {
      ok: false,
      error: 'API-ключ или folder ID не сохранены. Выполните setWordstatCredentials("ВАШ_API_KEY", "ВАШ_FOLDER_ID") в редакторе Apps Script.',
    };
  }
  return { ok: true, apiKey: apiKey, folderId: folderId };
}

function loadCache_() {
  var raw = PropertiesService.getScriptProperties().getProperty(WORDSTAT_CONFIG.CACHE_PROPERTY);
  if (!raw) return {};
  try { return JSON.parse(raw) || {}; } catch (error) { return {}; }
}

function saveCache_(cache) {
  var entries = Object.entries(cache).sort(function(a, b) { return new Date(b[1].fetchedAt) - new Date(a[1].fetchedAt); }).slice(0, WORDSTAT_CONFIG.MAX_CACHE_ENTRIES);
  PropertiesService.getScriptProperties().setProperty(WORDSTAT_CONFIG.CACHE_PROPERTY, JSON.stringify(Object.fromEntries(entries)));
}

function isCacheFresh_(entry, now) {
  return entry && typeof entry.count === 'number' && entry.fetchedAt && (now - new Date(entry.fetchedAt)) / 36e5 < WORDSTAT_CONFIG.CACHE_TTL_HOURS;
}

function isSslError_(error) {
  return /ssl|certificate|сертификат|ошибка ssl/i.test(String(error && (error.message || error)));
}

function sanitizeError_(message) {
  var credentials = getCredentials_();
  var sanitized = String(message);
  if (credentials.ok) {
    sanitized = sanitized.split(credentials.apiKey).join('[API_KEY]').split(credentials.folderId).join('[FOLDER_ID]');
  }
  return sanitized;
}

function showAlert_(message) {
  SpreadsheetApp.getUi().alert(message);
}
