/**
 * Utility helpers for GPT custom functions.
 * These helpers are shared across the Apps Script project.
 */

/**
 * Checks whether the provided value is a Range object.
 * @param {*} value
 * @return {boolean}
 */
function isRangeObject_(value) {
  return value && typeof value.getValues === 'function';
}

/**
 * Checks whether the provided value is a plain object.
 * @param {*} value
 * @return {boolean}
 */
function isPlainObject_(value) {
  return Object.prototype.toString.call(value) === '[object Object]';
}

/**
 * Normalizes any range-like input into a two-dimensional array of values.
 * @param {*} input
 * @return {!Array<!Array<*>>}
 */
function normalizeRangeValues_(input) {
  if (input === undefined || input === null) {
    return [];
  }

  if (isRangeObject_(input)) {
    try {
      return input.getValues();
    } catch (error) {
      Logger.log('[normalizeRangeValues_] getValues failed: %s', error);
      return [];
    }
  }

  if (Array.isArray(input)) {
    if (!input.length) {
      return [];
    }
    if (Array.isArray(input[0])) {
      return input;
    }
    return [input];
  }

  return [[input]];
}

/**
 * Converts range values into a tab-separated string.
 * Empty rows are filtered out automatically.
 * @param {*} input
 * @return {string}
 */
function serializeRange_(input) {
  var values = normalizeRangeValues_(input);
  if (!values.length) {
    return '';
  }

  var rows = values
    .map(function (row) {
      var serialized = row
        .map(function (cell) {
          return formatCellValue_(cell);
        })
        .join('\t')
        .trim();
      return serialized;
    })
    .filter(function (row) {
      return row !== '';
    });

  return rows.join('\n');
}

/**
 * Formats a single cell value into a string representation.
 * @param {*} value
 * @return {string}
 */
function formatCellValue_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (value instanceof Date) {
    try {
      return Utilities.formatDate(
        value,
        Session.getScriptTimeZone() || 'Etc/GMT',
        "yyyy-MM-dd'T'HH:mm:ssXXX"
      );
    } catch (error) {
      return value.toISOString ? value.toISOString() : String(value);
    }
  }

  if (typeof value === 'object') {
    try {
      return JSON.stringify(value);
    } catch (error) {
      return String(value);
    }
  }

  return String(value);
}

/**
 * Builds a cache key string using prompt, context and options.
 * @param {string} prompt
 * @param {string} context
 * @param {!Object} options
 * @param {string} sheetName
 * @param {string} cacheVersion
 * @return {string}
 */
function buildCacheKey_(prompt, context, options, sheetName, cacheVersion) {
  var payload = {
    v: cacheVersion || 'v1',
    prompt: prompt || '',
    context: context || '',
    sheet: sheetName || '',
    options: sanitizeCacheOptions_(options)
  };

  var payloadString = JSON.stringify(payload);
  return Utilities.base64EncodeWebSafe(payloadString);
}

/**
 * Picks stable cache-affecting fields from the options object.
 * @param {!Object} options
 * @return {!Object}
 */
function sanitizeCacheOptions_(options) {
  if (!options || typeof options !== 'object') {
    return {};
  }

  return {
    model: options.model,
    format: options.format,
    temperature: options.temperature,
    topP: options.topP,
    maxTokens: options.maxTokens,
    systemPrompt: options.systemPrompt,
    responseCharLimit: options.responseCharLimit,
    frequencyPenalty: options.frequencyPenalty,
    presencePenalty: options.presencePenalty
  };
}

/**
 * Retrieves a value from the Apps Script cache.
 * @param {string} key
 * @return {string|null}
 */
function getCacheValue_(key) {
  if (!key) {
    return null;
  }

  try {
    return CacheService.getScriptCache().get(key);
  } catch (error) {
    Logger.log('[getCacheValue_] %s', error);
    return null;
  }
}

/**
 * Stores a value in the Apps Script cache.
 * Values longer than 90k characters are skipped to avoid cache limits.
 * @param {string} key
 * @param {string} value
 * @param {number} ttlSeconds
 */
function setCacheValue_(key, value, ttlSeconds) {
  if (!key || !value) {
    return;
  }

  if (value.length > 90000) {
    Logger.log('[setCacheValue_] Skip caching because value is too large (%s chars)', value.length);
    return;
  }

  var ttl = ttlSeconds && ttlSeconds > 0 ? Math.min(ttlSeconds, 21600) : 3600;

  try {
    CacheService.getScriptCache().put(key, value, ttl);
  } catch (error) {
    Logger.log('[setCacheValue_] %s', error);
  }
}

/**
 * Estimates token usage with a simple heuristic.
 * @param {string} prompt
 * @param {string} context
 * @param {!Object} options
 * @return {number}
 */
function estimateTokenCount_(prompt, context, options) {
  var systemPrompt = options && options.systemPrompt ? options.systemPrompt : '';
  var combined = [prompt || '', context || '', systemPrompt || ''].join('\n');
  var totalChars = combined.length;

  // Rough heuristic: average 4 characters per token, plus constants.
  var estimatedTokens = Math.ceil(totalChars / 4) + 10;
  if (options && options.maxTokens) {
    estimatedTokens += Number(options.maxTokens);
  }
  return estimatedTokens;
}

/**
 * Trims text to the desired length, appending an indicator when truncated.
 * @param {string} text
 * @param {number} limit
 * @return {string}
 */
function trimResponse_(text, limit) {
  if (!text) {
    return '';
  }

  var effectiveLimit = limit && limit > 0 ? limit : text.length;
  if (text.length <= effectiveLimit) {
    return text;
  }

  return text.substring(0, effectiveLimit).trim() + '...(truncated)';
}

/**
 * Creates an error object with a GPT-specific code.
 * @param {string} code
 * @param {string=} detail
 * @return {!Error}
 */
function createGptError_(code, detail) {
  var error = new Error(detail || code);
  error.gptCode = code;
  return error;
}

/**
 * Normalizes various error shapes into a user-facing error string.
 * @param {*} error
 * @return {string}
 */
function normalizeErrorResponse_(error) {
  if (!error) {
    return '#GPT_ERROR: REQUEST_FAILED';
  }

  if (typeof error === 'string') {
    if (error.indexOf('#GPT_ERROR:') === 0) {
      return error;
    }
    return '#GPT_ERROR: REQUEST_FAILED - ' + error;
  }

  if (error.gptCode) {
    if (error.message && error.message !== error.gptCode) {
      return error.gptCode + ' - ' + error.message;
    }
    return error.gptCode;
  }

  var message = error.message || String(error);
  if (message && message.indexOf('#GPT_ERROR:') === 0) {
    return message;
  }

  return '#GPT_ERROR: REQUEST_FAILED - ' + message;
}

/**
 * Extracts the sheet name from a range-like input.
 * @param {*} input
 * @return {string}
 */
function extractSheetName_(input) {
  if (isRangeObject_(input) && typeof input.getSheet === 'function') {
    try {
      return input.getSheet().getName() || '';
    } catch (error) {
      Logger.log('[extractSheetName_] %s', error);
    }
  }
  return '';
}

/**
 * Clamps a number between the provided bounds.
 * @param {number} value
 * @param {number} min
 * @param {number} max
 * @return {number}
 */
function clamp_(value, min, max) {
  var numeric = Number(value);
  if (isNaN(numeric)) {
    return min;
  }
  if (numeric < min) {
    return min;
  }
  if (numeric > max) {
    return max;
  }
  return numeric;
}

