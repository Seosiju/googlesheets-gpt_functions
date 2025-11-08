/**
 * Main entry points for GPT custom functions in Google Sheets.
 */

var DEFAULT_MODEL = 'gpt-4o-mini';
var DEFAULT_SYSTEM_PROMPT = 'You are a helpful data assistant for spreadsheets.';
var RESPONSE_CHAR_LIMIT = 2000;
var TOKEN_LIMIT = 16000;
var CACHE_TTL_SECONDS = 3600;
var API_ENDPOINT = 'https://api.openai.com/v1/chat/completions';

/**
 * 빠른 동작 확인 절차
 * 1. Apps Script 편집기에서 setApiKey('sk-...') 실행 후 권한 부여
 * 2. 시트에서 예시 데이터(A1:C5 등)를 입력하고 =GPT("이 표를 요약해줘", A1:C5) 호출
 * 3. 캐시 확인: 동일 함수 재호출 시 즉시 응답이 오는지 확인, flushCache()로 강제 갱신 가능
 */

/**
 * Google Sheets custom function for GPT completions.
 * @param {string} prompt
 * @param {*=} rangeInput
 * @param {*=} toolsOrOptions
 * @param {*=} rawOptions
 * @return {string}
 * @customfunction
 */
function GPT(prompt, rangeInput, toolsOrOptions, rawOptions) {
  try {
    if (!prompt) {
      return '#GPT_ERROR: Missing prompt';
    }

    var split = splitRangeAndOptions_(rangeInput, toolsOrOptions, rawOptions);
    var options = normalizeOptions_(split.optionsInput);
    var apiKey = getApiKey_();
    if (!apiKey) {
      return '#GPT_ERROR: API_KEY_MISSING';
    }

    var rangeValues = normalizeRangeValues_(split.rangeInput);
    var context = rangeValues.length ? serializeRange_(rangeValues) : '';

    var sheetName = extractSheetName_(split.rangeInput);
    var cacheVersion = getCacheVersion_();
    // 에이전트 모드와 일반 모드 모두 캐시 키를 생성합니다.
    var cacheKey = buildCacheKey_(prompt, context, options, sheetName, cacheVersion, split.toolsInput);
    var cached = getCacheValue_(cacheKey);
    if (cached) {
      return cached;
    }

    // --- 에이전트 모드 분기 ---
    if (typeof split.toolsInput === 'string' && split.toolsInput.trim() !== '') {
      // 에이전트 모드는 자체 로직을 가지므로 여기서 바로 실행하고 반환합니다.
      var agentResult = runAgentMode_(apiKey, prompt, context, split.toolsInput.trim(), options);
      setCacheValue_(cacheKey, agentResult, options.cacheTtlSeconds || CACHE_TTL_SECONDS);
      return agentResult;
    }

    var tokenLimit = options.tokenLimit ? Math.min(options.tokenLimit, TOKEN_LIMIT) : TOKEN_LIMIT;
    var tokenEstimate = estimateTokenCount_(prompt, context, options);
    if (tokenEstimate > tokenLimit) {
      return '#GPT_ERROR: TOKEN_LIMIT - estimated ' + tokenEstimate + ' tokens';
    }

    var responseText = callGptApi_({
      apiKey: apiKey,
      prompt: prompt,
      context: context,
      options: options
    });

    var trimmed = trimResponse_(responseText, options.responseCharLimit);
    var ttl = options.cacheTtlSeconds || CACHE_TTL_SECONDS;
    setCacheValue_(cacheKey, trimmed, ttl);
    return trimmed;
  } catch (error) {
    Logger.log('[GPT] %s', error && error.stack ? error.stack : error);
    return normalizeErrorResponse_(error);
  }
}

/**
 * Variant helper to request JSON-formatted responses.
 * @param {string} prompt
 * @param {*=} rangeInput
 * @return {string}
 * @customfunction
 */
function GPT_JSON(prompt, rangeInput) {
  // GPT_JSON은 에이전트 모드를 사용하지 않으므로 tools 인자는 null로 전달합니다.
  return GPT(prompt, rangeInput, null, JSON.stringify({ format: 'json' }));
}

/**
 * Splits range and options inputs, tolerating different invocation styles.
 * @param {*=} rangeArg
 * @param {*=} toolsOrOptionsArg
 * @param {*=} optionsArg
 * @return {{rangeInput: *, toolsInput: *, optionsInput: *}}
 */
function splitRangeAndOptions_(rangeArg, toolsOrOptionsArg, optionsArg) {
  var rangeInput = null;
  var toolsInput = null;
  var optionsInput = null;

  var args = [rangeArg, toolsOrOptionsArg, optionsArg].filter(function(arg) { return arg !== undefined; });

  // 첫 번째 인자가 범위인지 확인
  if (args.length > 0 && (isRangeObject_(args[0]) || Array.isArray(args[0]))) {
    rangeInput = rangeArg;
    args.shift(); // 범위 인자 처리 완료
  }

  // 남은 인자들로 tools와 options를 식별
  if (args.length > 0) {
    // 첫 번째 남은 인자가 툴킷 이름인지 확인
    if (typeof args[0] === 'string' && !args[0].startsWith('{')) {
      toolsInput = args[0];
      args.shift(); // 툴킷 인자 처리 완료
    }
  }

  if (args.length > 0) {
    // 마지막으로 남은 인자는 options
    optionsInput = args[0];
  }

  // GPT(prompt, options) 또는 GPT(prompt, tools) 같은 2-인자 케이스 처리
  if (rangeInput === null && toolsInput === null && optionsInput === null && rangeArg !== undefined) {
     if (typeof rangeArg === 'string' && !rangeArg.startsWith('{')) {
        // GPT(prompt, "web_tools")
        toolsInput = rangeArg;
     } else {
        // GPT(prompt, "{...}")
        optionsInput = rangeArg;
     }
  }

  return {
    rangeInput: rangeInput,
    toolsInput: toolsInput,
    optionsInput: optionsInput
  };
}

/**
 * Builds a cache key from the core components of a request.
 * @private
 */
function buildCacheKey_(prompt, context, options, sheetName, cacheVersion, tools) {
  var components = [
    prompt,
    context,
    sheetName,
    JSON.stringify(options),
    tools || '', // 툴킷 이름도 캐시 키에 포함
    cacheVersion
  ];
  return Utilities.base64Encode(components.join('||'));
}

/**
 * Normalizes optional options input into a strongly typed object.
 * @param {*=} rawOptions
 * @return {!Object}
 */
function normalizeOptions_(rawOptions) {
  var result = {
    model: DEFAULT_MODEL,
    temperature: 0.7,
    topP: null,
    maxTokens: null,
    format: 'text',
    responseCharLimit: RESPONSE_CHAR_LIMIT,
    cacheTtlSeconds: CACHE_TTL_SECONDS,
    systemPrompt: DEFAULT_SYSTEM_PROMPT,
    tokenLimit: TOKEN_LIMIT,
    frequencyPenalty: null,
    presencePenalty: null
  };

  var parsed = parseOptionsInput_(rawOptions);
  if (!parsed) {
    return result;
  }

  if (parsed.model) {
    result.model = String(parsed.model);
  }

  if (parsed.temperature !== undefined) {
    result.temperature = clamp_(parsed.temperature, 0, 2);
  }

  if (parsed.topP !== undefined) {
    result.topP = clamp_(parsed.topP, 0, 1);
  }

  if (parsed.top_p !== undefined) {
    result.topP = clamp_(parsed.top_p, 0, 1);
  }

  var maxTokensCandidate = parsed.maxTokens !== undefined ? parsed.maxTokens : parsed.max_tokens;
  if (maxTokensCandidate !== undefined) {
    var maxTokens = Number(maxTokensCandidate);
    if (!isNaN(maxTokens) && maxTokens > 0) {
      result.maxTokens = Math.min(Math.floor(maxTokens), TOKEN_LIMIT);
    }
  }

  if (parsed.format) {
    var format = String(parsed.format).toLowerCase();
    if (format === 'json') {
      result.format = 'json';
    }
  }

  if (parsed.response_format) {
    var responseFormat = String(parsed.response_format).toLowerCase();
    if (responseFormat === 'json') {
      result.format = 'json';
    }
  }

  var responseLimitCandidate = parsed.responseCharLimit !== undefined ? parsed.responseCharLimit : parsed.response_char_limit;
  if (responseLimitCandidate !== undefined) {
    var responseLimit = Number(responseLimitCandidate);
    if (!isNaN(responseLimit) && responseLimit > 0) {
      result.responseCharLimit = Math.floor(responseLimit);
    }
  }

  var cacheTtlCandidate = parsed.cacheTtlSeconds !== undefined ? parsed.cacheTtlSeconds : parsed.cache_ttl_seconds;
  if (cacheTtlCandidate !== undefined) {
    var ttl = Number(cacheTtlCandidate);
    if (!isNaN(ttl) && ttl > 0) {
      result.cacheTtlSeconds = Math.min(Math.floor(ttl), 21600);
    }
  }

  var tokenLimitCandidate = parsed.tokenLimit !== undefined ? parsed.tokenLimit : parsed.token_limit;
  if (tokenLimitCandidate !== undefined) {
    var limit = Number(tokenLimitCandidate);
    if (!isNaN(limit) && limit > 0) {
      result.tokenLimit = Math.min(Math.floor(limit), TOKEN_LIMIT);
    }
  }

  if (parsed.systemPrompt) {
    result.systemPrompt = String(parsed.systemPrompt);
  } else if (parsed.system_prompt) {
    result.systemPrompt = String(parsed.system_prompt);
  }

  var frequencyCandidate = parsed.frequencyPenalty !== undefined ? parsed.frequencyPenalty : parsed.frequency_penalty;
  if (frequencyCandidate !== undefined) {
    result.frequencyPenalty = clamp_(frequencyCandidate, -2, 2);
  }

  var presenceCandidate = parsed.presencePenalty !== undefined ? parsed.presencePenalty : parsed.presence_penalty;
  if (presenceCandidate !== undefined) {
    result.presencePenalty = clamp_(presenceCandidate, -2, 2);
  }

  return result;
}

/**
 * Parses raw options input, tolerating JSON strings from Sheets.
 * @param {*=} rawOptions
 * @return {?Object}
 */
function parseOptionsInput_(rawOptions) {
  if (rawOptions === undefined || rawOptions === null || rawOptions === '') {
    return null;
  }

  if (Array.isArray(rawOptions)) {
    if (rawOptions.length === 1) {
      return parseOptionsInput_(rawOptions[0]);
    }
    return null;
  }

  if (typeof rawOptions === 'string') {
    try {
      return JSON.parse(rawOptions);
    } catch (error) {
      Logger.log('[parseOptionsInput_] 옵션 JSON 파싱 실패: %s', error);
      return null;
    }
  }

  if (isPlainObject_(rawOptions)) {
    return rawOptions;
  }

  return null;
}

/**
 * Calls the OpenAI Chat Completions API.
 * @param {{apiKey: string, prompt: string, context: string, options: !Object}} payload
 * @return {string}
 */
function callGptApi_(payload) {
  var options = payload.options || {};
  var systemPrompt = options.systemPrompt || DEFAULT_SYSTEM_PROMPT;
  var userContent = buildUserContent_(payload.prompt, payload.context);

  var messages = [
    { role: 'system', content: systemPrompt },
    { role: 'user', content: userContent }
  ];

  var requestBody = {
    model: options.model || DEFAULT_MODEL,
    messages: messages,
    temperature: options.temperature
  };

  if (options.topP !== null && options.topP !== undefined) {
    requestBody.top_p = options.topP;
  }

  if (options.maxTokens) {
    requestBody.max_tokens = options.maxTokens;
  }

  if (options.frequencyPenalty !== null && options.frequencyPenalty !== undefined) {
    requestBody.frequency_penalty = options.frequencyPenalty;
  }

  if (options.presencePenalty !== null && options.presencePenalty !== undefined) {
    requestBody.presence_penalty = options.presencePenalty;
  }

  if (options.format === 'json') {
    requestBody.response_format = { type: 'json_object' };
  }

  var requestOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + payload.apiKey
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(requestBody)
  };

  var response = UrlFetchApp.fetch(API_ENDPOINT, requestOptions);
  var status = response.getResponseCode();
  var bodyText = response.getContentText();
  var parsed;

  try {
    parsed = bodyText ? JSON.parse(bodyText) : null;
  } catch (error) {
    throw createGptError_('#GPT_ERROR: REQUEST_FAILED', 'API 응답 JSON 파싱 실패');
  }

  if (status >= 400) {
    var apiMessage = parsed && parsed.error && parsed.error.message ? parsed.error.message : (bodyText || 'Unknown error');
    throw createGptError_('#GPT_ERROR: REQUEST_FAILED', apiMessage);
  }

  var message = parsed && parsed.choices && parsed.choices.length ? parsed.choices[0].message : null;
  var content = message && message.content ? message.content.trim() : '';
  if (!content) {
    throw createGptError_('#GPT_ERROR: No response', 'Empty completion text');
  }

  if (options.format === 'json') {
    try {
      var json = JSON.parse(content);
      return JSON.stringify(json, null, 2);
    } catch (error) {
      throw createGptError_('#GPT_ERROR: REQUEST_FAILED', '유효한 JSON 응답이 아닙니다.');
    }
  }

  return content;
}

/**
 * Builds the user message content combining prompt and data context.
 * @param {string} prompt
 * @param {string} context
 * @return {string}
 */
function buildUserContent_(prompt, context) {
  if (context) {
    return prompt + '\n\nData:\n' + context;
  }
  return prompt;
}

/**
 * Runs the agentic logic using toolkits.
 * @private
 */
function runAgentMode_(apiKey, prompt, context, toolkitName, options) {
  var userContent = buildUserContent_(prompt, context);
  var messages = [
    { role: 'system', content: 'You are a helpful agent in a spreadsheet. You can use tools to answer questions. First, think about what tools you need from the provided toolkit. Then, call them. Finally, answer the user\'s question based on the tool results.' },
    { role: 'user', content: userContent }
  ];

  var toolSpecs = getToolkitSpecs_(toolkitName);
  if (toolSpecs.length === 0) {
    return '#AGENT_ERROR: Invalid or empty toolkit: ' + toolkitName;
  }

  const MAX_LOOPS = 5; // Prevent infinite loops
  for (let i = 0; i < MAX_LOOPS; i++) {
    const requestBody = {
      model: options.model || DEFAULT_MODEL,
      messages: messages,
      tools: toolSpecs,
      tool_choice: 'auto',
      temperature: options.temperature,
      top_p: options.topP,
      frequency_penalty: options.frequencyPenalty,
      presence_penalty: options.presencePenalty
    };

    const response = UrlFetchApp.fetch(API_ENDPOINT, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      muteHttpExceptions: true,
      payload: JSON.stringify(requestBody)
    });

    const responseData = JSON.parse(response.getContentText());
    if (response.getResponseCode() >= 400) {
      const apiMessage = responseData && responseData.error && responseData.error.message ? responseData.error.message : 'Unknown API error';
      return '#AGENT_ERROR: ' + apiMessage;
    }

    const message = responseData.choices[0].message;

    // Case 1: LLM provides a final answer without calling a tool.
    if (!message.tool_calls) {
      return message.content;
    }

    // Case 2: LLM requests to call one or more tools.
    messages.push(message); // Add AI's decision to call tools to the conversation history.

    for (const toolCall of message.tool_calls) {
      const functionName = toolCall.function.name;
      const args = JSON.parse(toolCall.function.arguments);

      let toolResult;
      if (GptTools[functionName]) {
        // Execute the actual tool function.
        toolResult = GptTools[functionName](...Object.values(args));
      } else {
        toolResult = `Error: Tool "${functionName}" not found.`;
      }

      // Add the tool's result back into the conversation history.
      messages.push({
        tool_call_id: toolCall.id,
        role: 'tool',
        name: functionName,
        content: String(toolResult)
      });
    }
    // Continue the loop to let the LLM process the tool results and generate a final answer.
  }

  return '#AGENT_ERROR: Max loops reached. The agent could not find an answer.';
}
