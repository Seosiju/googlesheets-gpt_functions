/**
 * Configuration helpers for GPT custom functions.
 */

var SCRIPT_PROPERTY_API_KEY = 'OPENAI_API_KEY';
var SCRIPT_PROPERTY_CACHE_VERSION = 'CACHE_VERSION';

/**
 * Stores the OpenAI API key in Script Properties.
 * Run manually from the Apps Script editor.
 * 입력 파라미터 UI가 없는 환경을 위해 configureApiKey()를 함께 제공한다.
 * @param {string} key
 * @return {string}
 */
function setApiKey(key) {
  if (!key || typeof key !== 'string') {
    throw new Error('유효한 OpenAI API 키를 입력하세요.');
  }

  var trimmed = key.trim();
  if (!trimmed) {
    throw new Error('빈 문자열은 API 키로 사용할 수 없습니다.');
  }

  clearApiKey(); // 먼저 기존 키를 모두 제거하여 일관성을 유지합니다.
  PropertiesService.getScriptProperties().setProperty(SCRIPT_PROPERTY_API_KEY, trimmed);
  return 'OpenAI API 키가 Script 속성에 안전하게 저장되었습니다.';
}

/**
 * Retrieves the stored OpenAI API key.
 * @return {string|null}
 */
function getApiKey_() {
  var scriptKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROPERTY_API_KEY);
  if (scriptKey) {
    return scriptKey;
  }

  var userKey = PropertiesService.getUserProperties().getProperty(SCRIPT_PROPERTY_API_KEY);
  if (userKey) {
    return userKey;
  }

  var documentKey = PropertiesService.getDocumentProperties().getProperty(SCRIPT_PROPERTY_API_KEY);
  if (documentKey) {
    return documentKey;
  }

  return null;
}

/**
 * Resets the cache version to invalidate cached entries.
 * @return {string}
 */
function flushCache() {
  var newVersion = 'v' + Date.now();
  PropertiesService.getScriptProperties().setProperty(SCRIPT_PROPERTY_CACHE_VERSION, newVersion);
  return 'Cache version updated: ' + newVersion;
}

/**
 * Retrieves the current cache version, initializing if absent.
 * @return {string}
 */
function getCacheVersion_() {
  var props = PropertiesService.getScriptProperties();
  var version = props.getProperty(SCRIPT_PROPERTY_CACHE_VERSION);
  if (!version) {
    version = 'v1';
    props.setProperty(SCRIPT_PROPERTY_CACHE_VERSION, version);
  }
  return version;
}

/**
 * Prompts the user to input an API key through a UI dialog.
 * Use this when running as a spreadsheet-bound script.
 * @return {string}
 */
function configureApiKey() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'OpenAI API 키 설정',
    'sk- 로 시작하는 OpenAI API 키를 입력하세요.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return '취소되었습니다.';
  }

  return setApiKey(response.getResponseText());
}

/**
 * Clears the stored API key from all property stores.
 * @return {string}
 */
function clearApiKey() {
  PropertiesService.getScriptProperties().deleteProperty(SCRIPT_PROPERTY_API_KEY);
  PropertiesService.getUserProperties().deleteProperty(SCRIPT_PROPERTY_API_KEY);
  PropertiesService.getDocumentProperties().deleteProperty(SCRIPT_PROPERTY_API_KEY);
  return 'OpenAI API 키가 제거되었습니다.';
}
