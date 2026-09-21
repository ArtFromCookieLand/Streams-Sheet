/**
 * ===================================================================
 * ENVIRONMENT
 * -------------------------------------------------------------------
 * The same code is pushed to two Apps Script projects:
 *   PROD - bound to the live spreadsheet
 *   DEV  - bound to a copy of it, used for testing changes
 *
 * Which one we are in is derived from the ID of the spreadsheet the
 * script is bound to, so a freshly made copy is recognised as DEV
 * with no setup at all. A script property 'ENV' set to 'prod' or
 * 'dev' overrides the detection if it is ever needed.
 * ===================================================================
 */

const ENV_PROD = 'prod';
const ENV_DEV = 'dev';

var _envCache = null;

/**
 * @return {string} 'prod' or 'dev'.
 */
function getEnvironment() {
  if (_envCache) return _envCache;

  const override = PropertiesService.getScriptProperties().getProperty('ENV');
  if (override === ENV_PROD || override === ENV_DEV) {
    _envCache = override;
    return _envCache;
  }

  const currentId = SpreadsheetApp.getActiveSpreadsheet().getId();
  _envCache = (currentId === CONFIG.ENV.PROD_SPREADSHEET_ID) ? ENV_PROD : ENV_DEV;
  return _envCache;
}

/**
 * @return {boolean} True when running against anything other than the live sheet.
 */
function isDev() {
  return getEnvironment() === ENV_DEV;
}

/**
 * Suffix for dialog titles, so it is always obvious which sheet is being changed.
 * @return {string} " [DEV]" in the dev copy, "" in production.
 */
function envTag() {
  return isDev() ? ' [DEV]' : '';
}
