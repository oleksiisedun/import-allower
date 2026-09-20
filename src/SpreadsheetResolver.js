/**
 * Opens the spreadsheet with the given ID, or falls back to the active
 * spreadsheet when none is given.
 * @param {string|undefined} spreadsheetId - Falsy means "use the active spreadsheet" (only resolvable from a container-bound script or an installable trigger).
 * @param {string} callerName - Public function name, used in the error message when no spreadsheet can be resolved.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function resolveSpreadsheet_(spreadsheetId, callerName) {
  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(`${callerName}: spreadsheetId was not provided and there is no active spreadsheet.`);
  }
  return ss;
}
