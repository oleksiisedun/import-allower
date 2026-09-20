/**
 * Scans a spreadsheet for every IMPORTRANGE source spreadsheet referenced
 * anywhere in it (across all sheets, including formulas that combine
 * multiple IMPORTRANGE calls) and silently grants this spreadsheet access
 * to each — the same one-time "Allow access" grant a user would otherwise
 * have to click through manually, once per source, every time a new
 * cross-spreadsheet import is added.
 * @param {string} [spreadsheetId] - Defaults to the active spreadsheet (only resolvable from a container-bound script or an installable trigger).
 * @returns {ImportRangeGrantResult[]} One result per unique source spreadsheet found; empty if none were found.
 */
function autoApproveImportRanges(spreadsheetId) {
  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('autoApproveImportRanges: spreadsheetId was not provided and there is no active spreadsheet.');
  }

  const sourceIds = findImportRangeSourceIdsForSpreadsheet_(ss);
  if (sourceIds.length === 0) {
    Logger.log(`autoApproveImportRanges: no IMPORTRANGE sources found in "${ss.getName()}".`);
    return [];
  }

  return grantImportRangeAccess_(ss.getId(), sourceIds);
}

/**
 * Installs an installable onEdit trigger that calls `handlerFunctionName` on
 * every edit. Apps Script always resolves an installable trigger's handler
 * against the CALLING project's own global scope — so `handlerFunctionName`
 * must name a real top-level function defined in the CONSUMING project, not
 * a library-qualified reference like `ImportAllower.autoApproveImportRanges`.
 * A one-line wrapper in the consuming project is enough:
 *
 *   function onEditAutoApprove(e) { ImportAllower.autoApproveImportRanges(); }
 *   ImportAllower.installAutoApproveTrigger('onEditAutoApprove');
 *
 * @param {string} handlerFunctionName - Name of a function defined in the CALLING project.
 * @param {string} [spreadsheetId] - Defaults to the active spreadsheet.
 * @returns {GoogleAppsScript.Script.Trigger}
 */
function installAutoApproveTrigger(handlerFunctionName, spreadsheetId) {
  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('installAutoApproveTrigger: spreadsheetId was not provided and there is no active spreadsheet.');
  }
  return ScriptApp.newTrigger(handlerFunctionName).forSpreadsheet(ss).onEdit().create();
}

/**
 * Logs every cell containing IMPORTRANGE — its raw formula and the source
 * IDs detected in it — for troubleshooting when a source isn't being
 * detected/granted as expected.
 * @param {string} [spreadsheetId] - Defaults to the active spreadsheet.
 * @returns {void}
 */
function logImportRangeFormulas(spreadsheetId) {
  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('logImportRangeFormulas: spreadsheetId was not provided and there is no active spreadsheet.');
  }

  forEachImportRangeFormula_(ss, (sheet, formula, row, column) => {
    const args = extractImportRangeFirstArgs_(formula);
    const a1Notation = sheet.getRange(row, column).getA1Notation();
    Logger.log(`${sheet.getName()}!${a1Notation} -> ${formula} (args: ${JSON.stringify(args)})`);
  });
}
