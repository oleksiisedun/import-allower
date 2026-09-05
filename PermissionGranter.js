/**
 * @typedef {Object} ImportRangeGrantResult
 * @property {string} sourceId - The source spreadsheet ID the grant was requested for.
 * @property {boolean} ok - Whether the request returned HTTP 200.
 * @property {number} httpStatus
 * @property {string} responseText - Raw response body, useful when ok is false.
 */

/**
 * Silently grants a destination spreadsheet access to import from each given
 * source spreadsheet, by calling the same undocumented endpoint the "Allow
 * access" button in the Sheets UI itself calls. This only records the
 * one-time IMPORTRANGE access grant — it can't bypass Drive sharing, so the
 * running user (via ScriptApp.getOAuthToken()) must already have at least
 * view access to every source spreadsheet, or its grant request will fail.
 * @param {string} destinationSpreadsheetId
 * @param {string[]} sourceSpreadsheetIds
 * @returns {ImportRangeGrantResult[]} One result per source ID, in the same order; empty if sourceSpreadsheetIds is empty.
 */
function grantImportRangeAccess_(destinationSpreadsheetId, sourceSpreadsheetIds) {
  if (sourceSpreadsheetIds.length === 0) return [];

  const token = ScriptApp.getOAuthToken();
  const requests = sourceSpreadsheetIds.map(donorId => ({
    url: `https://docs.google.com/spreadsheets/d/${destinationSpreadsheetId}/externaldata/addimportrangepermissions?donorDocId=${donorId}&includes_info_params=true&cros_files=false`,
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  }));

  const responses = UrlFetchApp.fetchAll(requests);
  return responses.map((response, i) => ({
    sourceId: sourceSpreadsheetIds[i],
    ok: response.getResponseCode() === 200,
    httpStatus: response.getResponseCode(),
    responseText: response.getContentText(),
  }));
}
