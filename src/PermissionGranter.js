/**
 * @typedef {Object} ImportRangeGrantResult
 * @property {string} sourceId - The source spreadsheet ID the grant was requested for.
 * @property {boolean} ok - Whether the request returned HTTP 200.
 * @property {number} httpStatus
 * @property {string} responseText - Raw response body, useful when ok is false.
 */

/** Number of grant requests sent per `UrlFetchApp.fetchAll` batch — trades off parallelism against how soon progress shows up in the log. */
const GRANT_BATCH_SIZE_ = 10;

/**
 * Silently grants a destination spreadsheet access to import from each given
 * source spreadsheet, by calling the same undocumented endpoint the "Allow
 * access" button in the Sheets UI itself calls. This only records the
 * one-time IMPORTRANGE access grant — it can't bypass Drive sharing, so the
 * running user (via ScriptApp.getOAuthToken()) must already have at least
 * view access to every source spreadsheet, or its grant request will fail.
 *
 * Requests are sent via `UrlFetchApp.fetchAll` in batches of
 * `GRANT_BATCH_SIZE_` (parallel within a batch) rather than all at once, so
 * each batch's results can be logged as soon as they're known instead of
 * only after every source has been processed — useful when there are many
 * sources and the whole scan takes a while — while still keeping most of the
 * speed benefit of batched fetches.
 * @param {string} destinationSpreadsheetId
 * @param {string[]} sourceSpreadsheetIds
 * @returns {ImportRangeGrantResult[]} One result per source ID, in the same order; empty if sourceSpreadsheetIds is empty.
 */
function grantImportRangeAccess_(destinationSpreadsheetId, sourceSpreadsheetIds) {
  if (sourceSpreadsheetIds.length === 0) return [];

  const token = ScriptApp.getOAuthToken();
  const results = [];
  for (let i = 0; i < sourceSpreadsheetIds.length; i += GRANT_BATCH_SIZE_) {
    const batch = sourceSpreadsheetIds.slice(i, i + GRANT_BATCH_SIZE_);
    const requests = batch.map(sourceId => ({
      url: `https://docs.google.com/spreadsheets/d/${destinationSpreadsheetId}/externaldata/addimportrangepermissions?donorDocId=${sourceId}&includes_info_params=true&cros_files=false`,
      method: 'post',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    }));

    UrlFetchApp.fetchAll(requests).forEach((response, j) => {
      const result = {
        sourceId: batch[j],
        ok: response.getResponseCode() === 200,
        httpStatus: response.getResponseCode(),
        responseText: response.getContentText(),
      };
      Logger.log(`Source ${result.sourceId} -> HTTP ${result.httpStatus}${result.ok ? '' : ' (FAILED: ' + result.responseText + ')'}`);
      results.push(result);
    });
  }
  return results;
}
