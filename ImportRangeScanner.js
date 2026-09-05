/**
 * Matches a spreadsheet ID out of a full Google Sheets URL.
 * @type {RegExp}
 */
const SPREADSHEET_URL_ID_REGEX = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;

/**
 * Matches a bare spreadsheet ID literal (no surrounding URL).
 * @type {RegExp}
 */
const BARE_SPREADSHEET_ID_REGEX = /^[a-zA-Z0-9-_]{20,}$/;

/**
 * Finds every unique source spreadsheet ID referenced by IMPORTRANGE
 * anywhere in a spreadsheet — across all sheets, and including formulas
 * that combine multiple IMPORTRANGE calls (e.g. inside VLOOKUP, joined with
 * &, or as separate array items).
 * @param {string} [spreadsheetId] - Defaults to the active spreadsheet (only resolvable from a container-bound script or an installable trigger).
 * @returns {string[]}
 */
function findImportRangeSourceIds(spreadsheetId) {
  const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('findImportRangeSourceIds: spreadsheetId was not provided and there is no active spreadsheet.');
  }
  return findImportRangeSourceIdsForSpreadsheet_(ss);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {string[]}
 */
function findImportRangeSourceIdsForSpreadsheet_(ss) {
  const cells = ss.createTextFinder('IMPORTRANGE')
    .matchCase(false)
    .matchFormulaText(true)
    .findAll();

  const ids = new Set();
  cells.forEach(cell => {
    const formula = cell.getFormula();
    extractImportRangeFirstArgs_(formula).forEach(ref => {
      const id = resolveImportRangeSourceId_(cell, ref);
      if (id) ids.add(id);
    });
  });

  return Array.from(ids);
}

/**
 * Resolves one IMPORTRANGE call's raw first argument (a URL/ID literal, or
 * an unresolved cell-reference expression) down to a spreadsheet ID.
 * @param {GoogleAppsScript.Spreadsheet.Range} cell - The cell whose formula the argument came from, used to resolve cell references relative to its own sheet.
 * @param {string} ref
 * @returns {string|null}
 */
function resolveImportRangeSourceId_(cell, ref) {
  let resolved = ref;
  if (!/^https?:\/\//.test(ref) && !BARE_SPREADSHEET_ID_REGEX.test(ref)) {
    // not already a URL or a bare spreadsheet ID literal — must be a cell reference (e.g. B1); resolve its value
    try { resolved = String(cell.getSheet().getRange(ref).getValue()); }
    catch (e) { return null; }
  }

  const urlMatch = resolved.match(SPREADSHEET_URL_ID_REGEX);
  if (urlMatch) return urlMatch[1];
  if (BARE_SPREADSHEET_ID_REGEX.test(resolved)) return resolved;
  return null;
}

/**
 * Finds every IMPORTRANGE(...) call in a formula — including formulas that
 * combine several of them (e.g. inside VLOOKUP, joined with &, or as
 * separate array items) — and returns each call's raw first argument
 * (spreadsheet URL/ID literal, or an unresolved cell-reference expression).
 * @param {string} formula
 * @returns {string[]}
 */
function extractImportRangeFirstArgs_(formula) {
  const args = [];
  const callRegex = /importrange\s*\(/gi;
  let m;
  while ((m = callRegex.exec(formula)) !== null) {
    const arg = parseFirstArg_(formula, m.index + m[0].length);
    if (arg) args.push(arg);
  }
  return args;
}

/**
 * Parses the first argument of a function call, starting just after its
 * opening "(". Handles a quoted string literal (with "" as an escaped
 * quote) or an unquoted expression, stopping at the top-level "," or ";"
 * argument separator while respecting nested parentheses.
 * @param {string} formula
 * @param {number} start - Index just after the opening "(".
 * @returns {string} The trimmed raw argument text (empty if unparseable).
 */
function parseFirstArg_(formula, start) {
  const len = formula.length;
  let i = start;
  while (i < len && /\s/.test(formula[i])) i++;

  if (formula[i] === '"') {
    let str = '';
    let j = i + 1;
    while (j < len) {
      if (formula[j] === '"') {
        if (formula[j + 1] === '"') { str += '"'; j += 2; continue; }
        break;
      }
      str += formula[j];
      j++;
    }
    return str.trim();
  }

  let depth = 0;
  let buf = '';
  let j = i;
  while (j < len) {
    const c = formula[j];
    if (c === '(') depth++;
    else if (c === ')') { if (depth === 0) break; depth--; }
    else if ((c === ',' || c === ';') && depth === 0) break;
    buf += c;
    j++;
  }
  return buf.trim();
}
