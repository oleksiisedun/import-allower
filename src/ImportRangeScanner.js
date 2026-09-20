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
 * Name of the spreadsheet function whose calls this library detects.
 * @type {string}
 */
const IMPORTRANGE_FUNCTION_NAME_ = 'IMPORTRANGE';

/**
 * Finds every unique source spreadsheet ID referenced by IMPORTRANGE
 * anywhere in a spreadsheet — across all sheets, and including formulas
 * that combine multiple IMPORTRANGE calls (e.g. inside VLOOKUP, joined with
 * &, or as separate array-literal items like `={IMPORTRANGE(...);IMPORTRANGE(...)}`).
 * @param {string} [spreadsheetId] - Defaults to the active spreadsheet (only resolvable from a container-bound script or an installable trigger).
 * @returns {string[]}
 */
function findImportRangeSourceIds(spreadsheetId) {
  return findImportRangeSourceIdsForSpreadsheet_(resolveSpreadsheet_(spreadsheetId, 'findImportRangeSourceIds'));
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {string[]}
 */
function findImportRangeSourceIdsForSpreadsheet_(ss) {
  const ids = new Set();
  forEachImportRangeFormula_(ss, (sheet, formula) => {
    extractImportRangeFirstArgs_(formula).forEach(ref => {
      const id = resolveImportRangeSourceId_(sheet, ref);
      if (id) ids.add(id);
    });
  });
  return Array.from(ids);
}

/**
 * Walks every cell in every sheet and invokes `callback` for each one whose
 * formula contains IMPORTRANGE. Reads formulas directly via getFormulas()
 * rather than via TextFinder — TextFinder can attribute a match inside a
 * spilled array-literal formula (e.g. `={IMPORTRANGE(...);IMPORTRANGE(...)}`)
 * to one of the spilled result cells instead of the anchor cell that
 * actually holds the formula text, and getFormula() on a spilled cell
 * returns '', silently losing whichever IMPORTRANGE call landed there.
 * Reading getFormulas() straight from each sheet's data range sidesteps
 * that entirely: only the anchor cell ever has non-empty formula text.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {(sheet: GoogleAppsScript.Spreadsheet.Sheet, formula: string, row: number, column: number) => void} callback - row/column are 1-indexed sheet coordinates of the formula cell.
 * @returns {void}
 */
function forEachImportRangeFormula_(ss, callback) {
  ss.getSheets().forEach(sheet => {
    const dataRange = sheet.getDataRange();
    const startRow = dataRange.getRow();
    const startColumn = dataRange.getColumn();
    const formulas = dataRange.getFormulas();
    formulas.forEach((rowFormulas, rowOffset) => {
      rowFormulas.forEach((formula, colOffset) => {
        if (!formula || !formula.toUpperCase().includes(IMPORTRANGE_FUNCTION_NAME_)) return;
        callback(sheet, formula, startRow + rowOffset, startColumn + colOffset);
      });
    });
  });
}

/**
 * Resolves one IMPORTRANGE call's raw first argument (a URL/ID literal, or
 * an unresolved cell-reference expression) down to a spreadsheet ID.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet the formula lives on, used to resolve cell references.
 * @param {string} ref
 * @returns {string|null}
 */
function resolveImportRangeSourceId_(sheet, ref) {
  let resolved = ref;
  if (!/^https?:\/\//.test(ref) && !BARE_SPREADSHEET_ID_REGEX.test(ref)) {
    // not already a URL or a bare spreadsheet ID literal — must be a cell reference (e.g. B1); resolve its value
    try { resolved = String(sheet.getRange(ref).getValue()); }
    catch { return null; }
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
  const callRegex = new RegExp(`${IMPORTRANGE_FUNCTION_NAME_}\\s*\\(`, 'gi');
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
