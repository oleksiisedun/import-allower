import { describe, test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import vm from 'node:vm';

// The Apps Script sources share one global scope and have no exports, so load
// the scanner into a fresh context and read its function declarations back off it.
const scanner = vm.createContext({});
vm.runInContext(readFileSync(join(import.meta.dirname, '../src/ImportRangeScanner.js'), 'utf8'), scanner);
const { parseFirstArg_, resolveImportRangeSourceId_ } = scanner;
// Arrays built inside the vm context have another realm's Array.prototype, which deepStrictEqual rejects — copy into a host array.
const extractImportRangeFirstArgs_ = formula => [...scanner.extractImportRangeFirstArgs_(formula)];

const ID_A = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa';
const ID_B = 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb';
const URL_A = `https://docs.google.com/spreadsheets/d/${ID_A}/edit#gid=0`;
const URL_B = `https://docs.google.com/spreadsheets/d/${ID_B}/edit`;

/**
 * @param {Record<string, unknown>} cells - Values keyed by A1 reference; any other reference throws like an invalid range would.
 */
const stubSheet = cells => ({
  getRange: ref => {
    if (!(ref in cells)) throw new Error(`Range not found: ${ref}`);
    return { getValue: () => cells[ref] };
  },
});

// parseFirstArg_ takes the index just past the opening paren of the call.
const parseArgOf = formula => parseFirstArg_(formula, formula.indexOf('(') + 1);

describe('extractImportRangeFirstArgs_', () => {
  test('finds a single call', () => {
    assert.deepEqual(extractImportRangeFirstArgs_(`=IMPORTRANGE("${URL_A}", "Sheet1!A:B")`), [URL_A]);
  });

  test('finds every call in a formula joined with &', () => {
    const formula = `=IMPORTRANGE("${URL_A}", "A:B") & IMPORTRANGE("${URL_B}", "A:B")`;
    assert.deepEqual(extractImportRangeFirstArgs_(formula), [URL_A, URL_B]);
  });

  test('finds every call nested inside other functions', () => {
    const formula = `=VLOOKUP(A1, IMPORTRANGE("${URL_A}", "A:B"), 2, FALSE) + SUM(IMPORTRANGE("${URL_B}", "C:C"))`;
    assert.deepEqual(extractImportRangeFirstArgs_(formula), [URL_A, URL_B]);
  });

  test('finds every item of an array-literal formula', () => {
    const formula = `={IMPORTRANGE("${URL_A}", "A1");IMPORTRANGE("${URL_B}", "A1")}`;
    assert.deepEqual(extractImportRangeFirstArgs_(formula), [URL_A, URL_B]);
  });

  test('matches lowercase importrange', () => {
    assert.deepEqual(extractImportRangeFirstArgs_(`=importrange("${ID_A}", "A1")`), [ID_A]);
  });

  test('tolerates whitespace before the paren', () => {
    assert.deepEqual(extractImportRangeFirstArgs_(`=IMPORTRANGE ("${ID_A}", "A1")`), [ID_A]);
  });

  test('returns an empty array for formulas without IMPORTRANGE', () => {
    assert.deepEqual(extractImportRangeFirstArgs_('=SUM(A1:A3)'), []);
  });

  test('skips a call with an empty first argument', () => {
    assert.deepEqual(extractImportRangeFirstArgs_(`=IMPORTRANGE(, "A1") & IMPORTRANGE("${ID_A}", "A1")`), [ID_A]);
  });
});

describe('parseFirstArg_', () => {
  test('keeps commas and semicolons inside a quoted string', () => {
    assert.equal(parseArgOf('IMPORTRANGE("https://x/a,b;c", "A1")'), 'https://x/a,b;c');
  });

  test('unescapes doubled quotes inside a quoted string', () => {
    assert.equal(parseArgOf('IMPORTRANGE("say ""hi""", "A1")'), 'say "hi"');
  });

  test('returns the partial text of an unterminated string', () => {
    assert.equal(parseArgOf('IMPORTRANGE("abc'), 'abc');
  });

  test('reads an unquoted expression up to the top-level comma, respecting nested parens', () => {
    assert.equal(parseArgOf('IMPORTRANGE(INDEX(A1:A2, 1), "A1")'), 'INDEX(A1:A2, 1)');
  });

  test('stops at a semicolon separator', () => {
    assert.equal(parseArgOf('IMPORTRANGE(B1; "A1")'), 'B1');
  });

  test('stops at the closing paren of a one-argument call', () => {
    assert.equal(parseArgOf('IMPORTRANGE(B1)'), 'B1');
  });
});

describe('resolveImportRangeSourceId_', () => {
  test('extracts the ID from a full URL', () => {
    assert.equal(resolveImportRangeSourceId_(stubSheet({}), URL_A), ID_A);
  });

  test('passes a bare spreadsheet ID through', () => {
    assert.equal(resolveImportRangeSourceId_(stubSheet({}), ID_A), ID_A);
  });

  test('resolves a cell reference holding a URL', () => {
    assert.equal(resolveImportRangeSourceId_(stubSheet({ B1: URL_B }), 'B1'), ID_B);
  });

  test('resolves a cell reference holding a bare ID', () => {
    assert.equal(resolveImportRangeSourceId_(stubSheet({ B1: ID_B }), 'B1'), ID_B);
  });

  test('returns null for an unresolvable reference', () => {
    assert.equal(resolveImportRangeSourceId_(stubSheet({}), 'INDEX(A1:A2, 1)'), null);
  });

  test('returns null for a cell holding a non-ID value', () => {
    assert.equal(resolveImportRangeSourceId_(stubSheet({ B1: 'not an id' }), 'B1'), null);
  });

  test('returns null for an empty cell', () => {
    assert.equal(resolveImportRangeSourceId_(stubSheet({ B1: '' }), 'B1'), null);
  });
});
