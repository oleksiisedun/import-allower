import { test } from 'node:test';
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

test('extractImportRangeFirstArgs_ finds a single call', () => {
  assert.deepEqual(extractImportRangeFirstArgs_(`=IMPORTRANGE("${URL_A}", "Sheet1!A:B")`), [URL_A]);
});

test('extractImportRangeFirstArgs_ finds every call in a combined formula', () => {
  const joined = `=IMPORTRANGE("${URL_A}", "A:B") & IMPORTRANGE("${URL_B}", "A:B")`;
  assert.deepEqual(extractImportRangeFirstArgs_(joined), [URL_A, URL_B]);

  const nested = `=VLOOKUP(A1, IMPORTRANGE("${URL_A}", "A:B"), 2, FALSE) + SUM(IMPORTRANGE("${URL_B}", "C:C"))`;
  assert.deepEqual(extractImportRangeFirstArgs_(nested), [URL_A, URL_B]);
});

test('extractImportRangeFirstArgs_ finds every item of an array-literal formula', () => {
  const formula = `={IMPORTRANGE("${URL_A}", "A1");IMPORTRANGE("${URL_B}", "A1")}`;
  assert.deepEqual(extractImportRangeFirstArgs_(formula), [URL_A, URL_B]);
});

test('extractImportRangeFirstArgs_ is case-insensitive and tolerates whitespace before the paren', () => {
  assert.deepEqual(extractImportRangeFirstArgs_(`=importrange ("${ID_A}", "A1")`), [ID_A]);
});

test('extractImportRangeFirstArgs_ returns nothing for formulas without IMPORTRANGE', () => {
  assert.deepEqual(extractImportRangeFirstArgs_('=SUM(A1:A3)'), []);
});

test('extractImportRangeFirstArgs_ skips a call with an empty first argument', () => {
  assert.deepEqual(extractImportRangeFirstArgs_(`=IMPORTRANGE(, "A1") & IMPORTRANGE("${ID_A}", "A1")`), [ID_A]);
});

test('parseFirstArg_ keeps commas and semicolons inside a quoted string', () => {
  const formula = 'IMPORTRANGE("https://x/a,b;c", "A1")';
  assert.equal(parseFirstArg_(formula, formula.indexOf('(') + 1), 'https://x/a,b;c');
});

test('parseFirstArg_ unescapes doubled quotes inside a quoted string', () => {
  const formula = 'IMPORTRANGE("say ""hi""", "A1")';
  assert.equal(parseFirstArg_(formula, formula.indexOf('(') + 1), 'say "hi"');
});

test('parseFirstArg_ returns the text so far for an unterminated string', () => {
  const formula = 'IMPORTRANGE("abc';
  assert.equal(parseFirstArg_(formula, formula.indexOf('(') + 1), 'abc');
});

test('parseFirstArg_ reads an unquoted expression up to the top-level separator, respecting nested parens', () => {
  const nested = 'IMPORTRANGE(INDEX(A1:A2, 1), "A1")';
  assert.equal(parseFirstArg_(nested, nested.indexOf('(') + 1), 'INDEX(A1:A2, 1)');

  const semicolon = 'IMPORTRANGE(B1; "A1")';
  assert.equal(parseFirstArg_(semicolon, semicolon.indexOf('(') + 1), 'B1');
});

test('parseFirstArg_ stops at the closing paren of a one-argument call', () => {
  const formula = 'IMPORTRANGE(B1)';
  assert.equal(parseFirstArg_(formula, formula.indexOf('(') + 1), 'B1');
});

test('resolveImportRangeSourceId_ extracts the ID from a full URL', () => {
  assert.equal(resolveImportRangeSourceId_(stubSheet({}), URL_A), ID_A);
});

test('resolveImportRangeSourceId_ passes a bare spreadsheet ID through', () => {
  assert.equal(resolveImportRangeSourceId_(stubSheet({}), ID_A), ID_A);
});

test('resolveImportRangeSourceId_ resolves a cell reference holding a URL or a bare ID', () => {
  assert.equal(resolveImportRangeSourceId_(stubSheet({ B1: URL_B }), 'B1'), ID_B);
  assert.equal(resolveImportRangeSourceId_(stubSheet({ B1: ID_B }), 'B1'), ID_B);
});

test('resolveImportRangeSourceId_ returns null for an unresolvable reference or a non-ID value', () => {
  assert.equal(resolveImportRangeSourceId_(stubSheet({}), 'INDEX(A1:A2, 1)'), null);
  assert.equal(resolveImportRangeSourceId_(stubSheet({ B1: 'not an id' }), 'B1'), null);
  assert.equal(resolveImportRangeSourceId_(stubSheet({ B1: '' }), 'B1'), null);
});
