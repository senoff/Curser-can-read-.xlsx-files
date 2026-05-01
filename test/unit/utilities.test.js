// Unit tests for the pure-function helpers in index.js.
// These tests exercise functions directly via require() — no subprocess,
// no fixture files, fast (sub-second). The XLSX_FOR_AI_RESPAWNED env var
// is set in the test runner so requiring index.js doesn't trigger the
// self-respawn.

'use strict';

process.env.XLSX_FOR_AI_RESPAWNED = '1';

const test = require('node:test');
const assert = require('node:assert/strict');
const lib = require('../../index.js');

// ---------------------------------------------------------------------------
// colLetter / colNum — bidirectional column ref conversion
// ---------------------------------------------------------------------------

test('colLetter: single-letter columns', () => {
  assert.equal(lib.colLetter(1), 'A');
  assert.equal(lib.colLetter(26), 'Z');
});

test('colLetter: double-letter columns', () => {
  assert.equal(lib.colLetter(27), 'AA');
  assert.equal(lib.colLetter(52), 'AZ');
  assert.equal(lib.colLetter(53), 'BA');
  assert.equal(lib.colLetter(702), 'ZZ');
});

test('colLetter: triple-letter columns (xlsx max is XFD = 16384)', () => {
  assert.equal(lib.colLetter(703), 'AAA');
  assert.equal(lib.colLetter(16384), 'XFD');
});

test('colNum: round-trips with colLetter', () => {
  for (const n of [1, 5, 26, 27, 52, 53, 100, 702, 703, 16384]) {
    assert.equal(lib.colNum(lib.colLetter(n)), n, `round-trip failed for n=${n}`);
  }
});

test('colNum: lowercase input', () => {
  assert.equal(lib.colNum('a'), 1);
  assert.equal(lib.colNum('aa'), 27);
});

// ---------------------------------------------------------------------------
// parseRange — A1-style range parsing
// ---------------------------------------------------------------------------

test('parseRange: rectangular range', () => {
  const r = lib.parseRange('A1:D5');
  assert.deepEqual(r, { startCol: 1, startRow: 1, endCol: 4, endRow: 5 });
});

test('parseRange: single cell expands to 1x1', () => {
  const r = lib.parseRange('B3');
  assert.deepEqual(r, { startCol: 2, startRow: 3, endCol: 2, endRow: 3 });
});

test('parseRange: high-numbered columns', () => {
  const r = lib.parseRange('AA10:ZZ100');
  assert.equal(r.startCol, 27);
  assert.equal(r.endCol, 702);
});

test('parseRange: returns null for null/empty input', () => {
  assert.equal(lib.parseRange(null), null);
  assert.equal(lib.parseRange(''), null);
});

test('parseRange: throws on malformed input', () => {
  assert.throws(() => lib.parseRange('garbage'), /Invalid range/);
  assert.throws(() => lib.parseRange('A1:garbage'), /Invalid range/);
});

// ---------------------------------------------------------------------------
// formatValue / plainValue / jsonValue — cell-value rendering
// ---------------------------------------------------------------------------

test('formatValue: primitive types', () => {
  assert.equal(lib.formatValue(null), '""');
  assert.equal(lib.formatValue(''), '""');
  assert.equal(lib.formatValue('hello'), '"hello"');
  assert.equal(lib.formatValue(42), '42');
  assert.equal(lib.formatValue(3.14), '3.14');
});

test('formatValue: dates render as ISO date (no time)', () => {
  const d = new Date('2026-04-28T15:30:00.000Z');
  assert.equal(lib.formatValue(d), '"2026-04-28"');
});

test('formatValue: formula with cached numeric result', () => {
  assert.equal(lib.formatValue({ formula: 'A1+B1', result: 42 }), '42');
});

test('formatValue: formula with no cached result is empty', () => {
  assert.equal(lib.formatValue({ formula: 'A1+B1' }), '""');
});

test('formatValue: shared formula follower with cached result', () => {
  // ExcelJS read shape
  assert.equal(lib.formatValue({ sharedFormula: 'A1', result: 42 }), '42');
});

test('formatValue: error cell renders as #ERROR', () => {
  assert.equal(lib.formatValue({ error: 'REF' }), '"#REF"');
  assert.equal(lib.formatValue({ formula: 'X', result: { error: 'DIV/0' } }), '"#DIV/0"');
});

test('formatValue: rich text concatenates runs', () => {
  const rt = { richText: [{ text: 'Hello ' }, { text: 'world' }] };
  assert.equal(lib.formatValue(rt), '"Hello world"');
});

test('formatValue: hyperlink cell uses display text', () => {
  assert.equal(lib.formatValue({ text: 'click', hyperlink: 'https://example.com' }), '"click"');
});

test('plainValue: returns null for empty', () => {
  assert.equal(lib.plainValue(null), null);
  assert.equal(lib.plainValue(''), null);
});

test('plainValue: handles all formula shapes', () => {
  assert.equal(lib.plainValue({ formula: 'X', result: 5 }), '5');
  assert.equal(lib.plainValue({ sharedFormula: 'A1', result: 5 }), '5');
  // --json output's shape:
  assert.equal(lib.plainValue({ sharedFormulaRef: 'A1', result: 5 }), '5');
});

test('plainValue: Date inside formula result is YYYY-MM-DD', () => {
  const d = new Date('2026-04-28T15:30:00.000Z');
  assert.equal(lib.plainValue({ formula: 'X', result: d }), '2026-04-28');
});

test('jsonValue: dates serialize as ISO string (full timestamp)', () => {
  const d = new Date('2026-04-28T00:00:00.000Z');
  assert.equal(lib.jsonValue(d), '2026-04-28T00:00:00.000Z');
});

test('jsonValue: formula objects serialize with formula + result', () => {
  const v = lib.jsonValue({ formula: 'A1+B1', result: 42 });
  assert.deepEqual(v, { formula: 'A1+B1', result: 42 });
});

test('jsonValue: shared formula renames sharedFormula → sharedFormulaRef', () => {
  const v = lib.jsonValue({ sharedFormula: 'A1', result: 42 });
  assert.equal(v.sharedFormulaRef, 'A1');
  assert.equal(v.result, 42);
});

// ---------------------------------------------------------------------------
// coerceMaybeDate — heuristic for ISO-string → Date conversion
// ---------------------------------------------------------------------------

test('coerceMaybeDate: pure date string → Date', () => {
  const r = lib.coerceMaybeDate('2026-04-28');
  assert.ok(r instanceof Date);
  assert.equal(r.toISOString().slice(0, 10), '2026-04-28');
});

test('coerceMaybeDate: midnight-UTC ISO → Date (Excel-shape)', () => {
  const r = lib.coerceMaybeDate('2026-01-01T00:00:00.000Z');
  assert.ok(r instanceof Date);
});

test('coerceMaybeDate: non-midnight UTC ISO → Date', () => {
  // Non-midnight UTC ISO is what JSON.stringify(Date) produces for Excel-source dates with time
  const r = lib.coerceMaybeDate('2026-04-28T15:30:42.000Z');
  assert.ok(r instanceof Date);
});

test('coerceMaybeDate: timezone-offset string stays string (user-typed timestamp)', () => {
  // This is the bug pattern from Proposable Transactions — strings with TZ offsets
  // should NOT be coerced to dates.
  const r = lib.coerceMaybeDate('2018-04-16T11:29:02-07:00');
  assert.equal(typeof r, 'string');
  assert.equal(r, '2018-04-16T11:29:02-07:00');
});

test('coerceMaybeDate: non-string input passes through', () => {
  assert.equal(lib.coerceMaybeDate(42), 42);
  assert.equal(lib.coerceMaybeDate(null), null);
  const d = new Date();
  assert.equal(lib.coerceMaybeDate(d), d);
});

test('coerceMaybeDate: invalid date string passes through unchanged', () => {
  // Regex matches the shape, but new Date() returns NaN — should fall back to string.
  assert.equal(lib.coerceMaybeDate('2026-13-45'), '2026-13-45');
});

// ---------------------------------------------------------------------------
// SQL value rendering
// ---------------------------------------------------------------------------

test('sqlIdent: quotes identifier and escapes embedded quotes', () => {
  assert.equal(lib.sqlIdent('Sheet1'), '"Sheet1"');
  assert.equal(lib.sqlIdent('a"b'), '"a""b"');
});

test('sqlVal: NULL for empty', () => {
  assert.equal(lib.sqlVal(null, 'TEXT'), 'NULL');
  assert.equal(lib.sqlVal('', 'TEXT'), 'NULL');
});

test('sqlVal: integer/numeric types', () => {
  assert.equal(lib.sqlVal(42, 'INTEGER'), '42');
  assert.equal(lib.sqlVal('3.14', 'NUMERIC'), '3.14');
  assert.equal(lib.sqlVal('1,234.56', 'NUMERIC'), '1234.56'); // commas stripped
});

test('sqlVal: boolean rendering', () => {
  assert.equal(lib.sqlVal(true, 'BOOLEAN'), 'TRUE');
  assert.equal(lib.sqlVal('true', 'BOOLEAN'), 'TRUE');
  assert.equal(lib.sqlVal('false', 'BOOLEAN'), 'FALSE');
});

test('sqlVal: text escapes single quotes', () => {
  assert.equal(lib.sqlVal("don't", 'TEXT'), "'don''t'");
});

test('sqlVal: date types', () => {
  const d = new Date('2026-04-28T15:30:00.000Z');
  assert.equal(lib.sqlVal(d, 'DATE'), "'2026-04-28'");
});

// ---------------------------------------------------------------------------
// inferType — type inference from a sample of values
// ---------------------------------------------------------------------------

test('inferType: all integers → INTEGER, not nullable', () => {
  const r = lib.inferType([1, 2, 3, 4, 5]);
  assert.equal(r.type, 'INTEGER');
  assert.equal(r.nullable, false);
});

test('inferType: mixed numbers → INTEGER if majority int', () => {
  const r = lib.inferType([1, 2, 3, 4.5]);
  assert.equal(r.type, 'INTEGER'); // 3 int beats 1 float
});

test('inferType: text dominant → TEXT', () => {
  const r = lib.inferType(['hello', 'world', '123', 'foo']);
  assert.equal(r.type, 'TEXT');
});

test('inferType: nulls increment nullable flag', () => {
  const r = lib.inferType([1, null, 2, '', 3]);
  assert.equal(r.nullable, true);
});

test('inferType: ISO date strings → DATE', () => {
  const r = lib.inferType(['2026-01-01', '2026-02-15', '2026-03-30']);
  assert.equal(r.type, 'DATE');
});

test('inferType: empty input → unknown', () => {
  const r = lib.inferType([null, '', null]);
  assert.equal(r.type, 'unknown');
});

// ---------------------------------------------------------------------------
// escapeMd — markdown table cell escaping
// ---------------------------------------------------------------------------

test('escapeMd: pipes escaped', () => {
  assert.equal(lib.escapeMd('a|b'), 'a\\|b');
});

test('escapeMd: newlines collapsed to spaces', () => {
  assert.equal(lib.escapeMd('line1\nline2'), 'line1 line2');
});

test('escapeMd: null/undefined → empty string', () => {
  assert.equal(lib.escapeMd(null), '');
  assert.equal(lib.escapeMd(undefined), '');
});

// ---------------------------------------------------------------------------
// coerceMarkdownValue — markdown cell → typed JS
// ---------------------------------------------------------------------------

test('coerceMarkdownValue: empty → null', () => {
  assert.equal(lib.coerceMarkdownValue(''), null);
  assert.equal(lib.coerceMarkdownValue(null), null);
});

test('coerceMarkdownValue: integer string → number', () => {
  assert.equal(lib.coerceMarkdownValue('42'), 42);
  assert.equal(lib.coerceMarkdownValue('-7'), -7);
});

test('coerceMarkdownValue: float string → number', () => {
  assert.equal(lib.coerceMarkdownValue('3.14'), 3.14);
});

test('coerceMarkdownValue: boolean strings → boolean', () => {
  assert.equal(lib.coerceMarkdownValue('true'), true);
  assert.equal(lib.coerceMarkdownValue('false'), false);
});

test('coerceMarkdownValue: ISO date string → Date', () => {
  const r = lib.coerceMarkdownValue('2026-04-28');
  assert.ok(r instanceof Date);
});

test('coerceMarkdownValue: backtick-fenced formula → formula object', () => {
  const r = lib.coerceMarkdownValue('`=SUM(A1:A10)`');
  assert.deepEqual(r, { formula: 'SUM(A1:A10)' });
});

test('coerceMarkdownValue: escaped pipe is unescaped', () => {
  assert.equal(lib.coerceMarkdownValue('a\\|b'), 'a|b');
});

// ---------------------------------------------------------------------------
// applyTokenBudget — output truncation
// ---------------------------------------------------------------------------

test('applyTokenBudget: short text passes through unchanged', () => {
  const out = lib.applyTokenBudget('hello world', 1000);
  assert.equal(out, 'hello world');
});

test('applyTokenBudget: long text gets truncated with summary tail', () => {
  const longText = Array.from({ length: 1000 }, (_, i) => `line ${i}`).join('\n');
  const out = lib.applyTokenBudget(longText, 100);
  assert.ok(out.length < longText.length);
  assert.match(out, /truncated/);
});

// ---------------------------------------------------------------------------
// validateSpec — spec validation
// ---------------------------------------------------------------------------

test('validateSpec: rejects non-object', () => {
  assert.throws(() => lib.validateSpec(null), /must be an object/);
  assert.throws(() => lib.validateSpec('string'), /must be an object/);
});

test('validateSpec: rejects empty sheets', () => {
  assert.throws(() => lib.validateSpec({}), /at least one sheet/);
  assert.throws(() => lib.validateSpec({ sheets: [] }), /at least one sheet/);
});

test('validateSpec: rejects sheet without name', () => {
  assert.throws(
    () => lib.validateSpec({ sheets: [{ rows: [] }] }),
    /needs a "name"/
  );
});

test('validateSpec: rejects sheet without rows or cells', () => {
  assert.throws(
    () => lib.validateSpec({ sheets: [{ name: 'X' }] }),
    /needs "rows" array or "cells" array/
  );
});

test('validateSpec: rejects duplicate sheet names', () => {
  assert.throws(
    () => lib.validateSpec({ sheets: [
      { name: 'X', rows: [] },
      { name: 'X', rows: [] },
    ] }),
    /Duplicate sheet name/
  );
});

test('validateSpec: single-sheet shortcut wraps into sheets array', () => {
  const spec = { name: 'X', rows: [[1, 2, 3]] };
  const r = lib.validateSpec(spec);
  assert.equal(r.sheets.length, 1);
  assert.equal(r.sheets[0].name, 'X');
});

test('validateSpec: array form (--json multi-sheet output) wraps', () => {
  const spec = [
    { name: 'A', rows: [] },
    { name: 'B', rows: [] },
  ];
  const r = lib.validateSpec(spec);
  assert.equal(r.sheets.length, 2);
});

// ---------------------------------------------------------------------------
// parseMarkdownSpec — markdown table → spec
// ---------------------------------------------------------------------------

test('parseMarkdownSpec: single table without heading', () => {
  const md = `| A | B |
|---|---|
| 1 | 2 |
| 3 | 4 |`;
  const r = lib.parseMarkdownSpec(md);
  assert.equal(r.sheets.length, 1);
  assert.deepEqual(r.sheets[0].headers, ['A', 'B']);
  assert.deepEqual(r.sheets[0].rows, [[1, 2], [3, 4]]);
});

test('parseMarkdownSpec: ## headings split into multiple sheets', () => {
  const md = `## First
| A | B |
|---|---|
| 1 | 2 |

## Second
| X | Y |
|---|---|
| a | b |`;
  const r = lib.parseMarkdownSpec(md);
  assert.equal(r.sheets.length, 2);
  assert.equal(r.sheets[0].name, 'First');
  assert.equal(r.sheets[1].name, 'Second');
});

test('parseMarkdownSpec: rejects input with no table', () => {
  assert.throws(() => lib.parseMarkdownSpec('just text\nno tables'), /No markdown table/);
});

// ---------------------------------------------------------------------------
// trySimpleEval — formula evaluation hook
// ---------------------------------------------------------------------------

test('trySimpleEval: SUM with literal args', () => {
  assert.equal(lib.trySimpleEval('=SUM(1,2,3)'), 6);
});

test('trySimpleEval: returns null for cell-ref formulas', () => {
  assert.equal(lib.trySimpleEval('=A1+B1'), null);
});

test('trySimpleEval: returns null for unknown function', () => {
  assert.equal(lib.trySimpleEval('=NOSUCH(1,2)'), null);
});

// ---------------------------------------------------------------------------
// describeNote — comment/note rendering
// ---------------------------------------------------------------------------

test('describeNote: null/empty', () => {
  assert.equal(lib.describeNote(null), null);
  assert.equal(lib.describeNote(undefined), null);
});

test('describeNote: string note returned as-is', () => {
  assert.equal(lib.describeNote('hello'), 'hello');
});

test('describeNote: rich-text note concatenated', () => {
  const note = { texts: [{ text: 'a' }, 'b', { text: 'c' }] };
  assert.equal(lib.describeNote(note), 'abc');
});
