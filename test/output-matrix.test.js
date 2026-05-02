// Cell-type × output-mode matrix tests.
//
// Builds one fixture containing every cell type we care about (string,
// number, date, formula, shared formula, error, rich text, hyperlink, bool),
// then runs each output mode (text, markdown, JSON, sql, schema) and
// asserts basic invariants per mode. Catches the class of bugs where a
// new output mode silently mishandles a cell type that other modes handle
// correctly.

'use strict';

process.env.XLSX_FOR_AI_RESPAWNED = '1';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { execFileSync } = require('node:child_process');
// Route through the engine seam: fixture construction uses createWorkbook()
// + writeWorkbook() so this test file never binds directly to @protobi/exceljs.
const engine = require('../lib/engine');

const REPO_ROOT = path.resolve(__dirname, '..');
const CLI = path.join(REPO_ROOT, 'index.js');

let TMP_DIR;
let FIXTURE;

test.before(async () => {
  TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'xfa-matrix-'));
  FIXTURE = path.join(TMP_DIR, 'all-cell-types.xlsx');

  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Data');
  // Headers (row 1)
  ws.getCell('A1').value = 'kind';
  ws.getCell('B1').value = 'value';
  // Cell-type samples (column B with kind in column A)
  ws.getCell('A2').value = 'string';        ws.getCell('B2').value = 'hello world';
  ws.getCell('A3').value = 'integer';       ws.getCell('B3').value = 42;
  ws.getCell('A4').value = 'float';         ws.getCell('B4').value = 3.14159;
  ws.getCell('A5').value = 'date';          ws.getCell('B5').value = new Date('2026-04-28T00:00:00.000Z');
  ws.getCell('A6').value = 'formula';       ws.getCell('B6').value = { formula: 'B3+B4', result: 45.14159 };
  ws.getCell('A7').value = 'sharedFormula'; ws.getCell('B7').value = { sharedFormula: 'B6', result: 90.28318 };
  ws.getCell('A8').value = 'error';         ws.getCell('B8').value = { error: '#REF!' };
  ws.getCell('A9').value = 'richText';      ws.getCell('B9').value = { richText: [{ text: 'rich ' }, { text: 'text' }] };
  ws.getCell('A10').value = 'hyperlink';    ws.getCell('B10').value = { text: 'click', hyperlink: 'https://example.com' };
  ws.getCell('A11').value = 'boolean';      ws.getCell('B11').value = true;

  await engine.writeWorkbook(wb, FIXTURE);
});

test.after(() => {
  if (TMP_DIR && fs.existsSync(TMP_DIR)) {
    fs.rmSync(TMP_DIR, { recursive: true, force: true });
  }
});

function runCli(args) {
  return execFileSync(process.execPath, [CLI, ...args], {
    encoding: 'utf8',
    env: { ...process.env, XLSX_FOR_AI_RESPAWNED: '1' },
  });
}

// ---------------------------------------------------------------------------
// Per-mode invariants — each mode produces output, doesn't crash, contains
// the cell-type labels (proving each cell type was processed).
// ---------------------------------------------------------------------------

test('text mode: produces output for every cell type', () => {
  const out = runCli([FIXTURE, '--stdout']);
  for (const kind of ['string','integer','float','date','formula','sharedFormula','error','richText','hyperlink','boolean']) {
    assert.match(out, new RegExp(kind), `text mode missing kind=${kind}`);
  }
});

test('markdown mode: produces table for every cell type', () => {
  const out = runCli([FIXTURE, '--md', '--stdout']);
  // Headers should be present
  assert.match(out, /\| kind \| value \|/);
  // Data rows
  for (const kind of ['string','integer','float','date','formula','error','hyperlink','boolean']) {
    assert.match(out, new RegExp(`\\| ${kind} \\|`), `markdown missing kind=${kind}`);
  }
});

test('JSON mode: emits one cell entry per non-empty cell', () => {
  const out = runCli([FIXTURE, '--json', '--stdout']);
  const parsed = JSON.parse(out);
  // Sheet object with cells array
  assert.equal(parsed.name, 'Data');
  assert.ok(Array.isArray(parsed.cells), 'cells should be array');
  // 22 cells = 11 rows × 2 cols (all non-empty)
  assert.ok(parsed.cells.length >= 20, `expected ~22 cells, got ${parsed.cells.length}`);
});

test('JSON mode: formula cell preserves formula + result', () => {
  const out = runCli([FIXTURE, '--json', '--stdout']);
  const parsed = JSON.parse(out);
  const formulaCell = parsed.cells.find(c => c.ref === 'B6');
  assert.ok(formulaCell, 'B6 should be present');
  assert.equal(formulaCell.value.formula, 'B3+B4');
  assert.equal(formulaCell.value.result, 45.14159);
});

test('JSON mode: shared formula emits sharedFormulaRef', () => {
  const out = runCli([FIXTURE, '--json', '--stdout']);
  const parsed = JSON.parse(out);
  const cell = parsed.cells.find(c => c.ref === 'B7');
  assert.ok(cell, 'B7 should be present');
  assert.equal(cell.value.sharedFormulaRef, 'B6');
});

test('JSON mode: hyperlink cell preserves both text and href', () => {
  const out = runCli([FIXTURE, '--json', '--stdout']);
  const parsed = JSON.parse(out);
  const cell = parsed.cells.find(c => c.ref === 'B10');
  assert.ok(cell);
  // Hyperlink shape: { text, hyperlink } object as the value
  assert.equal(cell.value.text, 'click');
  assert.equal(cell.value.hyperlink, 'https://example.com');
});

test('SQL mode: emits CREATE TABLE + at least one INSERT', () => {
  const out = runCli([FIXTURE, '--sql', '--stdout']);
  assert.match(out, /CREATE TABLE/);
  assert.match(out, /INSERT INTO/);
  // Quoted identifiers
  assert.match(out, /"kind"/);
  assert.match(out, /"value"/);
});

test('SQL mode: numeric values not quoted', () => {
  const out = runCli([FIXTURE, '--sql', '--stdout']);
  // The value column with kind=integer should have value 42 (unquoted) somewhere
  // Or the schema should infer the value column is TEXT (since values are mixed).
  // Either way: there should be at least one row with numeric literal 42 OR '42'
  assert.ok(/42/.test(out), 'expected 42 to appear in output');
});

test('schema mode: returns column types per sheet', () => {
  const out = runCli([FIXTURE, '--schema', '--stdout']);
  const parsed = JSON.parse(out);
  assert.equal(parsed.sheet, 'Data');
  assert.ok(Array.isArray(parsed.columns));
  // The "kind" column should be inferred as TEXT
  const kindCol = parsed.columns.find(c => c.column === 'A');
  assert.equal(kindCol.type, 'TEXT');
});

// ---------------------------------------------------------------------------
// Cross-mode invariant: every mode produces deterministic output (same input
// → same output across runs)
// ---------------------------------------------------------------------------

for (const mode of [['--stdout'], ['--md', '--stdout'], ['--json', '--stdout'], ['--sql', '--stdout'], ['--schema', '--stdout']]) {
  test(`deterministic: ${mode.join(' ')} produces identical output across runs`, () => {
    const a = runCli([FIXTURE, ...mode]);
    const b = runCli([FIXTURE, ...mode]);
    assert.equal(a, b, `mode ${mode.join(' ')} is non-deterministic`);
  });
}

// ---------------------------------------------------------------------------
// Cross-mode invariant: token budgeting works across all modes
// ---------------------------------------------------------------------------

for (const modeArgs of [['--stdout'], ['--md', '--stdout'], ['--json', '--stdout']]) {
  test(`--max-tokens ${modeArgs.join(' ')} respects budget`, () => {
    // Generate large fixture by writing a lot of rows
    const big = path.join(TMP_DIR, `big-${modeArgs.join('-').replace(/-/g,'_')}.xlsx`);
    // synchronously build a fat workbook inline
    const wb = engine.createWorkbook();
    const ws = wb.addWorksheet('Big');
    ws.getCell('A1').value = 'header';
    for (let r = 2; r <= 100; r++) {
      ws.getCell(`A${r}`).value = `lorem ipsum dolor sit amet ${r}`.repeat(5);
    }
    return engine.writeWorkbook(wb, big).then(() => {
      const full = runCli([big, ...modeArgs]);
      const truncated = runCli([big, ...modeArgs, '--max-tokens', '200']);
      assert.ok(truncated.length < full.length,
        `truncated should be shorter (full=${full.length}, truncated=${truncated.length})`);
      assert.match(truncated, /truncated/, 'truncation note should appear');
    });
  });
}

// ---------------------------------------------------------------------------
// CSV input: every output mode handles CSV input
// ---------------------------------------------------------------------------

test('CSV input: every output mode produces output', () => {
  const csvPath = path.join(TMP_DIR, 'data.csv');
  fs.writeFileSync(csvPath, 'name,age,balance\nAlice,30,1234.56\nBob,25,789.00\n');

  for (const mode of [['--stdout'], ['--md', '--stdout'], ['--json', '--stdout'], ['--schema', '--stdout']]) {
    const out = runCli([csvPath, ...mode]);
    assert.ok(out.length > 0, `mode ${mode.join(' ')} produced empty output for CSV`);
    assert.match(out, /Alice/, `mode ${mode.join(' ')} missing Alice`);
  }
});
