// Tests for --export-redacted-workbook.
//
// We build the same synthetic xlsx, run the redactor, then:
//   1. Open the result with ExcelJS and verify structure preservation
//      (sheet count, sheet names, merges, defined names, formulas).
//   2. Verify cell values are typed placeholders.
//   3. Grep the raw zip for the cell-content sentinel and confirm
//      it's gone.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const JSZip = require('jszip');
const ExcelJS = require('exceljs');

const { build } = require('./fixtures/build');
const { exportRedactedWorkbook } = require('../lib/redactWorkbook');

let fixturePath;
let redactedPath;
let workdir;

test.before(async () => {
  workdir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-for-ai-redact-'));
  fixturePath = path.join(workdir, 'bug-fixture.xlsx');
  redactedPath = path.join(workdir, 'bug-fixture-redacted.xlsx');
  await build(fixturePath);
  await exportRedactedWorkbook(fixturePath, redactedPath);
});

test.after(() => {
  if (workdir) fs.rmSync(workdir, { recursive: true, force: true });
});

test('redacted file exists and is a valid zip', async () => {
  assert.ok(fs.existsSync(redactedPath));
  const buf = fs.readFileSync(redactedPath);
  // ZIP local file header magic
  assert.equal(buf[0], 0x50);
  assert.equal(buf[1], 0x4b);
  await JSZip.loadAsync(buf); // throws if not a valid zip
});

test('structure preserved: sheet count, names, merges, defined names', async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(redactedPath);
  assert.equal(wb.worksheets.length, 3);
  assert.deepEqual(wb.worksheets.map((w) => w.name), ['Sales', 'Config', 'Empty']);

  const sales = wb.getWorksheet('Sales');
  assert.ok(sales.model.merges.length >= 1, 'expected ≥1 merge in Sales');

  const dn = wb.definedNames.model;
  assert.ok(dn.some((d) => d.name === 'Totals'), 'defined name "Totals" must survive redaction');
});

test('formulas preserved: D2, D3 still carry their formula text', async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(redactedPath);
  const sales = wb.getWorksheet('Sales');
  const d2 = sales.getCell('D2').value;
  const d3 = sales.getCell('D3').value;
  // ExcelJS surfaces formulas as { formula, result } or { sharedFormula, result }.
  const f2 = d2 && (d2.formula || d2.sharedFormula);
  const f3 = d3 && (d3.formula || d3.sharedFormula);
  assert.ok(f2, 'D2 must still be a formula cell');
  assert.ok(f3, 'D3 must still be a formula cell');
});

test('numeric cells redacted to 0', async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(redactedPath);
  const sales = wb.getWorksheet('Sales');
  // B2 was 100 in the fixture.
  assert.equal(sales.getCell('B2').value, 0, 'B2 must be redacted to 0');
  assert.equal(sales.getCell('C2').value, 0, 'C2 must be redacted to 0');
});

test('string cells redacted to "x"', async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(redactedPath);
  const sales = wb.getWorksheet('Sales');
  // A1 was "Region" in the fixture.
  const v = sales.getCell('A1').value;
  // Could be "x" string, or rich-text {richText:[{text:"x"}]}
  const flat = typeof v === 'string' ? v : (v && v.richText ? v.richText.map((r) => r.text).join('') : v);
  assert.equal(flat, 'x', 'A1 must be redacted to "x"');
});

test('boolean cell redacted to false', async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(redactedPath);
  const config = wb.getWorksheet('Config');
  // B2 was `true` in the fixture.
  // ExcelJS reads numeric 0 as 0; booleans live in t="b" cells which we
  // rewrite to <v>0</v>. After our rewrite the cell is a numeric 0,
  // which is the documented placeholder for boolean false.
  const v = config.getCell('B2').value;
  assert.ok(v === 0 || v === false, `B2 expected 0 or false, got ${JSON.stringify(v)}`);
});

test('raw zip contains no fixture cell-content sentinel', async () => {
  const buf = fs.readFileSync(redactedPath);
  const zip = await JSZip.loadAsync(buf);
  const sentinels = ['Sensitive Customer Notes Here', 'Region', 'TaxRate'];
  for (const name of Object.keys(zip.files)) {
    const file = zip.file(name);
    if (!file || file.dir) continue;
    // Only check the parts likely to hold user text.
    if (!/\.(xml|rels)$/i.test(name)) continue;
    const xml = await file.async('string');
    for (const s of sentinels) {
      assert.equal(
        xml.includes(s),
        false,
        `redacted zip part ${name} still contains sentinel ${JSON.stringify(s)}`
      );
    }
  }
});
