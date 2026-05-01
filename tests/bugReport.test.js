// Tests for --report-bug.
//
// We build a small synthetic xlsx, run generateBugReport on it, then
// check:
//   1. The output validates against the expected JSON shape.
//   2. The output contains zero strings sourced from cell content
//      (the fixture deliberately puts a unique sentinel in a merged
//      cell — the report MUST NOT contain it anywhere).

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');

const { build } = require('./fixtures/build');
const { generateBugReport, writeBugReport } = require('../lib/bugReport');

// Per-test sentinels MUST NOT appear in the bug report. The fixture
// uses these as cell values and named-range targets.
const CELL_VALUE_SENTINELS = [
  'Sensitive Customer Notes Here',
  'Region', 'North', 'South',
  'TaxRate', 'Active', 'Today',
  // Numeric values written into cells:
  '0.075',
  // Formulas written into cells:
  'B2+C2', 'B3+C3',
];

// Defined-name TARGETS (the formula side) MUST be absent. The defined
// NAME ("Totals") is allowed to appear because that's what we
// deliberately export.
const DEFINED_NAME_TARGET = 'Sales!$D$2:$D$3';

let fixturePath;
let workdir;

test.before(async () => {
  workdir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-for-ai-bugreport-'));
  fixturePath = path.join(workdir, 'bug-fixture.xlsx');
  await build(fixturePath);
});

test.after(() => {
  if (workdir) fs.rmSync(workdir, { recursive: true, force: true });
});

test('generateBugReport returns a v1 report with expected top-level shape', async () => {
  const r = await generateBugReport(fixturePath);
  assert.equal(r.schema, 'xlsx-for-ai/bug-report/v1');
  assert.match(r.generatedAt, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/);
  assert.equal(r.tool.name, 'xlsx-for-ai');
  assert.match(r.tool.version, /^\d+\.\d+\.\d+/);
  assert.ok(r.runtime.node.startsWith('v'));
  assert.ok(typeof r.runtime.platform === 'string');
  assert.ok(typeof r.runtime.arch === 'string');
});

test('file section reports basename + size, never the absolute path', async () => {
  const r = await generateBugReport(fixturePath);
  assert.equal(r.file.basename, 'bug-fixture.xlsx');
  assert.equal(r.file.ext, '.xlsx');
  assert.ok(r.file.sizeBytes > 0);
  // No leak of the absolute path / tmpdir name.
  const blob = JSON.stringify(r);
  assert.equal(blob.includes(fixturePath), false, 'report must not contain absolute fixture path');
  assert.equal(blob.includes(workdir), false, 'report must not contain tmpdir path');
});

test('workbook section: 3 sheets, ≥1 merge, defined-name name only', async () => {
  const r = await generateBugReport(fixturePath);
  assert.equal(r.workbook.sheetCount, 3);
  assert.ok(r.workbook.mergedRangeCountTotal >= 1, 'expected ≥1 merge');
  assert.equal(r.workbook.namedRangesCount, 1);
  assert.deepEqual(r.workbook.definedNames, ['Totals']);
  assert.equal(r.workbook.perSheet.length, 3);
  for (const s of r.workbook.perSheet) {
    assert.ok(typeof s.rows === 'number');
    assert.ok(typeof s.cols === 'number');
    assert.ok(typeof s.merges === 'number');
  }
});

test('report contains ZERO cell-value strings from the fixture', async () => {
  const r = await generateBugReport(fixturePath);
  const blob = JSON.stringify(r);
  for (const sentinel of CELL_VALUE_SENTINELS) {
    assert.equal(
      blob.includes(sentinel),
      false,
      `bug report leaked cell value: ${JSON.stringify(sentinel)}`
    );
  }
});

test('report contains ZERO defined-name TARGET formulas (only the name)', async () => {
  const r = await generateBugReport(fixturePath);
  const blob = JSON.stringify(r);
  assert.equal(
    blob.includes(DEFINED_NAME_TARGET),
    false,
    `bug report leaked defined-name target: ${DEFINED_NAME_TARGET}`
  );
});

test('writeBugReport writes a JSON file named with an ISO timestamp', async () => {
  const r = await generateBugReport(fixturePath);
  const outPath = writeBugReport(r, workdir);
  assert.ok(fs.existsSync(outPath));
  assert.match(path.basename(outPath), /^xlsx-for-ai-bugreport-\d{4}-\d{2}-\d{2}T.+\.json$/);
  const parsed = JSON.parse(fs.readFileSync(outPath, 'utf8'));
  assert.equal(parsed.schema, 'xlsx-for-ai/bug-report/v1');
});
