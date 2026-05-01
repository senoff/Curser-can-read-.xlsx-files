// Round-trip metadata fidelity tests.
//
// For each synthetic fixture: read xlsx → emit JSON via xlsx-for-ai → write
// back to a new xlsx → snapshot both → compare. Asserts that values AND
// metadata (column widths, merges, named ranges, frozen panes, hidden rows,
// auto-filter) survive the round-trip.
//
// This is the test that catches the class of bug we found in 1.4.2 (column
// widths silently dropped) and protects against future regressions.
//
// Run: npm test

'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { execFileSync } = require('node:child_process');

const { generateAll } = require('./helpers/synth');
const { loadWorkbook, snapshot, compareSnapshots } = require('./helpers/metadata');

const REPO_ROOT = path.resolve(__dirname, '..');
const CLI = path.join(REPO_ROOT, 'index.js');

let TMP_DIR;
let FIXTURE_NAMES;

test.before(async () => {
  TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'xfa-rt-'));
  // Set sentinel so the CLI doesn't try to self-respawn during the test
  // (otherwise child node spawns add 200ms × N tests of overhead).
  process.env.XLSX_FOR_AI_RESPAWNED = '1';
  FIXTURE_NAMES = await generateAll(path.join(TMP_DIR, 'fixtures'));
});

test.after(() => {
  if (TMP_DIR && fs.existsSync(TMP_DIR)) {
    fs.rmSync(TMP_DIR, { recursive: true, force: true });
  }
});

function runCli(args, opts = {}) {
  return execFileSync(process.execPath, [CLI, ...args], {
    cwd: opts.cwd || TMP_DIR,
    encoding: 'utf8',
    env: { ...process.env, XLSX_FOR_AI_RESPAWNED: '1' },
    ...opts,
  });
}

async function roundTrip(fixturePath, outPath) {
  // Read → JSON spec → write
  const json = runCli([fixturePath, '--json', '--stdout']);
  const specPath = path.join(TMP_DIR, 'spec.json');
  fs.writeFileSync(specPath, json);
  runCli(['write', specPath, '-o', outPath, '--no-report']);
}

// Fixtures and the bugs they currently surface (todo means: known issue, fix
// pending on a separate branch — the test stays in the suite so once fixed,
// removing the `todo` flag re-enables the assertion).
const FIXTURE_TESTS = [
  { name: 'basic-values.xlsx',   todo: false },
  { name: 'widths-layout.xlsx',  todo: false },
  { name: 'merges-names.xlsx',   todo: false },
  { name: 'multi-sheet.xlsx',    todo: false },
  { name: 'annotations.xlsx',    todo: false },
];

for (const { name, todo } of FIXTURE_TESTS) {
  const opts = todo ? { todo } : {};
  test(`round-trip preserves metadata: ${name}`, opts, async () => {
    const inPath = path.join(TMP_DIR, 'fixtures', name);
    const outPath = path.join(TMP_DIR, `out-${name}`);
    await roundTrip(inPath, outPath);

    const wbA = await loadWorkbook(inPath);
    const wbB = await loadWorkbook(outPath);
    const snapA = snapshot(wbA);
    const snapB = snapshot(wbB);
    const diffs = compareSnapshots(snapA, snapB);
    if (diffs.length > 0) {
      assert.fail(
        `Round-trip drift in ${name} (${diffs.length} diff${diffs.length === 1 ? '' : 's'}):\n` +
        diffs.map(d => `  - ${d}`).join('\n')
      );
    }
  });
}

test('--version returns the package version', () => {
  const out = runCli(['--version']).trim();
  const pkg = JSON.parse(fs.readFileSync(path.join(REPO_ROOT, 'package.json'), 'utf8'));
  assert.equal(out, pkg.version);
});

test('-v alias also returns the version', () => {
  const out = runCli(['-v']).trim();
  const pkg = JSON.parse(fs.readFileSync(path.join(REPO_ROOT, 'package.json'), 'utf8'));
  assert.equal(out, pkg.version);
});

test('empty file rejected with friendly error', () => {
  const emptyPath = path.join(TMP_DIR, 'empty.xlsx');
  fs.writeFileSync(emptyPath, '');
  try {
    runCli([emptyPath, '--list-sheets'], { stdio: 'pipe' });
    assert.fail('Expected non-zero exit');
  } catch (err) {
    const stderr = err.stderr ? err.stderr.toString() : '';
    assert.match(stderr, /empty|0 bytes/i, 'should mention empty/0 bytes');
  }
});

test('--list-sheets enumerates sheets without parsing values', async () => {
  const inPath = path.join(TMP_DIR, 'fixtures', 'multi-sheet.xlsx');
  const out = runCli([inPath, '--list-sheets']);
  assert.match(out, /Detail/);
  assert.match(out, /Summary/);
});

test('CSV input is accepted', () => {
  const csvPath = path.join(TMP_DIR, 'sample.csv');
  fs.writeFileSync(csvPath, 'name,age\nAlice,30\nBob,25\n');
  const out = runCli([csvPath, '--md', '--stdout']);
  assert.match(out, /Alice/);
  assert.match(out, /Bob/);
});
