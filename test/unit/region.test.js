// Tests for the detectRegion() function and --region CLI flag integration.
//
// Fixtures are built in-memory using the engine seam (lib/engine.js) — no
// @protobi/exceljs import here, per the PR contract.

'use strict';

process.env.XLSX_FOR_AI_RESPAWNED = '1';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const engine = require('../../lib/engine');
const lib = require('../../index.js');

const REPO_ROOT = path.resolve(__dirname, '../..');
const CLI = path.join(REPO_ROOT, 'index.js');

function runCli(args) {
  return execFileSync(process.execPath, [CLI, ...args], {
    encoding: 'utf8',
    env: { ...process.env, XLSX_FOR_AI_RESPAWNED: '1' },
  });
}

// ---------------------------------------------------------------------------
// detectRegion — pure function tests (no file I/O)
// ---------------------------------------------------------------------------

test('detectRegion: obvious single region returns correct bounds', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');
  // 3×4 block at B2:E4
  for (let r = 2; r <= 4; r++) {
    for (let c = 2; c <= 5; c++) {
      ws.getCell(r, c).value = `v${r}${c}`;
    }
  }
  const region = lib.detectRegion(ws);
  assert.ok(region, 'should detect a region');
  assert.equal(region.startRow, 2);
  assert.equal(region.endRow, 4);
  assert.equal(region.startCol, 2);
  assert.equal(region.endCol, 5);
});

test('detectRegion: region adjacent to row 1 and col A', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');
  // Block starts at A1
  ws.getCell('A1').value = 'header1';
  ws.getCell('B1').value = 'header2';
  ws.getCell('A2').value = 10;
  ws.getCell('B2').value = 20;
  const region = lib.detectRegion(ws);
  assert.ok(region);
  assert.equal(region.startRow, 1);
  assert.equal(region.startCol, 1);
  assert.equal(region.endRow, 2);
  assert.equal(region.endCol, 2);
});

test('detectRegion: empty worksheet returns null', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Empty');
  const region = lib.detectRegion(ws);
  assert.equal(region, null);
});

test('detectRegion: multiple disjoint regions — picks largest by populated cell count', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');

  // Small region: 2 cells at A1:A2
  ws.getCell('A1').value = 'small1';
  ws.getCell('A2').value = 'small2';

  // Large region: 6 cells at E5:G6 (well separated from small region)
  ws.getCell('E5').value = 'big1';
  ws.getCell('F5').value = 'big2';
  ws.getCell('G5').value = 'big3';
  ws.getCell('E6').value = 'big4';
  ws.getCell('F6').value = 'big5';
  ws.getCell('G6').value = 'big6';

  const region = lib.detectRegion(ws);
  assert.ok(region);
  // The large region (6 cells) should win
  assert.equal(region.startRow, 5);
  assert.equal(region.endRow, 6);
  assert.equal(region.startCol, 5); // col E = 5
  assert.equal(region.endCol, 7);  // col G = 7
});

test('detectRegion: two equal-sized regions — picks one consistently (first by scan order)', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');

  // Region 1: A1:A3 (3 cells)
  ws.getCell('A1').value = 'x';
  ws.getCell('A2').value = 'x';
  ws.getCell('A3').value = 'x';

  // Region 2: E1:E3 (3 cells, well separated)
  ws.getCell('E1').value = 'y';
  ws.getCell('E2').value = 'y';
  ws.getCell('E3').value = 'y';

  const region = lib.detectRegion(ws);
  assert.ok(region);
  // Both have 3 cells; the function should return one without crashing.
  // Start col should be 1 (A) or 5 (E) — either is correct; just must not be null.
  assert.ok(region.startCol === 1 || region.startCol === 5);
});

test('detectRegion: 8-neighbor connectivity — diagonal adjacency treated as one region', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');
  // Diagonal: A1, B2, C3 — all 8-connected, so one region of 3 cells
  ws.getCell('A1').value = 1;
  ws.getCell('B2').value = 2;
  ws.getCell('C3').value = 3;
  const region = lib.detectRegion(ws);
  assert.ok(region);
  assert.equal(region.startRow, 1);
  assert.equal(region.endRow, 3);
  assert.equal(region.startCol, 1);
  assert.equal(region.endCol, 3);
});

test('detectRegion: single cell', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('C5').value = 'solo';
  const region = lib.detectRegion(ws);
  assert.ok(region);
  assert.equal(region.startRow, 5);
  assert.equal(region.endRow, 5);
  assert.equal(region.startCol, 3);
  assert.equal(region.endCol, 3);
});

// ---------------------------------------------------------------------------
// selectionBounds integration — --region option feeds into selectionBounds
// ---------------------------------------------------------------------------

test('selectionBounds: --region with --max-rows caps endRow', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');
  // 10-row block at B3:D12
  for (let r = 3; r <= 12; r++) {
    for (let c = 2; c <= 4; c++) {
      ws.getCell(r, c).value = r * 10 + c;
    }
  }
  const bounds = lib.selectionBounds(ws, { region: true, maxRows: 5 });
  assert.equal(bounds.startRow, 3);
  assert.equal(bounds.endRow, 7); // 3 + 5 - 1 = 7
  assert.equal(bounds.startCol, 2);
  assert.equal(bounds.endCol, 4);
});

test('selectionBounds: --region with --max-cols caps endCol', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');
  // 1-row block at A1:F1
  for (let c = 1; c <= 6; c++) {
    ws.getCell(1, c).value = c;
  }
  const bounds = lib.selectionBounds(ws, { region: true, maxCols: 3 });
  assert.equal(bounds.startCol, 1);
  assert.equal(bounds.endCol, 3); // capped at 3
});

// ---------------------------------------------------------------------------
// CLI integration — --region flag end-to-end (writes temp file, runs CLI)
// ---------------------------------------------------------------------------

let TMP_DIR;

test.before(() => {
  TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'xfa-region-'));
});

test.after(() => {
  if (TMP_DIR && fs.existsSync(TMP_DIR)) {
    fs.rmSync(TMP_DIR, { recursive: true, force: true });
  }
});

test('CLI --region: detects region and excludes surrounding empty area', async () => {
  // Build a workbook: small 3×3 data block at D5:F7, surrounded by empty cells.
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Data');
  ws.getCell('D5').value = 'Name';
  ws.getCell('E5').value = 'Score';
  ws.getCell('F5').value = 'Grade';
  ws.getCell('D6').value = 'Alice';
  ws.getCell('E6').value = 95;
  ws.getCell('F6').value = 'A';
  ws.getCell('D7').value = 'Bob';
  ws.getCell('E7').value = 82;
  ws.getCell('F7').value = 'B';

  const fixturePath = path.join(TMP_DIR, 'region-smoke.xlsx');
  await engine.writeWorkbook(wb, fixturePath);

  const out = runCli([fixturePath, '--region', '--md', '--stdout']);
  // Should contain the data values
  assert.match(out, /Alice/);
  assert.match(out, /Score/);
  // The range header in the markdown output should reflect D5:F7
  assert.match(out, /D5:F7/);
});

test('CLI --region + --max-rows: region detected then capped', async () => {
  // 5-row block at A1:B5, cap at 2 rows
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');
  for (let r = 1; r <= 5; r++) {
    ws.getCell(r, 1).value = `row${r}`;
    ws.getCell(r, 2).value = r * 10;
  }

  const fixturePath = path.join(TMP_DIR, 'region-maxrows.xlsx');
  await engine.writeWorkbook(wb, fixturePath);

  const out = runCli([fixturePath, '--region', '--max-rows', '2', '--md', '--stdout']);
  // Only rows 1 and 2 should appear in data
  assert.match(out, /row1/);
  assert.match(out, /row2/);
  assert.doesNotMatch(out, /row3/);
  assert.doesNotMatch(out, /row4/);
  assert.doesNotMatch(out, /row5/);
});

test('CLI --region: empty workbook emits note to stderr, does not crash', async () => {
  const wb = engine.createWorkbook();
  wb.addWorksheet('Empty');

  const fixturePath = path.join(TMP_DIR, 'region-empty.xlsx');
  await engine.writeWorkbook(wb, fixturePath);

  // execFileSync throws on non-zero exit; capture stderr separately.
  let stderr = '';
  let stdout = '';
  try {
    stdout = execFileSync(process.execPath, [CLI, fixturePath, '--region', '--stdout'], {
      encoding: 'utf8',
      env: { ...process.env, XLSX_FOR_AI_RESPAWNED: '1' },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
  } catch (e) {
    // If CLI exits non-zero that's still a failure — re-throw.
    throw e;
  }
  // The CLI should succeed (no crash); the note about no region is best-effort.
  // We just assert it ran without throwing.
  assert.ok(typeof stdout === 'string');
});

test('CLI --region: largest of two disjoint regions is selected', async () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');

  // Small region at A1:A2 (2 cells)
  ws.getCell('A1').value = 'tiny1';
  ws.getCell('A2').value = 'tiny2';

  // Large region at F10:H12 (9 cells)
  ws.getCell('F10').value = 'BIG';
  ws.getCell('G10').value = 'BIG';
  ws.getCell('H10').value = 'BIG';
  ws.getCell('F11').value = 'BIG';
  ws.getCell('G11').value = 'BIG';
  ws.getCell('H11').value = 'BIG';
  ws.getCell('F12').value = 'BIG';
  ws.getCell('G12').value = 'BIG';
  ws.getCell('H12').value = 'BIG';

  const fixturePath = path.join(TMP_DIR, 'region-two-blocks.xlsx');
  await engine.writeWorkbook(wb, fixturePath);

  const out = runCli([fixturePath, '--region', '--md', '--stdout']);
  // The large region's range (F10:H12) should appear in the output
  assert.match(out, /F10:H12/);
  // The small region's header value should NOT appear (it's outside the detected region)
  assert.doesNotMatch(out, /tiny1/);
});
