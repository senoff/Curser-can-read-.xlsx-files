// Regression test for detectRegion DoS surface (M1 in 2026-05-02 audit).
//
// ExcelJS reports rowCount/columnCount as the highest USED cell, not actual
// storage. A workbook with one cell at the far corner reports >17B coordinates;
// the old detectRegion implementation iterated the full 2D space and would
// hang the CLI on a malicious or pathologically-shaped workbook. The fix caps
// the scan: rowCount × colCount > 5_000_000 → return null + console.warn.

'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { detectRegion } = require('../index.js');
const engine = require('../lib/engine');

test('detectRegion: workbook reporting pathological dimensions returns null fast', () => {
  // Build a workbook where rowCount/columnCount are very high (one populated
  // cell at the far corner) so the old implementation would have iterated
  // billions of coordinates.
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = 'start';
  // XFD1048576 is the last addressable cell in modern Excel.
  ws.getCell('XFD1048576').value = 'end';

  // Sanity: ExcelJS should report dimensions well over the 5M cap.
  assert.ok(
    ws.rowCount * ws.columnCount > 5_000_000,
    `expected pathological dims, got ${ws.rowCount}×${ws.columnCount}`,
  );

  // Capture warnings so we can verify the audit-trail message.
  const origWarn = console.warn;
  const warnings = [];
  console.warn = (msg) => warnings.push(String(msg));
  let result;
  let elapsed;
  try {
    const start = Date.now();
    result = detectRegion(ws);
    elapsed = Date.now() - start;
  } finally {
    console.warn = origWarn;
  }

  assert.equal(result, null, 'expected detectRegion to return null for over-cap workbook');
  assert.ok(elapsed < 2000, `expected detectRegion to return in <2s, took ${elapsed}ms`);
  assert.ok(
    warnings.some((w) => w.includes('skipping region detection')),
    `expected stderr note about skipping; got ${JSON.stringify(warnings)}`,
  );
});

test('detectRegion: small in-bounds workbook still returns a real bounding box', () => {
  const wb = engine.createWorkbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('B2').value = 'a';
  ws.getCell('B3').value = 'b';
  ws.getCell('C2').value = 'c';

  const result = detectRegion(ws);
  assert.ok(result, 'expected non-null bounds for a small populated sheet');
  // Bounds wraps the populated 2×2 cluster anchored at B2.
  assert.equal(result.startRow, 2);
  assert.equal(result.startCol, 2);
  assert.equal(result.endRow, 3);
  assert.equal(result.endCol, 3);
});
