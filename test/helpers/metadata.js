// Metadata-snapshot helper: capture the structural metadata of a workbook for
// round-trip comparison. The standard --diff only compares cell values; this
// helper captures column widths, merges, named ranges, frozen panes, hidden
// rows/cols, and auto-filter ranges so we can assert structural fidelity in
// addition to value fidelity.

'use strict';

// Route through the engine seam so metadata.js never binds directly to
// @protobi/exceljs. The returned workbook is still an ExcelJS object; the
// seam centralises which engine produces it.
const engine = require('../../lib/engine');

function colLetter(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

async function loadWorkbook(filePath) {
  return engine.loadWorkbook(filePath);
}

function snapshotSheet(ws) {
  const widths = {};
  const hiddenCols = [];
  for (let c = 1; c <= ws.columnCount; c++) {
    const col = ws.getColumn(c);
    if (col.width != null) widths[colLetter(c)] = col.width;
    if (col.hidden) hiddenCols.push(colLetter(c));
  }
  // Walk past columnCount for empty-but-styled trailing columns
  const allCols = ws.columns || [];
  for (let i = ws.columnCount; i < allCols.length; i++) {
    const col = allCols[i];
    if (!col) continue;
    const letter = colLetter(i + 1);
    if (col.width != null) widths[letter] = col.width;
    if (col.hidden) hiddenCols.push(letter);
  }
  const hiddenRows = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    if (ws.getRow(r).hidden) hiddenRows.push(r);
  }
  const frozen = (ws.views || []).find(v => v.state === 'frozen');
  const frozenSig = frozen ? `${frozen.ySplit ?? 0}/${frozen.xSplit ?? 0}` : null;
  const merges = Object.keys(ws._merges || {}).sort();
  const af = ws.autoFilter
    ? (typeof ws.autoFilter === 'string' ? ws.autoFilter : ws.autoFilter.ref || null)
    : null;
  return {
    name: ws.name,
    state: ws.state || 'visible',
    rowCount: ws.rowCount,
    columnCount: ws.columnCount,
    widths,
    hiddenCols: hiddenCols.sort(),
    hiddenRows,
    frozen: frozenSig,
    merges,
    autoFilter: af,
  };
}

function snapshotNamedRanges(wb) {
  try {
    const m = wb.definedNames?.model;
    if (!Array.isArray(m)) return [];
    return m
      .map(d => `${d.name}=${(d.ranges || []).join(',')}`)
      .sort();
  } catch (_) { return []; }
}

// Returns the cell values keyed by sheet!ref so we can compare round-trip.
// Values are normalized via plainValue-like logic for comparison.
function snapshotCells(ws) {
  const cells = {};
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 1; c <= ws.columnCount; c++) {
      const cell = ws.getRow(r).getCell(c);
      const v = cellPlain(cell.value);
      if (v != null) {
        cells[`${colLetter(c)}${r}`] = v;
      }
    }
  }
  return cells;
}

function cellPlain(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(r => r.text).join('');
    if (v.hyperlink) return v.text || v.hyperlink;
    if (v.formula || v.sharedFormula || v.sharedFormulaRef) {
      const r = v.result;
      if (r == null) return null;
      if (r instanceof Date) return r.toISOString().slice(0, 10);
      if (typeof r === 'object') {
        if (r.error) return `#${r.error}`;
        if (r.richText) return r.richText.map(x => x.text).join('');
        return String(r);
      }
      return String(r);
    }
    if (v.error) return `#${v.error}`;
    return JSON.stringify(v);
  }
  return String(v);
}

function snapshot(wb, opts = {}) {
  const skipReportSheet = opts.skipReportSheet !== false; // default true
  const userSheets = wb.worksheets.filter(s => !skipReportSheet || s.name !== '_xlsx-for-ai');
  return {
    sheets: Object.fromEntries(
      userSheets.map(s => [s.name, snapshotSheet(s)])
    ),
    cells: Object.fromEntries(
      userSheets.map(s => [s.name, snapshotCells(s)])
    ),
    namedRanges: snapshotNamedRanges(wb),
  };
}

// Returns a list of human-readable diff descriptions, empty if identical.
function compareSnapshots(a, b) {
  const diffs = [];
  const allSheets = new Set([...Object.keys(a.sheets), ...Object.keys(b.sheets)]);
  for (const name of allSheets) {
    const sa = a.sheets[name];
    const sb = b.sheets[name];
    if (!sa) { diffs.push(`SHEET-ADDED ${name}`); continue; }
    if (!sb) { diffs.push(`SHEET-REMOVED ${name}`); continue; }
    if (sa.frozen !== sb.frozen) diffs.push(`FROZEN ${name}: ${sa.frozen} → ${sb.frozen}`);
    if (sa.state !== sb.state) diffs.push(`STATE ${name}: ${sa.state} → ${sb.state}`);
    if (sa.autoFilter !== sb.autoFilter) diffs.push(`AUTOFILTER ${name}: ${sa.autoFilter} → ${sb.autoFilter}`);
    // Widths
    const wAll = new Set([...Object.keys(sa.widths), ...Object.keys(sb.widths)]);
    for (const L of wAll) {
      if (sa.widths[L] !== sb.widths[L]) {
        diffs.push(`WIDTH ${name}!${L}: ${sa.widths[L] ?? '-'} → ${sb.widths[L] ?? '-'}`);
      }
    }
    // Hidden cols
    if (sa.hiddenCols.join(',') !== sb.hiddenCols.join(',')) {
      diffs.push(`HIDDEN-COLS ${name}: [${sa.hiddenCols.join(',')}] → [${sb.hiddenCols.join(',')}]`);
    }
    // Hidden rows
    if (sa.hiddenRows.join(',') !== sb.hiddenRows.join(',')) {
      diffs.push(`HIDDEN-ROWS ${name}: [${sa.hiddenRows.join(',')}] → [${sb.hiddenRows.join(',')}]`);
    }
    // Merges
    const mLost = sa.merges.filter(m => !sb.merges.includes(m));
    const mAdded = sb.merges.filter(m => !sa.merges.includes(m));
    if (mLost.length) diffs.push(`MERGE-LOST ${name}: ${mLost.slice(0, 3).join(', ')}${mLost.length > 3 ? '...' : ''}`);
    if (mAdded.length) diffs.push(`MERGE-ADDED ${name}: ${mAdded.slice(0, 3).join(', ')}${mAdded.length > 3 ? '...' : ''}`);
    // Cell values
    const ca = a.cells[name] || {};
    const cb = b.cells[name] || {};
    const refsAll = new Set([...Object.keys(ca), ...Object.keys(cb)]);
    let cellChanges = 0;
    for (const ref of refsAll) {
      if (ca[ref] !== cb[ref]) {
        cellChanges++;
        if (cellChanges <= 5) {
          diffs.push(`CELL ${name}!${ref}: ${ca[ref] ?? '-'} → ${cb[ref] ?? '-'}`);
        }
      }
    }
    if (cellChanges > 5) diffs.push(`... and ${cellChanges - 5} more cell changes in ${name}`);
  }
  // Named ranges
  const nrLost = a.namedRanges.filter(n => !b.namedRanges.includes(n));
  const nrAdded = b.namedRanges.filter(n => !a.namedRanges.includes(n));
  if (nrLost.length) diffs.push(`NAMED-LOST: ${nrLost.slice(0, 3).join('; ')}${nrLost.length > 3 ? '...' : ''}`);
  if (nrAdded.length) diffs.push(`NAMED-ADDED: ${nrAdded.slice(0, 3).join('; ')}${nrAdded.length > 3 ? '...' : ''}`);
  return diffs;
}

module.exports = { loadWorkbook, snapshot, compareSnapshots, colLetter };
