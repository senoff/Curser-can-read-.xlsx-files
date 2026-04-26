#!/usr/bin/env node

// Self-respawn with a larger V8 heap before loading anything else.
// Some real-world .xlsx files (sub-1MB on disk but with huge calc chains or
// shared-string tables) blow Node's default ~4GB heap during parse. Re-execing
// with --max-old-space-size=8192 fixes this transparently. The sentinel env
// var prevents an infinite respawn loop.
if (!process.env.XLSX_FOR_AI_RESPAWNED) {
  const v8 = require('v8');
  const heapLimitMB = v8.getHeapStatistics().heap_size_limit / 1024 / 1024;
  if (heapLimitMB < 8000) {
    const { spawnSync } = require('child_process');
    const r = spawnSync(
      process.execPath,
      ['--max-old-space-size=8192', __filename, ...process.argv.slice(2)],
      { stdio: 'inherit', env: { ...process.env, XLSX_FOR_AI_RESPAWNED: '1' } }
    );
    process.exit(r.status ?? 1);
  }
}

const path = require('path');
const fs   = require('fs');
const ExcelJS = require('exceljs');

// Lazy-load heavy deps only when their feature is used (keeps cold start fast
// for the common --stdout / --json / --md path that needs none of them).
let _xlsxLib, _papaLib, _formulaJsLib, _tokenizerLib;
const lazyXlsx       = () => (_xlsxLib       ??= require('xlsx'));
const lazyPapa       = () => (_papaLib       ??= require('papaparse'));
const lazyFormulaJs  = () => (_formulaJsLib  ??= require('@formulajs/formulajs'));
const lazyTokenizer  = () => (_tokenizerLib  ??= require('gpt-tokenizer'));

// ---------------------------------------------------------------------------
// Argument parsing
// ---------------------------------------------------------------------------

function parseArgs(argv) {
  const opts = {
    positional: [],
    listSheets: false,
    stdout: false,
    json: false,
    md: false,
    sql: false,
    schema: false,
    compact: false,
    evaluate: false,
    stream: false,
    diff: null,
    range: null,
    namedRange: null,
    maxRows: null,
    maxCols: null,
    maxTokens: null,
    help: false,
  };
  let i = 0;
  while (i < argv.length) {
    const arg = argv[i];
    if      (arg === '--list-sheets')   opts.listSheets = true;
    else if (arg === '--stdout')        opts.stdout = true;
    else if (arg === '--json')          opts.json = true;
    else if (arg === '--md')            opts.md = true;
    else if (arg === '--sql')           opts.sql = true;
    else if (arg === '--schema')        opts.schema = true;
    else if (arg === '--compact')       opts.compact = true;
    else if (arg === '--evaluate')      opts.evaluate = true;
    else if (arg === '--stream')        opts.stream = true;
    else if (arg === '--diff')        { opts.diff = argv[++i]; }
    else if (arg === '--range')       { opts.range = argv[++i]; }
    else if (arg === '--named-range') { opts.namedRange = argv[++i]; }
    else if (arg === '--max-rows')    { opts.maxRows = parseInt(argv[++i], 10); }
    else if (arg === '--max-cols')    { opts.maxCols = parseInt(argv[++i], 10); }
    else if (arg === '--max-tokens')  { opts.maxTokens = parseInt(argv[++i], 10); }
    else if (arg === '-h' || arg === '--help') opts.help = true;
    else                                opts.positional.push(arg);
    i++;
  }
  return opts;
}

function printHelp() {
  console.log(`Usage: npx xlsx-for-ai <file> [sheetName] [options]

Converts spreadsheets to text, markdown, JSON, SQL, or schema dumps that AI
coding agents can read. Preserves values, formulas, formatting, layout.

Input formats: .xlsx .xls .xlsb .ods .csv .tsv

Output modes (mutually exclusive; default = text):
  --md              Markdown tables — best LLM comprehension per token
  --json            Structured JSON, one object per cell
  --sql             SQL CREATE TABLE + INSERT statements (uses --schema)
  --schema          Inferred per-column schema (name, type, sample) as JSON

Selection:
  [sheetName]       Positional second arg, dump only this sheet
  --range A1:D50    Dump only this rectangular range
  --named-range NM  Dump only the cells covered by this defined name
  --max-rows N      Limit to first N rows per sheet
  --max-cols N      Limit to first N columns per sheet

Output control:
  --stdout          Print to stdout instead of writing files in .xlsx-read/
  --list-sheets     Print sheet names + dimensions and exit
  --compact         Suppress noisy default tags (default colors, General fmt)
  --max-tokens N    Truncate output to ~N tokens (cl100k_base proxy);
                    appends a tail summary noting what was dropped
  --evaluate        Promote cached formula results to primary value;
                    re-evaluate simple formulas via formulajs

Other modes:
  --diff OTHER      Diff this workbook vs OTHER, emit changed cells/sheets
  --stream          Streaming reader for huge .xlsx files (>100MB);
                    emits row-by-row, drops some sheet metadata

Misc:
  -h, --help        Show this help

Examples:
  npx xlsx-for-ai data.xlsx
  npx xlsx-for-ai data.xlsx --md --stdout
  npx xlsx-for-ai data.xlsx --json --max-tokens 8000 --stdout
  npx xlsx-for-ai data.csv --md --stdout
  npx xlsx-for-ai data.xlsx --range B2:F100 --stdout
  npx xlsx-for-ai data.xlsx --named-range MyTotals --stdout
  npx xlsx-for-ai data.xlsx --sql --stdout > schema.sql
  npx xlsx-for-ai old.xlsx --diff new.xlsx --stdout
  npx xlsx-for-ai huge.xlsx --stream --stdout

Note: this package was previously published as 'cursor-reads-xlsx';
that command name still works as an alias.`);
}

// ---------------------------------------------------------------------------
// Helpers (ref math, formatting)
// ---------------------------------------------------------------------------

function colLetter(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function colNum(letters) {
  let n = 0;
  const u = letters.toUpperCase();
  for (let i = 0; i < u.length; i++) {
    n = n * 26 + (u.charCodeAt(i) - 64);
  }
  return n;
}

// Parse "A1:D50" or "B2" into {startCol, startRow, endCol, endRow} (1-indexed).
function parseRange(s) {
  if (!s) return null;
  const parts = s.split(':');
  const m1 = /^([A-Z]+)(\d+)$/i.exec(parts[0]);
  if (!m1) throw new Error(`Invalid range: ${s}`);
  const startCol = colNum(m1[1]);
  const startRow = parseInt(m1[2], 10);
  if (parts.length === 1) {
    return { startCol, startRow, endCol: startCol, endRow: startRow };
  }
  const m2 = /^([A-Z]+)(\d+)$/i.exec(parts[1]);
  if (!m2) throw new Error(`Invalid range: ${s}`);
  return {
    startCol,
    startRow,
    endCol: colNum(m2[1]),
    endRow: parseInt(m2[2], 10),
  };
}

const DEFAULT_TEXT_COLORS = new Set([
  'FF000000', 'FF1F1F1F', 'FF222120', 'FF333333',
]);
function isDefaultTextColor(argb) {
  return argb && DEFAULT_TEXT_COLORS.has(argb.toUpperCase());
}

function describeFill(fill, compact) {
  if (!fill || (fill.type === 'pattern' && fill.pattern === 'none')) return null;
  if (fill.type === 'pattern' && fill.fgColor?.argb) {
    if (compact && /^FF?FFFFFF$/i.test(fill.fgColor.argb)) return null;
    return `fill:${fill.fgColor.argb}`;
  }
  return null;
}

function describeFont(font, compact) {
  const parts = [];
  if (font?.bold)   parts.push('bold');
  if (font?.italic) parts.push('italic');
  if (font?.color?.argb && !(compact && isDefaultTextColor(font.color.argb))) {
    parts.push(`color:${font.color.argb}`);
  }
  return parts;
}

function formatValue(v) {
  if (v == null) return '""';
  if (v instanceof Date) return `"${v.toISOString().slice(0, 10)}"`;
  if (typeof v === 'object' && v.richText) {
    return `"${v.richText.map(r => r.text).join('')}"`;
  }
  if (typeof v === 'object' && v.hyperlink) {
    return `"${v.text || v.hyperlink}"`;
  }
  if (typeof v === 'object' && (v.formula || v.sharedFormula)) {
    const result = v.result;
    if (result == null) return '""';
    if (typeof result === 'object') {
      if (result.error)    return `"#${result.error}"`;
      if (result.richText) return `"${result.richText.map(r => r.text).join('')}"`;
      return JSON.stringify(result);
    }
    if (typeof result === 'string') return `"${result}"`;
    return String(result);
  }
  if (typeof v === 'object' && v.error) return `"#${v.error}"`;
  if (typeof v === 'string') return `"${v}"`;
  return String(v);
}

// Plain (unquoted) value extraction — for markdown/SQL/schema where we don't
// want JSON quoting. Returns string or null for empty cells.
function plainValue(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(r => r.text).join('');
    if (v.hyperlink) return v.text || v.hyperlink;
    if (v.formula || v.sharedFormula) {
      const r = v.result;
      if (r == null) return null;
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

function describeNote(note) {
  if (!note) return null;
  if (typeof note === 'string') return note;
  if (note.texts) {
    return note.texts.map(t => (typeof t === 'string' ? t : t.text || '')).join('');
  }
  return null;
}

// ---------------------------------------------------------------------------
// Named ranges
// ---------------------------------------------------------------------------

function getNamedRanges(wb, sheetName) {
  const results = [];
  try {
    const model = wb.definedNames?.model;
    if (!Array.isArray(model)) return results;
    for (const def of model) {
      if (!def.ranges?.length) continue;
      if (sheetName) {
        const relevant = def.ranges.filter(r => r.includes(sheetName + '!'));
        if (relevant.length) results.push({ name: def.name, ranges: relevant });
      } else {
        results.push({ name: def.name, ranges: def.ranges });
      }
    }
  } catch (_) {}
  return results;
}

// Resolve a named range to {sheet, range} pieces. Excel names look like
// 'Sheet1!$A$1:$D$10' (absolute) or 'Sheet1!A1:D10'.
function resolveNamedRange(wb, name) {
  const model = wb.definedNames?.model;
  if (!Array.isArray(model)) return null;
  const def = model.find(d => d.name === name);
  if (!def || !def.ranges?.length) return null;
  const ref = def.ranges[0];
  const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(ref);
  if (!m) return null;
  const sheetName = (m[1] || m[2]).trim();
  const rangeStr = m[3].replace(/\$/g, '');
  return { sheet: sheetName, range: parseRange(rangeStr) };
}

// ---------------------------------------------------------------------------
// Selection bounds — combines --range, --named-range, --max-rows/cols, sheet
// dimensions into a single {startRow, startCol, endRow, endCol}.
// ---------------------------------------------------------------------------

function selectionBounds(ws, opts) {
  let bounds = null;
  if (opts.range) {
    bounds = parseRange(opts.range);
  } else if (opts.namedRangeBounds) {
    bounds = opts.namedRangeBounds;
  }
  const startRow = bounds ? bounds.startRow : 1;
  const startCol = bounds ? bounds.startCol : 1;
  let endRow = bounds ? bounds.endRow : ws.rowCount;
  let endCol = bounds ? bounds.endCol : ws.columnCount;
  if (opts.maxRows) endRow = Math.min(endRow, startRow + opts.maxRows - 1);
  if (opts.maxCols) endCol = Math.min(endCol, startCol + opts.maxCols - 1);
  return { startRow, startCol, endRow, endCol };
}

// ---------------------------------------------------------------------------
// Sheet dump (text)
// ---------------------------------------------------------------------------

function dumpSheet(ws, wb, opts = {}) {
  const { compact = false } = opts;
  const { startRow, startCol, endRow, endCol } = selectionBounds(ws, opts);
  const lines = [];

  lines.push(`=== Sheet: ${ws.name} ===`);

  const frozen = (ws.views || []).find(v => v.state === 'frozen');
  if (frozen) lines.push(`Frozen: row ${frozen.ySplit ?? 0}, col ${frozen.xSplit ?? 0}`);

  // Columns
  const colWidths = [];
  const hiddenCols = [];
  for (let c = startCol; c <= endCol; c++) {
    const col = ws.getColumn(c);
    const letter = colLetter(c);
    if (col.hidden) hiddenCols.push(letter);
    if (col.width) colWidths.push(`${letter}(${Math.round(col.width)})`);
  }
  if (colWidths.length) lines.push(`Columns: ${colWidths.join(' ')}`);
  if (hiddenCols.length) lines.push(`Hidden columns: ${hiddenCols.join(', ')}`);
  if (opts.maxCols && ws.columnCount > endCol) {
    lines.push(`(${ws.columnCount - endCol} more columns truncated)`);
  }

  const merges = Object.keys(ws._merges || {});
  if (merges.length) lines.push(`Merged: ${merges.join(', ')}`);

  if (ws.autoFilter) {
    const af = typeof ws.autoFilter === 'string'
      ? ws.autoFilter
      : (ws.autoFilter.ref || JSON.stringify(ws.autoFilter));
    lines.push(`Auto-filter: ${af}`);
  }

  try { if (ws.pageSetup?.printArea) lines.push(`Print area: ${ws.pageSetup.printArea}`); } catch (_) {}

  const namedRanges = getNamedRanges(wb, ws.name);
  if (namedRanges.length) {
    lines.push(`Named ranges:`);
    for (const nr of namedRanges) lines.push(`  ${nr.name}: ${nr.ranges.join(', ')}`);
  }

  // Tables
  try {
    const tableMap = ws.tables;
    if (tableMap && typeof tableMap === 'object') {
      const tables = typeof tableMap.forEach === 'function'
        ? (() => { const a = []; tableMap.forEach(t => a.push(t)); return a; })()
        : Object.values(tableMap);
      for (const t of tables) {
        const model = t.table || t.model || t;
        const name = model.name || model.displayName || '(unnamed)';
        const ref = model.ref || model.tableRef || '';
        const cols = (model.columns || []).map(c => c.name).filter(Boolean);
        let desc = `Table: "${name}" ${ref}`;
        if (cols.length) desc += ` — columns: ${cols.join(', ')}`;
        lines.push(desc);
      }
    }
  } catch (_) {}

  try {
    const images = typeof ws.getImages === 'function' ? ws.getImages() : [];
    for (const img of images) {
      if (img.range?.tl) {
        const tl = img.range.tl, br = img.range.br;
        if (br) lines.push(`Image: ${colLetter(Math.floor(tl.col)+1)}${Math.floor(tl.row)+1} to ${colLetter(Math.floor(br.col)+1)}${Math.floor(br.row)+1}`);
        else    lines.push(`Image at: ${colLetter(Math.floor(tl.col)+1)}${Math.floor(tl.row)+1}`);
      }
    }
  } catch (_) {}

  lines.push('');

  for (let r = startRow; r <= endRow; r++) {
    const row = ws.getRow(r);
    const cells = [];
    const isHidden = row.hidden;
    for (let c = startCol; c <= endCol; c++) {
      const cell = row.getCell(c);
      const raw = cell.value;
      if (raw == null || raw === '') continue;
      const ref = `${colLetter(c)}${r}`;
      const tags = [];
      if (cell.type === ExcelJS.ValueType.Formula && typeof raw === 'object') {
        if (raw.formula) tags.push(`formula: =${raw.formula}`);
        else if (raw.sharedFormula) tags.push(`shared formula ref: ${raw.sharedFormula}`);
      }
      if (cell.numFmt && cell.numFmt !== 'General') tags.push(`numFmt: ${cell.numFmt}`);
      const fontTags = describeFont(cell.font, compact);
      if (fontTags.length) tags.push(...fontTags);
      const fillDesc = describeFill(cell.fill, compact);
      if (fillDesc) tags.push(fillDesc);
      if (cell.alignment?.horizontal && cell.alignment.horizontal !== 'general') tags.push(`align:${cell.alignment.horizontal}`);
      if (cell.hyperlink) tags.push(`link: ${cell.hyperlink}`);
      else if (typeof raw === 'object' && raw.hyperlink) tags.push(`link: ${raw.hyperlink}`);
      const noteText = describeNote(cell.note);
      if (noteText) tags.push(`note: ${noteText.replace(/\n/g, ' ').trim()}`);
      if (cell.dataValidation) {
        const dv = cell.dataValidation;
        if (dv.type === 'list' && dv.formulae?.length) tags.push(`validation: list [${dv.formulae[0]}]`);
        else if (dv.type) {
          const parts = [dv.type];
          if (dv.operator) parts.push(dv.operator);
          if (dv.formulae?.length) parts.push(dv.formulae.join(', '));
          tags.push(`validation: ${parts.join(' ')}`);
        }
      }
      const displayVal = formatValue(raw);
      const tagStr = tags.length ? `  [${tags.join('] [')}]` : '';
      cells.push(`  ${ref}: ${displayVal}${tagStr}`);
    }
    if (cells.length === 0) {
      const hiddenTag = isHidden ? ' [hidden]' : '';
      lines.push(`--- Row ${r} (empty)${hiddenTag} ---`);
    } else {
      const rowBold = row.font?.bold ? ' [bold]' : '';
      const hiddenTag = isHidden ? ' [hidden]' : '';
      lines.push(`--- Row ${r}${rowBold}${hiddenTag} ---`);
      lines.push(...cells);
    }
  }

  if (opts.maxRows && ws.rowCount > endRow) {
    lines.push('');
    lines.push(`... ${ws.rowCount - endRow} more rows (truncated)`);
  }

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Markdown dump (LLM-friendly tables)
// ---------------------------------------------------------------------------

function escapeMd(s) {
  if (s == null) return '';
  return String(s).replace(/\|/g, '\\|').replace(/\n/g, ' ');
}

function dumpSheetMarkdown(ws, wb, opts = {}) {
  const { startRow, startCol, endRow, endCol } = selectionBounds(ws, opts);
  const out = [];
  out.push(`## ${ws.name}`);

  // Frontmatter context
  const meta = [];
  meta.push(`Range: ${colLetter(startCol)}${startRow}:${colLetter(endCol)}${endRow}`);
  meta.push(`Total: ${ws.rowCount} rows × ${ws.columnCount} cols`);
  const frozen = (ws.views || []).find(v => v.state === 'frozen');
  if (frozen) meta.push(`Frozen: row ${frozen.ySplit ?? 0}, col ${frozen.xSplit ?? 0}`);
  const merges = Object.keys(ws._merges || {});
  if (merges.length) meta.push(`Merged: ${merges.slice(0, 6).join(', ')}${merges.length > 6 ? ', ...' : ''}`);
  const namedRanges = getNamedRanges(wb, ws.name);
  if (namedRanges.length) meta.push(`Named ranges: ${namedRanges.map(n => n.name).join(', ')}`);
  out.push(`*${meta.join(' · ')}*`);
  out.push('');

  // Header detection: use first row in selection if it looks like text headers,
  // otherwise fall back to column letters.
  const firstRow = ws.getRow(startRow);
  const headers = [];
  let textHeaders = 0, totalHeaders = 0;
  for (let c = startCol; c <= endCol; c++) {
    const v = plainValue(firstRow.getCell(c).value);
    if (v != null && v !== '') {
      totalHeaders++;
      if (isNaN(parseFloat(v))) textHeaders++;
    }
    headers.push(v);
  }
  const useFirstRowAsHeader = totalHeaders > 0 && (textHeaders / totalHeaders) > 0.5;
  let dataStart = startRow;
  let cols;
  if (useFirstRowAsHeader) {
    cols = headers.map((h, i) => h != null && h !== '' ? String(h) : colLetter(startCol + i));
    dataStart = startRow + 1;
  } else {
    cols = [];
    for (let c = startCol; c <= endCol; c++) cols.push(colLetter(c));
  }

  // Render table
  out.push('| ' + cols.map(escapeMd).join(' | ') + ' |');
  out.push('|' + cols.map(() => '---').join('|') + '|');

  for (let r = dataStart; r <= endRow; r++) {
    const row = ws.getRow(r);
    const cells = [];
    let nonEmpty = 0;
    for (let c = startCol; c <= endCol; c++) {
      const v = plainValue(row.getCell(c).value);
      if (v != null && v !== '') nonEmpty++;
      // Wrap formulas in backticks so the model knows it's a formula
      const raw = row.getCell(c).value;
      if (raw && typeof raw === 'object' && (raw.formula || raw.sharedFormula)) {
        const display = v != null ? `${v} \`=${raw.formula || raw.sharedFormula}\`` : `\`=${raw.formula || raw.sharedFormula}\``;
        cells.push(escapeMd(display));
      } else {
        cells.push(escapeMd(v ?? ''));
      }
    }
    if (nonEmpty > 0) out.push('| ' + cells.join(' | ') + ' |');
  }

  if (opts.maxRows && ws.rowCount > endRow) {
    out.push('');
    out.push(`*... ${ws.rowCount - endRow} more rows truncated*`);
  }
  return out.join('\n');
}

// ---------------------------------------------------------------------------
// JSON dump
// ---------------------------------------------------------------------------

function jsonValue(v) {
  if (v == null) return null;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(r => r.text).join('');
    if (v.hyperlink) return { text: v.text || v.hyperlink, hyperlink: v.hyperlink };
    if (v.formula || v.sharedFormula) {
      const out = {};
      if (v.formula) out.formula = v.formula;
      if (v.sharedFormula) out.sharedFormulaRef = v.sharedFormula;
      const result = v.result;
      if (result == null) out.result = null;
      else if (typeof result === 'object') {
        if (result.error) out.result = `#${result.error}`;
        else if (result.richText) out.result = result.richText.map(r => r.text).join('');
        else out.result = result;
      } else out.result = result;
      return out;
    }
    if (v.error) return `#${v.error}`;
  }
  return v;
}

function dumpSheetJSON(ws, wb, opts = {}) {
  const { startRow, startCol, endRow, endCol } = selectionBounds(ws, opts);

  const out = {
    name: ws.name,
    state: ws.state || 'visible',
    rowCount: ws.rowCount,
    columnCount: ws.columnCount,
    selection: { startRef: `${colLetter(startCol)}${startRow}`, endRef: `${colLetter(endCol)}${endRow}` },
    frozen: null,
    columns: [],
    hiddenColumns: [],
    merges: Object.keys(ws._merges || {}),
    autoFilter: null,
    printArea: null,
    namedRanges: getNamedRanges(wb, ws.name),
    tables: [],
    images: [],
    cells: [],
  };

  const frozen = (ws.views || []).find(v => v.state === 'frozen');
  if (frozen) out.frozen = { rowSplit: frozen.ySplit ?? 0, colSplit: frozen.xSplit ?? 0 };

  for (let c = startCol; c <= endCol; c++) {
    const col = ws.getColumn(c);
    const letter = colLetter(c);
    if (col.hidden) out.hiddenColumns.push(letter);
    out.columns.push({ letter, width: col.width || null, hidden: !!col.hidden });
  }

  if (ws.autoFilter) out.autoFilter = typeof ws.autoFilter === 'string' ? ws.autoFilter : (ws.autoFilter.ref || null);
  try { if (ws.pageSetup?.printArea) out.printArea = ws.pageSetup.printArea; } catch (_) {}

  try {
    const tableMap = ws.tables;
    if (tableMap && typeof tableMap === 'object') {
      const tables = typeof tableMap.forEach === 'function'
        ? (() => { const a = []; tableMap.forEach(t => a.push(t)); return a; })()
        : Object.values(tableMap);
      for (const t of tables) {
        const model = t.table || t.model || t;
        out.tables.push({
          name: model.name || model.displayName || null,
          ref: model.ref || model.tableRef || null,
          columns: (model.columns || []).map(c => c.name).filter(Boolean),
        });
      }
    }
  } catch (_) {}

  try {
    const images = typeof ws.getImages === 'function' ? ws.getImages() : [];
    for (const img of images) {
      if (img.range) {
        const tl = img.range.tl, br = img.range.br;
        out.images.push({
          tl: tl ? `${colLetter(Math.floor(tl.col)+1)}${Math.floor(tl.row)+1}` : null,
          br: br ? `${colLetter(Math.floor(br.col)+1)}${Math.floor(br.row)+1}` : null,
        });
      }
    }
  } catch (_) {}

  for (let r = startRow; r <= endRow; r++) {
    const row = ws.getRow(r);
    for (let c = startCol; c <= endCol; c++) {
      const cell = row.getCell(c);
      const raw = cell.value;
      if (raw == null || raw === '') continue;
      const entry = { ref: `${colLetter(c)}${r}`, row: r, col: c, value: jsonValue(raw) };
      if (cell.numFmt && cell.numFmt !== 'General') entry.numFmt = cell.numFmt;
      if (cell.font?.bold) entry.bold = true;
      if (cell.font?.italic) entry.italic = true;
      if (cell.font?.color?.argb) entry.color = cell.font.color.argb;
      if (cell.fill?.type === 'pattern' && cell.fill.fgColor?.argb) entry.fill = cell.fill.fgColor.argb;
      if (cell.alignment?.horizontal && cell.alignment.horizontal !== 'general') entry.align = cell.alignment.horizontal;
      if (cell.hyperlink) entry.hyperlink = cell.hyperlink;
      if (cell.note) entry.note = describeNote(cell.note);
      if (cell.dataValidation) entry.dataValidation = cell.dataValidation;
      if (row.hidden) entry.rowHidden = true;
      out.cells.push(entry);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Schema inference (#5)
// ---------------------------------------------------------------------------

function inferType(values) {
  let n = 0, ints = 0, floats = 0, dates = 0, bools = 0, strs = 0, nulls = 0;
  for (const v of values) {
    if (v == null || v === '') { nulls++; continue; }
    n++;
    if (v instanceof Date) { dates++; continue; }
    if (typeof v === 'boolean') { bools++; continue; }
    if (typeof v === 'number') {
      if (Number.isInteger(v)) ints++; else floats++;
      continue;
    }
    if (typeof v === 'string') {
      const s = v.trim();
      if (/^-?\d+$/.test(s)) { ints++; continue; }
      if (/^-?\d+\.\d+$/.test(s)) { floats++; continue; }
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) { dates++; continue; }
      if (/^(true|false)$/i.test(s)) { bools++; continue; }
      strs++; continue;
    }
    strs++;
  }
  if (n === 0) return { type: 'unknown', nullable: nulls > 0 };
  // Pick majority type
  const counts = { int: ints, float: floats, date: dates, bool: bools, str: strs };
  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  const top = sorted[0][0];
  const types = { int: 'INTEGER', float: 'NUMERIC', date: 'DATE', bool: 'BOOLEAN', str: 'TEXT' };
  return { type: types[top], nullable: nulls > 0, nonNull: n, total: n + nulls };
}

function inferSchema(ws, wb, opts = {}) {
  const { startRow, startCol, endRow, endCol } = selectionBounds(ws, opts);
  const headerRow = ws.getRow(startRow);
  const cols = [];
  for (let c = startCol; c <= endCol; c++) {
    const headerVal = plainValue(headerRow.getCell(c).value);
    const name = headerVal != null && headerVal !== '' ? String(headerVal) : colLetter(c);
    const sampleVals = [];
    for (let r = startRow + 1; r <= endRow && sampleVals.length < 200; r++) {
      const raw = ws.getRow(r).getCell(c).value;
      sampleVals.push(plainValue(raw));
    }
    const typeInfo = inferType(sampleVals);
    cols.push({
      name,
      column: colLetter(c),
      ...typeInfo,
      sample: sampleVals.filter(v => v != null && v !== '').slice(0, 3),
    });
  }
  return { sheet: ws.name, columns: cols };
}

// ---------------------------------------------------------------------------
// SQL export (#10)
// ---------------------------------------------------------------------------

function sqlIdent(s) {
  return '"' + String(s).replace(/"/g, '""') + '"';
}

function sqlVal(v, type) {
  if (v == null || v === '') return 'NULL';
  if (type === 'INTEGER' || type === 'NUMERIC') {
    const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
    return Number.isFinite(n) ? String(n) : 'NULL';
  }
  if (type === 'BOOLEAN') {
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return /^true$/i.test(String(v)) ? 'TRUE' : 'FALSE';
  }
  if (type === 'DATE') {
    if (v instanceof Date) return `'${v.toISOString().slice(0,10)}'`;
    return `'${String(v).slice(0,10)}'`;
  }
  return "'" + String(v).replace(/'/g, "''") + "'";
}

function dumpSheetSQL(ws, wb, opts = {}) {
  const schema = inferSchema(ws, wb, opts);
  const tableName = ws.name.replace(/[^a-zA-Z0-9_]/g, '_').replace(/^(\d)/, '_$1');
  const out = [];
  out.push(`-- Sheet: ${ws.name}`);
  out.push(`CREATE TABLE ${sqlIdent(tableName)} (`);
  const colDefs = schema.columns.map(c => `  ${sqlIdent(c.name)} ${c.type}${c.nullable ? '' : ' NOT NULL'}`);
  out.push(colDefs.join(',\n'));
  out.push(');');
  out.push('');

  const { startRow, startCol, endRow, endCol } = selectionBounds(ws, opts);
  const colNames = schema.columns.map(c => sqlIdent(c.name)).join(', ');
  for (let r = startRow + 1; r <= endRow; r++) {
    const row = ws.getRow(r);
    const values = [];
    let hasAny = false;
    for (let i = 0; i < schema.columns.length; i++) {
      const c = startCol + i;
      const v = plainValue(row.getCell(c).value);
      if (v != null && v !== '') hasAny = true;
      values.push(sqlVal(v, schema.columns[i].type));
    }
    if (hasAny) {
      out.push(`INSERT INTO ${sqlIdent(tableName)} (${colNames}) VALUES (${values.join(', ')});`);
    }
  }
  return out.join('\n');
}

// ---------------------------------------------------------------------------
// Formula evaluation (#4) — pragmatic: promote cached results, optionally
// recompute simple literal/arithmetic formulas via formulajs.
// ---------------------------------------------------------------------------

function evaluateWorkbook(wb) {
  // Most .xlsx files saved by Excel/LibreOffice/etc. carry cached formula
  // results in cell.value.result. ExcelJS exposes those, so promotion is
  // mostly a no-op — formatValue already uses .result. The work this function
  // does is compute results for formulas that do NOT have a cached value
  // (typically machine-generated xlsx files). For these we attempt a simple
  // arithmetic eval using formulajs.
  const formulaJs = lazyFormulaJs();
  let computed = 0, missing = 0;
  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (!v || typeof v !== 'object') return;
        if (!v.formula && !v.sharedFormula) return;
        if (v.result != null) return; // already cached
        const f = v.formula;
        if (!f) { missing++; return; }
        // Attempt: =SUM(literal,literal,...) or =A1+B1 (very narrow set)
        const m = /^([A-Z]+)\(([^()]+)\)$/i.exec(f);
        if (m) {
          const fn = m[1].toUpperCase();
          const args = m[2].split(',').map(s => parseFloat(s));
          if (typeof formulaJs[fn] === 'function' && args.every(Number.isFinite)) {
            try {
              const r = formulaJs[fn](...args);
              v.result = r;
              computed++;
              return;
            } catch (_) {}
          }
        }
        missing++;
      });
    });
  }
  return { computed, missing };
}

// ---------------------------------------------------------------------------
// Workbook diff (#7)
// ---------------------------------------------------------------------------

function diffWorkbooks(wbA, wbB, opts = {}) {
  const out = [];
  const sheetsA = new Map(wbA.worksheets.map(s => [s.name, s]));
  const sheetsB = new Map(wbB.worksheets.map(s => [s.name, s]));
  const allNames = new Set([...sheetsA.keys(), ...sheetsB.keys()]);

  for (const name of allNames) {
    const a = sheetsA.get(name);
    const b = sheetsB.get(name);
    if (!a) { out.push(`+ Sheet added: ${name}`); continue; }
    if (!b) { out.push(`- Sheet removed: ${name}`); continue; }
    out.push(`~ Sheet: ${name}`);
    const rows = Math.max(a.rowCount, b.rowCount);
    const cols = Math.max(a.columnCount, b.columnCount);
    let changes = 0;
    for (let r = 1; r <= rows; r++) {
      for (let c = 1; c <= cols; c++) {
        const va = plainValue(a.getRow(r).getCell(c).value);
        const vb = plainValue(b.getRow(r).getCell(c).value);
        if (va === vb) continue;
        const ref = `${colLetter(c)}${r}`;
        if (va == null && vb != null)      out.push(`  + ${ref}: ${escapeMd(vb)}`);
        else if (vb == null && va != null) out.push(`  - ${ref}: ${escapeMd(va)}`);
        else                                out.push(`  ~ ${ref}: ${escapeMd(va)} → ${escapeMd(vb)}`);
        changes++;
        if (opts.maxRows && changes >= opts.maxRows) {
          out.push(`  ... (more changes; raise --max-rows to see all)`);
          r = rows + 1; break;
        }
      }
    }
    if (changes === 0) out.push('  (no cell changes)');
  }
  return out.join('\n');
}

// ---------------------------------------------------------------------------
// Token budget (#2)
// ---------------------------------------------------------------------------

function applyTokenBudget(text, maxTokens) {
  const tk = lazyTokenizer();
  const totalTokens = tk.encode(text).length;
  if (totalTokens <= maxTokens) return text;
  // Truncate by lines (preserve table structure) until under budget.
  const lines = text.split('\n');
  let lo = 0, hi = lines.length;
  while (lo < hi) {
    const mid = Math.floor((lo + hi + 1) / 2);
    const candidate = lines.slice(0, mid).join('\n');
    const ct = tk.encode(candidate).length;
    if (ct <= maxTokens - 60 /* leave room for tail summary */) lo = mid;
    else hi = mid - 1;
  }
  const kept = lines.slice(0, lo).join('\n');
  const droppedLines = lines.length - lo;
  return kept + `\n\n... [truncated to ~${maxTokens} tokens; ${droppedLines} of ${lines.length} lines / ${totalTokens} of ${totalTokens} input tokens dropped]`;
}

// ---------------------------------------------------------------------------
// Multi-format input (#3)
// ---------------------------------------------------------------------------

async function loadAnyWorkbook(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.xlsx') {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    return wb;
  }
  if (ext === '.csv' || ext === '.tsv') {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(path.basename(filePath, ext));
    const text = fs.readFileSync(filePath, 'utf8');
    const papa = lazyPapa();
    const delimiter = ext === '.tsv' ? '\t' : ',';
    const parsed = papa.parse(text, { delimiter, skipEmptyLines: true });
    for (const row of parsed.data) ws.addRow(row);
    return wb;
  }
  if (ext === '.xls' || ext === '.xlsb' || ext === '.ods') {
    return loadViaSheetJS(filePath);
  }
  throw new Error(`Unsupported extension: ${ext}. Supported: .xlsx .xls .xlsb .ods .csv .tsv`);
}

// Read a non-xlsx spreadsheet via SheetJS, materialize into an ExcelJS
// Workbook so the rest of the code (dump/markdown/json/sql/schema) works
// unchanged. Loses some formatting; preserves values + formulas.
function loadViaSheetJS(filePath) {
  const XLSX = lazyXlsx();
  const sheetJsWb = XLSX.readFile(filePath, { cellFormula: true, cellDates: true });
  const wb = new ExcelJS.Workbook();
  for (const name of sheetJsWb.SheetNames) {
    const sjsSheet = sheetJsWb.Sheets[name];
    const ws = wb.addWorksheet(name);
    if (!sjsSheet['!ref']) continue;
    const range = XLSX.utils.decode_range(sjsSheet['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sjsSheet[addr];
        if (!cell) continue;
        const target = ws.getRow(r + 1).getCell(c + 1);
        if (cell.f) {
          target.value = { formula: cell.f, result: cell.v };
        } else if (cell.t === 'd') {
          target.value = cell.v instanceof Date ? cell.v : new Date(cell.v);
        } else {
          target.value = cell.v;
        }
      }
    }
  }
  return wb;
}

// ---------------------------------------------------------------------------
// Streaming (#9) — for files too large to fit in memory.
// Uses ExcelJS WorkbookReader; emits a simplified per-row text dump to stdout.
// ---------------------------------------------------------------------------

async function streamDump(filePath, opts) {
  const wb = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    worksheets: 'emit',
    styles: 'cache',
  });
  const sheetFilter = opts.positional[1] || null;
  let sheetIdx = 0;
  for await (const ws of wb) {
    sheetIdx++;
    const name = ws.name || `Sheet${sheetIdx}`;
    if (sheetFilter && name !== sheetFilter) continue;
    process.stdout.write(`=== Sheet: ${name} (streaming) ===\n`);
    let rowCount = 0;
    for await (const row of ws) {
      rowCount++;
      if (opts.maxRows && rowCount > opts.maxRows) {
        process.stdout.write(`... more rows truncated at --max-rows ${opts.maxRows}\n`);
        break;
      }
      const cells = [];
      row.eachCell({ includeEmpty: false }, (cell, col) => {
        if (opts.maxCols && col > opts.maxCols) return;
        const ref = `${colLetter(col)}${row.number}`;
        // Streaming cells sometimes carry raw model objects; cell.text is the
        // already-rendered string and is more reliable here than cell.value.
        const display = (cell.text != null && cell.text !== '')
          ? `"${cell.text}"`
          : formatValue(cell.value);
        cells.push(`  ${ref}: ${display}`);
      });
      if (cells.length) {
        process.stdout.write(`--- Row ${row.number} ---\n` + cells.join('\n') + '\n');
      }
    }
    process.stdout.write('\n');
  }
}

// ---------------------------------------------------------------------------
// List sheets
// ---------------------------------------------------------------------------

function listSheets(wb) {
  const lines = [];
  for (const ws of wb.worksheets) {
    const vis = ws.state === 'hidden' ? ' [hidden]'
              : ws.state === 'veryHidden' ? ' [very hidden]' : '';
    lines.push(`${ws.name}  ${ws.rowCount} rows × ${ws.columnCount} cols${vis}`);
  }
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const opts = parseArgs(process.argv.slice(2));

  if (opts.help) { printHelp(); process.exit(0); }
  if (opts.positional.length < 1) { printHelp(); process.exit(1); }

  const filePath = path.resolve(opts.positional[0]);
  const sheetFilter = opts.positional[1] || null;

  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    process.exit(1);
  }
  const stat = fs.statSync(filePath);
  if (stat.size === 0) {
    console.error(`File is empty (0 bytes), not a valid spreadsheet: ${filePath}`);
    process.exit(1);
  }
  // Min 22 bytes (zip EOCD) only meaningful for binary formats; CSV/TSV can be smaller.
  const ext = path.extname(filePath).toLowerCase();
  const isBinary = ext === '.xlsx' || ext === '.xls' || ext === '.xlsb' || ext === '.ods';
  if (isBinary && stat.size < 22) {
    console.error(`File is too small (${stat.size} bytes) to be a valid spreadsheet: ${filePath}`);
    process.exit(1);
  }

  // Streaming mode: bypass full-workbook load.
  if (opts.stream) {
    if (ext !== '.xlsx') {
      console.error(`--stream only supports .xlsx (got ${ext})`);
      process.exit(1);
    }
    await streamDump(filePath, opts);
    return;
  }

  let wb;
  try {
    wb = await loadAnyWorkbook(filePath);
  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    console.error(`Failed to read ${filePath}: ${msg}`);
    if (/End of data reached|Corrupted zip|invalid signature/i.test(msg)) {
      console.error('Hint: file may be truncated or not a real spreadsheet. Try opening it in Excel to confirm.');
    } else if (/Cannot read propert/i.test(msg)) {
      console.error('Hint: file parsed as a zip but a workbook part is malformed. Try --list-sheets for a lighter probe.');
    } else if (/Unsupported extension/.test(msg)) {
      console.error('Hint: rename or convert to a supported extension.');
    }
    process.exit(1);
  }

  // Diff mode
  if (opts.diff) {
    const otherPath = path.resolve(opts.diff);
    if (!fs.existsSync(otherPath)) {
      console.error(`Diff target not found: ${otherPath}`);
      process.exit(1);
    }
    const wbB = await loadAnyWorkbook(otherPath);
    const out = diffWorkbooks(wb, wbB, opts);
    if (opts.maxTokens) process.stdout.write(applyTokenBudget(out, opts.maxTokens) + '\n');
    else process.stdout.write(out + '\n');
    return;
  }

  // --list-sheets
  if (opts.listSheets) {
    console.log(listSheets(wb));
    return;
  }

  // --evaluate: promote cached results / compute simple formulas.
  if (opts.evaluate) {
    const r = evaluateWorkbook(wb);
    if (process.env.XLSX_FOR_AI_DEBUG) console.error(`evaluate: computed=${r.computed} missing=${r.missing}`);
  }

  // Resolve --named-range to a sheet+bounds; overrides sheetFilter if provided.
  let sheets;
  let perSheetOpts = { ...opts };
  if (opts.namedRange) {
    const resolved = resolveNamedRange(wb, opts.namedRange);
    if (!resolved) {
      console.error(`Named range "${opts.namedRange}" not found.`);
      process.exit(1);
    }
    const ws = wb.getWorksheet(resolved.sheet);
    if (!ws) {
      console.error(`Named range references sheet "${resolved.sheet}" which is missing.`);
      process.exit(1);
    }
    sheets = [ws];
    perSheetOpts.namedRangeBounds = resolved.range;
  } else {
    sheets = sheetFilter
      ? [wb.getWorksheet(sheetFilter)].filter(Boolean)
      : wb.worksheets;
  }

  if (sheets.length === 0) {
    if (sheetFilter) {
      console.error(`Sheet "${sheetFilter}" not found. Available: ${wb.worksheets.map(s => s.name).join(', ')}`);
    } else {
      console.error('No sheets in workbook.');
      console.error('Hint: this can happen when a non-Excel tool wrote the file with backslashes in zip entry paths (e.g. xl\\worksheets\\sheet1.xml). ExcelJS only recognizes forward-slash entries.');
    }
    process.exit(1);
  }

  const baseName = path.basename(filePath, path.extname(filePath));

  // Pick output formatter.
  const renderText  = (ws) => dumpSheet(ws, wb, perSheetOpts);
  const renderMd    = (ws) => dumpSheetMarkdown(ws, wb, perSheetOpts);
  const renderJSON  = (ws) => dumpSheetJSON(ws, wb, perSheetOpts);
  const renderSQL   = (ws) => dumpSheetSQL(ws, wb, perSheetOpts);
  const renderSchema = (ws) => inferSchema(ws, wb, perSheetOpts);

  // Schema mode (always JSON-shaped, may be array)
  if (opts.schema) {
    const payload = sheets.map(renderSchema);
    const json = JSON.stringify(sheets.length === 1 ? payload[0] : payload, null, 2);
    const final = opts.maxTokens ? applyTokenBudget(json, opts.maxTokens) : json;
    if (opts.stdout) { process.stdout.write(final + '\n'); return; }
    const outDir = path.join(process.cwd(), '.xlsx-read');
    fs.mkdirSync(outDir, { recursive: true });
    const outFile = path.join(outDir, `${baseName}--schema.json`);
    fs.writeFileSync(outFile, final, 'utf8');
    console.log(outFile);
    return;
  }

  // SQL mode
  if (opts.sql) {
    const text = sheets.map(renderSQL).join('\n\n');
    const final = opts.maxTokens ? applyTokenBudget(text, opts.maxTokens) : text;
    if (opts.stdout) { process.stdout.write(final + '\n'); return; }
    const outDir = path.join(process.cwd(), '.xlsx-read');
    fs.mkdirSync(outDir, { recursive: true });
    const outFile = path.join(outDir, `${baseName}.sql`);
    fs.writeFileSync(outFile, final, 'utf8');
    console.log(outFile);
    return;
  }

  // JSON mode
  if (opts.json) {
    const payload = sheets.map(renderJSON);
    const json = JSON.stringify(sheets.length === 1 ? payload[0] : payload, null, 2);
    const final = opts.maxTokens ? applyTokenBudget(json, opts.maxTokens) : json;
    if (opts.stdout) { process.stdout.write(final + '\n'); return; }
    const outDir = path.join(process.cwd(), '.xlsx-read');
    fs.mkdirSync(outDir, { recursive: true });
    const outFile = path.join(outDir, `${baseName}.json`);
    fs.writeFileSync(outFile, final, 'utf8');
    console.log(outFile);
    return;
  }

  // Markdown mode
  if (opts.md) {
    const text = sheets.map(renderMd).join('\n\n');
    const final = opts.maxTokens ? applyTokenBudget(text, opts.maxTokens) : text;
    if (opts.stdout) { process.stdout.write(final + '\n'); return; }
    const outDir = path.join(process.cwd(), '.xlsx-read');
    fs.mkdirSync(outDir, { recursive: true });
    const outFile = path.join(outDir, `${baseName}.md`);
    fs.writeFileSync(outFile, final, 'utf8');
    console.log(outFile);
    return;
  }

  // Default: text dump
  if (opts.stdout) {
    let combined = '';
    for (const ws of sheets) combined += renderText(ws) + '\n\n';
    const final = opts.maxTokens ? applyTokenBudget(combined, opts.maxTokens) : combined;
    process.stdout.write(final);
    return;
  }
  const outDir = path.join(process.cwd(), '.xlsx-read');
  fs.mkdirSync(outDir, { recursive: true });
  for (const ws of sheets) {
    const content = renderText(ws);
    const final = opts.maxTokens ? applyTokenBudget(content, opts.maxTokens) : content;
    const safeName = ws.name.replace(/[^a-zA-Z0-9_-]/g, '_');
    const outFile = path.join(outDir, `${baseName}--${safeName}.txt`);
    fs.writeFileSync(outFile, final, 'utf8');
    console.log(outFile);
  }
}

main().catch((err) => {
  const msg = err && err.message ? err.message : String(err);
  console.error(msg);
  if (/Invalid string length/i.test(msg)) {
    console.error('Hint: this sheet renders to a text dump larger than V8\'s 512MB string limit. Try --max-rows N, --max-cols N, --max-tokens N, --range A1:..., or --stream.');
  }
  process.exit(1);
});
