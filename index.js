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

// ---------------------------------------------------------------------------
// Argument parsing
// ---------------------------------------------------------------------------

function parseArgs(argv) {
  const opts = {
    positional: [],
    listSheets: false,
    stdout: false,
    json: false,
    compact: false,
    maxRows: null,
    maxCols: null,
    help: false,
  };
  let i = 0;
  while (i < argv.length) {
    const arg = argv[i];
    if (arg === '--list-sheets')       opts.listSheets = true;
    else if (arg === '--stdout')       opts.stdout = true;
    else if (arg === '--json')         opts.json = true;
    else if (arg === '--compact')      opts.compact = true;
    else if (arg === '--max-rows')   { opts.maxRows = parseInt(argv[++i], 10); }
    else if (arg === '--max-cols')   { opts.maxCols = parseInt(argv[++i], 10); }
    else if (arg === '-h' || arg === '--help') opts.help = true;
    else                               opts.positional.push(arg);
    i++;
  }
  return opts;
}

function printHelp() {
  console.log(`Usage: npx xlsx-for-ai <file.xlsx> [sheetName] [options]

Converts .xlsx to rich text (or JSON) that AI coding agents can read.

Options:
  --list-sheets   List sheet names, dimensions, and visibility then exit
  --stdout        Print output to stdout instead of writing files
  --json          Emit structured JSON instead of the human-readable text dump
                  (one object per cell with value/formula/format/style)
  --compact       Suppress noisy default tags (default text color, default font,
                  General number format, etc.) to reduce token usage
  --max-rows N    Limit output to the first N rows per sheet
  --max-cols N    Limit output to the first N columns per sheet
  -h, --help      Show this help message

Examples:
  npx xlsx-for-ai data.xlsx
  npx xlsx-for-ai data.xlsx "Sheet1"
  npx xlsx-for-ai data.xlsx --list-sheets
  npx xlsx-for-ai data.xlsx --stdout --max-rows 100
  npx xlsx-for-ai data.xlsx --stdout --compact
  npx xlsx-for-ai data.xlsx --json --stdout > out.json

Note: this package was previously published as 'cursor-reads-xlsx';
that command name still works as an alias.`);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function colLetter(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

// Colors that are visually indistinguishable from the default Excel text (near-black).
// In --compact mode we suppress these to reduce token noise.
const DEFAULT_TEXT_COLORS = new Set([
  'FF000000', // pure black
  'FF1F1F1F', // dark gray (Excel auto-text on some themes)
  'FF222120', // dark gray variant
  'FF333333',
]);

function isDefaultTextColor(argb) {
  return argb && DEFAULT_TEXT_COLORS.has(argb.toUpperCase());
}

function describeFill(fill, compact) {
  if (!fill || (fill.type === 'pattern' && fill.pattern === 'none')) return null;
  if (fill.type === 'pattern' && fill.fgColor?.argb) {
    // White / no-fill patterns are noise in compact mode
    if (compact && /^FF?FFFFFF$/i.test(fill.fgColor.argb)) return null;
    return `fill:${fill.fgColor.argb}`;
  }
  return null;
}

function describeFont(font, compact) {
  const parts = [];
  if (font?.bold)   parts.push('bold');
  if (font?.italic) parts.push('italic');
  if (font?.color?.argb) {
    if (!(compact && isDefaultTextColor(font.color.argb))) {
      parts.push(`color:${font.color.argb}`);
    }
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
  // Formula cells: 'formula' on master, 'sharedFormula' on follow-ups, 'result' is the computed value
  if (typeof v === 'object' && (v.formula || v.sharedFormula)) {
    const result = v.result;
    if (result == null) return '""';
    // Result may itself be a rich object (error, richText, etc.)
    if (typeof result === 'object') {
      if (result.error) return `"#${result.error}"`;
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

function describeNote(note) {
  if (!note) return null;
  if (typeof note === 'string') return note;
  if (note.texts) {
    return note.texts.map(t => (typeof t === 'string' ? t : t.text || '')).join('');
  }
  return null;
}

// ---------------------------------------------------------------------------
// Named ranges (workbook-level, filtered to a sheet if name provided)
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

// ---------------------------------------------------------------------------
// Sheet dump
// ---------------------------------------------------------------------------

function dumpSheet(ws, wb, opts = {}) {
  const { maxRows = null, maxCols = null, compact = false } = opts;
  const lines = [];

  lines.push(`=== Sheet: ${ws.name} ===`);

  // Frozen panes
  const views = ws.views || [];
  const frozen = views.find(v => v.state === 'frozen');
  if (frozen) {
    lines.push(`Frozen: row ${frozen.ySplit ?? 0}, col ${frozen.xSplit ?? 0}`);
  }

  // Column widths + hidden columns
  const totalCols = maxCols ? Math.min(ws.columnCount, maxCols) : ws.columnCount;
  const colWidths = [];
  const hiddenCols = [];
  for (let c = 1; c <= totalCols; c++) {
    const col = ws.getColumn(c);
    const letter = colLetter(c);
    if (col.hidden) hiddenCols.push(letter);
    if (col.width) colWidths.push(`${letter}(${Math.round(col.width)})`);
  }
  if (colWidths.length) lines.push(`Columns: ${colWidths.join(' ')}`);
  if (hiddenCols.length) lines.push(`Hidden columns: ${hiddenCols.join(', ')}`);
  if (maxCols && ws.columnCount > maxCols) {
    lines.push(`(${ws.columnCount - maxCols} more columns truncated at --max-cols ${maxCols})`);
  }

  // Merged cells
  const merges = Object.keys(ws._merges || {});
  if (merges.length) lines.push(`Merged: ${merges.join(', ')}`);

  // Auto-filter
  if (ws.autoFilter) {
    const af = typeof ws.autoFilter === 'string'
      ? ws.autoFilter
      : (ws.autoFilter.ref || JSON.stringify(ws.autoFilter));
    lines.push(`Auto-filter: ${af}`);
  }

  // Print area
  try {
    if (ws.pageSetup?.printArea) {
      lines.push(`Print area: ${ws.pageSetup.printArea}`);
    }
  } catch (_) {}

  // Named ranges relevant to this sheet
  const namedRanges = getNamedRanges(wb, ws.name);
  if (namedRanges.length) {
    lines.push(`Named ranges:`);
    for (const nr of namedRanges) {
      lines.push(`  ${nr.name}: ${nr.ranges.join(', ')}`);
    }
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

  // Images
  try {
    const images = typeof ws.getImages === 'function' ? ws.getImages() : [];
    for (const img of images) {
      if (img.range) {
        const tl = img.range.tl;
        const br = img.range.br;
        if (tl && br) {
          lines.push(`Image: ${colLetter(Math.floor(tl.col) + 1)}${Math.floor(tl.row) + 1} to ${colLetter(Math.floor(br.col) + 1)}${Math.floor(br.row) + 1}`);
        } else if (tl) {
          lines.push(`Image at: ${colLetter(Math.floor(tl.col) + 1)}${Math.floor(tl.row) + 1}`);
        }
      } else {
        lines.push(`Image: (position unknown)`);
      }
    }
  } catch (_) {}

  lines.push('');

  // Rows
  const rowLimit = maxRows ? Math.min(ws.rowCount, maxRows) : ws.rowCount;

  for (let r = 1; r <= rowLimit; r++) {
    const row = ws.getRow(r);
    const cells = [];
    const isHidden = row.hidden;

    for (let c = 1; c <= totalCols; c++) {
      const cell = row.getCell(c);
      const raw = cell.value;
      if (raw == null || raw === '') continue;

      const ref = `${colLetter(c)}${r}`;
      const tags = [];

      // Formula (handle both standalone and shared formulas)
      if (cell.type === ExcelJS.ValueType.Formula) {
        if (typeof raw === 'object') {
          if (raw.formula) {
            tags.push(`formula: =${raw.formula}`);
          } else if (raw.sharedFormula) {
            tags.push(`shared formula ref: ${raw.sharedFormula}`);
          }
        }
      }

      // Number format
      if (cell.numFmt && cell.numFmt !== 'General') {
        tags.push(`numFmt: ${cell.numFmt}`);
      }

      // Font
      const fontTags = describeFont(cell.font, compact);
      if (fontTags.length) tags.push(...fontTags);

      // Fill
      const fillDesc = describeFill(cell.fill, compact);
      if (fillDesc) tags.push(fillDesc);

      // Alignment
      if (cell.alignment?.horizontal && cell.alignment.horizontal !== 'general') {
        tags.push(`align:${cell.alignment.horizontal}`);
      }

      // Hyperlink
      if (cell.hyperlink) {
        tags.push(`link: ${cell.hyperlink}`);
      } else if (typeof raw === 'object' && raw.hyperlink) {
        tags.push(`link: ${raw.hyperlink}`);
      }

      // Comment / note
      const noteText = describeNote(cell.note);
      if (noteText) {
        tags.push(`note: ${noteText.replace(/\n/g, ' ').trim()}`);
      }

      // Data validation
      if (cell.dataValidation) {
        const dv = cell.dataValidation;
        if (dv.type === 'list' && dv.formulae?.length) {
          tags.push(`validation: list [${dv.formulae[0]}]`);
        } else if (dv.type) {
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

  if (maxRows && ws.rowCount > maxRows) {
    lines.push('');
    lines.push(`... ${ws.rowCount - maxRows} more rows (truncated at --max-rows ${maxRows})`);
  }

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// JSON dump (structured per-cell output)
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
      if (result == null) {
        out.result = null;
      } else if (typeof result === 'object') {
        if (result.error) out.result = `#${result.error}`;
        else if (result.richText) out.result = result.richText.map(r => r.text).join('');
        else out.result = result;
      } else {
        out.result = result;
      }
      return out;
    }
    if (v.error) return `#${v.error}`;
  }
  return v;
}

function dumpSheetJSON(ws, wb, opts = {}) {
  const { maxRows = null, maxCols = null } = opts;
  const totalCols = maxCols ? Math.min(ws.columnCount, maxCols) : ws.columnCount;
  const rowLimit = maxRows ? Math.min(ws.rowCount, maxRows) : ws.rowCount;

  const out = {
    name: ws.name,
    state: ws.state || 'visible',
    rowCount: ws.rowCount,
    columnCount: ws.columnCount,
    truncated: {
      rows: maxRows && ws.rowCount > maxRows ? ws.rowCount - maxRows : 0,
      cols: maxCols && ws.columnCount > maxCols ? ws.columnCount - maxCols : 0,
    },
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

  // Frozen panes
  const frozen = (ws.views || []).find(v => v.state === 'frozen');
  if (frozen) out.frozen = { rowSplit: frozen.ySplit ?? 0, colSplit: frozen.xSplit ?? 0 };

  // Columns
  for (let c = 1; c <= totalCols; c++) {
    const col = ws.getColumn(c);
    const letter = colLetter(c);
    if (col.hidden) out.hiddenColumns.push(letter);
    out.columns.push({ letter, width: col.width || null, hidden: !!col.hidden });
  }

  // Auto-filter
  if (ws.autoFilter) {
    out.autoFilter = typeof ws.autoFilter === 'string'
      ? ws.autoFilter
      : (ws.autoFilter.ref || null);
  }

  // Print area
  try { if (ws.pageSetup?.printArea) out.printArea = ws.pageSetup.printArea; } catch (_) {}

  // Tables
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

  // Images
  try {
    const images = typeof ws.getImages === 'function' ? ws.getImages() : [];
    for (const img of images) {
      if (img.range) {
        const tl = img.range.tl, br = img.range.br;
        out.images.push({
          tl: tl ? `${colLetter(Math.floor(tl.col) + 1)}${Math.floor(tl.row) + 1}` : null,
          br: br ? `${colLetter(Math.floor(br.col) + 1)}${Math.floor(br.row) + 1}` : null,
        });
      }
    }
  } catch (_) {}

  // Cells
  for (let r = 1; r <= rowLimit; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= totalCols; c++) {
      const cell = row.getCell(c);
      const raw = cell.value;
      if (raw == null || raw === '') continue;

      const entry = {
        ref: `${colLetter(c)}${r}`,
        row: r,
        col: c,
        value: jsonValue(raw),
      };
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
// List sheets mode
// ---------------------------------------------------------------------------

function listSheets(wb) {
  const lines = [];
  for (const ws of wb.worksheets) {
    const vis = ws.state === 'hidden' ? ' [hidden]'
              : ws.state === 'veryHidden' ? ' [very hidden]'
              : '';
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

  const xlsxPath = path.resolve(opts.positional[0]);
  const sheetFilter = opts.positional[1] || null;

  if (!fs.existsSync(xlsxPath)) {
    console.error(`File not found: ${xlsxPath}`);
    process.exit(1);
  }

  const stat = fs.statSync(xlsxPath);
  if (stat.size === 0) {
    console.error(`File is empty (0 bytes), not a valid xlsx: ${xlsxPath}`);
    process.exit(1);
  }
  // Minimum valid zip is a 22-byte end-of-central-directory record. Anything
  // smaller cannot be an xlsx; ExcelJS would crash with a misleading
  // "Corrupted zip" error from deep in its parser.
  if (stat.size < 22) {
    console.error(`File is too small (${stat.size} bytes) to be a valid xlsx: ${xlsxPath}`);
    process.exit(1);
  }

  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.readFile(xlsxPath);
  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    console.error(`Failed to read ${xlsxPath}: ${msg}`);
    if (/End of data reached|Corrupted zip|invalid signature/i.test(msg)) {
      console.error('Hint: file may be truncated or not a real xlsx. Try opening it in Excel to confirm.');
    } else if (/Cannot read propert/i.test(msg)) {
      console.error('Hint: file parsed as a zip but a workbook part is malformed. Try --list-sheets for a lighter probe.');
    }
    process.exit(1);
  }

  // --list-sheets: print summary and exit
  if (opts.listSheets) {
    console.log(listSheets(wb));
    process.exit(0);
  }

  const sheets = sheetFilter
    ? [wb.getWorksheet(sheetFilter)].filter(Boolean)
    : wb.worksheets;

  if (sheets.length === 0) {
    if (sheetFilter) {
      console.error(`Sheet "${sheetFilter}" not found. Available: ${wb.worksheets.map(s => s.name).join(', ')}`);
    } else {
      console.error('No sheets in workbook.');
      console.error('Hint: this can happen when a non-Excel tool wrote the file with backslashes in zip entry paths (e.g. xl\\worksheets\\sheet1.xml). ExcelJS only recognizes forward-slash entries.');
    }
    process.exit(1);
  }

  const baseName = path.basename(xlsxPath, path.extname(xlsxPath));

  const dumpOpts = { maxRows: opts.maxRows, maxCols: opts.maxCols, compact: opts.compact };

  // --json mode: structured per-cell output
  if (opts.json) {
    const payload = sheets.map(ws => dumpSheetJSON(ws, wb, dumpOpts));
    const json = JSON.stringify(sheets.length === 1 ? payload[0] : payload, null, 2);

    if (opts.stdout) {
      process.stdout.write(json + '\n');
      process.exit(0);
    }
    const outDir = path.join(process.cwd(), '.xlsx-read');
    fs.mkdirSync(outDir, { recursive: true });
    const outFile = path.join(outDir, `${baseName}.json`);
    fs.writeFileSync(outFile, json, 'utf8');
    console.log(outFile);
    process.exit(0);
  }

  // --stdout: print text dump to console
  if (opts.stdout) {
    for (const ws of sheets) {
      console.log(dumpSheet(ws, wb, dumpOpts));
      console.log('');
    }
    process.exit(0);
  }

  // Default: write text dump to .xlsx-read/ files
  const outDir = path.join(process.cwd(), '.xlsx-read');
  fs.mkdirSync(outDir, { recursive: true });

  for (const ws of sheets) {
    const content = dumpSheet(ws, wb, dumpOpts);
    const safeName = ws.name.replace(/[^a-zA-Z0-9_-]/g, '_');
    const outFile = path.join(outDir, `${baseName}--${safeName}.txt`);
    fs.writeFileSync(outFile, content, 'utf8');
    console.log(outFile);
  }
}

main().catch((err) => {
  const msg = err && err.message ? err.message : String(err);
  console.error(msg);
  if (/Invalid string length/i.test(msg)) {
    console.error('Hint: this sheet renders to a text dump larger than V8\'s 512MB string limit. Try --max-rows N or --max-cols N to bound the output, or --json which streams per cell.');
  }
  process.exit(1);
});
