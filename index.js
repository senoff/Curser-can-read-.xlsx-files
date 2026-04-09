#!/usr/bin/env node

const path = require('path');
const fs   = require('fs');
const ExcelJS = require('exceljs');

// ---------------------------------------------------------------------------
// Argument parsing
// ---------------------------------------------------------------------------

function parseArgs(argv) {
  const opts = { positional: [], listSheets: false, stdout: false, maxRows: null, help: false };
  let i = 0;
  while (i < argv.length) {
    const arg = argv[i];
    if (arg === '--list-sheets')       opts.listSheets = true;
    else if (arg === '--stdout')       opts.stdout = true;
    else if (arg === '--max-rows')   { opts.maxRows = parseInt(argv[++i], 10); }
    else if (arg === '-h' || arg === '--help') opts.help = true;
    else                               opts.positional.push(arg);
    i++;
  }
  return opts;
}

function printHelp() {
  console.log(`Usage: npx cursor-reads-xlsx <file.xlsx> [sheetName] [options]

Converts .xlsx to rich text that AI coding agents can read.

Options:
  --list-sheets   List sheet names, dimensions, and visibility then exit
  --stdout        Print output to stdout instead of writing files
  --max-rows N    Limit output to the first N rows per sheet
  -h, --help      Show this help message

Examples:
  npx cursor-reads-xlsx data.xlsx
  npx cursor-reads-xlsx data.xlsx "Sheet1"
  npx cursor-reads-xlsx data.xlsx --list-sheets
  npx cursor-reads-xlsx data.xlsx --stdout --max-rows 100`);
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

function describeFill(fill) {
  if (!fill || (fill.type === 'pattern' && fill.pattern === 'none')) return null;
  if (fill.type === 'pattern' && fill.fgColor?.argb) return `fill:${fill.fgColor.argb}`;
  return null;
}

function describeFont(font) {
  const parts = [];
  if (font?.bold)   parts.push('bold');
  if (font?.italic) parts.push('italic');
  if (font?.color?.argb) parts.push(`color:${font.color.argb}`);
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
  if (typeof v === 'object' && v.formula) {
    return String(v.result ?? '');
  }
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

function dumpSheet(ws, wb, maxRows) {
  const lines = [];

  lines.push(`=== Sheet: ${ws.name} ===`);

  // Frozen panes
  const views = ws.views || [];
  const frozen = views.find(v => v.state === 'frozen');
  if (frozen) {
    lines.push(`Frozen: row ${frozen.ySplit ?? 0}, col ${frozen.xSplit ?? 0}`);
  }

  // Column widths + hidden columns
  const colWidths = [];
  const hiddenCols = [];
  for (let c = 1; c <= ws.columnCount; c++) {
    const col = ws.getColumn(c);
    const letter = colLetter(c);
    if (col.hidden) hiddenCols.push(letter);
    if (col.width) colWidths.push(`${letter}(${Math.round(col.width)})`);
  }
  if (colWidths.length) lines.push(`Columns: ${colWidths.join(' ')}`);
  if (hiddenCols.length) lines.push(`Hidden columns: ${hiddenCols.join(', ')}`);

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

    for (let c = 1; c <= ws.columnCount; c++) {
      const cell = row.getCell(c);
      const raw = cell.value;
      if (raw == null || raw === '') continue;

      const ref = `${colLetter(c)}${r}`;
      const tags = [];

      // Formula
      if (cell.type === ExcelJS.ValueType.Formula) {
        const formula = typeof raw === 'object' ? raw.formula : null;
        if (formula) tags.push(`formula: =${formula}`);
      }

      // Number format
      if (cell.numFmt && cell.numFmt !== 'General') {
        tags.push(`numFmt: ${cell.numFmt}`);
      }

      // Font
      const fontTags = describeFont(cell.font);
      if (fontTags.length) tags.push(...fontTags);

      // Fill
      const fillDesc = describeFill(cell.fill);
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

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);

  // --list-sheets: print summary and exit
  if (opts.listSheets) {
    console.log(listSheets(wb));
    process.exit(0);
  }

  const sheets = sheetFilter
    ? [wb.getWorksheet(sheetFilter)].filter(Boolean)
    : wb.worksheets;

  if (sheets.length === 0) {
    console.error(sheetFilter
      ? `Sheet "${sheetFilter}" not found. Available: ${wb.worksheets.map(s => s.name).join(', ')}`
      : 'No sheets in workbook');
    process.exit(1);
  }

  const baseName = path.basename(xlsxPath, path.extname(xlsxPath));

  // --stdout: print to console
  if (opts.stdout) {
    for (const ws of sheets) {
      console.log(dumpSheet(ws, wb, opts.maxRows));
      console.log('');
    }
    process.exit(0);
  }

  // Default: write to .xlsx-read/ files
  const outDir = path.join(process.cwd(), '.xlsx-read');
  fs.mkdirSync(outDir, { recursive: true });

  for (const ws of sheets) {
    const content = dumpSheet(ws, wb, opts.maxRows);
    const safeName = ws.name.replace(/[^a-zA-Z0-9_-]/g, '_');
    const outFile = path.join(outDir, `${baseName}--${safeName}.txt`);
    fs.writeFileSync(outFile, content, 'utf8');
    console.log(outFile);
  }
}

main().catch((err) => {
  console.error(err.message);
  process.exit(1);
});
