// Bug-report generator for xlsx-for-ai.
//
// Produces a JSON blob describing the *structure* of an .xlsx workbook
// (sheet count + shape, used-features inventory, env) with ZERO user
// content (no cell values, no formulas, no shared strings, no
// named-range formulas, no comment text). Designed to be safe for a
// reporter to attach to a public GitHub issue.
//
// Implementation:
//   1. Read the .xlsx as a ZIP via JSZip (already a transitive dep
//      of exceljs). Walk the OOXML parts to detect features by
//      filename pattern + targeted ContentType / relationship lookups.
//   2. Use ExcelJS only for sheet shape (rowCount, columnCount),
//      merge counts, and named-range *names* (not their refs/formulas).
//
// We deliberately avoid emitting anything sourced from cell text,
// shared strings, or formula expressions. The bug-report consumer
// should be able to grep the output for any user content and find none.

const fs = require('fs');
const path = require('path');
const os = require('os');
const JSZip = require('jszip');
const ExcelJS = require('exceljs');

const PKG_VERSION = require('../package.json').version;

// OOXML feature detectors. Each entry maps a feature key to a predicate
// over the list of zip entry filenames. We choose names + content-type
// matches that are stable across Excel versions.
//
// References:
//   ECMA-376 part-1 (OOXML) section 18.x for sheet parts
//   MS-OE376 for vendor extension parts
const FEATURE_PATTERNS = [
  // Pivot tables: xl/pivotTables/pivotTable*.xml + xl/pivotCache/*
  { key: 'pivotTables',      test: (n) => /^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(n) },
  { key: 'pivotCaches',      test: (n) => /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i.test(n) },

  // Charts (drawing-based + chartsheets)
  { key: 'charts',           test: (n) => /^xl\/charts\/chart\d+\.xml$/i.test(n) },
  { key: 'chartsheets',      test: (n) => /^xl\/chartsheets\/sheet\d+\.xml$/i.test(n) },
  { key: 'drawings',         test: (n) => /^xl\/drawings\/drawing\d+\.xml$/i.test(n) },

  // Threaded comments (modern; Office 365). Plain comments are detected separately.
  { key: 'threadedComments', test: (n) => /^xl\/threadedComments\/threadedComment\d+\.xml$/i.test(n) },
  { key: 'comments',         test: (n) => /^xl\/comments\d+\.xml$/i.test(n) },
  { key: 'persons',          test: (n) => /^xl\/persons\/person\.xml$/i.test(n) },

  // Sensitivity labels (MIP). docMetadata folder + LabelInfo part.
  { key: 'sensitivityLabel', test: (n) => /^docMetadata\/LabelInfo\.xml$/i.test(n) },

  // Linked / rich data types (the "Stocks", "Geography" data types).
  { key: 'richValueData',    test: (n) => /^xl\/richData\/rdRichValues\.xml$/i.test(n) },
  { key: 'richValueRel',     test: (n) => /^xl\/richData\/richValueRel\.xml$/i.test(n) },

  // Power Query / Data Model
  { key: 'powerQuery',       test: (n) => /^xl\/queryTables\/queryTable\d+\.xml$/i.test(n)
                                       || /^customXml\/item\d+\.xml$/i.test(n) && false /* refined below */ },
  { key: 'dataModel',        test: (n) => /^xl\/model\/item\.data$/i.test(n) },
  { key: 'connections',      test: (n) => /^xl\/connections\.xml$/i.test(n) },

  // Slicers / Timelines (modern PivotTable controls)
  { key: 'slicers',          test: (n) => /^xl\/slicers\/slicer\d+\.xml$/i.test(n) },
  { key: 'slicerCaches',     test: (n) => /^xl\/slicerCaches\/slicerCache\d+\.xml$/i.test(n) },
  { key: 'timelines',        test: (n) => /^xl\/timelines\/timeline\d+\.xml$/i.test(n) },
  { key: 'timelineCaches',   test: (n) => /^xl\/timelineCaches\/timelineCache\d+\.xml$/i.test(n) },

  // Tables (Excel ListObjects)
  { key: 'tables',           test: (n) => /^xl\/tables\/table\d+\.xml$/i.test(n) },

  // External links / workbook references
  { key: 'externalLinks',    test: (n) => /^xl\/externalLinks\/externalLink\d+\.xml$/i.test(n) },

  // Macros / VBA
  { key: 'vbaProject',       test: (n) => /^xl\/vbaProject\.bin$/i.test(n) },

  // Custom XML parts (often used by enterprise add-ins / SharePoint)
  { key: 'customXml',        test: (n) => /^customXml\/item\d+\.xml$/i.test(n) },

  // Embedded objects (OLE)
  { key: 'embeddings',       test: (n) => /^xl\/embeddings\/.+/i.test(n) },

  // Theme + custom properties (low signal but cheap)
  { key: 'customProps',      test: (n) => /^docProps\/custom\.xml$/i.test(n) },
];

// Detect dynamic arrays + sparklines from sheet XML. These don't have
// dedicated parts — they're attributes on cell / extLst inside sheetN.xml.
// We do a coarse string scan (no value extraction) just to flag presence.
//
// Dynamic arrays: <f t="array" ...> with <ext> CT_ExtensionList for
//   x14ac:cm, or modern: presence of <ext> with namespace x17 + cm attr.
// Sparklines: <ext><x14:sparklineGroups> inside <extLst>.
async function detectInSheetFeatures(zip, sheetNames) {
  const flags = { dynamicArrays: false, sparklines: false, conditionalFormatting: false };
  for (const name of sheetNames) {
    const file = zip.file(name);
    if (!file) continue;
    const xml = await file.async('string');
    // Coarse but conservative — we look for tag names only, never values.
    if (!flags.dynamicArrays && /\bcm="\d+"/.test(xml))           flags.dynamicArrays = true;
    if (!flags.dynamicArrays && /<f[^>]*\bt="array"/.test(xml))   flags.dynamicArrays = true;
    if (!flags.sparklines    && /sparklineGroup/.test(xml))       flags.sparklines    = true;
    if (!flags.conditionalFormatting && /<conditionalFormatting/.test(xml))
      flags.conditionalFormatting = true;
  }
  return flags;
}

function inventoryFeatures(filenames) {
  const out = {};
  for (const { key, test } of FEATURE_PATTERNS) {
    const count = filenames.filter(test).length;
    if (count > 0) out[key] = count;
  }
  return out;
}

// Given the workbook.xml, extract the sheet relationship Ids and order
// without reading any user content. We just need names and rIds so we
// can pair them with worksheet parts to compute per-sheet stats.
function listSheetPartNames(zip) {
  // Resolve via workbook rels: xl/_rels/workbook.xml.rels.
  const out = [];
  const relsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (!relsFile) return out;
  // Sync — we already have the file in memory inside JSZip.
  // We use a lightweight regex; structural only, no values inside.
  // Each Relationship: <Relationship Id="rId1" Type="..." Target="worksheets/sheet1.xml"/>
  // We can't do sync read without loading; caller already loaded.
  return out;
}

async function generateBugReport(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== '.xlsx' && ext !== '.xlsm') {
    throw new Error(`--report-bug only supports .xlsx / .xlsm (got ${ext})`);
  }

  const stat = fs.statSync(filePath);
  const buf = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buf);
  const filenames = Object.keys(zip.files).filter((n) => !zip.files[n].dir);

  const features = inventoryFeatures(filenames);

  // Sheet parts list — derived from filename pattern, not content.
  const sheetParts = filenames.filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(n));

  // In-sheet feature flags (string scan, no extraction).
  const inSheet = await detectInSheetFeatures(zip, sheetParts);
  if (inSheet.dynamicArrays)         features.dynamicArrays         = true;
  if (inSheet.sparklines)            features.sparklines            = true;
  if (inSheet.conditionalFormatting) features.conditionalFormatting = true;

  // Use ExcelJS for sheet shape, merges, and *names* of named ranges.
  // We never read cell values or named-range formulas — only enumerate.
  let sheetCount = 0;
  let mergedTotal = 0;
  let namedRangesCount = 0;
  let definedNames = [];
  const perSheet = [];
  let exceljsError = null;

  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    sheetCount = wb.worksheets.length;
    for (const ws of wb.worksheets) {
      const merges = ws.model && ws.model.merges ? ws.model.merges.length : 0;
      mergedTotal += merges;
      perSheet.push({
        index: ws.id,
        rows: ws.rowCount || 0,
        cols: ws.columnCount || 0,
        merges,
        hidden: ws.state && ws.state !== 'visible' ? ws.state : null,
      });
    }
    // Defined names — names ONLY (deliberately drop ranges/formulas).
    const dnModel = wb.definedNames && wb.definedNames.model;
    if (Array.isArray(dnModel)) {
      namedRangesCount = dnModel.length;
      definedNames = dnModel
        .map((d) => (d && typeof d.name === 'string' ? d.name : null))
        .filter(Boolean);
    }
  } catch (err) {
    // ExcelJS may fail on edge-case files; report the error class but
    // don't include the message verbatim (could leak a path inside the
    // workbook). Sheet count falls back to part count.
    exceljsError = err && err.name ? err.name : 'Error';
    sheetCount = sheetParts.length;
  }

  const report = {
    schema: 'xlsx-for-ai/bug-report/v1',
    generatedAt: new Date().toISOString(),
    tool: {
      name: 'xlsx-for-ai',
      version: PKG_VERSION,
    },
    runtime: {
      node: process.version,
      platform: process.platform, // e.g. 'darwin', 'linux', 'win32'
      arch: process.arch,         // e.g. 'arm64', 'x64'
      osRelease: os.release(),
    },
    file: {
      // ONLY the basename + size — never the absolute path (could leak
      // user/dir names). The reporter knows what file they ran it on.
      basename: path.basename(filePath),
      ext,
      sizeBytes: stat.size,
    },
    workbook: {
      sheetCount,
      mergedRangeCountTotal: mergedTotal,
      namedRangesCount,
      // Names only — Excel defined-name *names* are user-chosen labels
      // ("Totals", "TaxRate"). We emit them because they're often the
      // hint a maintainer needs. If a reporter considers their names
      // sensitive, they should sanitize before attaching.
      definedNames,
      perSheet,
      featuresPresent: features,
    },
    notes: [
      'This report contains zero cell values, formulas, shared strings, named-range formulas, or comment text.',
      'Defined-name *labels* are included (e.g. "Totals") but their target ranges are not.',
      'Generated with --report-bug. Attach to a GitHub issue at https://github.com/senoff/xlsx-for-ai/issues',
    ],
  };
  if (exceljsError) {
    report.workbook.exceljsLoadError = exceljsError;
  }
  return report;
}

function writeBugReport(report, cwd) {
  const ts = report.generatedAt.replace(/[:.]/g, '-');
  const outPath = path.join(cwd, `xlsx-for-ai-bugreport-${ts}.json`);
  fs.writeFileSync(outPath, JSON.stringify(report, null, 2), 'utf8');
  return outPath;
}

module.exports = { generateBugReport, writeBugReport };
