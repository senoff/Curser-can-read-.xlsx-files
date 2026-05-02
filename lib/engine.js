// Engine abstraction layer.
//
// xlsx-for-ai's logic shouldn't depend directly on ExcelJS. This module is
// the *seam* between xlsx-for-ai's code and the underlying xlsx engine —
// today ExcelJS, tomorrow possibly a fork, a from-scratch JS port,
// xlsx-populate, or SheetJS Pro server-side.
//
// The exposed surface is intentionally narrow: file I/O entry points
// (load, stream, write), workbook construction, and the small set of
// ExcelJS constants the rest of the codebase uses. The in-memory workbook
// representation flows through this layer unchanged — at this stage the
// goal is to centralize *which engine produces the workbook objects*, not
// to define a fully-engine-agnostic in-memory model.
//
// To swap engines, replace this file. xlsx-for-ai's other modules import
// only from here; nothing else has a direct require('@protobi/exceljs').

'use strict';

const ExcelJS = require('@protobi/exceljs');

class ExcelJSEngine {
  /** Engine identifier — useful for diagnostics. */
  get name() { return 'exceljs'; }
  get version() {
    try { return require('@protobi/exceljs/package.json').version; } catch (_) { return 'unknown'; }
  }

  /**
   * Load a workbook from a file path. Returns the engine's workbook object
   * (currently an ExcelJS Workbook).
   */
  async loadWorkbook(filePath) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    return wb;
  }

  /** Construct an empty workbook (used by write mode and CSV/TSV/legacy load paths). */
  createWorkbook() {
    return new ExcelJS.Workbook();
  }

  /** Write a workbook to disk. */
  async writeWorkbook(wb, filePath) {
    return wb.xlsx.writeFile(filePath);
  }

  /** Streaming reader for huge files. Returns an async iterator of sheets. */
  streamReader(filePath, opts) {
    return new ExcelJS.stream.xlsx.WorkbookReader(filePath, opts);
  }

  /**
   * Constants the rest of the codebase needs. Keeping these here means
   * the rest of xlsx-for-ai never imports ExcelJS directly — only from
   * the engine.
   */
  get ValueType() { return ExcelJS.ValueType; }
}

// Singleton: the rest of the codebase imports this module and gets the
// active engine. To swap engines, replace `module.exports` with a different
// engine instance that implements the same surface.
module.exports = new ExcelJSEngine();
