// Synthetic workbook generator: build a small set of test xlsx files in
// memory and write them to a target directory. Each fixture targets a
// specific class of round-trip behavior we want to verify.

'use strict';

const ExcelJS = require('exceljs');
const path = require('path');

// Fixture #1: minimal — values, formulas, dates. The "happy path."
async function basicValues(outDir) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sales');
  ws.getCell('A1').value = 'Region';
  ws.getCell('B1').value = 'Q1';
  ws.getCell('C1').value = 'Q2';
  ws.getCell('D1').value = 'Total';
  ws.getCell('A2').value = 'North';
  ws.getCell('B2').value = 1000;
  ws.getCell('C2').value = 1500;
  ws.getCell('D2').value = { formula: 'B2+C2', result: 2500 };
  ws.getCell('A3').value = 'South';
  ws.getCell('B3').value = 800;
  ws.getCell('C3').value = 1200;
  ws.getCell('D3').value = { formula: 'B3+C3', result: 2000 };
  ws.getCell('A5').value = new Date('2026-01-15');
  // Set column widths
  ws.getColumn(1).width = 12;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 10;
  ws.getColumn(4).width = 12;
  await wb.xlsx.writeFile(path.join(outDir, 'basic-values.xlsx'));
}

// Fixture #2: column widths + frozen panes + hidden columns
async function widthsAndLayout(outDir) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Layout');
  for (let c = 1; c <= 6; c++) {
    ws.getColumn(c).width = 8 + c * 2; // 10, 12, 14, 16, 18, 20
    ws.getCell(1, c).value = `Col${c}`;
  }
  ws.getColumn(3).hidden = true; // hide a middle column
  ws.views = [{ state: 'frozen', ySplit: 1, xSplit: 1 }];
  for (let r = 2; r <= 5; r++) {
    for (let c = 1; c <= 6; c++) {
      ws.getCell(r, c).value = r * 100 + c;
    }
  }
  await wb.xlsx.writeFile(path.join(outDir, 'widths-layout.xlsx'));
}

// Fixture #3: merged cells + named ranges + auto-filter
async function mergesAndNames(outDir) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Report');
  ws.getCell('A1').value = 'Q4 2025 Summary';
  ws.mergeCells('A1:D1');
  ws.getCell('A2').value = 'Category';
  ws.getCell('B2').value = 'Plan';
  ws.getCell('C2').value = 'Actual';
  ws.getCell('D2').value = 'Variance';
  ws.getCell('A3').value = 'Marketing';
  ws.getCell('B3').value = 50000;
  ws.getCell('C3').value = 47500;
  ws.getCell('D3').value = { formula: 'C3-B3', result: -2500 };
  ws.getCell('A4').value = 'Sales';
  ws.getCell('B4').value = 100000;
  ws.getCell('C4').value = 110000;
  ws.getCell('D4').value = { formula: 'C4-B4', result: 10000 };
  ws.autoFilter = 'A2:D4';
  // Workbook-level named ranges
  wb.definedNames.add('Report!$D$3:$D$4', 'Variances');
  wb.definedNames.add('Report!$B$3:$C$4', 'PlanActual');
  await wb.xlsx.writeFile(path.join(outDir, 'merges-names.xlsx'));
}

// Fixture #4: multi-sheet with cross-sheet formulas
async function multiSheet(outDir) {
  const wb = new ExcelJS.Workbook();
  const detail = wb.addWorksheet('Detail');
  detail.getCell('A1').value = 'Item';
  detail.getCell('B1').value = 'Amount';
  detail.getCell('A2').value = 'A';
  detail.getCell('B2').value = 100;
  detail.getCell('A3').value = 'B';
  detail.getCell('B3').value = 250;
  detail.getCell('A4').value = 'C';
  detail.getCell('B4').value = 175;

  const summary = wb.addWorksheet('Summary');
  summary.getCell('A1').value = 'Total';
  summary.getCell('B1').value = { formula: 'SUM(Detail!B2:B4)', result: 525 };
  summary.getCell('A2').value = 'Count';
  summary.getCell('B2').value = { formula: 'COUNTA(Detail!A2:A4)', result: 3 };

  await wb.xlsx.writeFile(path.join(outDir, 'multi-sheet.xlsx'));
}

// Fixture #5: hidden rows + comments + hyperlinks
async function annotations(outDir) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Annotated');
  ws.getCell('A1').value = 'Item';
  ws.getCell('B1').value = 'Link';
  ws.getCell('A2').value = 'Visible';
  ws.getCell('B2').value = { text: 'Example', hyperlink: 'https://example.com' };
  ws.getCell('A3').value = 'Hidden row below';
  ws.getCell('A4').value = 'this row is hidden';
  ws.getRow(4).hidden = true;
  ws.getCell('A5').value = 'Visible again';
  await wb.xlsx.writeFile(path.join(outDir, 'annotations.xlsx'));
}

const FIXTURES = {
  'basic-values.xlsx': basicValues,
  'widths-layout.xlsx': widthsAndLayout,
  'merges-names.xlsx': mergesAndNames,
  'multi-sheet.xlsx': multiSheet,
  'annotations.xlsx': annotations,
};

async function generateAll(outDir) {
  const fs = require('fs');
  fs.mkdirSync(outDir, { recursive: true });
  for (const [name, fn] of Object.entries(FIXTURES)) {
    await fn(outDir);
  }
  return Object.keys(FIXTURES);
}

module.exports = { FIXTURES, generateAll };

// Allow running directly: node test/helpers/synth.js [outDir]
if (require.main === module) {
  const outDir = process.argv[2] || path.join(__dirname, '..', 'fixtures');
  generateAll(outDir).then(names => {
    console.log(`Generated ${names.length} fixtures in ${outDir}:`);
    for (const n of names) console.log(`  ${n}`);
  }).catch(err => {
    console.error('Fixture generation failed:', err);
    process.exit(1);
  });
}
