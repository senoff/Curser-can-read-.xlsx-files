// Build a small synthetic .xlsx fixture for the bug-report and
// redacted-workbook tests. Emits ~/tests/fixtures/bug-fixture.xlsx.
//
// The fixture exercises:
//   - 3 sheets, varying shapes
//   - merged ranges
//   - a defined name (named range)
//   - a formula
//   - one of every cell type that the redactor cares about
//     (number, string, boolean, formula-string, formula-number)
//
// Run: node tests/fixtures/build.js

const path = require('path');
const ExcelJS = require('exceljs');

async function build(outPath) {
  const wb = new ExcelJS.Workbook();

  const s1 = wb.addWorksheet('Sales');
  s1.getCell('A1').value = 'Region';
  s1.getCell('B1').value = 'Q1';
  s1.getCell('C1').value = 'Q2';
  s1.getCell('D1').value = 'Total';
  s1.getCell('A2').value = 'North';
  s1.getCell('B2').value = 100;
  s1.getCell('C2').value = 200;
  s1.getCell('D2').value = { formula: 'B2+C2', result: 300 };
  s1.getCell('A3').value = 'South';
  s1.getCell('B3').value = 50;
  s1.getCell('C3').value = 75;
  s1.getCell('D3').value = { formula: 'B3+C3', result: 125 };
  s1.mergeCells('A5:D5');
  s1.getCell('A5').value = 'Sensitive Customer Notes Here';

  const s2 = wb.addWorksheet('Config');
  s2.getCell('A1').value = 'TaxRate';
  s2.getCell('B1').value = 0.075;
  s2.getCell('A2').value = 'Active';
  s2.getCell('B2').value = true;
  s2.getCell('A3').value = 'Today';
  s2.getCell('B3').value = new Date('2026-04-30');

  const s3 = wb.addWorksheet('Empty');
  s3.getCell('A1').value = null;

  // Defined name pointing into Sales — name "Totals" is captured by
  // bug-report; its formula is NOT.
  wb.definedNames.add('Sales!$D$2:$D$3', 'Totals');

  await wb.xlsx.writeFile(outPath);
  return outPath;
}

if (require.main === module) {
  const out = path.join(__dirname, 'bug-fixture.xlsx');
  build(out).then(() => {
    console.log(out);
  }).catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

module.exports = { build };
