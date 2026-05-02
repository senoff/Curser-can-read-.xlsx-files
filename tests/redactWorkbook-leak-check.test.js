// End-to-end leak verification for --export-redacted-workbook.
//
// Builds a "loaded" fixture containing realistic PII (email, password, SSN,
// API key, hidden sheets, document properties with real author names) then
// runs the redactor and asserts that none of the sensitive strings survive in
// the output ZIP.
//
// Phase 1c of the 2026-05-02 security review — per task specification.

'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const JSZip = require('jszip');
const engine = require('../lib/engine');
const {
  exportRedactedWorkbook,
  _redactCoreXml,
  _redactAppXml,
  _redactCustomPropsXml,
  _redactCommentsXml,
  _redactThreadedCommentsXml,
  _redactPersonsXml,
  _TRANSPARENT_1X1_PNG,
} = require('../lib/redactWorkbook');

// ---------------------------------------------------------------------------
// Fixture helpers
// ---------------------------------------------------------------------------

// Synthetic PNG: the PNG magic bytes + a recognizable payload we can grep for.
// We don't need a fully valid PNG here — just something with PNG magic bytes
// and a distinguishable body that the redactor must not pass through.
const PNG_MAGIC = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); // \x89PNG\r\n\x1a\n
// Add a fake unique payload after the magic bytes so we can confirm the
// original bytes don't survive (the placeholder is a different buffer).
const FAKE_PNG_PAYLOAD = Buffer.concat([PNG_MAGIC, Buffer.from('FAKE-IMAGE-DATA-DO-NOT-LEAK', 'utf8')]);

const SENSITIVE_EMAIL = 'alice@example.com';
const SENSITIVE_PASSWORD = 'hunter2';
const SENSITIVE_SSN = '123-45-6789';
const SENSITIVE_COMBINED = `${SENSITIVE_EMAIL} password=${SENSITIVE_PASSWORD} ssn=${SENSITIVE_SSN}`;
const SENSITIVE_API_KEY = 'secret-api-key-abc123';
const SENSITIVE_AUTHOR = 'Alice Smith';
const SENSITIVE_LAST_MOD = 'Bob Reporter';
const SENSITIVE_TITLE = 'Q4 Confidential Budget';
const SENSITIVE_COMPANY = 'Acme Corp';

// All strings that must NEVER appear in redacted output.
const MUST_NOT_APPEAR = [
  SENSITIVE_EMAIL,
  SENSITIVE_PASSWORD,
  SENSITIVE_SSN,
  SENSITIVE_API_KEY,
  SENSITIVE_AUTHOR,
  SENSITIVE_LAST_MOD,
  SENSITIVE_TITLE,
  SENSITIVE_COMPANY,
];

async function buildLoadedFixture(outPath) {
  const wb = engine.createWorkbook();

  // Sheet 1: PII data in multiple cell types + merged range
  const s1 = wb.addWorksheet('PII Sheet');
  s1.getCell('A1').value = 'Name';
  s1.getCell('B1').value = 'Email';
  s1.getCell('C1').value = 'Password';
  s1.getCell('D1').value = 'SSN';
  s1.getCell('A2').value = 'Alice';
  s1.getCell('B2').value = SENSITIVE_EMAIL;
  s1.getCell('C2').value = SENSITIVE_PASSWORD;
  s1.getCell('D2').value = SENSITIVE_SSN;
  s1.getCell('A3').value = SENSITIVE_COMBINED; // combined sentinel
  // Formula cell referencing sensitive numeric data
  s1.getCell('E2').value = { formula: 'COUNTA(A2:D2)', result: 4 };
  // Boolean and date
  s1.getCell('F1').value = 'Active';
  s1.getCell('F2').value = true;
  s1.getCell('G1').value = 'Created';
  s1.getCell('G2').value = new Date('2024-01-15');
  // Merged cell with sensitive content
  s1.mergeCells('A4:D4');
  s1.getCell('A4').value = SENSITIVE_COMBINED;
  // Defined name
  wb.definedNames.add("'PII Sheet'!$B$2:$D$2", 'SensitiveData');

  // Sheet 2: Hidden sheet — must still be redacted even though hidden
  const s2 = wb.addWorksheet('Hidden Config');
  s2.state = 'hidden';
  s2.getCell('A1').value = 'API Key';
  s2.getCell('B1').value = SENSITIVE_API_KEY;
  s2.getCell('A2').value = 'Password';
  s2.getCell('B2').value = SENSITIVE_PASSWORD;
  s2.getCell('A3').value = SENSITIVE_COMBINED;

  // Sheet 3: Numeric data
  const s3 = wb.addWorksheet('Numbers');
  s3.getCell('A1').value = 42;
  s3.getCell('B1').value = 3.14;
  s3.getCell('C1').value = { formula: 'A1*B1', result: 131.88 };

  await wb.xlsx.writeFile(outPath);
}

async function injectDocProps(xlsxPath) {
  // Inject real author/company metadata into docProps/core.xml and
  // docProps/app.xml via JSZip — ExcelJS writes "Unknown" by default so
  // we manually set realistic PII here to test the redactor.
  const buf = fs.readFileSync(xlsxPath);
  const zip = await JSZip.loadAsync(buf);

  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${SENSITIVE_AUTHOR}</dc:creator><dc:title>${SENSITIVE_TITLE}</dc:title><cp:lastModifiedBy>${SENSITIVE_LAST_MOD}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T09:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2024-04-30T17:30:00Z</dcterms:modified></cp:coreProperties>`);

  zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application><Company>${SENSITIVE_COMPANY}</Company><Manager>${SENSITIVE_LAST_MOD}</Manager></Properties>`);

  // Inject a synthetic media file with recognisable bytes to test xl/media/ stripping.
  zip.file('xl/media/image1.png', FAKE_PNG_PAYLOAD);

  const out = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  fs.writeFileSync(xlsxPath, out);
}

// ---------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------

let workdir;
let fixturePath;
let redactedPath;

test.before(async () => {
  workdir = fs.mkdtempSync(path.join(os.tmpdir(), 'xfa-leak-check-'));
  fixturePath = path.join(workdir, 'loaded.xlsx');
  redactedPath = path.join(workdir, 'loaded-redacted.xlsx');

  await buildLoadedFixture(fixturePath);
  await injectDocProps(fixturePath);
  await exportRedactedWorkbook(fixturePath, redactedPath);
});

test.after(() => {
  if (workdir) fs.rmSync(workdir, { recursive: true, force: true });
});

// ---------------------------------------------------------------------------
// Test 1 — raw ZIP grep: none of the sensitive strings survive
// ---------------------------------------------------------------------------

test('raw ZIP: no sensitive strings in any XML/rels part', async () => {
  const buf = fs.readFileSync(redactedPath);
  const zip = await JSZip.loadAsync(buf);

  for (const name of Object.keys(zip.files)) {
    const file = zip.file(name);
    if (!file || file.dir) continue;
    if (!/\.(xml|rels)$/i.test(name)) continue;

    const xml = await file.async('string');
    for (const s of MUST_NOT_APPEAR) {
      assert.equal(
        xml.includes(s),
        false,
        `[LEAK] "${s}" found in ${name}`
      );
    }
  }
});

// ---------------------------------------------------------------------------
// Test 2 — cell values in PII sheet redacted
// ---------------------------------------------------------------------------

test('cell values: PII sheet string cells all become "x"', async () => {
  const wb = await engine.loadWorkbook(redactedPath);
  const pii = wb.getWorksheet('PII Sheet');
  assert.ok(pii, 'PII Sheet must be present');

  for (const ref of ['B2', 'C2', 'D2', 'A3']) {
    const v = pii.getCell(ref).value;
    const flat = typeof v === 'string' ? v : (v && v.richText ? v.richText.map((r) => r.text).join('') : v);
    assert.equal(flat, 'x', `${ref} expected "x", got ${JSON.stringify(v)}`);
  }
});

// ---------------------------------------------------------------------------
// Test 3 — merged cell redacted
// ---------------------------------------------------------------------------

test('merged cell A4: sensitive string replaced with "x"', async () => {
  const wb = await engine.loadWorkbook(redactedPath);
  const pii = wb.getWorksheet('PII Sheet');
  const v = pii.getCell('A4').value;
  const flat = typeof v === 'string' ? v : (v && v.richText ? v.richText.map((r) => r.text).join('') : v);
  assert.equal(flat, 'x', `A4 (merged) expected "x", got ${JSON.stringify(v)}`);
});

// ---------------------------------------------------------------------------
// Test 4 — hidden sheet content redacted
// ---------------------------------------------------------------------------

test('hidden sheet: cell values redacted even on hidden sheet', async () => {
  const wb = await engine.loadWorkbook(redactedPath);
  const hidden = wb.getWorksheet('Hidden Config');
  assert.ok(hidden, 'Hidden Config sheet must be present');
  assert.equal(hidden.state, 'hidden', 'sheet must remain hidden (state preserved)');

  const b1 = hidden.getCell('B1').value;
  const b2 = hidden.getCell('B2').value;
  const a3 = hidden.getCell('A3').value;

  const flatB1 = typeof b1 === 'string' ? b1 : (b1 && b1.richText ? b1.richText.map((r) => r.text).join('') : b1);
  const flatB2 = typeof b2 === 'string' ? b2 : (b2 && b2.richText ? b2.richText.map((r) => r.text).join('') : b2);
  const flatA3 = typeof a3 === 'string' ? a3 : (a3 && a3.richText ? a3.richText.map((r) => r.text).join('') : a3);

  assert.equal(flatB1, 'x', `Hidden B1 (was API key) expected "x", got ${JSON.stringify(b1)}`);
  assert.equal(flatB2, 'x', `Hidden B2 (was password) expected "x", got ${JSON.stringify(b2)}`);
  assert.equal(flatA3, 'x', `Hidden A3 (was combined sentinel) expected "x", got ${JSON.stringify(a3)}`);
});

// ---------------------------------------------------------------------------
// Test 5 — document properties stripped
// ---------------------------------------------------------------------------

test('docProps/core.xml: author, title, lastModifiedBy stripped', async () => {
  const buf = fs.readFileSync(redactedPath);
  const zip = await JSZip.loadAsync(buf);
  const core = zip.file('docProps/core.xml');
  assert.ok(core, 'docProps/core.xml must exist');
  const xml = await core.async('string');

  assert.ok(!xml.includes(SENSITIVE_AUTHOR), `dc:creator "${SENSITIVE_AUTHOR}" must be stripped`);
  assert.ok(!xml.includes(SENSITIVE_LAST_MOD), `cp:lastModifiedBy "${SENSITIVE_LAST_MOD}" must be stripped`);
  assert.ok(!xml.includes(SENSITIVE_TITLE), `dc:title "${SENSITIVE_TITLE}" must be stripped`);
  // Timestamps must survive (structural, non-identifying)
  assert.ok(xml.includes('dcterms:created'), 'timestamp elements must survive');
});

// ---------------------------------------------------------------------------
// Test 6 — app.xml: Company and Manager stripped
// ---------------------------------------------------------------------------

test('docProps/app.xml: Company and Manager stripped', async () => {
  const buf = fs.readFileSync(redactedPath);
  const zip = await JSZip.loadAsync(buf);
  const app = zip.file('docProps/app.xml');
  assert.ok(app, 'docProps/app.xml must exist');
  const xml = await app.async('string');

  assert.ok(!xml.includes(SENSITIVE_COMPANY), `Company "${SENSITIVE_COMPANY}" must be stripped`);
  assert.ok(!xml.includes(SENSITIVE_LAST_MOD), `Manager "${SENSITIVE_LAST_MOD}" must be stripped`);
  // Application name is structural, must survive
  assert.ok(xml.includes('<Application>'), 'Application element must survive');
});

// ---------------------------------------------------------------------------
// Test 7 — formula cells preserved
// ---------------------------------------------------------------------------

test('formula cells preserved through redaction', async () => {
  const wb = await engine.loadWorkbook(redactedPath);
  const pii = wb.getWorksheet('PII Sheet');
  const e2 = pii.getCell('E2').value;
  const formula = e2 && (e2.formula || e2.sharedFormula);
  assert.ok(formula, `E2 must still be a formula cell, got ${JSON.stringify(e2)}`);

  const nums = wb.getWorksheet('Numbers');
  const c1 = nums.getCell('C1').value;
  const f2 = c1 && (c1.formula || c1.sharedFormula);
  assert.ok(f2, `Numbers C1 must still be a formula cell, got ${JSON.stringify(c1)}`);
});

// ---------------------------------------------------------------------------
// Test 8 — numeric cells become 0
// ---------------------------------------------------------------------------

test('numeric cells redacted to 0', async () => {
  const wb = await engine.loadWorkbook(redactedPath);
  const nums = wb.getWorksheet('Numbers');
  assert.equal(nums.getCell('A1').value, 0, 'A1 (was 42) must be 0');
  assert.equal(nums.getCell('B1').value, 0, 'B1 (was 3.14) must be 0');
});

// ---------------------------------------------------------------------------
// Test 9 — boolean cell becomes false/0
// ---------------------------------------------------------------------------

test('boolean cell redacted to false/0', async () => {
  const wb = await engine.loadWorkbook(redactedPath);
  const pii = wb.getWorksheet('PII Sheet');
  const f2 = pii.getCell('F2').value;
  assert.ok(f2 === false || f2 === 0, `F2 (was true) expected false or 0, got ${JSON.stringify(f2)}`);
});

// ---------------------------------------------------------------------------
// Unit tests for XML redactors
// ---------------------------------------------------------------------------

test('_redactCoreXml: strips all PII fields, preserves timestamps', () => {
  const input = `<cp:coreProperties><dc:creator>Alice Smith</dc:creator><dc:title>Budget 2024</dc:title><dc:subject>Finance</dc:subject><dc:description>Confidential</dc:description><cp:keywords>budget finance</cp:keywords><cp:category>Reports</cp:category><cp:lastModifiedBy>Bob Reporter</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2024-04-30T00:00:00Z</dcterms:modified></cp:coreProperties>`;
  const out = _redactCoreXml(input);

  const stripped = ['Alice Smith', 'Budget 2024', 'Finance', 'Confidential', 'budget finance', 'Reports', 'Bob Reporter'];
  for (const s of stripped) {
    assert.ok(!out.includes(s), `"${s}" must be stripped from core.xml`);
  }
  assert.ok(out.includes('dcterms:created'), 'dcterms:created must survive');
  assert.ok(out.includes('dcterms:modified'), 'dcterms:modified must survive');
});

test('_redactAppXml: strips Company and Manager, preserves Application', () => {
  const input = `<Properties><Application>Microsoft Excel</Application><Company>Acme Corp</Company><Manager>Jane Doe</Manager><AppVersion>16.0</AppVersion></Properties>`;
  const out = _redactAppXml(input);
  assert.ok(!out.includes('Acme Corp'), 'Company must be stripped');
  assert.ok(!out.includes('Jane Doe'), 'Manager must be stripped');
  assert.ok(out.includes('Microsoft Excel'), 'Application must survive');
  assert.ok(out.includes('16.0'), 'AppVersion must survive');
});

test('_redactCustomPropsXml: strips custom property values', () => {
  const input = `<Properties><property name="ProjectCode" fmtid="{...}" pid="2"><vt:lpwstr>SECRET-PROJECT-X</vt:lpwstr></property></Properties>`;
  const out = _redactCustomPropsXml(input);
  assert.ok(!out.includes('SECRET-PROJECT-X'), 'custom property value must be stripped');
  assert.ok(out.includes('ProjectCode'), 'property name must survive');
});

// C1 regression: numeric custom-property types (vt:r4/r8/i4/etc.) must be redacted.
test('_redactCustomPropsXml: numeric vt types (r8, i4, ui8, filetime) are stripped', () => {
  const cases = [
    ['vt:r8', '123456.78'],
    ['vt:r4', '3.14'],
    ['vt:i4', '987654321'],
    ['vt:i8', '1234567890123'],
    ['vt:ui8', '9999999999'],
    ['vt:bool', 'true'],
    ['vt:filetime', '2024-01-15T09:00:00Z'],
  ];
  for (const [tag, value] of cases) {
    const input = `<Properties><property name="X"><${tag}>${value}</${tag}></property></Properties>`;
    const out = _redactCustomPropsXml(input);
    assert.ok(
      !out.includes(value),
      `${tag} value "${value}" must be stripped, got ${JSON.stringify(out)}`,
    );
    assert.ok(out.includes(`<${tag}>`), `${tag} open tag must survive`);
    assert.ok(out.includes(`</${tag}>`), `${tag} close tag must survive`);
  }
});

// C1 regression: nested vt:variant > vt:lpwstr must produce well-formed XML.
test('_redactCustomPropsXml: nested vt:variant > vt:lpwstr stays well-formed', () => {
  const input = `<Properties><property name="V"><vt:variant><vt:lpwstr>SECRET</vt:lpwstr></vt:variant></property></Properties>`;
  const out = _redactCustomPropsXml(input);
  assert.ok(!out.includes('SECRET'), 'nested vt:lpwstr value must be stripped');
  // Old regex would have matched <vt:variant>...</vt:lpwstr> producing mangled tags.
  // Verify the structure is intact: <vt:variant><vt:lpwstr></vt:lpwstr></vt:variant>
  assert.ok(out.includes('<vt:variant>'), 'outer vt:variant open tag survives');
  assert.ok(out.includes('</vt:variant>'), 'outer vt:variant close tag survives');
  assert.ok(out.includes('<vt:lpwstr>'), 'inner vt:lpwstr open tag survives');
  assert.ok(out.includes('</vt:lpwstr>'), 'inner vt:lpwstr close tag survives');
  // No mangled tag like </vt:variant><vt:variant> or stray attributes.
  const variantOpens = (out.match(/<vt:variant\b/g) || []).length;
  const variantCloses = (out.match(/<\/vt:variant>/g) || []).length;
  assert.equal(variantOpens, variantCloses, 'vt:variant open/close counts must match');
  const lpwstrOpens = (out.match(/<vt:lpwstr\b/g) || []).length;
  const lpwstrCloses = (out.match(/<\/vt:lpwstr>/g) || []).length;
  assert.equal(lpwstrOpens, lpwstrCloses, 'vt:lpwstr open/close counts must match');
});

// C2 regression: legacy comment authors must be scrubbed.
test('_redactCommentsXml: <author> display names are scrubbed', () => {
  const input = `<comments><authors><author>Alice Smith</author><author>Bob Reporter</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>note</t></r></text></comment></commentList></comments>`;
  const out = _redactCommentsXml(input);
  assert.ok(!out.includes('Alice Smith'), 'first author must be scrubbed');
  assert.ok(!out.includes('Bob Reporter'), 'second author must be scrubbed');
  assert.ok(out.includes('<author>x</author>'), 'author tag should retain placeholder "x"');
  // <t> payload "note" must also be redacted by the existing branch.
  assert.ok(!out.includes('>note<'), 'comment body text must remain redacted');
});

// C3 regression: xl/persons/person.xml identifying attributes scrubbed.
test('_redactPersonsXml: displayName, userId, providerId attributes scrubbed', () => {
  const input = `<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person displayName="Alice Smith" id="{abcd-ef}" userId="alice@company.com" providerId="AzureAD"/></personList>`;
  const out = _redactPersonsXml(input);
  assert.ok(!out.includes('Alice Smith'), 'displayName must be scrubbed');
  assert.ok(!out.includes('alice@company.com'), 'userId must be scrubbed');
  assert.ok(!out.includes('AzureAD'), 'providerId must be scrubbed');
  // id (UUID) must survive so threadedComment authorId references resolve.
  assert.ok(out.includes('{abcd-ef}'), 'id (UUID) must survive');
  assert.ok(out.includes('displayName="x"'), 'displayName placeholder present');
  assert.ok(out.includes('userId="x"'), 'userId placeholder present');
  assert.ok(out.includes('providerId="x"'), 'providerId placeholder present');
});

// H1 regression: single-quoted text="..." attribute on threadedComments.
test('_redactThreadedCommentsXml: handles both double-quoted and single-quoted text=', () => {
  const dq = `<threadedComment id="x" text="DOUBLE-QUOTED-SECRET"/>`;
  const sq = `<threadedComment id="x" text='SINGLE-QUOTED-SECRET'/>`;
  const out1 = _redactThreadedCommentsXml(dq);
  const out2 = _redactThreadedCommentsXml(sq);
  assert.ok(!out1.includes('DOUBLE-QUOTED-SECRET'), 'double-quoted body must be scrubbed');
  assert.ok(!out2.includes('SINGLE-QUOTED-SECRET'), 'single-quoted body must be scrubbed');
  assert.ok(out1.includes('text="x"'), 'double-quoted form replaced with text="x"');
  assert.ok(out2.includes('text="x"'), 'single-quoted form replaced with text="x"');
});

// ---------------------------------------------------------------------------
// Test 10 — xl/media/ binaries stripped (follow-up to #20)
// ---------------------------------------------------------------------------

test('xl/media/: embedded image bytes do not survive redaction', async () => {
  // The fixture already has xl/media/image1.png injected with FAKE_PNG_PAYLOAD
  // (see injectDocProps). Verify the output contains the entry (ZIP remains
  // structurally intact) but the original bytes are gone.
  const buf = fs.readFileSync(redactedPath);
  const zip = await JSZip.loadAsync(buf);

  const mediaEntry = zip.file('xl/media/image1.png');
  assert.ok(mediaEntry, 'xl/media/image1.png must still exist in ZIP (structural integrity)');

  const outBytes = await mediaEntry.async('nodebuffer');

  // Original fake payload must not survive.
  const originalPayload = Buffer.from('FAKE-IMAGE-DATA-DO-NOT-LEAK', 'utf8');
  assert.equal(
    outBytes.indexOf(originalPayload),
    -1,
    'Original image bytes must not survive in redacted output',
  );

  // Replacement must be the transparent PNG placeholder (or at least have
  // PNG magic bytes — confirming it was replaced with valid PNG, not just
  // wiped to empty which would break drawing rels).
  assert.ok(
    outBytes.length > 0,
    'Replacement entry must not be empty (would break drawing relationships)',
  );
  assert.equal(
    outBytes.slice(0, 4).toString('hex'),
    '89504e47', // PNG magic bytes \x89PNG
    'Replacement must start with PNG magic bytes',
  );

  // Confirm the placeholder is the _TRANSPARENT_1X1_PNG constant.
  assert.deepEqual(
    outBytes,
    _TRANSPARENT_1X1_PNG,
    'Replacement must equal the TRANSPARENT_1X1_PNG placeholder',
  );
});
