// Redacted-workbook exporter.
//
// Reads an .xlsx as a ZIP, mutates only the *value* portions of each
// cell (and the shared-string + comment payloads) to typed placeholders,
// then repacks. Everything else — formulas, styles, sheet names, named
// ranges, feature parts (pivots / charts / queries / vba) — is passed
// through byte-for-byte where possible.
//
// Why ZIP-passthrough rather than ExcelJS round-trip:
//   ExcelJS write() is lossy for many features (pivots, slicers,
//   queries, conditional formatting, sparklines, threaded comments).
//   For a bug-repro artifact we want maximum structural fidelity, so
//   we operate at the XML-fragment level inside the existing ZIP.
//
// Placeholders:
//   numbers   → 0
//   strings   → "x"
//   booleans  → false (0)
//   dates     → 1900-01-01 (numeric date cells render to default date
//                under their existing format; t="d" cells get the
//                literal ISO string)
//   errors    → preserved as-is
//
// Comments and shared strings are also rewritten to "x" because they
// contain user text. Defined-name formulas are preserved (per spec).

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

// Match each <c ...>...</c> or self-closing <c .../> element.
// We deliberately restrict to a single regex pass per sheet — this is
// fragile only if a cell contains a nested <c> in user-supplied XML,
// which OOXML cells do not.
const CELL_RE = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;

// Cell type attribute extractor.
function getAttr(attrs, name) {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(attrs);
  return m ? m[1] : null;
}
function setAttr(attrs, name, value) {
  if (new RegExp(`\\b${name}="`).test(attrs)) {
    return attrs.replace(new RegExp(`\\b${name}="[^"]*"`), `${name}="${value}"`);
  }
  return `${attrs} ${name}="${value}"`;
}
function removeAttr(attrs, name) {
  return attrs.replace(new RegExp(`\\s*\\b${name}="[^"]*"`), '');
}

// Extract first <f ...>...</f> or <f .../> from a cell body. Preserve verbatim.
const F_RE = /<f\b[^>]*(?:\/>|>[\s\S]*?<\/f>)/;

function redactCell(match, attrs, selfOrBody, body) {
  // Self-closing <c r="A1"/> — empty cell, nothing to redact.
  if (selfOrBody === '/>') return match;

  const t = getAttr(attrs, 't');
  const fMatch = body.match(F_RE);
  const formulaXml = fMatch ? fMatch[0] : '';

  // Errors: preserve the value as-is. Cell type is "e".
  if (t === 'e') {
    return match;
  }

  // Inline string: rebuild as <is><t>x</t></is>.
  if (t === 'inlineStr') {
    return `<c${attrs}>${formulaXml}<is><t>x</t></is></c>`;
  }

  // Shared string: convert to inline string so we don't depend on the
  // sst index meaning anything. (We also rewrite sst payloads to "x"
  // for defense-in-depth, but this avoids index-collision worries.)
  if (t === 's') {
    let newAttrs = setAttr(attrs, 't', 'inlineStr');
    return `<c${newAttrs}>${formulaXml}<is><t>x</t></is></c>`;
  }

  // Formula returning a literal string.
  if (t === 'str') {
    return `<c${attrs}>${formulaXml}<v>x</v></c>`;
  }

  // Boolean → false (0).
  if (t === 'b') {
    return `<c${attrs}>${formulaXml}<v>0</v></c>`;
  }

  // ISO-date typed cell.
  if (t === 'd') {
    return `<c${attrs}>${formulaXml}<v>1900-01-01</v></c>`;
  }

  // Default = number (no t attribute, or t="n"). Whether it's a date
  // is encoded in the *style* (numFmt), not the cell type. By
  // replacing the numeric value with 0, a date-styled cell will render
  // as 1900-01-00 / 1900-01-01 depending on the date system in use,
  // which is the documented placeholder.
  return `<c${attrs}>${formulaXml}<v>0</v></c>`;
}

function redactSheetXml(xml) {
  return xml.replace(CELL_RE, redactCell);
}

// Shared strings: every <t>...</t> payload becomes "x". Preserves the
// number of unique strings + their indices so cells that happen to
// reference sst still resolve to a valid (redacted) string.
function redactSharedStringsXml(xml) {
  // Replace inner text of every <t> element (handles <t>x</t> and
  // <t xml:space="preserve">x</t>). Empty payloads stay empty.
  return xml.replace(/(<t\b[^>]*>)([\s\S]*?)(<\/t>)/g, (m, open, payload, close) => {
    return open + (payload === '' ? '' : 'x') + close;
  });
}

// Comments: <comment><text><r>...<t>USER TEXT</t></r></text></comment>
// Replace every <t> payload with "x". Also strips <author>NAME</author>
// display names in <authors>; the numeric authorId on each <comment>
// references the (now redacted) author entry.
function redactCommentsXml(xml) {
  let out = xml.replace(/(<t\b[^>]*>)([\s\S]*?)(<\/t>)/g, (m, open, payload, close) => {
    return open + (payload === '' ? '' : 'x') + close;
  });
  out = out.replace(/(<author\b[^>]*>)([\s\S]*?)(<\/author>)/g, (m, open, payload, close) => {
    return open + (payload === '' ? '' : 'x') + close;
  });
  return out;
}

// Threaded comments: <threadedComment ... text="USER TEXT" .../>
// Excel encodes the body as an attribute — must redact in place. Both
// double-quoted and single-quoted attribute values are valid XML and we
// must scrub both forms.
function redactThreadedCommentsXml(xml) {
  return xml.replace(/\btext=("[^"]*"|'[^']*')/g, 'text="x"');
}

// xl/persons/person.xml — author registry for threaded comments.
// <person displayName="Alice" id="..." userId="alice@co.com" providerId="AzureAD"/>
// Strip the three identifying attributes; leave id (a UUID) so threaded comment
// authorId references still resolve.
function redactPersonsXml(xml) {
  return xml
    .replace(/\bdisplayName="[^"]*"/g, 'displayName="x"')
    .replace(/\buserId="[^"]*"/g, 'userId="x"')
    .replace(/\bproviderId="[^"]*"/g, 'providerId="x"');
}

// docProps/core.xml — strip author, title, subject, description, keywords,
// category, lastModifiedBy, and any other user-text elements.
// The timestamp elements (dcterms:created / dcterms:modified) and structural
// elements (the xmlns declarations, DocSecurity, etc.) are left alone because
// they're non-identifying metadata needed for round-trip fidelity.
//
// Elements scrubbed:
//   dc:creator         → the file's original author name
//   dc:title           → document title set by author
//   dc:subject         → subject field
//   dc:description     → description field
//   cp:keywords        → keyword tags
//   cp:category        → category field
//   cp:lastModifiedBy  → last editor's name
//   cp:contentStatus   → rarely set, but can contain user text
const CORE_SCRUB_TAGS = [
  'dc:creator',
  'dc:title',
  'dc:subject',
  'dc:description',
  'cp:keywords',
  'cp:category',
  'cp:lastModifiedBy',
  'cp:contentStatus',
];

function redactCoreXml(xml) {
  let out = xml;
  for (const tag of CORE_SCRUB_TAGS) {
    // Replace inner content: <dc:creator>...</dc:creator> → <dc:creator></dc:creator>
    // Handles attributes on the opening tag and multi-line content.
    out = out.replace(
      new RegExp(`(<${tag}\\b[^>]*>)[\\s\\S]*?(<\\/${tag}>)`, 'g'),
      '$1$2'
    );
  }
  return out;
}

// docProps/app.xml — strip Company, Manager, and HyperlinkBase which can
// contain user-identifying strings. The Application, AppVersion, DocSecurity,
// HeadingPairs, and TitlesOfParts (sheet names) fields are structural and left
// alone — sheet names are part of workbook structure, not cell values.
const APP_SCRUB_TAGS = [
  'Company',
  'Manager',
  'HyperlinkBase',
];

function redactAppXml(xml) {
  let out = xml;
  for (const tag of APP_SCRUB_TAGS) {
    out = out.replace(
      new RegExp(`(<${tag}\\b[^>]*>)[\\s\\S]*?(<\\/${tag}>)`, 'g'),
      '$1$2'
    );
  }
  return out;
}

// docProps/custom.xml — custom properties are arbitrary user-defined key/value
// pairs. Strip the value payloads; keep the property names so the file remains
// structurally valid.
function redactCustomPropsXml(xml) {
  // Custom property values live inside <vt:*> typed-value elements.
  // Replace their inner text with empty string (preserves type nodes).
  //
  // The character class includes digits so OOXML numeric type names
  // (vt:r4, vt:r8, vt:i1/i2/i4/i8, vt:ui1/ui2/ui4/ui8, vt:filetime) match.
  // The \2 backreference forces the open and close tag names to match,
  // so a payload that contains nested elements (e.g.
  // <vt:variant><vt:lpwstr>X</vt:lpwstr></vt:variant>) doesn't produce
  // mangled output. The inner text class [^<] keeps the match strictly
  // text-only; the outer wrapper is left structurally intact and any
  // nested vt:* elements get scrubbed by the same regex on overlapping
  // passes.
  return xml.replace(/(<(vt:[a-zA-Z0-9]+)\b[^>]*>)[^<]*(<\/\2>)/g, '$1$3');
}

// 1×1 transparent PNG — minimum valid PNG bytes. Used as a safe placeholder
// when stripping xl/media/ binary blobs so the ZIP remains structurally valid
// and drawing relationships don't point to missing entries.
// (96 bytes: PNG sig + IHDR + IDAT with one transparent pixel + IEND)
const TRANSPARENT_1X1_PNG = Buffer.from(
  '89504e470d0a1a0a' + // PNG signature
  '0000000d49484452' + // IHDR length + type
  '00000001' +         // width = 1
  '00000001' +         // height = 1
  '08060000' +         // 8-bit RGBA
  '001f15c4' +         // IHDR CRC
  '89' +               // IHDR chunk footer padding
  '0000000a49444154' + // IDAT length + type
  '789c6260' +         // zlib header + deflate block
  '0000000200' +       // deflate end
  '01e221bc33' +       // IDAT CRC
  '0000000049454e44' + // IEND length + type
  'ae426082',          // IEND CRC
  'hex',
);

async function exportRedactedWorkbook(inputPath, outputPath) {
  if (!fs.existsSync(inputPath)) {
    throw new Error(`File not found: ${inputPath}`);
  }
  const ext = path.extname(inputPath).toLowerCase();
  if (ext !== '.xlsx' && ext !== '.xlsm') {
    throw new Error(`--export-redacted-workbook only supports .xlsx / .xlsm (got ${ext})`);
  }

  const buf = fs.readFileSync(inputPath);
  const zip = await JSZip.loadAsync(buf);

  const filenames = Object.keys(zip.files).filter((n) => !zip.files[n].dir);

  for (const name of filenames) {
    const file = zip.file(name);
    if (!file || file.dir) continue;

    if (/^xl\/worksheets\/sheet\d+\.xml$/i.test(name)) {
      const xml = await file.async('string');
      zip.file(name, redactSheetXml(xml));
    } else if (/^xl\/sharedStrings\.xml$/i.test(name)) {
      const xml = await file.async('string');
      zip.file(name, redactSharedStringsXml(xml));
    } else if (/^xl\/comments\d+\.xml$/i.test(name)) {
      const xml = await file.async('string');
      zip.file(name, redactCommentsXml(xml));
    } else if (/^xl\/threadedComments\/threadedComment\d+\.xml$/i.test(name)) {
      const xml = await file.async('string');
      zip.file(name, redactThreadedCommentsXml(xml));
    } else if (/^docProps\/core\.xml$/i.test(name)) {
      const xml = await file.async('string');
      zip.file(name, redactCoreXml(xml));
    } else if (/^docProps\/app\.xml$/i.test(name)) {
      const xml = await file.async('string');
      zip.file(name, redactAppXml(xml));
    } else if (/^docProps\/custom\.xml$/i.test(name)) {
      const xml = await file.async('string');
      zip.file(name, redactCustomPropsXml(xml));
    } else if (/^xl\/persons\/person\.xml$/i.test(name)) {
      const xml = await file.async('string');
      zip.file(name, redactPersonsXml(xml));
    } else if (/^xl\/media\//i.test(name)) {
      // Embedded images / media — replace with a 1×1 transparent PNG so
      // drawing relationships remain valid and the ZIP is structurally intact,
      // but no user-supplied binary data survives in the output.
      zip.file(name, TRANSPARENT_1X1_PNG);
    }
    // All other parts pass through untouched.
  }

  // Use store-or-deflate matching Excel's defaults (deflate level 6).
  const out = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  fs.writeFileSync(outputPath, out);
  return outputPath;
}

module.exports = {
  exportRedactedWorkbook,
  // exported for unit testing
  _redactSheetXml: redactSheetXml,
  _redactSharedStringsXml: redactSharedStringsXml,
  _redactCommentsXml: redactCommentsXml,
  _redactThreadedCommentsXml: redactThreadedCommentsXml,
  _redactCoreXml: redactCoreXml,
  _redactAppXml: redactAppXml,
  _redactCustomPropsXml: redactCustomPropsXml,
  _redactPersonsXml: redactPersonsXml,
  _TRANSPARENT_1X1_PNG: TRANSPARENT_1X1_PNG,
};
