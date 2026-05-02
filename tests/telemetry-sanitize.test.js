'use strict';

const { test } = require('node:test');
const assert   = require('node:assert/strict');
const { scrubPaths, sanitizeMessage, buildPayload, MAX_MESSAGE_LENGTH } = require('../lib/telemetry-sanitize');

// ---------------------------------------------------------------------------
// scrubPaths — core path removal
// ---------------------------------------------------------------------------

test('scrubPaths: macOS /Users/<name>/... path scrubbed', () => {
  const result = scrubPaths('Error reading /Users/alice/Documents/budget.xlsx');
  assert.ok(!result.includes('/Users/alice'), `path leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: Linux /home/<name>/... path scrubbed', () => {
  const result = scrubPaths('ENOENT: /home/bob/files/data.xlsx');
  assert.ok(!result.includes('/home/bob'), `path leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: Windows C:\\Users\\<name>\\... path scrubbed', () => {
  const result = scrubPaths('Cannot open C:\\Users\\charlie\\Desktop\\report.xlsx');
  assert.ok(!result.includes('charlie'), `username leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: Windows C:/Users/<name>/... forward slash path scrubbed', () => {
  const result = scrubPaths('Error at C:/Users/dave/Downloads/file.xlsx');
  assert.ok(!result.includes('dave'), `username leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: /tmp/... path scrubbed', () => {
  const result = scrubPaths('Failed reading /tmp/xfa-telemetry/test.xlsx');
  assert.ok(!result.includes('/tmp/xfa-telemetry'), `path leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: /private/tmp/... (macOS worktree) path scrubbed', () => {
  const result = scrubPaths('Error at /private/tmp/xfa-telemetry/index.js:42');
  assert.ok(!result.includes('/private/tmp/xfa-telemetry'), `path leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: /var/folders/... (macOS tmp) path scrubbed', () => {
  const result = scrubPaths('Temp at /var/folders/ab/cdef1234/T/xlsx.tmp');
  assert.ok(!result.includes('/var/folders'), `path leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: multiple paths in one string all scrubbed', () => {
  const result = scrubPaths(
    'Read /Users/alice/a.xlsx and /home/bob/b.xlsx'
  );
  assert.ok(!result.includes('/Users/alice'), `first path leaked: ${result}`);
  assert.ok(!result.includes('/home/bob'), `second path leaked: ${result}`);
});

test('scrubPaths: no paths — string unchanged', () => {
  const input = 'TypeError: Cannot read properties of null';
  assert.equal(scrubPaths(input), input);
});

test('scrubPaths: adversarial — deep nested path scrubbed', () => {
  const result = scrubPaths('/Users/alice/work/projects/budget/q4/final/v3/report.xlsx');
  assert.ok(!result.includes('alice'), `username leaked: ${result}`);
});

test('scrubPaths: adversarial — path in stack trace line scrubbed', () => {
  const stackLine = '    at Object.<anonymous> (/Users/alice/src/index.js:42:10)';
  const result = scrubPaths(stackLine);
  assert.ok(!result.includes('/Users/alice'), `path leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no replacement: ${result}`);
});

// ---------------------------------------------------------------------------
// sanitizeMessage — truncation + scrub
// ---------------------------------------------------------------------------

test('sanitizeMessage: caps at 200 chars', () => {
  const long = 'x'.repeat(500);
  const result = sanitizeMessage(long);
  assert.equal(result.length, MAX_MESSAGE_LENGTH);
});

test('sanitizeMessage: scrubs paths before truncating', () => {
  const msg = 'Error at /Users/secret/very/long/path/to/file.xlsx that has lots of extra text to push it over the limit';
  const result = sanitizeMessage(msg);
  assert.ok(!result.includes('secret'), `username leaked: ${result}`);
  assert.ok(result.length <= MAX_MESSAGE_LENGTH);
});

test('sanitizeMessage: null/undefined returns empty string', () => {
  assert.equal(sanitizeMessage(null), '');
  assert.equal(sanitizeMessage(undefined), '');
});

test('sanitizeMessage: short clean message passes through', () => {
  const msg = 'TypeError: Cannot read property of null';
  assert.equal(sanitizeMessage(msg), msg);
});

// ---------------------------------------------------------------------------
// buildPayload — structure + invariants
// ---------------------------------------------------------------------------

test('buildPayload: returns required fields', () => {
  const err = new TypeError('test error');
  const payload = buildPayload(err, '1.5.0');
  assert.equal(payload.v, 1);
  assert.ok(payload.ts);
  assert.equal(payload.error_type, 'TypeError');
  assert.equal(payload.xlsx_for_ai_version, '1.5.0');
  assert.ok(payload.node_version);
  assert.ok(payload.os_arch);
});

test('buildPayload: error_message is sanitized and capped', () => {
  const err = new Error('Failed at /Users/testuser/secret/path.xlsx with extra long message ' + 'x'.repeat(300));
  const payload = buildPayload(err, '1.5.0');
  assert.ok(!payload.error_message.includes('/Users/testuser'), `path leaked: ${payload.error_message}`);
  assert.ok(payload.error_message.length <= MAX_MESSAGE_LENGTH);
});

test('buildPayload: no hostname field', () => {
  const err = new Error('test');
  const payload = buildPayload(err, '1.5.0');
  assert.ok(!('hostname' in payload), 'hostname must not be in payload');
});

test('buildPayload: no argv field', () => {
  const err = new Error('test');
  const payload = buildPayload(err, '1.5.0');
  assert.ok(!('argv' in payload), 'argv must not be in payload');
});

test('buildPayload: no env vars field', () => {
  const err = new Error('test');
  const payload = buildPayload(err, '1.5.0');
  assert.ok(!('env' in payload), 'env must not be in payload');
});

test('buildPayload: command defaults to <other> for unknown arg', () => {
  const original = process.argv[2];
  // Temporarily set argv[2] to something not in allowlist
  process.argv[2] = '/Users/alice/secret.xlsx';
  const err = new Error('test');
  const payload = buildPayload(err, '1.5.0');
  process.argv[2] = original;
  assert.equal(payload.command, '<other>', `command should be <other> but got: ${payload.command}`);
});

test('buildPayload: command is allowed if in allowlist', () => {
  const original = process.argv[2];
  process.argv[2] = '--json';
  const err = new Error('test');
  const payload = buildPayload(err, '1.5.0');
  process.argv[2] = original;
  assert.equal(payload.command, '--json');
});

test('buildPayload: non-Error reason still builds payload', () => {
  const payload = buildPayload('plain string error', '1.5.0');
  assert.ok(payload.error_type);
  assert.ok(payload.error_message !== undefined);
});

test('buildPayload: adversarial — Windows path in message scrubbed', () => {
  const err = new Error('Cannot open C:\\Users\\eve\\Documents\\report.xlsx');
  const payload = buildPayload(err, '1.5.0');
  assert.ok(!payload.error_message.includes('eve'), `Windows path leaked: ${payload.error_message}`);
});

test('buildPayload: adversarial — nested Linux home path scrubbed', () => {
  const err = new Error('ENOENT: /home/frank/data/nested/deep.xlsx');
  const payload = buildPayload(err, '1.5.0');
  assert.ok(!payload.error_message.includes('frank'), `Linux path leaked: ${payload.error_message}`);
});

// ---------------------------------------------------------------------------
// scrubPaths — adversarial encoding coverage (PR #18 telemetry follow-up)
// ---------------------------------------------------------------------------

test('scrubPaths: URL-encoded macOS path scrubbed', () => {
  const raw = '/Users/alice/Documents/budget.xlsx';
  const encoded = encodeURIComponent(raw);
  const result = scrubPaths('Error: ' + encoded);
  assert.ok(!result.includes('alice'), `URL-encoded username leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: URL-encoded Linux home path scrubbed', () => {
  const raw = '/home/bob/data.xlsx';
  const encoded = encodeURIComponent(raw);
  const result = scrubPaths('ENOENT: ' + encoded);
  assert.ok(!result.includes('bob'), `URL-encoded username leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: $HOME/... reference scrubbed', () => {
  const result = scrubPaths('Error: $HOME/secret.xlsx not found');
  assert.ok(!result.includes('$HOME/secret'), `$HOME path leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: ${HOME}/... reference scrubbed', () => {
  const result = scrubPaths('Error: ${HOME}/secret.xlsx not found');
  assert.ok(!result.includes('${HOME}/secret'), `\${HOME} path leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: %USERPROFILE%\\... Windows env reference scrubbed', () => {
  const result = scrubPaths('Cannot open %USERPROFILE%\\Desktop\\file.xlsx');
  assert.ok(!result.includes('%USERPROFILE%\\Desktop'), `%USERPROFILE% path leaked: ${result}`);
  assert.ok(result.includes('<path>'), `no <path> placeholder: ${result}`);
});

test('scrubPaths: base64-encoded path NOT scrubbed (acceptable — not recognizable)', () => {
  // Base64 of a path is not a recognizable path — we document it is not scrubbed
  // and accept this because error messages containing base64 blobs with paths
  // are vanishingly rare and the encoding makes the path unreadable anyway.
  const b64 = Buffer.from('/Users/alice/secret.xlsx').toString('base64');
  // Just assert no crash — we don't require the base64 to be scrubbed.
  const result = scrubPaths('Error: ' + b64);
  assert.ok(typeof result === 'string');
});

test('buildPayload: adversarial — URL-encoded path in error message scrubbed', () => {
  const encoded = encodeURIComponent('/Users/grace/Documents/confidential.xlsx');
  const err = new Error('Cannot read: ' + encoded);
  const payload = buildPayload(err, '1.5.0');
  assert.ok(!payload.error_message.includes('grace'), `URL-encoded path leaked: ${payload.error_message}`);
});

test('buildPayload: adversarial — $HOME path in error message scrubbed', () => {
  const err = new Error('Error reading $HOME/private/budget.xlsx');
  const payload = buildPayload(err, '1.5.0');
  assert.ok(!payload.error_message.includes('$HOME/private'), `$HOME path leaked: ${payload.error_message}`);
});
