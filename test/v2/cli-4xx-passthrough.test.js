'use strict';

// XLS-657 — the CLI must surface a 4xx server validation message inline,
// the same way the MCP path already does (friendlyErrorMessage). Before
// this fix, `friendlyCliError` (index.js) had no `API_CLIENT_ERROR` case,
// so every 4xx collapsed to the generic `request failed (code=API_CLIENT_
// ERROR)` — a caller who typed a wrong sheet name never saw
// `Sheet "Ghost" not found. Available sheets: ...`.
//
// These are REAL-CLI assertions: they spawn index.js against a mock HTTP
// server returning the 4xx (XLSX_FOR_AI_CI=1 forces the no-network
// registration sentinel, so the ONLY network call is the tool POST — no
// live API, deterministic in CI too). The discriminating 5xx case proves
// we did not over-open the boundary, and the path-leak case proves the
// shared sanitizer runs on the CLI surface.

const { test, before, after } = require('node:test');
const assert = require('node:assert');
const { spawn } = require('node:child_process');
const http = require('node:http');
const path = require('node:path');
const os = require('node:os');
const fs = require('node:fs');

const INDEX = path.join(__dirname, '..', '..', 'index.js');

// A tiny workbook fixture — bytes are arbitrary; the CLI base64-encodes and
// ships them without parsing (validation happens server-side, which we mock).
let TMP_XLSX;
let server;
let port;
// Mutable per-test response the mock returns for POST /api/v1/tools/*.
let nextResponse = { status: 200, body: { content: [{ text: 'ok' }] } };

before(async () => {
  TMP_XLSX = path.join(os.tmpdir(), `xls657-fixture-${process.pid}.xlsx`);
  fs.writeFileSync(TMP_XLSX, Buffer.from('PK not-a-real-xlsx-but-extension-passes'));
  server = http.createServer((req, res) => {
    // Drain the body, then answer with whatever the current test queued.
    req.on('data', () => {});
    req.on('end', () => {
      res.writeHead(nextResponse.status, { 'content-type': 'application/json' });
      res.end(JSON.stringify(nextResponse.body));
    });
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  port = server.address().port;
});

after(() => {
  if (server) server.close();
  if (TMP_XLSX) { try { fs.unlinkSync(TMP_XLSX); } catch (_) {} }
});

// MUST be async spawn, not spawnSync: the mock server runs in THIS process,
// and spawnSync would block the event loop so the server could never answer
// the child's request (self-deadlock). Async spawn keeps the loop free.
function runCli(args, extraEnv = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn('node', [INDEX, ...args], {
      env: {
        ...process.env,
        XLSX_FOR_AI_CI: '1',                        // no-network registration sentinel
        XLSX_FOR_AI_API: `http://127.0.0.1:${port}`, // callTool hits the mock
        ...extraEnv,
      },
    });
    let stdout = '', stderr = '';
    const killer = setTimeout(() => { child.kill('SIGKILL'); reject(new Error('CLI timed out')); }, 60_000);
    child.stdout.on('data', (d) => (stdout += d));
    child.stderr.on('data', (d) => (stderr += d));
    child.on('error', (e) => { clearTimeout(killer); reject(e); });
    child.on('close', (code) => { clearTimeout(killer); resolve({ code, stdout, stderr }); });
  });
}

test('4xx sheet-not-found: stderr surfaces the server message, NOT the generic', async () => {
  nextResponse = {
    status: 400,
    body: { error: { code: 'sheet_not_found', message: 'Sheet "Ghost" not found. Available sheets: Sheet1, Sheet2' } },
  };
  const { code, stderr } = await runCli([TMP_XLSX, '--sheet', 'Ghost']);
  assert.ok(
    stderr.includes('Sheet "Ghost" not found. Available sheets: Sheet1, Sheet2'),
    `expected the inline server validation message on stderr; got: ${stderr}`,
  );
  assert.ok(
    !stderr.includes('request failed (code=API_CLIENT_ERROR)'),
    `4xx must NOT collapse to the generic message; got: ${stderr}`,
  );
  assert.equal(code, 1, `4xx is a caller error → exit 1; got ${code}`);
});

test('4xx with an absolute path in the message: path is scrubbed, not leaked', async () => {
  nextResponse = {
    status: 400,
    body: { error: { message: 'validation failed reading /Users/bob/Desktop/secret.xlsx at row 3' } },
  };
  const { stderr } = await runCli([TMP_XLSX, '--sheet', 'X']);
  assert.ok(!stderr.includes('/Users/bob/Desktop/secret.xlsx'),
    `absolute path must be scrubbed from CLI stderr; got: ${stderr}`);
  assert.ok(stderr.includes('<path>'),
    `expected the <path> placeholder; got: ${stderr}`);
  // The caller-actionable shape survives.
  assert.ok(stderr.includes('validation failed reading') && stderr.includes('at row 3'),
    `actionable signal should survive scrubbing; got: ${stderr}`);
});

test('5xx stays generic (discriminating case — boundary not over-opened)', async () => {
  nextResponse = {
    status: 500,
    body: { error: { message: 'internal stack trace at server.js:42' } },
  };
  const { code, stderr } = await runCli([TMP_XLSX, '--sheet', 'X']);
  assert.ok(!stderr.includes('internal stack trace'),
    `5xx must NOT surface payload internals on the CLI; got: ${stderr}`);
  assert.ok(stderr.includes('server error'),
    `5xx should show the generic server-error text; got: ${stderr}`);
  assert.equal(code, 3, `5xx → exit 3 (retryable); got ${code}`);
});

test('XFA_DEBUG=1 still appends the raw underlying message for 4xx', async () => {
  nextResponse = {
    status: 400,
    body: { error: { message: 'Sheet "Ghost" not found. Available sheets: Sheet1' } },
  };
  const { stderr } = await runCli([TMP_XLSX, '--sheet', 'Ghost'], { XFA_DEBUG: '1' });
  assert.ok(stderr.includes('Raw: xlsx-for-ai API error 400:'),
    `XFA_DEBUG=1 should append the raw wrapped message; got: ${stderr}`);
});
