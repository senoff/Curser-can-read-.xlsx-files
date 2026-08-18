'use strict';

// XLS-958 leg-3 — client_rtt_ms piggyback.
//
// The client remembers the user-felt round-trip of each successful call and
// stamps the PREVIOUS one onto the NEXT request as X-XFA-Client-Rtt-Ms, so the
// server can log what the user actually waited. Pins: (1) the first call carries
// no header (nothing measured yet); (2) the second call carries an integer-ms
// header; (3) strict privacy suppresses it, same lane as X-MCP-Client-*.

const { test } = require('node:test');
const assert = require('node:assert');
const http = require('node:http');
const path = require('node:path');

const CLIENT_PATH = path.join(__dirname, '..', '..', 'lib', 'client.js');
const CONFIG_PATH = require.resolve('../../lib/config');

function freshClient() {
  delete require.cache[CLIENT_PATH];
  delete require.cache[CONFIG_PATH];
  return require('../../lib/client.js');
}

function startServer(seenHeaders) {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      seenHeaders.push(req.headers);
      res.writeHead(200, { 'content-type': 'application/json' });
      res.end(JSON.stringify({ ok: true }));
    });
    server.listen(0, '127.0.0.1', () => resolve({ server, port: server.address().port }));
  });
}

test('first call omits the RTT header; the second call carries the first call RTT (integer ms)', async () => {
  const seen = [];
  const { server, port } = await startServer(seen);
  process.env.XLSX_FOR_AI_API = `http://127.0.0.1:${port}`;
  const { post } = freshClient();
  try {
    await post('/api/v1/tools/xlsx_list_sheets', { file_b64: 'x' }, { auth: false });
    await post('/api/v1/tools/xlsx_list_sheets', { file_b64: 'y' }, { auth: false });

    assert.equal(seen.length, 2, 'both calls reached the server');
    assert.equal(seen[0]['x-xfa-client-rtt-ms'], undefined, 'first request has no prior RTT to report');
    const rtt = seen[1]['x-xfa-client-rtt-ms'];
    assert.ok(rtt !== undefined, 'second request carries the RTT header');
    assert.ok(/^\d+$/.test(rtt), `RTT header is integer ms, got "${rtt}"`);
  } finally {
    server.close();
    delete process.env.XLSX_FOR_AI_API;
  }
});

test('strict privacy suppresses the RTT header', async () => {
  const seen = [];
  const { server, port } = await startServer(seen);
  process.env.XLSX_FOR_AI_API = `http://127.0.0.1:${port}`;
  const { post } = freshClient();
  try {
    // First call (non-strict) establishes a measured RTT.
    await post('/api/v1/tools/xlsx_list_sheets', { file_b64: 'x' }, { auth: false });
    // Second call strict — must NOT carry the header even though one is available.
    await post('/api/v1/tools/xlsx_list_sheets', { file_b64: 'y' }, { auth: false, privacyStrict: true });

    assert.equal(seen[1]['x-xfa-client-rtt-ms'], undefined, 'strict privacy drops the RTT header');
  } finally {
    server.close();
    delete process.env.XLSX_FOR_AI_API;
  }
});
