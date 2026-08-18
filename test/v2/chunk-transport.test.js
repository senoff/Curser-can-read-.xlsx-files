'use strict';

// XLS-958 leg-3 — client chunked upload/download transport.
//
// These drive the REAL client.post() path with `fetch` stubbed, so they pin the
// wire contract (routes hit, chunk math, reassembly) end-to-end, not just a
// mocked-out helper. The assertion that matters most is byte-identity: what
// uploadWorkbook streams out must come back from downloadWorkbook unchanged.

const { test } = require('node:test');
const assert = require('node:assert');
const path = require('node:path');
const crypto = require('node:crypto');

const CLIENT_PATH = path.join(__dirname, '..', '..', 'lib', 'client.js');
const TRANSPORT_PATH = path.join(__dirname, '..', '..', 'lib', 'chunk-transport.js');
const CONFIG_PATH = require.resolve('../../lib/config');

function freshTransport() {
  delete require.cache[CLIENT_PATH];
  delete require.cache[CONFIG_PATH];
  delete require.cache[TRANSPORT_PATH];
  return require('../../lib/chunk-transport.js');
}

function jsonResponse(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

test('uploadWorkbook streams the buffer in chunks, finalizes, returns the handle', async () => {
  process.env.XFA_UPLOAD_CHUNK_BYTES = '1024'; // force multiple chunks on a small buffer
  const { uploadWorkbook } = freshTransport();

  const original = crypto.randomBytes(3000); // 3 chunks of 1024 (1024 + 1024 + 952)
  const store = {}; // handle -> [chunk_b64...]
  const wire = [];
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    const u = String(url);
    const body = JSON.parse(init.body);
    wire.push(u);
    if (u.endsWith('/api/v1/cache/upload-chunk')) {
      store[body.handle] = store[body.handle] || [];
      store[body.handle][body.chunk_index] = body.chunk_b64;
      const received = store[body.handle].filter((x) => x !== undefined).length;
      return jsonResponse({ handle: body.handle, received, total: body.total_chunks, complete: received === body.total_chunks });
    }
    if (u.endsWith('/api/v1/cache/finalize')) {
      const bytes = store[body.handle].reduce((n, b64) => n + Buffer.from(b64, 'base64').length, 0);
      return jsonResponse({ handle: body.handle, size_bytes: bytes });
    }
    throw new Error(`unexpected route ${u}`);
  };

  try {
    const handle = await uploadWorkbook(original);
    assert.ok(handle && typeof handle === 'string', 'a handle string is returned');
    const uploads = wire.filter((u) => u.endsWith('/upload-chunk')).length;
    const finals = wire.filter((u) => u.endsWith('/finalize')).length;
    assert.equal(uploads, 3, 'exactly 3 chunks uploaded');
    assert.equal(finals, 1, 'exactly one finalize');
    // The stored chunks reassemble to the original bytes.
    const reassembled = Buffer.concat(store[handle].map((b64) => Buffer.from(b64, 'base64')));
    assert.ok(reassembled.equals(original), 'stored chunks reassemble byte-identically');
  } finally {
    globalThis.fetch = realFetch;
    delete process.env.XFA_UPLOAD_CHUNK_BYTES;
  }
});

test('downloadWorkbook walks the chunk grid to is_last and reassembles byte-identically', async () => {
  const { downloadWorkbook } = freshTransport();

  const original = crypto.randomBytes(5000);
  const GRID = 2048;
  const total = Math.ceil(original.length / GRID); // 3
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    const u = String(url);
    const body = JSON.parse(init.body);
    assert.ok(u.endsWith('/api/v1/cache/download-chunk'), 'only download-chunk is called');
    const start = body.chunk_index * GRID;
    const slice = original.subarray(start, start + GRID);
    return jsonResponse({
      handle: body.handle,
      size_bytes: original.length,
      chunk_index: body.chunk_index,
      total_chunks: total,
      chunk_b64: slice.toString('base64'),
      is_last: body.chunk_index === total - 1,
    });
  };

  try {
    const got = await downloadWorkbook('some-handle');
    assert.ok(got.equals(original), 'reassembled download is byte-identical to the source');
  } finally {
    globalThis.fetch = realFetch;
  }
});

test('uploadWorkbook fails loud when the server received-count desyncs', async () => {
  process.env.XFA_UPLOAD_CHUNK_BYTES = '1024';
  const { uploadWorkbook } = freshTransport();
  const original = crypto.randomBytes(3000);
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    const u = String(url);
    const body = JSON.parse(init.body);
    if (u.endsWith('/upload-chunk')) {
      // Lie: always report received=1, so chunk 1 (expecting received=2) trips the guard.
      return jsonResponse({ handle: body.handle, received: 1, total: body.total_chunks, complete: false });
    }
    return jsonResponse({ handle: body.handle, size_bytes: 3000 });
  };
  try {
    await assert.rejects(uploadWorkbook(original), (e) => e.code === 'CHUNK_UPLOAD_DESYNC');
  } finally {
    globalThis.fetch = realFetch;
    delete process.env.XFA_UPLOAD_CHUNK_BYTES;
  }
});

test('uploadWorkbook rejects a finalize size mismatch (truncation guard)', async () => {
  const { uploadWorkbook } = freshTransport();
  const original = crypto.randomBytes(2000);
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    const u = String(url);
    const body = JSON.parse(init.body);
    if (u.endsWith('/upload-chunk')) {
      return jsonResponse({ handle: body.handle, received: 1, total: 1, complete: true });
    }
    // Report a wrong size.
    return jsonResponse({ handle: body.handle, size_bytes: 999 });
  };
  try {
    await assert.rejects(uploadWorkbook(original), (e) => e.code === 'FINALIZE_SIZE_MISMATCH');
  } finally {
    globalThis.fetch = realFetch;
  }
});

test('downloadWorkbook stops with an error if is_last never arrives within total_chunks', async () => {
  const { downloadWorkbook } = freshTransport();
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    const body = JSON.parse(init.body);
    // Advertise total_chunks=2 but never set is_last → overrun guard must fire.
    return jsonResponse({
      handle: body.handle,
      size_bytes: 100,
      chunk_index: body.chunk_index,
      total_chunks: 2,
      chunk_b64: Buffer.from('xx').toString('base64'),
      is_last: false,
    });
  };
  try {
    await assert.rejects(downloadWorkbook('h'), (e) => e.code === 'CHUNK_DOWNLOAD_OVERRUN');
  } finally {
    globalThis.fetch = realFetch;
  }
});

test('maybeUploadForRead: a big non-evaluate read uploads and returns an xlsx_read_handle body', async () => {
  process.env.XFA_CHUNK_THRESHOLD_BYTES = '1024'; // small threshold so a modest buffer trips it
  process.env.XFA_UPLOAD_CHUNK_BYTES = '4096';
  const { maybeUploadForRead } = freshTransport();
  const bytes = crypto.randomBytes(5000); // > 1024 threshold
  const fileB64 = bytes.toString('base64');
  const realFetch = globalThis.fetch;
  const wire = [];
  globalThis.fetch = async (url, init) => {
    const u = String(url);
    const body = JSON.parse(init.body);
    wire.push(u);
    if (u.endsWith('/upload-chunk')) return jsonResponse({ handle: body.handle, received: body.chunk_index + 1, total: body.total_chunks, complete: false });
    if (u.endsWith('/finalize')) return jsonResponse({ handle: body.handle, size_bytes: bytes.length });
    throw new Error(`unexpected ${u}`);
  };
  try {
    const out = await maybeUploadForRead(fileB64, { sheet: 'S1', format: 'json', evaluate: false });
    assert.ok(out && typeof out.workbook_handle === 'string', 'returns a workbook_handle body');
    assert.deepEqual(out.options, { sheet: 'S1', format: 'json' }, 'carries sheet+format (no evaluate)');
    assert.ok(wire.some((u) => u.endsWith('/finalize')), 'the bytes were uploaded + finalized');
  } finally {
    globalThis.fetch = realFetch;
    delete process.env.XFA_CHUNK_THRESHOLD_BYTES;
    delete process.env.XFA_UPLOAD_CHUNK_BYTES;
  }
});

test('maybeUploadForRead: a small read stays inline (returns null, no upload)', async () => {
  const { maybeUploadForRead } = freshTransport(); // default ~15MB threshold
  const fileB64 = crypto.randomBytes(2000).toString('base64');
  const realFetch = globalThis.fetch;
  let called = false;
  globalThis.fetch = async () => { called = true; return jsonResponse({}); };
  try {
    const out = await maybeUploadForRead(fileB64, { format: 'md' });
    assert.equal(out, null, 'small file returns null → caller sends inline xlsx_read');
    assert.equal(called, false, 'no upload happened for a small file');
  } finally {
    globalThis.fetch = realFetch;
  }
});

test('maybeUploadForRead: a big read WITH evaluate stays inline (evaluate is not on the handle route)', async () => {
  process.env.XFA_CHUNK_THRESHOLD_BYTES = '1024';
  const { maybeUploadForRead } = freshTransport();
  const fileB64 = crypto.randomBytes(5000).toString('base64');
  const realFetch = globalThis.fetch;
  let called = false;
  globalThis.fetch = async () => { called = true; return jsonResponse({}); };
  try {
    const out = await maybeUploadForRead(fileB64, { evaluate: true });
    assert.equal(out, null, 'evaluate reads stay inline even when big');
    assert.equal(called, false, 'no upload for an evaluate read');
  } finally {
    globalThis.fetch = realFetch;
    delete process.env.XFA_CHUNK_THRESHOLD_BYTES;
  }
});

test('shouldUseHandle honors the threshold: default keeps small inline, dial@0 forces handle', async () => {
  const t = freshTransport();
  // Default (~15MB): a 1MB file stays inline, a 30MB file uses a handle.
  assert.equal(t.shouldUseHandle(1 * 1024 * 1024), false, 'small file inline by default');
  assert.equal(t.shouldUseHandle(30 * 1024 * 1024), true, 'big file uses a handle by default');

  process.env.XFA_CHUNK_THRESHOLD_BYTES = '0';
  const t0 = freshTransport();
  assert.equal(t0.shouldUseHandle(1), true, 'dial@0 forces the handle flow for any non-empty file');
  assert.equal(t0.shouldUseHandle(0), false, 'a zero-byte file still stays inline at dial@0');
  delete process.env.XFA_CHUNK_THRESHOLD_BYTES;
});
