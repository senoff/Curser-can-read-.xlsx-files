'use strict';

/**
 * XLS-958 leg-3 — client-side chunked upload / download transport for the
 * big-file fast path.
 *
 * A single inline request tops out at the server body cap (~20 MB), and base64
 * inflates bytes ~1.33×, so a workbook larger than ~15 MB raw cannot be sent as
 * `file_b64` at all — it 413s before it does any work. This module streams such
 * a workbook to the server-side cache in bounded chunks and hands back an opaque
 * handle, so a handle-aware tool (today: `xlsx_read_handle`) can operate on the
 * bytes without them ever crossing the wire whole. Download is the inverse: walk
 * the server's chunk grid back and reassemble byte-identically.
 *
 * Server contract (frozen; XLS-958 server legs, live on api.xlsx-for-ai.dev):
 *   POST /api/v1/cache/upload-chunk
 *        {handle, chunk_index (0..N-1), total_chunks, chunk_b64}
 *        -> {handle, received, total, complete}
 *   POST /api/v1/cache/finalize
 *        {handle}
 *        -> {handle, size_bytes}
 *   POST /api/v1/cache/download-chunk
 *        {handle, chunk_index}
 *        -> {handle, size_bytes, chunk_index, total_chunks, chunk_b64, is_last}
 *
 * The client picks the handle (a UUID); the server keys the cache on it and
 * scopes it to the authenticated client. Every call goes through client.post(),
 * so it inherits auth, the retry/timeout ceiling, privacy headers, and the
 * X-XFA-Client-Rtt-Ms telemetry uniformly.
 */

const crypto = require('crypto');
const { post } = require('./client');

// Raw bytes per upload chunk. The server accepts <=16 MB raw/chunk; 8 MB raw
// (~10.7 MB once base64'd in the JSON body, comfortably under the 20 MB body
// cap) leaves headroom for the request envelope. Env-tunable to trade round-trip
// count against peak per-request memory; clamped to the server's 16 MB ceiling.
const DEFAULT_UPLOAD_CHUNK_BYTES = 8 * 1024 * 1024;
const MAX_UPLOAD_CHUNK_BYTES = 16 * 1024 * 1024;

function uploadChunkBytes() {
  const raw = parseInt(process.env.XFA_UPLOAD_CHUNK_BYTES, 10);
  if (Number.isFinite(raw) && raw > 0) return Math.min(raw, MAX_UPLOAD_CHUNK_BYTES);
  return DEFAULT_UPLOAD_CHUNK_BYTES;
}

// Decoded-byte size at or below which a workbook is small enough to send inline;
// above it, callers should route through the handle flow. Default is the inline
// ceiling (~15 MB raw ≈ 20 MB base64), i.e. "use a handle only when inline would
// not fit" — which is the safe, non-lossy default: the handle read route
// (xlsx_read_handle) does not carry every inline option, so redirecting a small
// read that fits inline would be pure cost. Set XFA_CHUNK_THRESHOLD_BYTES=0 to
// force the handle flow always (dial@0), or a larger value to raise the wall.
const DEFAULT_CHUNK_THRESHOLD_BYTES = 15 * 1024 * 1024;

function chunkThresholdBytes() {
  const raw = process.env.XFA_CHUNK_THRESHOLD_BYTES;
  if (raw === undefined || raw === '') return DEFAULT_CHUNK_THRESHOLD_BYTES;
  const n = parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 0) return DEFAULT_CHUNK_THRESHOLD_BYTES;
  return n;
}

/** True when `byteLength` decoded bytes should go through the handle flow. */
function shouldUseHandle(byteLength) {
  return byteLength > chunkThresholdBytes();
}

/**
 * Upload a workbook to the server cache in chunks and return its handle.
 *
 * @param {Buffer} buffer  the raw workbook bytes
 * @returns {Promise<string>} the finalized cache handle
 */
async function uploadWorkbook(buffer) {
  if (!Buffer.isBuffer(buffer)) {
    const e = new Error('uploadWorkbook expects a Buffer of workbook bytes');
    e.code = 'CHUNK_UPLOAD_BAD_INPUT';
    throw e;
  }
  const handle = crypto.randomUUID();
  const chunkSize = uploadChunkBytes();
  // A zero-byte buffer still finalizes as one (empty) chunk so the server's
  // total_chunks >= 1 contract holds; an empty workbook is invalid downstream
  // anyway, but the transport must not produce total_chunks = 0.
  const totalChunks = Math.max(1, Math.ceil(buffer.length / chunkSize));

  for (let i = 0; i < totalChunks; i += 1) {
    const slice = buffer.subarray(i * chunkSize, (i + 1) * chunkSize);
    const res = await post('/api/v1/cache/upload-chunk', {
      handle,
      chunk_index: i,
      total_chunks: totalChunks,
      chunk_b64: slice.toString('base64'),
    });
    // Sending chunks in order, the server's running `received` count must equal
    // the number of chunks sent so far (i + 1). A mismatch means a chunk was
    // dropped or double-counted — fail loud rather than finalize a blob that is
    // silently short.
    if (res && typeof res.received === 'number' && res.received !== i + 1) {
      const e = new Error(
        `chunk upload desynced: after chunk ${i} the server reports received=${res.received}, expected ${i + 1}`,
      );
      e.code = 'CHUNK_UPLOAD_DESYNC';
      throw e;
    }
  }

  const fin = await post('/api/v1/cache/finalize', { handle });
  // The finalized size is the server's own tally of the bytes it stored; if it
  // disagrees with what we uploaded, the blob is truncated/corrupt and must not
  // be handed on as a valid handle.
  if (fin && typeof fin.size_bytes === 'number' && fin.size_bytes !== buffer.length) {
    const e = new Error(
      `finalize size mismatch: uploaded ${buffer.length} bytes, server finalized ${fin.size_bytes}`,
    );
    e.code = 'FINALIZE_SIZE_MISMATCH';
    throw e;
  }
  return handle;
}

/**
 * Download a finalized workbook by handle, walking the server's chunk grid until
 * `is_last`, and return the reassembled bytes — byte-identical to what was
 * uploaded.
 *
 * @param {string} handle
 * @returns {Promise<Buffer>}
 */
async function downloadWorkbook(handle) {
  if (typeof handle !== 'string' || handle.length === 0) {
    const e = new Error('downloadWorkbook expects a non-empty handle string');
    e.code = 'CHUNK_DOWNLOAD_BAD_INPUT';
    throw e;
  }
  const parts = [];
  // Bound the walk by the server-advertised total so a server that never sets
  // is_last cannot spin us forever. Seeded to Infinity until the first response
  // tells us total_chunks.
  let guardMax = Infinity;
  for (let index = 0; ; index += 1) {
    const res = await post('/api/v1/cache/download-chunk', { handle, chunk_index: index });
    if (!res || typeof res.chunk_b64 !== 'string') {
      const e = new Error(`download-chunk returned no chunk_b64 for index ${index}`);
      e.code = 'CHUNK_DOWNLOAD_EMPTY';
      throw e;
    }
    parts.push(Buffer.from(res.chunk_b64, 'base64'));
    if (Number.isFinite(res.total_chunks) && res.total_chunks > 0) guardMax = res.total_chunks;
    if (res.is_last) break;
    if (index + 1 >= guardMax) {
      const e = new Error(
        `download-chunk walk reached total_chunks (${guardMax}) without is_last for handle`,
      );
      e.code = 'CHUNK_DOWNLOAD_OVERRUN';
      throw e;
    }
  }
  return Buffer.concat(parts);
}

module.exports = {
  uploadWorkbook,
  downloadWorkbook,
  uploadChunkBytes,
  chunkThresholdBytes,
  shouldUseHandle,
  DEFAULT_UPLOAD_CHUNK_BYTES,
  DEFAULT_CHUNK_THRESHOLD_BYTES,
};
