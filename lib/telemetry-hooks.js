'use strict';

/**
 * Process-level crash telemetry hooks for xlsx-for-ai.
 *
 * Registers uncaughtException + unhandledRejection handlers only when the user
 * has opted in (isTelemetryActive() === true). On crash, sends a minimal,
 * sanitized payload and then re-throws the original error so the user still
 * sees the stack trace and gets a non-zero exit code.
 *
 * Endpoint: XLSX_FOR_AI_TELEMETRY_ENDPOINT env var if set, else default below.
 * // Endpoint deployment tracked separately — see project memory
 * // project_xlsx_for_ai_telemetry_endpoint.md (TBD).
 */

const https  = require('https');
const http   = require('http');
const { URL } = require('url');

const { isTelemetryActive, telemetryStatus } = require('./telemetry-config');
const { buildPayload } = require('./telemetry-sanitize');

const SEND_TIMEOUT_MS = 2000;

// Endpoint deployment tracked separately — see project memory
// project_xlsx_for_ai_telemetry_endpoint.md (TBD).
const DEFAULT_ENDPOINT = 'https://telemetry.xlsx-for-ai.dev/v1/crash';

function resolveEndpoint() {
  return process.env.XLSX_FOR_AI_TELEMETRY_ENDPOINT || DEFAULT_ENDPOINT;
}

/**
 * Send payload to the telemetry endpoint. Returns a Promise that:
 *   - resolves on success (2xx)
 *   - resolves (with a warning) on non-2xx or send failure
 *   - resolves on timeout (after SEND_TIMEOUT_MS)
 *
 * The Promise ALWAYS resolves — never rejects. A hung send must not block exit.
 */
function sendPayload(payload) {
  return new Promise((resolve) => {
    const body = JSON.stringify(payload);
    const endpoint = resolveEndpoint();

    let parsed;
    try {
      parsed = new URL(endpoint);
    } catch (_) {
      resolve();
      return;
    }

    const transport = parsed.protocol === 'http:' ? http : https;
    const options = {
      hostname: parsed.hostname,
      port: parsed.port || (parsed.protocol === 'http:' ? 80 : 443),
      path: parsed.pathname + parsed.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body),
      },
    };

    const timer = setTimeout(() => {
      try { req.destroy(); } catch (_) { /* ignore */ }
      resolve();
    }, SEND_TIMEOUT_MS);

    const req = transport.request(options, (res) => {
      clearTimeout(timer);
      // Drain response body to free the socket.
      res.resume();
      res.on('end', resolve);
      res.on('error', resolve);
    });

    req.on('error', () => {
      clearTimeout(timer);
      resolve();
    });

    req.write(body);
    req.end();
  });
}

/**
 * Register process-level crash handlers if telemetry is active.
 * Call once at startup. No-op if telemetry is not enabled.
 *
 * Prints a one-line notice if telemetry was opted in but consent_version is stale.
 */
function registerCrashHooks(version) {
  const status = telemetryStatus();

  if (status === 'paused (consent_version mismatch)') {
    process.stderr.write(
      'xlsx-for-ai: telemetry has been updated. Run `xlsx-for-ai --enable-telemetry`' +
      ' to resume on the new shape, or `--telemetry-status` for details.\n'
    );
    return;
  }

  if (status !== 'enabled') return;

  async function handleCrash(err) {
    const payload = buildPayload(err, version);
    try {
      await sendPayload(payload);
    } catch (_) {
      // Never let telemetry mask the real error.
    }
    // Re-throw so the original stack + non-zero exit still happens.
    // We use process.exit(1) here because re-throwing from an
    // uncaughtException handler after it fires causes Node to call the
    // handler again, creating an infinite loop.
    process.stderr.write((err && (err.stack || err.message)) ? (err.stack || err.message) + '\n' : String(err) + '\n');
    process.exit(1);
  }

  process.on('uncaughtException', (err) => {
    handleCrash(err);
  });

  process.on('unhandledRejection', (reason) => {
    handleCrash(reason instanceof Error ? reason : new Error(String(reason)));
  });
}

module.exports = {
  registerCrashHooks,
  sendPayload,
  resolveEndpoint,
  DEFAULT_ENDPOINT,
  SEND_TIMEOUT_MS,
};
