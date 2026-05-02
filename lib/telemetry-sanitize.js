'use strict';

/**
 * Sanitization for crash telemetry payloads.
 *
 * INVARIANTS (non-negotiable):
 * - No file paths: scrub /Users/<x>/..., C:\Users\<x>\..., /home/<x>/...
 * - Cap error_message at 200 chars (after scrubbing)
 * - No cell values, no workbook structure (not available post-crash anyway)
 * - No env vars, no argv beyond a hardcoded allowlist
 * - No machine identifier (no hostname, MAC, install ID)
 *
 * Future maintainers: do NOT enrich this payload. The consent_version gates
 * any shape expansion. Bump CURRENT_CONSENT_VERSION in telemetry-config.js
 * before adding new fields.
 */

const os = require('os');

const MAX_MESSAGE_LENGTH = 200;

// Allowlisted first-arg values for the command field.
// Everything else becomes '<other>'.
const ALLOWED_COMMANDS = new Set([
  'xlsx-for-ai',
  'cursor-reads-xlsx',
  'write',
  '--json',
  '--md',
  '--stdout',
  '--sql',
  '--schema',
  '--compact',
  '--evaluate',
  '--stream',
  '--list-sheets',
  '--diff',
  '--range',
  '--named-range',
  '--max-rows',
  '--max-cols',
  '--max-tokens',
  '--report-bug',
  '--export-redacted-workbook',
  '--enable-telemetry',
  '--disable-telemetry',
  '--telemetry-status',
  '--help',
  '--version',
  '-h',
  '-v',
]);

/**
 * Scrub filesystem paths from a string.
 *
 * Covers:
 *   /Users/<name>/...       (macOS)
 *   /home/<name>/...        (Linux)
 *   C:\Users\<name>\...     (Windows, forward or back slash variants)
 *   /var/folders/...        (macOS temp)
 *   /tmp/...                (Linux/macOS tmp)
 *   /private/tmp/...        (macOS private tmp — e.g. worktrees)
 */
function scrubPaths(str) {
  if (typeof str !== 'string') return str;

  // Windows: C:\Users\<name>\... or C:/Users/<name>/...
  let s = str.replace(
    /[A-Za-z]:[/\\][Uu]sers[/\\][^/\\:\s]+([/\\][^\s]*)*/g,
    '<path>'
  );

  // Unix home dirs: /Users/<name>/... or /home/<name>/...
  s = s.replace(
    /\/(Users|home)\/[^/\s:]+([^\s:])*/g,
    '<path>'
  );

  // /tmp/... and /private/tmp/...
  s = s.replace(
    /\/(?:private\/)?tmp\/[^\s:]+/g,
    '<path>'
  );

  // /var/folders/...
  s = s.replace(
    /\/var\/folders\/[^\s:]+/g,
    '<path>'
  );

  return s;
}

/**
 * Sanitize error message: scrub paths, cap at 200 chars.
 */
function sanitizeMessage(message) {
  if (!message) return '';
  const scrubbed = scrubPaths(String(message));
  return scrubbed.slice(0, MAX_MESSAGE_LENGTH);
}

/**
 * Build the outgoing crash payload from an Error object.
 * Returns a plain object with only the allowed fields.
 */
function buildPayload(err, version) {
  const errorType  = (err && err.constructor && err.constructor.name) || 'Error';
  const rawMessage = (err && err.message) ? err.message : String(err);
  const message    = sanitizeMessage(rawMessage);

  // First arg from process.argv — only if in allowlist.
  let command = '<other>';
  try {
    const firstArg = process.argv[2];
    if (firstArg && ALLOWED_COMMANDS.has(firstArg)) {
      command = firstArg;
    }
  } catch (_) { /* ignore */ }

  return {
    v: 1,
    ts: new Date().toISOString(),
    error_type: errorType,
    error_message: message,
    command,
    xlsx_for_ai_version: version || 'unknown',
    node_version: process.version,
    os_arch: `${process.platform}-${process.arch}`,
  };
}

module.exports = {
  scrubPaths,
  sanitizeMessage,
  buildPayload,
  MAX_MESSAGE_LENGTH,
  ALLOWED_COMMANDS,
};
