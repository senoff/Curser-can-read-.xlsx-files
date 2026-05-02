'use strict';

/**
 * Persistent user-level telemetry config at ~/.xlsx-for-ai/config.json.
 *
 * Stored outside node_modules so consent survives `npm install -g xlsx-for-ai@latest`
 * upgrades. Path is resolved via os.homedir() for cross-platform support.
 *
 * Config shape:
 *   { "telemetry": true, "consented_at": "ISO-string", "consent_version": 1 }
 *
 * consent_version: bump CURRENT_CONSENT_VERSION when the telemetry shape changes.
 * If the file's version is older, telemetry is PAUSED until the user re-runs
 * --enable-telemetry. Never silently expand data shape under old consent.
 */

const fs   = require('fs');
const path = require('path');
const os   = require('os');

const CURRENT_CONSENT_VERSION = 1;

/**
 * Return the path to the config file. Uses XFA_CONFIG_DIR env var for test
 * isolation; otherwise defaults to ~/.xlsx-for-ai/config.json.
 */
function configDir() {
  return process.env.XFA_CONFIG_DIR || path.join(os.homedir(), '.xlsx-for-ai');
}

function configPath() {
  return path.join(configDir(), 'config.json');
}

/**
 * Read config from disk. Returns null if file doesn't exist or is unreadable.
 */
function readConfig() {
  try {
    const raw = fs.readFileSync(configPath(), 'utf8');
    return JSON.parse(raw);
  } catch (_) {
    return null;
  }
}

/**
 * Write config to disk atomically. Creates the directory if needed.
 */
function writeConfig(data) {
  const dir = configDir();
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(configPath(), JSON.stringify(data, null, 2) + '\n', 'utf8');
}

/**
 * Telemetry status as one of:
 *   'enabled'                  - opt-in, consent_version matches
 *   'disabled'                 - explicitly opted out
 *   'not configured'           - no config file yet
 *   'paused (consent_version mismatch)' - opted in but consent_version is stale
 */
function telemetryStatus() {
  const cfg = readConfig();
  if (!cfg) return 'not configured';
  if (cfg.telemetry === false) return 'disabled';
  if (cfg.telemetry === true) {
    if (cfg.consent_version !== CURRENT_CONSENT_VERSION) {
      return 'paused (consent_version mismatch)';
    }
    return 'enabled';
  }
  return 'not configured';
}

/**
 * Returns true only if telemetry is fully active (opted in AND version matches).
 */
function isTelemetryActive() {
  return telemetryStatus() === 'enabled';
}

/**
 * Enable telemetry — write consent with current version.
 * Idempotent.
 */
function enableTelemetry() {
  const existing = readConfig() || {};
  writeConfig({
    ...existing,
    telemetry: true,
    consented_at: new Date().toISOString(),
    consent_version: CURRENT_CONSENT_VERSION,
  });
}

/**
 * Disable telemetry — write explicit false (keeps the file so we can distinguish
 * "user said no" from "never asked").
 */
function disableTelemetry() {
  const existing = readConfig() || {};
  writeConfig({ ...existing, telemetry: false });
}

module.exports = {
  CURRENT_CONSENT_VERSION,
  configPath,
  readConfig,
  writeConfig,
  telemetryStatus,
  isTelemetryActive,
  enableTelemetry,
  disableTelemetry,
};
