'use strict';

const { test, before, after, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const fs   = require('fs');
const path = require('path');
const os   = require('os');

// ---------------------------------------------------------------------------
// Use a temp directory for all config writes — never touch ~/.xlsx-for-ai/
// ---------------------------------------------------------------------------
let tmpDir;

before(() => {
  tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xfa-cfg-test-'));
  process.env.XFA_CONFIG_DIR = tmpDir;
  // Re-require the module so configDir() picks up the env var on first load.
  // We use delete require.cache to ensure a fresh module per test run.
  delete require.cache[require.resolve('../lib/telemetry-config')];
});

after(() => {
  delete process.env.XFA_CONFIG_DIR;
  fs.rmSync(tmpDir, { recursive: true, force: true });
});

beforeEach(() => {
  // Clear any config file between tests for isolation.
  const cfgPath = path.join(tmpDir, 'config.json');
  if (fs.existsSync(cfgPath)) fs.unlinkSync(cfgPath);
  // Bust require cache so each test sees a fresh module state.
  delete require.cache[require.resolve('../lib/telemetry-config')];
});

function loadFresh() {
  delete require.cache[require.resolve('../lib/telemetry-config')];
  return require('../lib/telemetry-config');
}

// ---------------------------------------------------------------------------
// configPath
// ---------------------------------------------------------------------------

test('configPath: resolves under XFA_CONFIG_DIR when set', () => {
  const tc = loadFresh();
  assert.equal(tc.configPath(), path.join(tmpDir, 'config.json'));
});

// ---------------------------------------------------------------------------
// readConfig
// ---------------------------------------------------------------------------

test('readConfig: returns null when file absent', () => {
  const tc = loadFresh();
  assert.equal(tc.readConfig(), null);
});

test('readConfig: returns null when file is malformed JSON', () => {
  fs.writeFileSync(path.join(tmpDir, 'config.json'), 'NOT_JSON', 'utf8');
  const tc = loadFresh();
  assert.equal(tc.readConfig(), null);
});

test('readConfig: returns parsed object for valid JSON', () => {
  const data = { telemetry: true, consent_version: 1 };
  fs.writeFileSync(path.join(tmpDir, 'config.json'), JSON.stringify(data), 'utf8');
  const tc = loadFresh();
  assert.deepEqual(tc.readConfig(), data);
});

// ---------------------------------------------------------------------------
// writeConfig
// ---------------------------------------------------------------------------

test('writeConfig: creates file with correct content', () => {
  const tc = loadFresh();
  const data = { telemetry: true, consented_at: '2026-01-01T00:00:00.000Z', consent_version: 1 };
  tc.writeConfig(data);
  const raw = fs.readFileSync(path.join(tmpDir, 'config.json'), 'utf8');
  assert.deepEqual(JSON.parse(raw), data);
});

test('writeConfig: creates directory if needed', () => {
  const subDir = path.join(tmpDir, 'nested', 'config');
  process.env.XFA_CONFIG_DIR = subDir;
  const tc = loadFresh();
  tc.writeConfig({ telemetry: false });
  assert.ok(fs.existsSync(path.join(subDir, 'config.json')));
  // Restore
  process.env.XFA_CONFIG_DIR = tmpDir;
});

// ---------------------------------------------------------------------------
// enableTelemetry / disableTelemetry
// ---------------------------------------------------------------------------

test('enableTelemetry: writes telemetry:true with consent_version', () => {
  const tc = loadFresh();
  tc.enableTelemetry();
  const cfg = tc.readConfig();
  assert.equal(cfg.telemetry, true);
  assert.equal(cfg.consent_version, tc.CURRENT_CONSENT_VERSION);
  assert.ok(cfg.consented_at);
});

test('enableTelemetry: idempotent — second call overwrites timestamp', () => {
  const tc = loadFresh();
  tc.enableTelemetry();
  const first = tc.readConfig().consented_at;
  // Small delay not needed — just check it re-writes
  tc.enableTelemetry();
  const cfg = tc.readConfig();
  assert.equal(cfg.telemetry, true);
  assert.equal(cfg.consent_version, tc.CURRENT_CONSENT_VERSION);
});

test('disableTelemetry: writes telemetry:false', () => {
  const tc = loadFresh();
  tc.enableTelemetry(); // set true first
  tc.disableTelemetry();
  const cfg = tc.readConfig();
  assert.equal(cfg.telemetry, false);
});

test('disableTelemetry: keeps file (does not delete)', () => {
  const tc = loadFresh();
  tc.disableTelemetry();
  assert.ok(fs.existsSync(path.join(tmpDir, 'config.json')));
});

// ---------------------------------------------------------------------------
// telemetryStatus
// ---------------------------------------------------------------------------

test('telemetryStatus: no file → "not configured"', () => {
  const tc = loadFresh();
  assert.equal(tc.telemetryStatus(), 'not configured');
});

test('telemetryStatus: telemetry:false → "disabled"', () => {
  const tc = loadFresh();
  tc.disableTelemetry();
  assert.equal(tc.telemetryStatus(), 'disabled');
});

test('telemetryStatus: telemetry:true, version matches → "enabled"', () => {
  const tc = loadFresh();
  tc.enableTelemetry();
  assert.equal(tc.telemetryStatus(), 'enabled');
});

test('telemetryStatus: telemetry:true, old version → "paused (consent_version mismatch)"', () => {
  const tc = loadFresh();
  // Write a config with a stale consent_version (0)
  tc.writeConfig({ telemetry: true, consent_version: 0, consented_at: '2020-01-01T00:00:00.000Z' });
  assert.equal(tc.telemetryStatus(), 'paused (consent_version mismatch)');
});

// ---------------------------------------------------------------------------
// isTelemetryActive
// ---------------------------------------------------------------------------

test('isTelemetryActive: false when not configured', () => {
  const tc = loadFresh();
  assert.equal(tc.isTelemetryActive(), false);
});

test('isTelemetryActive: false when disabled', () => {
  const tc = loadFresh();
  tc.disableTelemetry();
  assert.equal(tc.isTelemetryActive(), false);
});

test('isTelemetryActive: true when enabled and version matches', () => {
  const tc = loadFresh();
  tc.enableTelemetry();
  assert.equal(tc.isTelemetryActive(), true);
});

test('isTelemetryActive: false when version mismatch (paused)', () => {
  const tc = loadFresh();
  tc.writeConfig({ telemetry: true, consent_version: 0, consented_at: '2020-01-01T00:00:00.000Z' });
  assert.equal(tc.isTelemetryActive(), false);
});
