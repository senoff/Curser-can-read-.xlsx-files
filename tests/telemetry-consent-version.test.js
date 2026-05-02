'use strict';

/**
 * Tests for consent_version mismatch behavior:
 *   - Old version → telemetry paused
 *   - Re-running --enable-telemetry (enableTelemetry()) writes new version → resumes
 *   - Future version (code is older than config) → paused (treated as mismatch)
 */

const { test, before, after, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const fs   = require('fs');
const path = require('path');
const os   = require('os');

let tmpDir;

before(() => {
  tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xfa-cv-test-'));
  process.env.XFA_CONFIG_DIR = tmpDir;
  delete require.cache[require.resolve('../lib/telemetry-config')];
});

after(() => {
  delete process.env.XFA_CONFIG_DIR;
  fs.rmSync(tmpDir, { recursive: true, force: true });
});

beforeEach(() => {
  const cfgPath = path.join(tmpDir, 'config.json');
  if (fs.existsSync(cfgPath)) fs.unlinkSync(cfgPath);
  delete require.cache[require.resolve('../lib/telemetry-config')];
});

function loadFresh() {
  delete require.cache[require.resolve('../lib/telemetry-config')];
  return require('../lib/telemetry-config');
}

test('consent_version 0 (stale) → status is paused', () => {
  const tc = loadFresh();
  tc.writeConfig({ telemetry: true, consent_version: 0, consented_at: '2020-01-01T00:00:00.000Z' });
  assert.equal(tc.telemetryStatus(), 'paused (consent_version mismatch)');
  assert.equal(tc.isTelemetryActive(), false);
});

test('consent_version mismatch → re-running enableTelemetry() resumes', () => {
  const tc = loadFresh();
  // Simulate old consent version
  tc.writeConfig({ telemetry: true, consent_version: 0, consented_at: '2020-01-01T00:00:00.000Z' });
  assert.equal(tc.isTelemetryActive(), false);

  // User re-runs --enable-telemetry
  tc.enableTelemetry();
  const tc2 = loadFresh();
  assert.equal(tc2.telemetryStatus(), 'enabled');
  assert.equal(tc2.isTelemetryActive(), true);
  assert.equal(tc2.readConfig().consent_version, tc2.CURRENT_CONSENT_VERSION);
});

test('consent_version matches current → telemetry active', () => {
  const tc = loadFresh();
  tc.enableTelemetry();
  const cfg = tc.readConfig();
  assert.equal(cfg.consent_version, tc.CURRENT_CONSENT_VERSION);
  assert.equal(tc.isTelemetryActive(), true);
});

test('future version in config file → treated as mismatch (paused)', () => {
  // If someone manually set consent_version to a future value, it mismatches.
  const tc = loadFresh();
  const futureVersion = tc.CURRENT_CONSENT_VERSION + 999;
  tc.writeConfig({ telemetry: true, consent_version: futureVersion, consented_at: '2030-01-01T00:00:00.000Z' });
  assert.equal(tc.telemetryStatus(), 'paused (consent_version mismatch)');
  assert.equal(tc.isTelemetryActive(), false);
});

test('disable after consent_version mismatch → status changes to disabled', () => {
  const tc = loadFresh();
  tc.writeConfig({ telemetry: true, consent_version: 0, consented_at: '2020-01-01T00:00:00.000Z' });
  tc.disableTelemetry();
  const tc2 = loadFresh();
  assert.equal(tc2.telemetryStatus(), 'disabled');
});

test('CURRENT_CONSENT_VERSION is a positive integer', () => {
  const tc = loadFresh();
  assert.equal(typeof tc.CURRENT_CONSENT_VERSION, 'number');
  assert.ok(tc.CURRENT_CONSENT_VERSION >= 1);
  assert.equal(Math.floor(tc.CURRENT_CONSENT_VERSION), tc.CURRENT_CONSENT_VERSION);
});
