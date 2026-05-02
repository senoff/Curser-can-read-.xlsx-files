'use strict';

/**
 * Tests for CLI flags: --enable-telemetry, --disable-telemetry, --telemetry-status.
 *
 * These tests exercise the CLI subprocess via node child_process.spawnSync so
 * we can verify both stdout content and config-file side effects. HTTP is never
 * invoked in these tests.
 */

const { test, before, after, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const fs   = require('fs');
const path = require('path');
const os   = require('os');
const { spawnSync } = require('child_process');

const INDEX = path.resolve(__dirname, '../index.js');

let tmpDir;

before(() => {
  tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xfa-flags-test-'));
});

after(() => {
  fs.rmSync(tmpDir, { recursive: true, force: true });
});

beforeEach(() => {
  // Clear config between tests
  const cfgPath = path.join(tmpDir, 'config.json');
  if (fs.existsSync(cfgPath)) fs.unlinkSync(cfgPath);
});

function run(args) {
  return spawnSync(process.execPath, [INDEX, ...args], {
    encoding: 'utf8',
    env: { ...process.env, XFA_CONFIG_DIR: tmpDir },
  });
}

function readConfig() {
  const cfgPath = path.join(tmpDir, 'config.json');
  if (!fs.existsSync(cfgPath)) return null;
  return JSON.parse(fs.readFileSync(cfgPath, 'utf8'));
}

// ---------------------------------------------------------------------------
// --enable-telemetry
// ---------------------------------------------------------------------------

test('--enable-telemetry: exits 0', () => {
  const r = run(['--enable-telemetry']);
  assert.equal(r.status, 0, `stderr: ${r.stderr}`);
});

test('--enable-telemetry: prints confirmation line', () => {
  const r = run(['--enable-telemetry']);
  assert.ok(r.stdout.includes('enabled'), `stdout: ${r.stdout}`);
});

test('--enable-telemetry: prints payload schema', () => {
  const r = run(['--enable-telemetry']);
  // Must show key payload fields so users see what gets sent
  assert.ok(r.stdout.includes('error_type'), `missing error_type in: ${r.stdout}`);
  assert.ok(r.stdout.includes('error_message'), `missing error_message in: ${r.stdout}`);
  assert.ok(r.stdout.includes('xlsx_for_ai_version'), `missing version in: ${r.stdout}`);
  assert.ok(r.stdout.includes('node_version'), `missing node_version in: ${r.stdout}`);
  assert.ok(r.stdout.includes('os_arch'), `missing os_arch in: ${r.stdout}`);
});

test('--enable-telemetry: writes config with telemetry:true', () => {
  run(['--enable-telemetry']);
  const cfg = readConfig();
  assert.ok(cfg, 'config file not created');
  assert.equal(cfg.telemetry, true);
});

test('--enable-telemetry: config has consent_version', () => {
  run(['--enable-telemetry']);
  const cfg = readConfig();
  assert.ok('consent_version' in cfg, 'consent_version missing');
  assert.equal(typeof cfg.consent_version, 'number');
});

test('--enable-telemetry: config has consented_at ISO timestamp', () => {
  run(['--enable-telemetry']);
  const cfg = readConfig();
  assert.ok(cfg.consented_at, 'consented_at missing');
  // Should be a parseable ISO date
  const d = new Date(cfg.consented_at);
  assert.ok(!isNaN(d.getTime()), `invalid date: ${cfg.consented_at}`);
});

test('--enable-telemetry: idempotent — second call still exits 0', () => {
  run(['--enable-telemetry']);
  const r = run(['--enable-telemetry']);
  assert.equal(r.status, 0);
  const cfg = readConfig();
  assert.equal(cfg.telemetry, true);
});

test('--enable-telemetry: prints config path', () => {
  const r = run(['--enable-telemetry']);
  assert.ok(r.stdout.includes(tmpDir), `config path not shown: ${r.stdout}`);
});

// ---------------------------------------------------------------------------
// --disable-telemetry
// ---------------------------------------------------------------------------

test('--disable-telemetry: exits 0', () => {
  const r = run(['--disable-telemetry']);
  assert.equal(r.status, 0, `stderr: ${r.stderr}`);
});

test('--disable-telemetry: prints confirmation', () => {
  const r = run(['--disable-telemetry']);
  assert.ok(r.stdout.toLowerCase().includes('disabled'), `stdout: ${r.stdout}`);
});

test('--disable-telemetry: writes telemetry:false', () => {
  run(['--enable-telemetry']);
  run(['--disable-telemetry']);
  const cfg = readConfig();
  assert.equal(cfg.telemetry, false);
});

test('--disable-telemetry: keeps config file (does not delete)', () => {
  run(['--enable-telemetry']);
  run(['--disable-telemetry']);
  assert.ok(fs.existsSync(path.join(tmpDir, 'config.json')));
});

test('--disable-telemetry: can disable without prior enable', () => {
  const r = run(['--disable-telemetry']);
  assert.equal(r.status, 0);
  const cfg = readConfig();
  assert.equal(cfg.telemetry, false);
});

// ---------------------------------------------------------------------------
// --telemetry-status
// ---------------------------------------------------------------------------

test('--telemetry-status: exits 0', () => {
  const r = run(['--telemetry-status']);
  assert.equal(r.status, 0);
});

test('--telemetry-status: shows "not configured" when no file', () => {
  const r = run(['--telemetry-status']);
  assert.ok(r.stdout.includes('not configured'), `stdout: ${r.stdout}`);
});

test('--telemetry-status: shows "enabled" after enable', () => {
  run(['--enable-telemetry']);
  const r = run(['--telemetry-status']);
  assert.ok(r.stdout.includes('enabled'), `stdout: ${r.stdout}`);
});

test('--telemetry-status: shows "disabled" after disable', () => {
  run(['--enable-telemetry']);
  run(['--disable-telemetry']);
  const r = run(['--telemetry-status']);
  assert.ok(r.stdout.includes('disabled'), `stdout: ${r.stdout}`);
});

test('--telemetry-status: shows config path', () => {
  const r = run(['--telemetry-status']);
  assert.ok(r.stdout.includes(tmpDir), `config path not shown: ${r.stdout}`);
});

test('--telemetry-status: shows "paused" for stale consent_version', () => {
  // Write a config with stale version directly
  fs.mkdirSync(tmpDir, { recursive: true });
  fs.writeFileSync(
    path.join(tmpDir, 'config.json'),
    JSON.stringify({ telemetry: true, consent_version: 0, consented_at: '2020-01-01T00:00:00.000Z' }),
    'utf8'
  );
  const r = run(['--telemetry-status']);
  assert.ok(r.stdout.includes('paused'), `stdout: ${r.stdout}`);
});
