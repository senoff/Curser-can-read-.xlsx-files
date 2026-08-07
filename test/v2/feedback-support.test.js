'use strict';

// XLS-264: `xfa feedback` and `xfa support`. Two guarantees pinned here:
//
// 1. Client-side validation rejects an empty message / missing-or-malformed
//    email / missing question BEFORE any network call (the DoD's "rejected
//    client-side" clause). We assert the injected post() was never called.
// 2. A valid submission posts the right body to the right path, unauthenticated
//    (auth:false), with the internal tag header only when XFA_INTERNAL=1, and
//    prints the spec's confirmation line + exits 0.
//
// post() and ensureRegistered() are injected, so the suite runs with no network
// and passes in CI (unlike the live-API cli-subcommands suite).

const { test } = require('node:test');
const assert = require('node:assert');

const {
  runFeedbackSubcommand,
  runSupportSubcommand,
  validateFeedbackArgs,
  validateSupportArgs,
} = require('../../lib/feedback');

const CLIENT_ID = '11111111-2222-3333-4444-555555555555';

// A deps harness: records post() calls and captures out/err, never touches the
// network.
function harness(overrides = {}) {
  const calls = [];
  const out = [];
  const err = [];
  const deps = {
    post: async (path, body, opts) => { calls.push({ path, body, opts }); return { ok: true }; },
    ensureRegistered: async () => ({ client_id: CLIENT_ID, api_key: 'xfa_test' }),
    out: (m) => out.push(m),
    err: (m) => err.push(m),
    ...overrides,
  };
  return { deps, calls, out, err };
}

// ---- pure validators -----------------------------------------------------

test('validateFeedbackArgs: rejects empty / whitespace, accepts real text', () => {
  assert.equal(validateFeedbackArgs([]).ok, false);
  assert.equal(validateFeedbackArgs(['   ']).ok, false);
  const v = validateFeedbackArgs(['love', 'the', 'diff', 'tool']);
  assert.equal(v.ok, true);
  assert.equal(v.message, 'love the diff tool');
});

test('validateSupportArgs: needs a valid email AND a question', () => {
  assert.equal(validateSupportArgs([]).ok, false, 'no email');
  assert.equal(validateSupportArgs(['not-an-email', 'help']).ok, false, 'malformed email');
  assert.equal(validateSupportArgs(['me@example.com']).ok, false, 'no question');
  const v = validateSupportArgs(['me@example.com', 'how', 'do', 'I', 'diff?']);
  assert.equal(v.ok, true);
  assert.equal(v.email, 'me@example.com');
  assert.equal(v.question, 'how do I diff?');
});

// ---- feedback subcommand -------------------------------------------------

test('feedback: valid message posts to /feedback and confirms, exit 0', async () => {
  const { deps, calls, out } = harness();
  const code = await runFeedbackSubcommand(['great', 'tool'], deps);
  assert.equal(code, 0);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].path, '/feedback');
  assert.deepEqual(calls[0].body, { client_id: CLIENT_ID, message: 'great tool' });
  assert.equal(calls[0].opts.auth, false, 'anonymous — no auth');
  assert.ok(out.join('').includes('feedback was sent'));
});

test('feedback: empty message is rejected client-side, no network', async () => {
  const { deps, calls, err } = harness();
  const code = await runFeedbackSubcommand([], deps);
  assert.equal(code, 1);
  assert.equal(calls.length, 0, 'must NOT post on invalid input');
  assert.ok(err.join('').includes('Usage: xfa feedback'));
});

test('feedback: XFA_INTERNAL=1 adds the internal tag header', async () => {
  const prev = process.env.XFA_INTERNAL;
  process.env.XFA_INTERNAL = '1';
  try {
    const { deps, calls } = harness();
    await runFeedbackSubcommand(['internal check'], deps);
    assert.equal(calls[0].opts.headers['X-XFA-Internal'], '1');
  } finally {
    if (prev === undefined) delete process.env.XFA_INTERNAL; else process.env.XFA_INTERNAL = prev;
  }
});

test('feedback: no internal header when XFA_INTERNAL is unset', async () => {
  const prev = process.env.XFA_INTERNAL;
  delete process.env.XFA_INTERNAL;
  try {
    const { deps, calls } = harness();
    await runFeedbackSubcommand(['normal user'], deps);
    assert.ok(!calls[0].opts.headers, 'no headers object when not internal');
  } finally {
    if (prev !== undefined) process.env.XFA_INTERNAL = prev;
  }
});

// ---- support subcommand --------------------------------------------------

test('support: valid email + question posts to /support and confirms, exit 0', async () => {
  const { deps, calls, out } = harness();
  const code = await runSupportSubcommand(['me@example.com', 'why', 'no', 'sheet?'], deps);
  assert.equal(code, 0);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].path, '/support');
  assert.deepEqual(calls[0].body, {
    client_id: CLIENT_ID,
    email: 'me@example.com',
    question: 'why no sheet?',
  });
  assert.equal(calls[0].opts.auth, false);
  assert.ok(out.join('').includes("we'll reply to me@example.com"), 'confirm names the reply-to email');
});

test('support: malformed email is rejected client-side, no network', async () => {
  const { deps, calls, err } = harness();
  const code = await runSupportSubcommand(['nope', 'my question'], deps);
  assert.equal(code, 1);
  assert.equal(calls.length, 0, 'must NOT post on a bad email');
  assert.ok(err.join('').includes('valid email'));
});

test('support: missing question is rejected client-side, no network', async () => {
  const { deps, calls } = harness();
  const code = await runSupportSubcommand(['me@example.com'], deps);
  assert.equal(code, 1);
  assert.equal(calls.length, 0);
});

// ---- network failure surfaces a friendly non-zero exit -------------------

test('feedback: a post() failure exits 1 with a friendly message', async () => {
  const { deps, out, err } = harness({
    post: async () => { const e = new Error('API unreachable'); throw e; },
  });
  const code = await runFeedbackSubcommand(['hi'], deps);
  assert.equal(code, 1);
  assert.equal(out.length, 0, 'no success confirmation on failure');
  assert.ok(err.join('').includes("Couldn't send your feedback"));
});
