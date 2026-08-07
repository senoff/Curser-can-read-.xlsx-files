'use strict';

/**
 * `xfa feedback` and `xfa support` — the two inbound channels (XLS-264).
 *
 * Both ride the anonymous client_id: no auth, no login. `feedback` is
 * anonymous (tied only to client_id); `support` additionally carries an email
 * the user types, stored server-side as the reply-to. Neither relays email
 * from the CLI — the server owns any reply.
 *
 *   feedback: POST /feedback  { client_id, message }
 *   support:  POST /support   { client_id, email, question }
 *
 * Validation is client-side FIRST: an empty message, a missing/malformed
 * email, or a missing question is rejected here with a friendly usage line
 * and NO network call (the server still enforces the authoritative caps and
 * rejects — this just spares the user a wasted round-trip on obvious misuse).
 *
 * Internal test traffic (the DoD harness / SPM's live check) sets
 * XFA_INTERNAL=1, which adds `X-XFA-Internal: 1` so the server tags the row
 * and it's excluded from adoption stats — same mechanism client registration
 * already uses.
 */

const { post } = require('./client');
const { ensureRegistered } = require('./register');

// Pragmatic reply-to sanity check — one @, a dot in the domain, no spaces.
// Deliberately NOT a full RFC 5322 grammar: this guards against fat-finger
// typos before we promise "we'll reply to <email>", not against every exotic
// legal address. The server is the authority on what it can actually deliver.
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// Client-side soft caps — a local reject before a wasted round-trip. The
// server enforces the authoritative hard cap and 400s past it; these are
// generous so only obvious abuse trips them here.
const MAX_MESSAGE_CHARS = 8000;
const MAX_QUESTION_CHARS = 8000;
const MAX_EMAIL_CHARS = 254; // RFC 5321 addr-spec ceiling

const FEEDBACK_USAGE = 'Usage: xfa feedback "<your message>"';
const SUPPORT_USAGE = 'Usage: xfa support "<your email>" "<your question>"';

// Args arrive already shell-split. We join with a space so both the quoted
// form (`xfa feedback "two words"` -> one arg) and the bare form
// (`xfa feedback two words` -> two args) produce the same message.
function joinArgs(args) {
  return (args || []).join(' ').trim();
}

function validateFeedbackArgs(args) {
  const message = joinArgs(args);
  if (!message) return { ok: false, error: FEEDBACK_USAGE };
  if (message.length > MAX_MESSAGE_CHARS) {
    return { ok: false, error: `Message too long (max ${MAX_MESSAGE_CHARS} characters).` };
  }
  return { ok: true, message };
}

function validateSupportArgs(args) {
  const list = args || [];
  const email = (list[0] || '').trim();
  const question = joinArgs(list.slice(1));
  if (!email) return { ok: false, error: SUPPORT_USAGE };
  if (email.length > MAX_EMAIL_CHARS || !EMAIL_RE.test(email)) {
    return { ok: false, error: `That doesn't look like a valid email address: ${email}` };
  }
  if (!question) {
    return { ok: false, error: `Please include your question. ${SUPPORT_USAGE}` };
  }
  if (question.length > MAX_QUESTION_CHARS) {
    return { ok: false, error: `Question too long (max ${MAX_QUESTION_CHARS} characters).` };
  }
  return { ok: true, email, question };
}

// Anonymous POST opts, plus the internal tag header when XFA_INTERNAL=1.
function requestOpts() {
  const opts = { auth: false };
  if (process.env.XFA_INTERNAL === '1') opts.headers = { 'X-XFA-Internal': '1' };
  return opts;
}

// Shared body of both subcommands: validate, resolve client_id, POST, confirm.
// `deps` injects post / ensureRegistered / writers so tests run with no network.
async function run(kind, args, deps = {}) {
  const out = deps.out || ((m) => process.stdout.write(m));
  const err = deps.err || ((m) => process.stderr.write(m));
  const postFn = deps.post || post;
  const register = deps.ensureRegistered || ensureRegistered;

  const validate = kind === 'feedback' ? validateFeedbackArgs : validateSupportArgs;
  const v = validate(args);
  if (!v.ok) {
    err(v.error + '\n');
    return 1;
  }

  let clientId;
  try {
    const reg = await register();
    clientId = reg && reg.client_id;
  } catch (e) {
    err(`Couldn't reach xlsx-for-ai to register this install: ${e.message}\n`);
    return 1;
  }
  // Defensive: never POST { client_id: undefined }. ensureRegistered() always
  // sets one, but if a future change ever handed back a blank id we'd rather
  // fail fast with a clear message than send a body the server just 400s.
  if (!clientId) {
    err("Couldn't get a client id for this install. Please try again.\n");
    return 1;
  }

  const path = kind === 'feedback' ? '/feedback' : '/support';
  const body = kind === 'feedback'
    ? { client_id: clientId, message: v.message }
    : { client_id: clientId, email: v.email, question: v.question };

  try {
    await postFn(path, body, requestOpts());
  } catch (e) {
    const verb = kind === 'feedback' ? 'send your feedback' : 'submit your support request';
    err(`Couldn't ${verb}: ${e.message}\n`);
    return 1;
  }

  if (kind === 'feedback') {
    out('Thanks — your feedback was sent.\n');
  } else {
    out(`Got it — we'll reply to ${v.email}.\n`);
  }
  return 0;
}

function runFeedbackSubcommand(args, deps) {
  return run('feedback', args, deps);
}

function runSupportSubcommand(args, deps) {
  return run('support', args, deps);
}

module.exports = {
  runFeedbackSubcommand,
  runSupportSubcommand,
  validateFeedbackArgs,
  validateSupportArgs,
  EMAIL_RE,
  MAX_MESSAGE_CHARS,
  MAX_QUESTION_CHARS,
};
