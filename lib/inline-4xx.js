'use strict';

// ---------------------------------------------------------------------------
// Shared 4xx inline-surface sanitizer + surfacer.
//
// Both the MCP path (mcp.js `friendlyErrorMessage`) and the CLI path
// (index.js `friendlyCliError`) need to turn an `API_CLIENT_ERROR` (4xx)
// into short, client-safe text that surfaces the server's validation
// message (the caller's own input shape — "Sheet \"X\" not found.
// Available sheets: ...", "spec.sheets must be an array") while scrubbing
// anything sensitive a wrapped 4xx path could carry. Lifted here so there
// is exactly ONE sanitizer — not a second, drifting fork per surface.
//
// 5xx / everything else stays generic and is handled by each caller; this
// module is 4xx-only.
// ---------------------------------------------------------------------------

// Defense in depth on the 4xx inline message. The SPEC's bet is that
// 4xx server messages describe the CALLER'S OWN INPUT (which field,
// what was expected) — but a wrapped 4xx path could still carry
// absolute file paths, emails, JWTs / Bearer tokens, Slack tokens,
// or other PII. Scrub those before surfacing, replace with `<…>`
// placeholders so the caller still sees the SHAPE of the message
// without the sensitive payload.
//
// `<…>` was picked over a more verbose `[redacted-x]` so it's
// visually compact and unambiguously not real input.
const PII_SCRUBBERS = [
  // Bearer / Authorization tokens — match before generic JWT pattern.
  [/\bBearer\s+[A-Za-z0-9._~+/-]{8,}=*/g, '<bearer>'],
  // JSON Web Tokens. Three dot-separated base64url segments, the first
  // starting with `eyJ` (the canonical JWT header prefix).
  [/\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\b/g, '<jwt>'],
  // Slack bot / user / app tokens.
  [/\bxox[bpoars]-[A-Za-z0-9-]{10,}\b/g, '<slack-token>'],
  // Our own API keys.
  [/\bxfa_[a-z]+_[A-Za-z0-9]{16,}\b/g, '<xfa-key>'],
  // Generic 32+ char hex (api keys / hashes).
  [/\b[a-f0-9]{32,}\b/gi, '<hex>'],
  // Emails.
  [/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g, '<email>'],
  // --- Absolute paths ---------------------------------------------------
  // A path CAN contain spaces AND parentheses ("/Users/Bob Smith/My File
  // (1).xlsx", "C:\Program Files (x86)\app.exe"), but so does the
  // surrounding prose ("... failed at row 3"). The ONLY genuinely ambiguous
  // case is "where does an EXTENSIONLESS path end in free prose"; when a path
  // ends in a file extension we can anchor there and preserve the prose. So
  // two categorical pass families (first match wins — the substring is
  // `<path>` by the time the next pass runs), each covering all three
  // dialects (POSIX, Windows drive-letter, Windows UNC \\host\share):
  //
  //   A. EXTENSION-ANCHORED — ROOT-AGNOSTIC. When a token ends in a file
  //      extension it is unambiguously a filename, so we do NOT gate on a
  //      root allowlist: any absolute path (any Windows drive, any UNC host,
  //      any POSIX root — `/data`, `/srv/www`, whatever) is redacted. The
  //      lazy inner `[^\r\n'"]*?` covers spaces, parens, and backticks so no
  //      char class can leak a tail; the extension anchor + boundary lookahead
  //      stop the match AT the extension, so trailing prose survives. The
  //      POSIX pass requires the leading `/` to sit on a word boundary
  //      (negative lookbehind), which keeps it from eating the `//` inside a
  //      URL — every `/` there is preceded by a non-boundary char.
  //   B. EXTENSIONLESS. Greedy from a known anchor to a hard stop (newline or
  //      a quote). Here the end is genuinely unknowable, and a root-agnostic
  //      leading-`/` match would swallow arbitrary prose, so B STAYS anchored
  //      to a system-root allowlist and over-redacts toward end-of-line — the
  //      safe default for a defense-in-depth PII net (these are rare, and the
  //      server that emits these 4xx validation messages never receives client
  //      paths in the first place; this is belt-and-suspenders). The residual —
  //      an extensionless path under a NON-listed root — is theoretical-only on
  //      this surface and is the one deliberately-open seam, not an oversight.
  //
  // A — extension-anchored, root-agnostic, all dialects, all inner chars:
  [/[A-Za-z]:\\[^\r\n'"]*?\.[A-Za-z0-9]{1,10}(?=[\s'"`)\]]|$)/g, '<path>'],   // Windows drive
  [/\\\\[^\r\n'"]*?\.[A-Za-z0-9]{1,10}(?=[\s'"`)\]]|$)/g, '<path>'],           // Windows UNC
  [/(?<![^\s'"([])\/[^\r\n'"]*?\.[A-Za-z0-9]{1,10}(?=[\s'"`)\]]|$)/g, '<path>'], // POSIX (any root)
  // B — extensionless, greedy to newline or quote (over-redact toward EOL):
  [/[A-Za-z]:\\[^\r\n'"]*/g, '<path>'],   // Windows drive
  [/\\\\[^\r\n'"]*/g, '<path>'],           // Windows UNC
  [/\/(?:Users|home|var|opt|tmp|etc|private|Volumes|Library|usr|mnt|srv)\/[^\r\n'"]*/g, '<path>'], // POSIX (system roots)
];

// Strip the well-known low-signal noise an inline 4xx surface message
// could carry: leading "xlsx-for-ai API error 4xx: " prefix from
// lib/client.js, scrub PII via PII_SCRUBBERS, bound the length so a
// pathological payload can't blow up the conversation log / CLI stderr.
const INLINE_4XX_MAX_LEN = 280;
function shapeInline4xxMessage(raw) {
  if (typeof raw !== 'string') return '';
  let s = raw.replace(/^xlsx-for-ai API error \d+:\s*/i, '').trim();
  for (const [pattern, replacement] of PII_SCRUBBERS) {
    s = s.replace(pattern, replacement);
  }
  if (s.length > INLINE_4XX_MAX_LEN) {
    s = s.slice(0, INLINE_4XX_MAX_LEN - 1) + '…';
  }
  return s;
}

// Pull the raw inline message out of a 4xx error. Prefer the structured
// `{error: {message}}` shape our server emits, fall through to the flat
// `message` / string `error`, finally the wrapped `err.message` (whose
// "API error 4xx:" prefix shapeInline4xxMessage strips).
function extractInline4xxRaw(payload, err) {
  let inline = '';
  if (payload && typeof payload === 'object') {
    const structured = payload.error;
    if (structured && typeof structured === 'object' && typeof structured.message === 'string') {
      inline = structured.message;
    } else if (typeof payload.message === 'string') {
      inline = payload.message;
    } else if (typeof payload.error === 'string') {
      inline = payload.error;
    }
  }
  if (!inline && err && typeof err.message === 'string') {
    inline = err.message;
  }
  return inline;
}

// Turn a 4xx (`API_CLIENT_ERROR`) error into the final client-safe line,
// prefixed by the caller's label (tool name on the MCP side, command
// prefix on the CLI side). Known specific statuses keep curated text
// (ordered first); the generic branch surfaces the sanitized server
// message; an empty/absent payload degrades gracefully — never
// `undefined`, never `[object Object]`.
function surface4xx(prefix, err) {
  // Coerce so a string status ("429") still hits the curated branches; a
  // missing/garbage status becomes NaN and falls through to the generic path.
  const status = Number(err && err.status);
  const payload = err && err.payload;

  if (status === 429) {
    return `${prefix}: monthly request cap reached — resets next month.`;
  }
  if (status === 402) {
    // Server emits 402 only on the not-yet-active full_bytes capture-consent
    // path; neutralize so its subscription wording never reaches the user.
    return `${prefix}: that capture mode is not available.`;
  }

  const shaped = shapeInline4xxMessage(extractInline4xxRaw(payload, err));
  if (shaped) {
    return `${prefix}: ${shaped}`;
  }
  return `${prefix}: invalid request (no detail provided).`;
}

module.exports = {
  PII_SCRUBBERS,
  INLINE_4XX_MAX_LEN,
  shapeInline4xxMessage,
  extractInline4xxRaw,
  surface4xx,
};
