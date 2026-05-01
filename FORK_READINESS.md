# FORK_READINESS — surviving a `@protobi/exceljs` compromise

`@protobi/exceljs` is a single-maintainer npm package
(`https://npm.im/@protobi/exceljs`, repo: `github.com/protobi/exceljs`).
It is a soft fork of upstream `exceljs` carrying pivot-table enhancements and
bug fixes pending upstream merge. We may adopt it as a direct dependency once
the upstream merge stalls long enough that the pivot work is load-bearing for
us.

A single-maintainer scoped package is a single point of failure: if the
`protobi` npm account is compromised, an attacker can publish a malicious
`@protobi/exceljs@x.y.z` that lands in `xlsx-for-ai` users' machines on the
next `npm install`. This document is the runbook for that scenario, and the
prep work we do *before* it happens so the response is hours, not days.

The procedure assumes you have already read `docs/INTEGRITY_PINNING.md` and
understand what `npm audit signatures` and the lockfile-resolve check
actually catch.

---

## Triggers — when to execute this runbook

Any one of:

1. `upgrade-verify.yml` fails on a `@protobi/exceljs` bump and re-running the
   workflow does not clear it.
2. `npm audit signatures` reports a signature failure on `@protobi/exceljs`
   that does not correspond to a documented npm signing-key rotation.
3. A new `@protobi/exceljs` version appears on npm with no matching tag,
   release, or commit on `github.com/protobi/exceljs`.
4. The `protobi` GitHub or npm profile shows signs of takeover (deleted
   repos, force-pushed `main`, pinned-issue notice from the maintainer,
   sudden ownership transfer).
5. Socket.dev or another supply-chain scanner flags the version as
   exfiltrating credentials, opening a reverse shell, or running install
   scripts that touch the network.
6. Bob's judgment call. Trust your gut here — the cost of a false alarm is
   one wasted afternoon; the cost of shipping a poisoned `npm install` to
   users is much higher.

If *any* trigger fires, treat the package as compromised until proven
otherwise. Stop merging Dependabot PRs that touch it. Open a `security`
issue on this repo to track the response.

---

## Pre-positioning — do these once, now

These steps cost nothing and shave hours off the response.

### 1. Mirror the repo locally

```sh
git clone --mirror https://github.com/protobi/exceljs.git \
  ~/src/mirrors/protobi-exceljs.git
```

Re-run weekly via cron or a scheduled GitHub Action. A mirror clone
preserves all branches, tags, and refs — including ones the attacker may
delete during a takeover. If `protobi/exceljs` disappears from GitHub, this
mirror is what we fork from.

### 2. Record the last-known-good commit and tarball hash

Pin a known-good baseline in this file. Update it whenever we adopt a new
`@protobi/exceljs` version after a clean `upgrade-verify` run.

| Field                          | Value                                                  |
|--------------------------------|--------------------------------------------------------|
| Last vetted version            | `4.4.0-protobi.9`                                      |
| Published                      | `2026-02-02` by `protobi`                              |
| Upstream `exceljs` base        | `4.4.0`                                                |
| Tarball SRI hash (sha512)      | _populate on adoption — `npm view @protobi/exceljs@4.4.0-protobi.9 dist.integrity`_ |
| Mirror commit (last vetted)    | _populate on adoption — `git -C ~/src/mirrors/protobi-exceljs.git rev-parse HEAD`_ |
| Vetted by                      | _name + date_                                          |
| Notes                          | _e.g. "diff vs upstream reviewed; pivot patch only"_   |

Anyone responding to a compromise should be able to look at this table and
know exactly which bytes are trustworthy.

### 3. Keep upstream `exceljs@4.4.0` warm

The fastest fallback is to revert to upstream `exceljs` and lose the pivot
features temporarily. We already depend on `exceljs ^4.4.0` directly today,
so the revert path is "remove `@protobi/exceljs` from `package.json`,
`npm install`, run the round-trip corpus." Do not let the codebase grow a
hard requirement on protobi-only APIs without a feature flag — see
*Coding constraints* below.

### 4. Reserve the npm fork name

Reserve `@senoff/exceljs` (or whichever scope we publish under) on npm now,
even if empty. A 1-byte placeholder publish costs nothing and prevents an
attacker from squatting the obvious replacement name during a fast-moving
incident.

```sh
mkdir /tmp/scope-reserve && cd /tmp/scope-reserve
npm init -y --scope=@senoff
# edit package.json: name=@senoff/exceljs, version=0.0.0, private=false
echo "// reserved" > index.js
npm publish --access=public
```

---

## Response — once a trigger fires

Time matters here. The longer the bad version sits on npm, the more
downstream users pull it.

### Step 1 — Freeze (minutes)

1. Open a `security` issue on `senoff/xlsx-for-ai` describing the trigger.
2. In `package.json`, pin `@protobi/exceljs` to the **last vetted version**
   from the table above with an exact match (no caret):
   ```json
   "@protobi/exceljs": "4.4.0-protobi.9"
   ```
   Add an `overrides` entry doing the same so transitives cannot pull a
   newer version. Commit on a branch named `freeze-protobi-<date>`.
3. Disable Dependabot for `@protobi/exceljs` temporarily by adding it to
   `.github/dependabot.yml` `ignore:`. Do not let an auto-PR re-introduce
   the bad version.
4. If a release is in flight, hold it. If 1.x.y is already on npm with the
   bad version, see *Step 5 — User-facing recovery*.

### Step 2 — Diagnose (under an hour)

Goal: decide whether this is a key rotation, a benign mistake, or an
actual compromise.

1. Compare the bad-version tarball against the last-vetted version:
   ```sh
   npm pack @protobi/exceljs@<bad-version>
   npm pack @protobi/exceljs@4.4.0-protobi.9
   diff -r <(tar -xOf bad.tgz) <(tar -xOf good.tgz) | less
   ```
   Look for: new `postinstall` / `preinstall` scripts, network calls in
   non-network code paths, base64 blobs, additions to `bin/`, modifications
   to `package.json` scripts, new top-level deps.
2. Check `socket.dev/npm/package/@protobi/exceljs` for the version's risk
   signals.
3. Check `github.com/protobi/exceljs` — does the version have a matching
   tag and commit? Does the commit history match what the local mirror has?
   Force-pushes to `main` are a strong signal of takeover.
4. Check `github.com/protobi` for any pinned issue, deleted repo, or
   ownership transfer.
5. If you can reach the maintainer (GitHub issue, email on commits), ask
   directly whether they published this version. Document the response.

### Step 3 — Decide

Three outcomes:

- **Benign** (key rotation, accidental publish, version with no malicious
  diff): note the diagnosis in the security issue, unfreeze, resume normal
  flow. Update the *last vetted version* table.
- **Compromised but recoverable** (the package is bad but the maintainer is
  responsive and intends to unpublish/republish): stay frozen on the last
  vetted version. Wait for the cleaned republish and re-verify before
  unfreezing.
- **Compromised, non-recoverable** (account takeover with no maintainer
  response, or the maintainer's repo is gone): proceed to Step 4.

### Step 4 — Fork

Cut our own scoped fork.

1. From the local mirror at `~/src/mirrors/protobi-exceljs.git`, check out
   the **mirror commit (last vetted)** SHA from the table.
2. Push it to a fresh `senoff/exceljs` repo on GitHub.
3. Update `package.json` of the fork: rename to `@senoff/exceljs`, bump
   patch, regenerate lockfile.
4. Publish to npm under `@senoff/exceljs` (the scope reserved in
   *Pre-positioning step 4*). Use a hardware-backed 2FA token.
5. In `xlsx-for-ai`:
   ```diff
   - "@protobi/exceljs": "4.4.0-protobi.9"
   + "@senoff/exceljs": "<published-version>"
   ```
   Update any `require('@protobi/exceljs')` or import paths. There should
   be only one or two — see *Coding constraints* below.
6. Run `npm ci`, full test suite, round-trip corpus. Open the PR; let
   `audit.yml` and `upgrade-verify.yml` pass.
7. Cut a patch release of `xlsx-for-ai`.

### Step 5 — User-facing recovery

If a poisoned `xlsx-for-ai` version made it to npm before we caught it:

1. `npm deprecate xlsx-for-ai@<bad-version> "compromised dependency, do not install — see GH issue #N"`.
2. Publish a clean patch release immediately (Step 4 above produces it).
3. Open a GitHub release with a "Security" header naming the affected
   versions.
4. Post a one-line note on the npm package README pointing to the security
   issue.
5. We do **not** have telemetry on installs. Assume some users pulled the
   bad version. The deprecate notice is what they will see on their next
   `npm outdated` / `npm update`.

We cannot `npm unpublish` a version older than 72 hours, and even within
72 hours unpublish is brittle. `npm deprecate` is the right tool.

---

## Coding constraints — keep the fork path cheap

The cost of executing this runbook scales with how deeply
`@protobi/exceljs` is wired into `xlsx-for-ai`. Keep the surface area small:

1. **Import the package in exactly one module**, e.g. `lib/exceljs.js`,
   which re-exports what the rest of the code uses. Swapping
   `@protobi/exceljs` for `@senoff/exceljs` or upstream `exceljs` should be
   a one-line change in that one file.
2. **Gate protobi-only features behind a capability check.** If the pivot
   API is missing (because we fell back to upstream `exceljs`), degrade to
   the non-pivot code path with a logged warning rather than crashing.
3. **Do not let protobi-only types leak into public API.** The CLI's
   external surface (text/JSON output schema, flags) must not depend on
   anything specific to the fork.
4. **The round-trip corpus must run green on upstream `exceljs` alone.**
   Any test that requires the protobi fork is marked as such. CI runs both
   matrices on release branches.

---

## What this doc is not

- Not a substitute for `docs/INTEGRITY_PINNING.md`. Read that first.
- Not a guarantee. A determined attacker who compromises both the
  maintainer account *and* GitHub *and* re-publishes within 72 hours can
  defeat parts of this. The point is to make that combination expensive
  and to give us a clean rollback when any single layer fails.
- Not Hollywood. There is no red phone. The response is "freeze, diagnose,
  decide, fork" and most of the work is already pre-positioned.

## Review cadence

Re-read and update this file:

- Whenever we adopt a new `@protobi/exceljs` version (update the
  *last vetted version* table).
- Quarterly, even if nothing changes (sanity check the mirror cron, the
  reserved scope, the freeze procedure).
- Immediately after any execution of the runbook, real or drill —
  document what worked and what did not.
