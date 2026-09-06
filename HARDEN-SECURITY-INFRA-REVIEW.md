# `harden-security-infra` merge-readiness review

**Branch:** `harden-security-infra` (6 commits, tip `057bef2`)
**Base:** `main`
**Reviewed:** 2026-04-30 (read-only, no checkout)
**Diff size:** 9 files changed, 927 insertions(+), 0 deletions (purely additive)

## Files touched (full diff)

| File | Lines added | Status |
|---|---|---|
| `.github/audit-allowlist.json` | +32 | new |
| `.github/dependabot.yml` | +88 | new |
| `.github/workflows/audit.yml` | +127 | new |
| `.github/workflows/upgrade-verify.yml` | +172 | new |
| `FORK_READINESS.md` | +248 | new |
| `README.md` | +8 | edit (added `## Security` section before `## License`) |
| `SECURITY.md` | +96 | new |
| `docs/INTEGRITY_PINNING.md` | +155 | new |
| `package.json` | +1 | edit (added `SECURITY.md` to `files` array) |

## Per-commit assessment

| SHA | Subject | Files | Complete unit? | Rationale |
|---|---|---|---|---|
| `314df05` | Stand up @protobi/exceljs adoption-readiness security infra | 5 | yes | Five-component scaffold (dependabot, audit.yml, upgrade-verify.yml, INTEGRITY_PINNING.md, FORK_READINESS.md) lands together as a coherent system. Workflows are syntactically complete; docs cross-reference each other; no broken refs. |
| `a8efdcd` | audit allowlist + lockfile re-resolve fixes | 5 | yes | Fixes two real bugs introduced by 314df05 that the commit message claims were caught by dry-run verification (high+ gate failed on unfixable xlsx CVEs; re-resolve script silently 404'd 21/153 entries). Adds `audit-allowlist.json` with 3 triaged entries and a documented allowlist policy in INTEGRITY_PINNING. |
| `f363551` | SECURITY.md entry point + correct stale doc claim | 2 | yes | Adds the GitHub-recognized `SECURITY.md` at repo root and corrects a stale claim in INTEGRITY_PINNING.md ("prepublish flow" referenced a `scripts/` dir that does not exist — verified: `scripts/` is absent on this branch). |
| `1eaf81d` | close silent-republish gap with daily registry re-resolve | 3 | yes | Adds `schedule:` and `workflow_dispatch:` triggers to upgrade-verify.yml, guards the PR-comment step with `github.event_name == 'pull_request'` (necessary for cron run not to crash), updates rule 7 in INTEGRITY_PINNING, and surfaces the security docs in README. Self-contained. |
| `bcf7383` | ship SECURITY.md to npm consumers | 1 | yes | One-line addition of `SECURITY.md` to `package.json` `files` array. Commit message documents `npm pack --dry-run` verification (4.3 kB, total 7 files, 32.8 kB tarball) and a CLI smoke test. Correctly ships SECURITY.md while keeping internal docs (INTEGRITY_PINNING, FORK_READINESS, allowlist) maintainer-only. |
| `057bef2` | harden audit allowlist schema validation | 1 | yes | Closes a real hole in the allowlist enforcement: previously `a.reassess && a.reassess < today` filtered missing-`reassess` entries OUT of the expired set, letting them suppress forever. Now enforces structural schema before the date check. Commit message documents three local test cases (current allowlist clean, missing reassess caught, malformed reassess caught). |

## Cross-cutting checks

| Check | Finding |
|---|---|
| (a) TODO/FIXME comments added | None. `grep -nE '^\+.*\b(TODO\|FIXME\|XXX\|HACK)\b'` over the full diff returns zero hits. |
| (b) Tests added but skipped/todo'd | None. Diff touches no test files (no `test/`, `tests/`, `__tests__/`, `*.test.js`, `*.spec.js`). The branch is pure infra/docs. |
| (c) Code paths referencing missing files/functions | None. SECURITY.md cross-links INTEGRITY_PINNING.md, FORK_READINESS.md, audit-allowlist.json, audit.yml, upgrade-verify.yml — all exist on the branch. README.md cross-links SECURITY.md, INTEGRITY_PINNING.md, FORK_READINESS.md — all exist. FORK_READINESS.md's coding-constraint claim that "exceljs is imported in exactly one place (index.js:24)" verified: `index.js:24 const ExcelJS = require('exceljs');`, no other imports. |
| (d) Pre-existing functionality removed | None. Diff is 927 insertions, 0 deletions on the README counts as a section insertion only. `package.json` edit only adds a line to the `files` array; nothing removed. |

## Other observations

- **No source-code changes.** Branch is entirely CI workflows + Markdown docs + JSON allowlist + one-line `package.json` `files` entry. Risk of regressing the CLI is essentially zero.
- **No lockfile changes.** No new npm deps; CI workflows install with `npm ci --ignore-scripts` against the existing lock.
- **`actionlint clean across both workflows`** is asserted in the `1eaf81d` commit message but not independently re-verified in this review.
- **Forward-looking name.** `FORK_READINESS.md` and the dependabot `@protobi/exceljs` group entry both reference a fork we have not adopted yet. They are forward-looking but harmless on `main` — neither activates anything until the dep is actually added.
- **`feat-bug-report` lineage.** This branch is the parent of `feat-bug-report`. Merging it as-is means `feat-bug-report` rebases cleanly onto a `main` that already carries the security infra. Reworking or peeling commits here forces a `feat-bug-report` rebase.

## Recommendation

**`merge-ready`**

## Rationale

All six commits are self-contained, well-scoped, and each one corrects a specific issue the prior commit's verification surfaced — exactly the iteration pattern that produces a clean merge. The cumulative diff is purely additive (927 insertions, 0 deletions), touches no source code or tests, and introduces no broken references, no TODOs, no skipped tests, and no removed functionality. The `feat-bug-report` lineage argues for a fast-forward merge so the child branch lands cleanly on top. Two minor caveats Bob may want to act on post-merge but neither blocks the merge: (1) the `last vetted version` table in FORK_READINESS.md has placeholders waiting for actual `@protobi/exceljs` adoption — fine since the fork is not yet a dep; (2) the `actionlint clean` claim in `1eaf81d` is taken on faith — Bob can re-run actionlint locally if he wants belt-and-suspenders before merging.

## Notes for the merger

- This is a fast-forwardable branch (linear history, 6 commits ahead of main, no merge commits). A `--ff-only` merge or a squash merge are both viable; I would lean toward preserving the 6 commits because the per-commit messages document the dry-run evidence behind each fix, which is useful future archaeology.
- After merge, `feat-bug-report` should rebase onto the new `main` rather than merge — keeps history linear.
