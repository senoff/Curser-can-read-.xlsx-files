# xlsx-for-ai test suite

Comprehensive test suite using Node's built-in test runner (`node --test`). Zero new dependencies. ~98 tests across three layers:

- **Unit** (`test/unit/utilities.test.js`) — 70 tests of internal helpers (column math, parsing, formatting, type inference, spec validation, formula eval, etc.)
- **Output matrix** (`test/output-matrix.test.js`) — 18 tests covering every cell type × every output mode (text, markdown, JSON, SQL, schema), plus determinism + token-budget invariants + CSV input
- **Round-trip metadata fidelity** (`test/round-trip.test.js`) — 10 tests against synthetic fixtures covering values, layout, merges, named ranges, multi-sheet, hidden rows + hyperlinks

## Run

```bash
node --test test/round-trip.test.js test/output-matrix.test.js test/unit/*.test.js
```

## What it covers

The CLI's `--diff` mode compares cell *values* but not metadata. This suite is the gap-filler — it specifically asserts that **column widths, hidden rows/columns, frozen panes, merged cells, named ranges, and auto-filter ranges** all survive a `--json → write` round-trip, in addition to cell values.

This is the test class that catches the bug we shipped in 1.4.2 (column widths silently dropped on round-trip) and protects against future regressions of the same shape.

## Architecture

- **`helpers/synth.js`** — generates synthetic test fixtures programmatically. Each fixture targets a specific class of behavior (basic values, widths/layout, merges/names, multi-sheet, hidden/annotations). No real-world data needed; fixtures are reproducible and tiny.
- **`helpers/metadata.js`** — captures workbook metadata as a comparable snapshot, then diffs two snapshots. The output is a list of human-readable diff strings ("WIDTH Sheet1!E: 12 → -", "MERGE-LOST: A1:D1", etc.) for clear failure messages.
- **`round-trip.test.js`** — for each fixture, runs `xlsx-for-ai --json → xlsx-for-ai write`, snapshots both files, and asserts no metadata drift.
- **`fixtures/`** — gitignored. Fixtures are generated fresh each test run; never committed.

## Adding a new fixture

In `helpers/synth.js`:
1. Add a function that builds the workbook and writes it to `outDir`
2. Register it in the `FIXTURES` object
3. Add the filename + (optional) `todo` marker to `FIXTURE_TESTS` in `round-trip.test.js`

If the new fixture surfaces a known bug, set `todo: 'description of the bug'` — that documents the issue without failing CI. Once the bug is fixed in xlsx-for-ai, remove the `todo` flag.

## Currently-known bugs (todo'd in tests)

- `merges-names.xlsx` — merges become top-left-cell refs (e.g. `A1`) on round-trip; the `--json` output needs to emit the full range form (`A1:D1`)
- `annotations.xlsx` — hidden rows aren't preserved; both `--json` output and write path need a `hiddenRows[]` field per sheet

These are tracked here rather than in a separate issue tracker because they were discovered *by* the test suite — the test is the canonical bug report.

## CI

The intent is to wire this into GitHub Actions (`.github/workflows/test.yml`) so every PR runs the full suite. Not done yet — separate branch.
