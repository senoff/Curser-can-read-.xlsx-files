---
name: Bug report
about: Something xlsx-for-ai did wrong on a workbook
title: ''
labels: bug
assignees: ''
---

## What happened

<!-- One sentence: what did you run, and what went wrong? -->

## Repro

```bash
# the exact command line you ran
npx xlsx-for-ai ...
```

## Expected vs. actual

<!-- What did you expect to see? What did you actually get? -->

## Privacy note

**This project does not collect telemetry. We never auto-send your data.**
To help us reproduce, please attach the artifacts described below — both
are generated locally and contain no cell values, formulas, or text from
your workbook.

## Required: bug-report JSON

Run this against the file that triggered the bug:

```bash
npx xlsx-for-ai --report-bug your-file.xlsx
```

It writes `xlsx-for-ai-bugreport-<timestamp>.json` in the current
directory. **Drag-drop that file into this issue.** The report contains
only structural info (sheet count, shape, feature inventory, env) — no
cell content. You can `cat` it before attaching to verify.

## Optional: redacted workbook

If the bug only repros on workbooks with a specific structure (a
particular pivot, a chart configuration, a cross-sheet formula chain),
generate a redacted copy:

```bash
npx xlsx-for-ai --export-redacted-workbook your-file.xlsx
```

It writes `your-file-redacted.xlsx` next to the original. Every cell
value is replaced with a typed placeholder (`0`, `"x"`, `false`,
`1900-01-01`); formulas, structure, styles, and feature usage are
preserved. Open it in Excel first to confirm it still triggers the bug
before attaching.

## Environment

<!-- This is captured in the bug-report JSON; only fill in if you skipped that step. -->

- xlsx-for-ai version:
- Node version:
- OS:
