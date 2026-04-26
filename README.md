# xlsx-for-ai

Converts spreadsheets into text, **markdown**, JSON, SQL, or schema dumps that AI coding agents can actually read.

AI tools — Claude, Cursor, Copilot, ChatGPT, and other LLM coding agents — can read text files but **not** `.xlsx` binaries. This CLI bridges the gap.

**Input formats:** `.xlsx` `.xls` `.xlsb` `.ods` `.csv` `.tsv`

**Output modes:** text dump, markdown tables (best LLM comprehension per token), JSON, SQL `CREATE TABLE`+`INSERT`, inferred schema, workbook diff.

It extracts everything a human would see in Excel:

- **Values** — strings, numbers, dates
- **Formulas** — the actual formula expression, plus shared-formula references
- **Formatting** — bold, italic, font colors, background fills
- **Number formats** — percentages, currency, custom patterns
- **Layout** — column widths, frozen panes, merged cells, alignment
- **Hyperlinks** — URLs embedded in cells
- **Comments / notes** — cell annotations
- **Named ranges** — workbook-defined names and their references
- **Hidden rows & columns** — flagged so the AI knows data is suppressed
- **Data validation** — dropdown lists, numeric constraints
- **Tables** — Excel Table objects with their names and column headers
- **Images & charts** — existence and position noted (content not rendered)
- **Auto-filters** — active filter ranges
- **Print areas** — defined print regions

> Previously published as **`cursor-reads-xlsx`**. The old name still works as an alias on the CLI, but please install the new package: `npm install -g xlsx-for-ai`.

## Install

```bash
npm install -g xlsx-for-ai
```

Or run directly with npx (no install needed):

```bash
npx xlsx-for-ai budget.xlsx
```

## Usage

```bash
# Dump all sheets
npx xlsx-for-ai data.xlsx

# Dump a specific sheet
npx xlsx-for-ai data.xlsx "Sheet1"

# List sheet names and dimensions without dumping
npx xlsx-for-ai data.xlsx --list-sheets

# Print to stdout instead of writing files
npx xlsx-for-ai data.xlsx --stdout

# Limit to first 200 rows per sheet (useful for huge files)
npx xlsx-for-ai data.xlsx --max-rows 200

# Limit to first 8 columns (useful for very wide sheets)
npx xlsx-for-ai data.xlsx --max-cols 8

# Suppress noisy default tags (default text colors, white fills, etc.)
npx xlsx-for-ai data.xlsx --stdout --compact

# Emit structured JSON (one entry per cell) instead of the text dump
npx xlsx-for-ai data.xlsx --json --stdout > out.json

# Combine flags
npx xlsx-for-ai data.xlsx "Sheet1" --stdout --max-rows 50 --compact
```

### Options

**Output modes** (mutually exclusive; default = text):

| Flag | Description |
|------|-------------|
| `--md` | Markdown tables — highest LLM comprehension per token |
| `--json` | Structured JSON, one object per cell |
| `--sql` | `CREATE TABLE` + `INSERT` statements (uses inferred schema) |
| `--schema` | Per-column schema (name, type, nullable, samples) as JSON |

**Selection:**

| Flag | Description |
|------|-------------|
| `[sheetName]` | Positional: dump only this sheet |
| `--range A1:D50` | Dump only this rectangular range |
| `--named-range NAME` | Dump only the cells covered by a workbook-defined name |
| `--max-rows N` | Cap at the first N rows per sheet |
| `--max-cols N` | Cap at the first N columns per sheet |

**Output control:**

| Flag | Description |
|------|-------------|
| `--list-sheets` | Print sheet names + dimensions and exit |
| `--stdout` | Print to stdout instead of writing files in `.xlsx-read/` |
| `--compact` | Suppress noisy default tags (default colors, "General" format) |
| `--max-tokens N` | Truncate output to ~N tokens; appends a tail summary noting what was dropped |
| `--evaluate` | Promote cached formula results to primary value; re-evaluate simple formulas via formulajs |

**Other modes:**

| Flag | Description |
|------|-------------|
| `--diff OTHER` | Diff this workbook vs `OTHER` — emit changed/added/removed cells and sheets |
| `--stream` | Streaming reader for huge `.xlsx` files (>100MB); emits row-by-row, drops some sheet metadata |
| `-h`, `--help` | Show help |

Output files are written to `.xlsx-read/` in the current working directory.
The path(s) are printed to stdout so your agent knows where to read.

## Output Format

### Text dump (default)

```
=== Sheet: Sales ===
Frozen: row 1, col 0
Columns: A(12) B(20) C(15) D(10)
Auto-filter: A1:D20
Named ranges:
  Totals: Sales!$D$2:$D$20
Table: "SalesTable" A1:D20 — columns: Region, Q1, Q2, Total

--- Row 1 [bold] ---
  A1: "Region"  [bold]
  B1: "Q1"  [bold] [align:center]
  C1: "Q2"  [bold] [align:center]
  D1: "Total"  [bold] [align:center]
--- Row 2 ---
  A2: "North"  [link: https://example.com/north]
  B2: 14500  [numFmt: #,##0]
  C2: 17200  [numFmt: #,##0]
  D2: 31700  [formula: =B2+C2] [numFmt: #,##0] [note: Includes returns]
--- Row 3 ---
  A3: "South"  [fill:FFFFFF00]
  B3: 9800  [numFmt: #,##0] [validation: list [North,South,East,West]]
  C3: 11050  [numFmt: #,##0]
  D3: 20850  [shared formula ref: D2] [numFmt: #,##0]
--- Row 4 (empty) [hidden] ---
```

### JSON dump (`--json`)

```json
{
  "name": "Sales",
  "rowCount": 4,
  "columnCount": 4,
  "frozen": { "rowSplit": 1, "colSplit": 0 },
  "columns": [{ "letter": "A", "width": 12 }, ...],
  "namedRanges": [{ "name": "Totals", "ranges": ["Sales!$D$2:$D$20"] }],
  "tables": [{ "name": "SalesTable", "ref": "A1:D20", "columns": ["Region", "Q1", "Q2", "Total"] }],
  "cells": [
    { "ref": "D2", "row": 2, "col": 4, "value": { "formula": "B2+C2", "result": 31700 }, "numFmt": "#,##0" },
    { "ref": "D3", "row": 3, "col": 4, "value": { "sharedFormulaRef": "D2", "result": 20850 }, "numFmt": "#,##0" }
  ]
}
```

### Sheet Metadata

| Line | Meaning |
|------|---------|
| `Frozen: row 1, col 2` | Frozen panes position |
| `Columns: A(12) B(20)` | Column widths (Excel character units) |
| `Hidden columns: E, F` | Columns hidden in the spreadsheet |
| `Merged: A1:B1` | Merged cell ranges |
| `Auto-filter: A1:D20` | Active auto-filter range |
| `Print area: A1:D50` | Defined print area |
| `Named ranges:` | Workbook-defined names referencing this sheet |
| `Table: "Name" A1:D20` | Excel Table objects with column headers |
| `Image: A1 to C5` | Embedded image position |

### Cell Tags

| Tag | Meaning |
|-----|---------|
| `[formula: =SUM(A1:A10)]` | Cell contains this formula (master cell) |
| `[shared formula ref: D2]` | Cell shares D2's formula (Excel "shared formula" — common when you drag-fill) |
| `[numFmt: 0.00%]` | Number format (when not "General") |
| `[bold]` | Bold font |
| `[italic]` | Italic font |
| `[color:FF8B0000]` | Font color (ARGB hex) |
| `[fill:FFFFFF00]` | Cell background color (ARGB hex) |
| `[align:center]` | Horizontal alignment (when not default) |
| `[link: https://...]` | Hyperlink URL |
| `[note: ...]` | Cell comment or note text |
| `[validation: list [...]]` | Data validation (dropdown values or constraints) |
| `[hidden]` | Row is hidden in the spreadsheet |

### `--list-sheets` Output

```
Sales  250 rows × 12 cols
Config  15 rows × 4 cols
Archive  1200 rows × 8 cols [hidden]
```

## Cursor / Claude / Agent Rule Template

Copy the included rule template into your project so your AI agent automatically uses this tool when it encounters `.xlsx` files:

```bash
mkdir -p .cursor/rules
cp node_modules/xlsx-for-ai/cursor-rule-template/read-xlsx.mdc .cursor/rules/
```

Or fetch it directly:

```bash
mkdir -p .cursor/rules
curl -o .cursor/rules/read-xlsx.mdc https://raw.githubusercontent.com/senoff/xlsx-for-ai/main/cursor-rule-template/read-xlsx.mdc
```

The same rule works for Claude Code (`.claude/rules/`), Copilot (`.github/copilot-instructions.md`), or any other agent — just adjust the path.

## Why This Exists

Spreadsheets are everywhere in real projects — financial models, data exports, config files, tax estimates. AI coding agents choke on binary formats. This tool makes spreadsheets legible to AI with zero information loss, including the tricky bits like shared formulas, named ranges, and merged cells that other tools drop.

## License

MIT
