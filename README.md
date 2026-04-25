# cursor-reads-xlsx

Converts `.xlsx` files into rich text dumps that AI coding agents can actually read.

AI tools like Cursor, Claude, Copilot, etc. can read text files but **not** `.xlsx` binaries. This CLI bridges the gap — it extracts everything a human would see in Excel and writes it to a plain text file:

- **Values** — strings, numbers, dates
- **Formulas** — the actual formula expression, not just the result
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

## Install

```bash
npm install -g cursor-reads-xlsx
```

Or run directly with npx (no install needed):

```bash
npx cursor-reads-xlsx budget.xlsx
```

## Usage

```bash
# Dump all sheets
npx cursor-reads-xlsx data.xlsx

# Dump a specific sheet
npx cursor-reads-xlsx data.xlsx "Sheet1"

# List sheet names and dimensions without dumping
npx cursor-reads-xlsx data.xlsx --list-sheets

# Print to stdout instead of writing files
npx cursor-reads-xlsx data.xlsx --stdout

# Limit to first 200 rows per sheet (useful for huge files)
npx cursor-reads-xlsx data.xlsx --max-rows 200

# Limit to first 8 columns (useful for very wide sheets)
npx cursor-reads-xlsx data.xlsx --max-cols 8

# Suppress noisy default tags (default text colors, white fills, etc.)
npx cursor-reads-xlsx data.xlsx --stdout --compact

# Emit structured JSON (one entry per cell) instead of the text dump
npx cursor-reads-xlsx data.xlsx --json --stdout > out.json

# Combine flags
npx cursor-reads-xlsx data.xlsx "Sheet1" --stdout --max-rows 50 --compact
```

### Options

| Flag | Description |
|------|-------------|
| `--list-sheets` | Print sheet names, row/column counts, and visibility — then exit |
| `--stdout` | Print output to stdout instead of writing `.txt` files |
| `--json` | Emit structured JSON (one object per cell with value/formula/format/style) |
| `--compact` | Suppress noisy default tags (default text color, white fill, etc.) — reduces token usage for AI agents |
| `--max-rows N` | Cap output at the first N rows per sheet |
| `--max-cols N` | Cap output at the first N columns per sheet |
| `-h`, `--help` | Show help message |

Output files are written to `.xlsx-read/` in the current working directory.
Each sheet produces a file named `<filename>--<sheetname>.txt`.
The path(s) are printed to stdout so your agent knows where to read.

## Output Format

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
  D3: 20850  [formula: =B3+C3] [numFmt: #,##0]
--- Row 4 (empty) [hidden] ---
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
| `[formula: =SUM(A1:A10)]` | Cell contains this formula |
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

## Cursor Rule Template

Copy the included rule template into your project so your AI agent automatically uses this tool when it encounters `.xlsx` files:

```bash
mkdir -p .cursor/rules
cp node_modules/cursor-reads-xlsx/cursor-rule-template/read-xlsx.mdc .cursor/rules/
```

Or if you installed globally / use npx, copy the template from the repo:

```bash
mkdir -p .cursor/rules
curl -o .cursor/rules/read-xlsx.mdc https://raw.githubusercontent.com/senoff/cursor-reads-xlsx/main/cursor-rule-template/read-xlsx.mdc
```

## Why This Exists

Spreadsheets are everywhere in real projects — financial models, data exports, config files. AI coding agents choke on binary formats. This tool makes spreadsheets legible to AI with zero information loss.

## License

MIT
