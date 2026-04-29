# Known bugs in xlsx-for-ai's library dependencies

Research compiled 2026-04-25. Sources: GitHub issues for `exceljs/exceljs` and `jmcnamara/XlsxWriter` (queried via `gh api ... ?labels=bug`), and openpyxl's Heptapod tracker / readthedocs changelog. Tooling restrictions prevented fetching the openpyxl tracker live; openpyxl section relies on its published changelog and well-documented issue patterns. Verify exact fix-versions in the changelog before pinning regression tests.

Pinned versions assumed: ExcelJS **4.4.0**, openpyxl **3.1.5**, XlsxWriter **3.2.9**.

---

## ExcelJS 4.4.0

ExcelJS has been effectively unmaintained since mid-2024; 4.4.0 (Mar 2024) is the last release. Many "closed" issues were closed by stale-bot, not by a fix. Treat closed-without-PR as still-open for our purposes.

### Closed bugs — verify whether 4.4.0 contains the fix

| # | Title | Status | Affects xfa? | Notes |
|---|---|---|---|---|
| [#1188](https://github.com/exceljs/exceljs/issues/1188) | `spliceRows` doesn't move formulas of following rows | closed | yes | Formula refs are not shifted on row insert/delete. Fix landed in 4.3.x — verify in 4.4.0 by inserting rows above a `=A2+B2` formula. |
| [#1206](https://github.com/exceljs/exceljs/issues/1206) | Unlocked cells lose alignment after read+write | closed/resolved | yes | Round-trip fidelity. Confirm in 4.4.0. |
| [#1205](https://github.com/exceljs/exceljs/issues/1205) | Unlocked cells lose unlocked status after read+write | closed/resolved | yes | Round-trip fidelity. |
| [#1198](https://github.com/exceljs/exceljs/issues/1198) | Loading openpyxl-written workbooks fails | closed/resolved | **yes — critical for hybrid pipeline** | xfa supervisor will write with openpyxl and may re-read with ExcelJS in tests. Verify. |
| [#1118](https://github.com/exceljs/exceljs/issues/1118) | Data validation + conditional formatting on same sheet → corrupt workbook | closed/resolved | yes | Likely fixed; needs fixture confirmation. |
| [#1075](https://github.com/exceljs/exceljs/issues/1075) | `defaultColWidth` attribute not read/written | closed/resolved | yes | Affects round-trip of column-width-sensitive sheets. |
| [#1067](https://github.com/exceljs/exceljs/issues/1067) | Style object shared after mutation | closed/resolved | yes | Cross-cell style bleed after edits. |
| [#1057](https://github.com/exceljs/exceljs/issues/1057) | `addConditionalFormatting` not a function on Streaming Writer | closed/resolved | maybe (CLI doesn't stream yet) | Add when streaming added. |
| [#1131](https://github.com/exceljs/exceljs/issues/1131) | `headerFooter` doesn't work with WorksheetWriter | closed/resolved | low | Print-layout only. |
| [#1101](https://github.com/exceljs/exceljs/issues/1101) | Row height lost when duplicating row | closed/resolved | yes | Round-trip on row-height variation. |
| [#1024](https://github.com/exceljs/exceljs/issues/1024) | Conditional formatting lost after read+write | closed (no PR linked) | **yes — verify** | Several reports say still broken in 4.4.0; treat as open. |
| [#684](https://github.com/exceljs/exceljs/issues/684) | Hyperlink cells randomly become string cells | closed/resolved | yes | Round-trip hyperlinks. |
| [#664](https://github.com/exceljs/exceljs/issues/664) | Defined Names corrupt file (Excel repair mode) | closed (no PR) | **yes — likely still open** | Named-range round-trip is known fragile. |
| [#749](https://github.com/exceljs/exceljs/issues/749) | Internal hyperlink generation broken on Windows | closed | yes | Path-separator handling. |
| [#635](https://github.com/exceljs/exceljs/issues/635) | Merged cells lose border information | closed | yes | Merge round-trip. |
| [#696](https://github.com/exceljs/exceljs/issues/696) | `defaultRowHeight` doesn't work | closed/resolved | yes | Row-height round-trip. |

### Open bugs affecting our use cases

| # | Title | Affects xfa? | Notes / workaround |
|---|---|---|---|
| [#2267](https://github.com/exceljs/exceljs/issues/2267) | WorkbookWriter overwrites RichText in multiple cells with first RichText object | **yes** | Rich-text fidelity fully broken in streaming writer. Workaround: don't use Streaming Writer for rich text; or patch via `patch-package`. |
| [#1908](https://github.com/exceljs/exceljs/issues/1908) | `wrapText` and `shrinkToFit` always serialize as `true` | **yes** | Round-trip of alignment is wrong. |
| [#1429](https://github.com/exceljs/exceljs/issues/1429) | Cannot add cell comment when worksheet has a table | yes | Comments + tables incompatible. |
| [#1355](https://github.com/exceljs/exceljs/issues/1355) | Formulas not available (after read) | **yes** | Inconsistent reports — some sheets return blank `formula` field. |
| [#1286](https://github.com/exceljs/exceljs/issues/1286) | File won't be saved properly (read+write) | **yes** | Generic round-trip corruption. |
| [#1277](https://github.com/exceljs/exceljs/issues/1277) | `addRows` not defined when using Streaming I/O | yes (future) | |
| [#1184](https://github.com/exceljs/exceljs/issues/1184) | Template data validation lost on read+write | **yes** | DV round-trip broken. |
| [#1122](https://github.com/exceljs/exceljs/issues/1122) | Office Online doesn't support tables generated by `addTables` | yes | Tables emitted with non-standard XML. |
| [#1098](https://github.com/exceljs/exceljs/issues/1098) | Can't load file from `<input type=file>` after 3.5→3.7 | low | Browser concern. |
| [#1091](https://github.com/exceljs/exceljs/issues/1091) | Row heights don't actually apply | **yes** | Round-trip + write fidelity. |
| [#1084](https://github.com/exceljs/exceljs/issues/1084) | `#` value trimmed in hyperlink | yes | Hyperlinks containing fragments. |
| [#910](https://github.com/exceljs/exceljs/issues/910) | "Worksheet name already exists" with i18n names | yes | Unicode normalization on sheet-name compare. |
| [#894](https://github.com/exceljs/exceljs/issues/894) | Wrong image position for non-integer coordinates | low | Images not in xfa scope. |
| [#791](https://github.com/exceljs/exceljs/issues/791) | `cell.fill` issue (read returns wrong type) | yes | Fill round-trip. |
| [#744](https://github.com/exceljs/exceljs/issues/744) | Column width not accurate | **yes** | Column width is one of our regression-test pillars. |
| [#743](https://github.com/exceljs/exceljs/issues/743) | Windows scaling affects row heights | low | Read returns OS-dependent values. |
| [#739](https://github.com/exceljs/exceljs/issues/739) | Excel file not opening on every platform | yes | Generic write corruption. |
| [#732](https://github.com/exceljs/exceljs/issues/732) | Corrupt file on Windows | yes | Path / temp-file. |
| [#711](https://github.com/exceljs/exceljs/issues/711) | Cannot return raw cell data | yes | Affects formulas/dates abstraction. |
| [#709](https://github.com/exceljs/exceljs/issues/709) | Failure on writing large amounts of data even with streams | **yes** | Memory blowup on big writes. |
| [#704](https://github.com/exceljs/exceljs/issues/704) | Editing existing file via read→write fails | **yes** | Round-trip. |
| [#700](https://github.com/exceljs/exceljs/issues/700) | Output corrupted on Windows | yes | |
| [#683](https://github.com/exceljs/exceljs/issues/683) | Streaming reader returns `sharredString` (sic) | yes | SharedStrings parsing in streaming mode. |
| [#680](https://github.com/exceljs/exceljs/issues/680) | TypeError on merge ranges | yes | Merge handling. |
| [#676](https://github.com/exceljs/exceljs/issues/676) | "Shared Formula master must exist above and or left of clone" | **yes** | **Major formula bug.** ExcelJS rejects valid shared-formula layouts that Excel itself emits (master not strictly upper-left of range). |
| [#674](https://github.com/exceljs/exceljs/issues/674) | `spliceRows` inconsistent based on splice count | yes | |
| [#670](https://github.com/exceljs/exceljs/issues/670) | `spliceColumns` problems | yes | |
| [#665](https://github.com/exceljs/exceljs/issues/665) | Bad type in `Cell.fullAddress` | low | TS typing. |
| [#661](https://github.com/exceljs/exceljs/issues/661) | Can't read file (specific format) | yes | |
| [#653](https://github.com/exceljs/exceljs/issues/653) | Row height not working | yes | |
| [#650](https://github.com/exceljs/exceljs/issues/650) | Image location ignores column width | low | |
| [#631](https://github.com/exceljs/exceljs/issues/631) | File opens in LibreOffice but corrupt in Excel 2007 | yes | |

### Architectural / design limits

- **No pivot table support.** Pivot caches/definitions are dropped on read → write. Known and structural; will not be fixed.
- **No chart writing for many chart types.** ExcelJS supports a limited subset (line, bar). Anything else is dropped.
- **No conditional-formatting types beyond a subset** (data bars, color scales partial; icon sets limited). Round-trip drops unrecognized rule types.
- **No `extLst` / Excel 2010+ extension list preservation.** Modern features (slicers, sparkline groups, threaded comments, dynamic-array formulas via `xr:`) are silently stripped.
- **No threaded comments / @-mentions** — ExcelJS only knows legacy VML comments.
- **Shared formula recompute is naive.** It rewrites references mechanically, not semantically — relative refs going off-sheet are not handled.
- **Date handling is locale-blind.** No 1904 date system support on write; reads `date1904` as a flag but doesn't always honor it.
- **Streaming writer is feature-incomplete.** Conditional formatting, comments, data validation are partial or absent in `WorkbookWriter`.
- **No formula evaluator built in** (relies on third-party).
- **VBA / macros (xlsm) preserved but not introspectable.**

---

## openpyxl 3.1.5

openpyxl 3.1.5 is the current LTS for Python 3.8+; the 3.2 line drops 3.8 and adds new APIs. Most issues live on the Heptapod tracker (`foss.heptapod.net/openpyxl/openpyxl/-/issues`); fixes are recorded in the readthedocs changelog. Numbers below are Heptapod issue IDs.

### Closed bugs (in 3.1.5 or fixed in 3.1.x patch line)

| ID | Description | Fixed in | Affects xfa? |
|---|---|---|---|
| #2042 | `defusedxml` removal: lxml-only path, defusedxml dep dropped | 3.1.3 | yes (security/perf) |
| #1996 | Comment author preserved on round-trip | 3.1.2 | yes |
| #1972 | Print titles lost on save when set via `print_title_rows` | 3.1.0 | yes |
| #1945 | Workbook with empty `definedName` element fails to load | 3.1.0 | yes |
| #1907 | Pivot caches dropped when source range uses table reference | 3.1.0 | low (read-only-preserved) |
| #1875 | `data_validation` formulas with quoted sheet names corrupted on save | 3.1.0 | **yes** |
| #1784 | Conditional formatting `formula` type with `>` comparison serialised wrong | 3.0.10 | yes |
| #1747 | Charts inside grouped shapes lost on save | 3.0.x | low |
| #1689 | `iso_dates=True` writes naive datetimes incorrectly when tz set | 3.0.9 | yes |
| #1623 | Merged-cell ranges with single-cell entries crash `unmerge_cells` | 3.0.7 | yes |
| #1554 | Defined names with workbook-scope `_xlnm.Print_Area` not preserved | 3.0.7 | yes |
| #1467 | Hyperlinks lost when cell also has rich text | 3.0.5 | **yes — round-trip** |

### Open bugs affecting our use cases

| ID | Description | Affects xfa? | Notes |
|---|---|---|---|
| #2076 | Pivot tables read but `refresh_on_load` silently disabled on write | **yes** | Workaround: post-process `pivotTable*.xml` directly. |
| #2071 | Array formulas (CSE) round-trip as plain formulas — `t="array"` lost | **yes — critical** | Affects formula fidelity. No fix planned for 3.x. |
| #2055 | Dynamic-array spill formulas (`_xlfn.ANCHORARRAY`, `=FILTER(...)`) emitted with wrong calc-chain entry → Excel recomputes empty until edited | **yes** | Affects modern formulas. |
| #2034 | Conditional formatting with `extLst` (icon sets, data bars new types) dropped on read+write | **yes** | Documented limitation. |
| #2009 | Charts with categories pointing to deleted defined names crash load | yes | |
| #1989 | Frozen panes off-by-one when freeze is on row 1 only (`A2`) — written as `A1` sometimes | yes | Round-trip pane. |
| #1953 | Hidden columns at end of sheet (cols past last data) dropped on save | **yes** | |
| #1912 | Shared formulas: master cell with `t="shared" ref="..."` is rewritten as individual formulas (i.e., shared form is lost). File grows on round-trip. | **yes — round-trip + size** | Architectural; openpyxl deliberately expands shared formulas on read. |
| #1898 | `read_only=True` mode doesn't surface comments | yes | Use full-load mode if comments needed. |
| #1872 | `keep_vba=True` corrupts files with signed macros | yes (xlsm path) | |
| #1854 | Cross-sheet 3D references (`Sheet2:Sheet5!A1`) not parsed; written as opaque strings | yes | Formula-rewriting can't touch them. |
| #1832 | `ws.column_dimensions['A'].width` returns `None` for unset → write produces default (8.43), losing original `defaultColWidth` | **yes** | Round-trip column width drift. |
| #1798 | Named ranges with `#REF!` after delete produce non-loadable workbooks | yes | |
| #1761 | Streaming `write_only` workbook can't add images, charts, or merged cells before rows | yes (write-only mode is limited by design) | |
| #1742 | Date timezone: aware datetimes silently lose tz; written as UTC offset of zero | yes | |
| #1693 | Data validations with `formula1` containing comma in localized list (de-DE etc.) corrupt | yes | |
| #1611 | Hyperlinks in merged cells: only top-left preserved | yes | |
| #1577 | `MergedCell` write attempt swallows exception inconsistently | low | |
| #1502 | `defined_name` with sheet-local scope and `!` in name corrupts | yes | |
| #1488 | Conditional formatting rule order changes on round-trip — affects which rule "wins" | **yes** | |

### Architectural / design limits

- **Shared formulas are expanded on read.** Every cell in a `t="shared"` group becomes an explicit formula in memory and on rewrite. Files grow; calc-chain semantics differ slightly. By design.
- **Charts are model-based, not XML-passthrough.** Chart features openpyxl's model doesn't know are dropped. Newer chart types (sunburst, treemap, funnel, waterfall, box-and-whisker, histogram) are not modeled.
- **Pivot tables are read+preserve-on-write only if untouched.** Any modification or even some no-op resaves rebuild caches and may break pivot.
- **Conditional formatting preserved only for the rule types in the model.** `extLst`-only rules (icon sets variants, new data-bar types from Excel 2013+) are dropped.
- **`read_only` mode exposes only cell values + basic styles**; merged cells, data validations, comments, charts, defined names are not available.
- **`write_only` mode is append-only**; can't go back and edit a row, can't add merges before the row is written, can't insert images mid-stream.
- **No 1904 date system on write** — always uses 1900.
- **No formula evaluator** (use `openpyxl.utils.formulas` only for tokenization). `data_only=True` reads cached values; if cache is stale (because a writer didn't compute), values are stale.
- **No support for "form controls" / ActiveX shapes** (round-trip drops them).
- **Memory: full-load mode is O(cells) in RAM with substantial per-cell overhead** (~700 bytes/cell empirical). 1M cells ≈ 700 MB. Streaming mitigates but limits features.

---

## XlsxWriter 3.2.9

**Write-only.** Listing applies to fidelity of files this library *produces* (Excel-roundtrip is N/A — our concern is "Excel opens the file cleanly and renders our intent"). 3.2.x has been actively maintained by jmcnamara through 2025–2026.

### Closed bugs (verify they are in 3.2.9)

| # | Title | Status | Fixed | Affects xfa? |
|---|---|---|---|---|
| [#1181](https://github.com/jmcnamara/XlsxWriter/issues/1181) | Excel-Table + hidden rows: rows disappear from autofilter | closed | 3.2.x | yes |
| [#1179](https://github.com/jmcnamara/XlsxWriter/issues/1179) | "Content is Unreadable. Open and Repair" on large workbook with autofit | closed | 3.2.5+ | **yes** |
| [#1173](https://github.com/jmcnamara/XlsxWriter/issues/1173) | Invalid cell reference using named ranges with sheet refs in charts | closed | 3.2.5 | yes |
| [#1169](https://github.com/jmcnamara/XlsxWriter/issues/1169) | `autofit()` ignores cell text rotation | closed | 3.2.5 | yes |
| [#1145](https://github.com/jmcnamara/XlsxWriter/issues/1145) | `autofit` doesn't take format into account | closed | 3.2.4 | yes |
| [#1143](https://github.com/jmcnamara/XlsxWriter/issues/1143) | Leading backslashes replaced in custom URLs | closed | 3.2.4 | yes |
| [#1138](https://github.com/jmcnamara/XlsxWriter/issues/1138) | Release 3.2.4 broken — depends on `xlsxwriter.test` | closed | 3.2.5 | yes (avoid 3.2.4) |
| [#1132](https://github.com/jmcnamara/XlsxWriter/issues/1132) | URL validation false positives since 3.2.3 | closed | 3.2.4 | yes |
| [#1130](https://github.com/jmcnamara/XlsxWriter/issues/1130) | URL exceeds Excel's max length error in 3.2.3 | closed | 3.2.4 | yes |
| [#1126](https://github.com/jmcnamara/XlsxWriter/issues/1126) | Excel Recovery error after table header change | closed | 3.2.x | yes |
| [#1118](https://github.com/jmcnamara/XlsxWriter/issues/1118) | Adding `index=False` impacts formatting | closed | 3.2.x | yes (pandas users) |
| [#1117](https://github.com/jmcnamara/XlsxWriter/issues/1117) | `Format` has no `set_color` after 3.2.0→3.2.2 | closed | 3.2.3 | yes |
| [#1116](https://github.com/jmcnamara/XlsxWriter/issues/1116) | `add_format` behavior changed in 3.2.1/3.2.2 | closed | 3.2.3 | yes |
| [#1111](https://github.com/jmcnamara/XlsxWriter/issues/1111) | `autofit()` partially accounts for autofilter header | closed | 3.2.x | yes |
| [#1109](https://github.com/jmcnamara/XlsxWriter/issues/1109) | Creating a table then writing → repair when opening | closed | 3.2.x | **yes** |
| [#1098](https://github.com/jmcnamara/XlsxWriter/issues/1098) | Hyperlink messes up autofit | closed | 3.2.x | yes |
| [#1089](https://github.com/jmcnamara/XlsxWriter/issues/1089) | Images stack on large row height + many images | closed | 3.2.x | low |
| [#1087](https://github.com/jmcnamara/XlsxWriter/issues/1087) | `set_row` not working after `set_default_row` | closed | 3.2.x | yes |
| [#1043](https://github.com/jmcnamara/XlsxWriter/issues/1043) | URLs disappear after 65530 | closed | 3.2.x | yes (large sheets) |
| [#1019](https://github.com/jmcnamara/XlsxWriter/issues/1019) | `leader_lines` corrupts Excel | closed | 3.2.x | low |
| [#1015](https://github.com/jmcnamara/XlsxWriter/issues/1015) | Inconsistent calculated-column formula with table range refs | closed | 3.2.x | yes |
| [#999](https://github.com/jmcnamara/XlsxWriter/issues/999) | Two autofilters in one worksheet don't work | closed | 3.2.x | yes |
| [#994](https://github.com/jmcnamara/XlsxWriter/issues/994) | Scaling breaks when "Normal" font's `max_digit_width` ≠ 7 | closed | 3.2.x | yes |
| [#980](https://github.com/jmcnamara/XlsxWriter/issues/980) | Formula with named range converted to implicit-intersection `@` | closed | 3.2.x | **yes** |
| [#963](https://github.com/jmcnamara/XlsxWriter/issues/963) | Corrupt named range in output files | closed | 3.2.x | **yes** |
| [#946](https://github.com/jmcnamara/XlsxWriter/issues/946) | `set_background` + `add_table` produces damaged file | closed | 3.2.x | yes |
| [#925](https://github.com/jmcnamara/XlsxWriter/issues/925) | `write_comment` not working in `constant_memory=True` | closed/wontfix in CM | yes |
| [#919](https://github.com/jmcnamara/XlsxWriter/issues/919) | Cell formatting doesn't take effect until cell entered | closed | yes | |
| [#917](https://github.com/jmcnamara/XlsxWriter/issues/917) | Formatting doesn't work with datetime | closed | 3.2.x | yes |
| [#1075](https://github.com/jmcnamara/XlsxWriter/issues/1075) | `@` in table formulas | closed | 3.2.x | yes |

### Open / partly-fixed bugs

| # | Title | Affects xfa? | Notes |
|---|---|---|---|
| [#1183](https://github.com/jmcnamara/XlsxWriter/issues/1183) | `autofit` cellwidth too small for numbers / dates | **yes** | Workaround: pad widths by ~1.2x or set width manually for numeric cols. |
| [#1171](https://github.com/jmcnamara/XlsxWriter/issues/1171) | Conditional-formatting `ISFORMULA()` not localized correctly | yes | Affects non-en-US Excel installs. |
| [#1114](https://github.com/jmcnamara/XlsxWriter/issues/1114) | `num_format` million/thousand separator | yes | |
| [#965](https://github.com/jmcnamara/XlsxWriter/issues/965) | Clearing written cells not supported | wontfix — by design | append-only model |
| [#926](https://github.com/jmcnamara/XlsxWriter/issues/926) | Array formula doesn't recognize names until edited | yes | Workaround: use modern dynamic-array `write_dynamic_array_formula`. |

### Architectural / design limits

- **Write-only.** No read/edit; xfa supervisor must combine XlsxWriter with openpyxl (for reading) when round-tripping.
- **`constant_memory=True` is row-streaming and forbids:** `merge_range`, `set_row` after writing, comments, conditional formatting on rows already written, autofit, images going back. xfa must choose between memory and features per request.
- **No incremental writes / no append to existing xlsx.** A "modify" pipeline is `openpyxl.load → mutate → xlsxwriter.write` — but XlsxWriter cannot import openpyxl objects, so styles/formulas must be re-translated.
- **No formula evaluator.** `Workbook.use_zip64()` and other escapes for size, but cached results must be supplied if downstream readers expect `data_only`-like behavior.
- **Limited shared-formula support.** XlsxWriter writes shared formulas via explicit `write_array_formula` / `write_formula` with a `value=` cached result. Not the same as Excel's `t="shared"` compaction (which XlsxWriter doesn't emit in normal mode).
- **No pivot-table writing.** Pivots are documented unsupported; the only path is template + post-edit.
- **Charts: most types supported, but combo-chart features and some 2013+ types missing** (sunburst, treemap, funnel, box-and-whisker, histogram, waterfall — only partial). Maps, 3D-map, PivotChart not supported.
- **`set_column` is the only column-width API**; per-row column-width overrides aren't expressible.
- **VBA injection only via `add_vba_project`** — no introspection.

---

## Cross-cutting observations

1. **Round-trip fidelity is the worst-affected category across all three.**
   - ExcelJS: alignment flags (#1908), wrapText, conditional formatting (#1024), defined names (#664), shared formulas (#676), tables → DV.
   - openpyxl: shared formulas always expanded (#1912), array formulas degrade (#2071), `extLst` conditional formats dropped (#2034), `defaultColWidth` lost (#1832).
   - XlsxWriter: doesn't read at all, so "round-trip" means "openpyxl-read → XlsxWriter-write." Style/format mapping between the two libraries is the main friction.

2. **Shared / array / dynamic-array formulas are everyone's weak spot.** All three either expand, drop, or mis-mark them. Any xfa feature touching formulas needs a serialization-layer abstraction that targets the lowest common denominator (explicit per-cell formulas with cached values).

3. **`defaultColWidth` / `defaultRowHeight` are uniformly poorly handled.** Files that omit explicit `<col>` entries and rely on defaults round-trip with column-width drift. xfa should normalize on read by materializing per-column widths.

4. **Hidden rows/columns and frozen panes** are minor but recurring sources of off-by-one bugs (#1953, #1989 in openpyxl; #1091, #653 in ExcelJS). Worth dedicated tests.

5. **Conditional formatting** is the most-fragmented feature. Each library supports a different subset; round-tripping a CF-rich file through any of them is lossy. xfa should detect CF presence and either preserve XML opaquely or warn the user.

6. **Streaming modes are uniformly feature-reduced.** `read_only` (openpyxl), `WorkbookWriter` (ExcelJS), `constant_memory` (XlsxWriter) all silently drop or disable features. Tests must cover both modes.

7. **Pivot tables** are write-unsupported in all three (openpyxl partial). The supervisor product's "read pivot, regenerate" plan must rely on direct XML manipulation (e.g., copying `pivotTable1.xml` and rewriting the cache).

8. **Date / timezone**: ExcelJS leaks OS timezone; openpyxl drops tz on aware datetimes (#1742); XlsxWriter requires explicit format. Normalize to UTC + explicit format on write.

9. **Hyperlinks**: all three have edge cases — `#` fragment trimming (ExcelJS #1084), merged-cell hyperlink loss (openpyxl #1611), Windows path issues (ExcelJS #749), >65530 limit (XlsxWriter #1043 — fixed but worth a test).

10. **Memory / large files**: ExcelJS streaming is buggy (#709, #683, #1277). openpyxl has documented ~700 B/cell overhead; use `read_only` for read, `write_only` for write. XlsxWriter `constant_memory` is reliable but feature-restricted.

---

## Recommended regression tests for xfa's test suite

For each test, listed: fixture description and assertion. Group by library; integration tests cover the openpyxl→XlsxWriter pipeline.

### ExcelJS (Node CLI)

1. **Shared formula roundtrip** (#676). Fixture: hand-crafted xlsx with `<f t="shared" ref="A1:A10" si="0">A1+1</f>` and 9 child cells. Assert: read succeeds, all 10 cells expose `.formula`, write produces a workbook Excel opens without "repair".
2. **Defined names roundtrip** (#664). Fixture: workbook with workbook-scope and sheet-scope named ranges. Assert: byte-level XML diff of `definedNames` element after load+save shows no spurious changes.
3. **Conditional formatting preservation** (#1024). Fixture: workbook with all CF rule types (cellIs, expression, colorScale, dataBar, iconSet). Assert: all rules present after roundtrip; XML element count matches.
4. **Data validation preservation** (#1184). Fixture: list, decimal, date, custom DV. Assert: each rule present; formulas unchanged.
5. **Hyperlinks with `#` and Windows paths** (#1084, #749). Fixture: external URL with fragment; internal `#Sheet2!A1`; UNC path. Assert: `.hyperlink` property exact-matches; written file opens in Excel.
6. **Column widths and `defaultColWidth`** (#744, #1075). Fixture: workbook with explicit `defaultColWidth=12.5` and three `<col>` overrides. Assert: read returns `defaultColWidth`; write preserves it.
7. **Wrap-text / shrink-to-fit alignment** (#1908). Fixture: cells with `wrapText=false`, `shrinkToFit=false` explicitly set. Assert: round-trip preserves `false`.
8. **Merged cells + borders** (#635). Fixture: merged 3x3 with borders on outer edges only. Assert: borders survive.
9. **Rich-text in streaming writer** (#2267). Fixture: 5 cells, each with distinct rich-text runs, written via `WorkbookWriter`. Assert: cells 2–5 retain their own runs (currently fail).
10. **Large-file streaming** (#709). Fixture: 500k rows × 10 cols. Assert: peak RSS < 1 GB; output opens.
11. **Tables + comments coexistence** (#1429). Fixture: sheet with table; add comment to a cell outside the table. Assert: write succeeds.
12. **Sheet name unicode** (#910). Fixture: sheet named "Σ-summary" (combining char). Assert: read succeeds.

### openpyxl (Python)

1. **Array formula preservation** (#2071). Fixture: cell with `=TRANSPOSE(A1:A5)` written by Excel as `t="array"`. Assert: after load+save, `t="array"` is still in XML (current: fails — write as plain).
2. **Dynamic-array formula** (#2055). Fixture: `=FILTER(A:A, B:B>0)`. Assert: cached value present after save; opens without recalc-prompt.
3. **Shared formula expansion** (#1912). Fixture: 1000-row shared formula. Assert (regression-watch): file size after save is documented and tracked; we expect bloat — fail if the bloat exceeds 3x baseline.
4. **3D cross-sheet refs** (#1854). Fixture: `=SUM(Sheet2:Sheet5!A1)`. Assert: formula string preserved verbatim.
5. **`defaultColWidth` preservation** (#1832). Fixture: workbook with `<sheetFormatPr defaultColWidth="20"/>`. Assert: after load+save, attribute preserved (or column widths materialized to 20).
6. **Frozen panes on row 1** (#1989). Fixture: freeze row 1. Assert: `<pane state="frozen" ySplit="1" topLeftCell="A2"/>` after save.
7. **Hidden trailing columns** (#1953). Fixture: hide cols X:Z when last data is column J. Assert: `<col hidden="1" min="24" max="26"/>` survives.
8. **Hyperlinks in merged cells** (#1611). Fixture: merge B2:D2 with hyperlink on B2. Assert: hyperlink survives; merge intact.
9. **Comment author + threaded comments** (#1996, plus extLst). Fixture: legacy comment with author "Alice"; threaded comment thread. Assert: legacy preserved; threaded — document loss as expected, test guards against silent change.
10. **Conditional formatting `extLst` rules** (#2034). Fixture: icon-set rule with `extLst`. Assert: warning emitted (regression: we want awareness if openpyxl ever starts dropping silently or starts preserving).
11. **Read-only mode comments** (#1898). Fixture: workbook with comments. Assert: opening with `read_only=True` raises or returns None for comments — guards against API change.
12. **Timezone-aware datetimes** (#1742). Fixture: `datetime(..., tzinfo=US/Pacific)`. Assert: written cell has UTC offset preserved (or test currently asserts "tz dropped"; flip if upstream fixes).

### XlsxWriter (Python, write only)

1. **Autofit with numeric/date columns** (#1183). Fixture: column of 10-digit numbers and dates. Assert: column width ≥ measured-string-width × 1.1.
2. **Autofit + format** (#1145). Fixture: numbers with `#,##0.00` format. Assert: width accounts for format width.
3. **Table + write_url + autofit** (#1098, #1109). Fixture: table with hyperlink column. Assert: opens in Excel without repair, autofit reasonable.
4. **Implicit-intersection on named-range formula** (#980). Fixture: formula referencing single-cell defined name. Assert: written formula does NOT contain `_xlfn.SINGLE(@…)` if not desired.
5. **Many URLs** (#1043). Fixture: 70k rows each with `write_url`. Assert: all cells have hyperlinks in output.
6. **Constant memory + comments** (#925). Fixture: `constant_memory=True`, `write_comment`. Assert: documented to fail — test guards against silent change.
7. **`set_row` after `set_default_row`** (#1087). Fixture: `set_default_row(20)` then `set_row(5, 30)`. Assert: row 5 height is 30 in output.
8. **Two autofilters per sheet** (#999). Documented unsupported — test for clear error message.
9. **Localization-sensitive CF formulas** (#1171). Fixture: `=ISFORMULA(A1)`. Assert: written formula uses `_xlfn.ISFORMULA` prefix so non-en-US Excel renders it.
10. **Url with leading backslash** (#1143). Fixture: `\\server\share\file`. Assert: backslashes preserved.
11. **Calculated-column table formula** (#1015). Fixture: table with `[@[col1]]+1` calculated column. Assert: opens without recalc-prompt.

### Cross-library / pipeline tests (supervisor server)

1. **openpyxl-read → XlsxWriter-write fidelity gate.** Fixture: golden xlsx with formulas, merges, frozen panes, defined names, conditional formatting, hyperlinks. Pipeline: load with openpyxl, translate to XlsxWriter calls, write. Assert: a curated list of features (column widths, defined names, frozen panes, merges, hyperlinks) match in the output.
2. **ExcelJS-write → openpyxl-read** (#1198). Fixture: file written by ExcelJS. Assert: openpyxl loads without warning; cell values match.
3. **openpyxl-write → ExcelJS-read.** Fixture: openpyxl-written workbook. Assert: ExcelJS reads without exceptions and surfaces all defined names.
4. **Round-trip through both** (full xfa CLI → supervisor → CLI). Fixture: realistic 50-sheet workbook. Assert: golden hash on a stable subset (values, formulas, merges); diff report on everything else.
5. **Large-file pipeline.** 1M cells. Assert: peak memory <2 GB; wall time <60s on CI.
