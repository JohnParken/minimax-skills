# Read-only data analysis and file identification

Resolve `<skill_dir>` from the actual SKILL.md path before running commands. Run the
[entrypoint self-check](../SKILL.md#执行环境约束与首步自检) first, require exit 0 and
`status: ready`, then confirm input readability. Use Python 3.9+ with `-I -S -B`;
no third-party packages, Office application, network or output-file permission is needed.
Run from the confirmed task directory with task-relative inputs. Commands emit JSON to
stdout and never modify inputs. Redirect stdout only to a new task report if needed.

## Profile a sheet or region

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_analysis.py" profile input.xlsx --sheet Sales
python3 -I -S -B "<skill_dir>/scripts/xlsx_analysis.py" profile input.xlsx --sheet Sales --range B3:F100 --no-header
```

By default the first selected row supplies headers and is excluded from statistics.
Column letters remain absolute coordinates. For explicit multi-level rows and path/ID
selectors, use [multi-level-headers.md](multi-level-headers.md).
Duplicate/empty header labels do not affect selection. Without a range, use A1 through
the maximum stored row/column. Explicit regions include sparse blank cells.

Reports include column types, missing cells, empty text, blank totals, distinct and
repeated nonblank values, duplicate rows, numeric count/sum/min/max/mean. Duplicate
counts exclude the first occurrence. Whole-row duplicates include blank rows. Whitespace
is text, not blank. Missing cells and empty strings are counted separately. Grouping and
duplicates distinguish text, number, boolean, blank and ISO date text; number 1 and 1.0
are equal, but true and text "1" are distinct. Text comparison is case-sensitive.

Numeric statistics use stored Decimal numbers only. Text "12" and booleans are never
coerced. Date serials remain numeric; number formats are not used to infer dates. Cells
stored as ISO dates have type `date_text`, without calendar conversion or validation.
JSON Decimal values are strings with accompanying type metadata. Sums are exact within
1000-digit arithmetic; exceeding that limit fails. Means use 34 significant digits and
ROUND_HALF_EVEN with `mean_rounded` reporting inexact division. Empty numeric sets return
null statistics, not zero. Hidden/filtered rows are included; merges are not expanded.

## Filter and summarize

Create task-local `query.json`:

```json
{"group_by":["A"],"measure":"C","mixed":"reject",
 "filters":[{"column":"C","op":"ge","value":10}]}
```

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_analysis.py" summarize input.xlsx --sheet Sales --query query.json
```

`group_by` is a list of unique column selectors inside the selected range; omit it for
one aggregate group. Omit `measure` for row counts only. One numeric measure is supported.
`filters` are combined with AND; eq/ne use typed equality, gt/ge/lt/le require JSON
numeric thresholds and select numeric cells only. No string-to-number conversion, expression evaluation, OR expressions or date comparisons
occur. Header path lookup requires explicit header rows; plain label guessing is not used.
Unknown keys and duplicate JSON keys fail. Null filters match missing cells; an empty
string matches empty text. ISO date cells do not match ordinary text filters.

A nonblank, nonnumeric measure fails with its cell address by default. Choose
`"mixed":"ignore"` explicitly to omit it from numeric statistics; reports count omitted
and blank measures per group. Groups appear in first-seen order with typed keys. Blank
and empty-text keys form separate groups. Reports show selected/excluded row counts;
no matches produce an empty group list. Header-only input also yields no groups.

Both analysis commands reject formulas by default, including formulas in header cells.
`--formulas cached` explicitly uses existing caches and reports their number as
`unverified_formula_caches`; missing caches and Excel error cells fail. No recalculation
is performed, and cache freshness is never assumed. These checks cover the selected
rectangle before filtering, so a filter cannot silently hide erroneous inputs.
Limits: 200,000 selected cells including blanks, 10,000 groups, query file 64 KiB,
and the underlying workbook package limits. This is bounded in-memory analysis.
Exit 0 indicates completed analysis; exit 1 returns JSON `status: failed`.

## Identify contents, including ET/ETT candidates

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_identify.py" input.et
```

The inspector reads at most 32 MiB plus one limit-detection byte. ZIP limits are 10,000
members and 128 MiB declared uncompressed data. It checks ZIP member safety, content
types and the maintained OOXML worksheet indexer. `supported_subset` means that indexer
accepted matching XLSX/XLSM contents; it does not validate formulas or application layout.
`extension_mismatch` can identify OOXML contents under another suffix, with the expected
extension reported; the maintained reader still rejects that suffix. A ZIP signature
alone never yields spreadsheet support. Templates or other OOXML document types may be
unsupported. Unrecognized data is not automatically declared corrupt.

`compound_binary` checks only basic compound-container header consistency. Streams are
not decoded: content might be a binary spreadsheet, encrypted OOXML or another document.
Truncated/invalid headers are reported; a consistent header does not prove a valid file.
Native ET/ETT parsing is not implemented, and reports always state that explicitly.
For a trusted file, open a copy locally in WPS Spreadsheets, Save As a new XLSX, preserve
the original template, then verify sheets, values, formulas and formatting. Renaming
is not conversion. No application is launched and no conversion happens automatically.
The inspector exits 0 for a completed diagnosis, including unsupported/invalid contents;
inspect JSON `status` and `parser_supported`. File I/O failure returns exit 1.
