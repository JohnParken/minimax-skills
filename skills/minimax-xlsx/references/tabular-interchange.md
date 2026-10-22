# CSV and TSV interchange

Use this path to import a rectangular text table into a new XLSX file or export raw
values from one worksheet. Run the self-check in [../SKILL.md](../SKILL.md) first.
Python 3.9+, standard library only; no Office application or network is used.
Resolve `<skill_dir>` from the actual SKILL.md path and substitute it before running
commands from the confirmed task directory. Inputs and outputs are task-relative.

## Import a table

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_tabular.py" import input.csv imported.xlsx --sheet Sales
python3 -I -S -B "<skill_dir>/scripts/xlsx_tabular.py" import input.tsv typed.xlsx --sheet Sales --numeric-columns B C --header-rows 1 --encoding gbk
```

The default encoding is UTF-8 with optional BOM. Delimiter comes from `.csv` or `.tsv`;
there is no encoding, delimiter, date or number guessing. Quoting and embedded newlines
are handled by Python's strict CSV parser. Every field is text by default, including
leading zeros, long identifiers and strings beginning `=`. No formulas are created.

For explicit numeric columns use letters (B, AA). Header rows remain text; the default
is one header row. Use `--header-rows 0` for headerless data. Empty fields stay empty
strings. Numeric syntax accepts decimal or scientific notation, not grouping separators,
currency symbols or percentages. Invalid numeric cells fail with their address.
More than 15 significant digits are refused in numeric columns; use text for identifiers.
Numeric display is General: `19.90` may display as `19.9`, although the stored number
is preserved. Text fields retain original whitespace and spelling.

Example input:

```csv
id,amount,note
00123,19.90,=1+2
```

Import with `--numeric-columns B`; A2 remains `00123`, B2 is numeric 19.90, C2 is
literal text `=1+2`. Existing output files are never overwritten. Empty inputs,
inconsistent row widths, out-of-table numeric columns and invalid XML text are rejected.
Input is bounded to 16 MiB, 200,000 cells, and Excel row/column limits; this is not a
streaming importer. Rows containing only delimiters represent empty fields; empty
physical rows are not silently discarded.

## Export one sheet or rectangle

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_tabular.py" export imported.xlsx values.csv --sheet Sales
python3 -I -S -B "<skill_dir>/scripts/xlsx_tabular.py" export imported.xlsx region.tsv --sheet Sales --range A1:C20
```

Select the exact sheet name. Without a range, export A1 through the maximum stored
row/column. Sparse cells become empty fields. A range is one cell or an ascending
rectangle such as B2:D10; export is limited to 200,000 cells including blanks.
Output uses UTF-8 BOM, CRLF line endings, standard CSV quoting, and the delimiter
selected by the output suffix. Workbook input is never modified.

Exports are raw values: formatting, charts, merged-cell display, hyperlinks, date
formats and formulas are not transferred. Date serials stay numeric; booleans become
TRUE/FALSE. Literal Excel errors block export rather than being mistaken for data.

By default, any formula in the selected range blocks export. `--formulas cached`
explicitly exports existing caches, refuses missing caches, and reports the number
as `unverified_formula_caches`. It does not recalculate or assert freshness. Use the
separate calculation workflow and inspect its report when updated results are needed.

The default `--text safe` prefixes an apostrophe to potentially executable text
(beginning with =, +, -, @ after leading whitespace, or a leading tab/CR/LF). The
report counts changed fields. This changes text; it reduces spreadsheet formula
interpretation risks but is not a guarantee for every consumer. For exact data
interchange with a trusted consumer, `--text raw` preserves text without that prefix.
CSV has no type metadata: import through this skill to retain leading zeros; opening
CSV directly in spreadsheet software can infer types and change values.

## Verification and diagnostics

Exit 0 plus `output_written: true` indicates successful conversion. Exit 1 and JSON
`status: failed` describes rejected data or I/O; no new output is retained. Reopen
XLSX with `stdlib_xlsx.py read` and compare representative strings and numeric cells.
For CSV/TSV, read with `csv.reader` using the documented delimiter and encoding.
Neither conversion nor ZIP/XML validation constitutes Excel/WPS application testing.
