# Filter, deduplicate and sort typed rows

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检) first; require exit 0
and `status: ready`. Confirm the task directory and input readability. Resolve
`<skill_dir>` from the actual SKILL.md path and replace it before execution. Use task-relative
input paths from the task directory, Python 3.9+ and `-I -S -B`. No packages, application,
network or workbook-write permission is required. The command emits JSON to stdout.

Create task-local `rows.json`:

```json
{
  "filters": [{"column":"C","op":"ge","value":10}],
  "deduplicate": {"columns":["A","B"],"keep":"first"},
  "sort": [{"column":"C","descending":true,"blanks":"last"}]
}
```

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_rows.py" input.xlsx --sheet Sales --query rows.json
python3 -I -S -B "<skill_dir>/scripts/xlsx_rows.py" input.xlsx --sheet Sales --query rows.json --range A3:C100 --no-header
```

Use absolute column letters or explicit IDs. Multi-level path selectors require
[explicit header rows](multi-level-headers.md); duplicate paths are rejected.
The first selected row is a header unless `--no-header` is supplied. Processing order
is filter, deduplicate, then sort. Each stage is optional; `{}` returns all data rows.
Output contains typed cells, column/header metadata, original worksheet row numbers,
excluded-row and removed-duplicate counts. Inputs are never modified. This is a JSON
data extraction result, not an edited XLSX preserving worksheet layout.

## Deduplication

Set a nonempty `columns` list. Composite keys compare exact typed values. `keep` is:

- `first`: keep the first source row for each key (default).
- `last`: keep the last source row for each key.
- `unique`: remove every row whose key appears more than once.

Retained rows remain in source order before sorting. Numeric 1 and 1.0 share a key;
boolean true, text "1", missing cells and empty text are distinct keys. Leading zeros,
long identifiers, text case and whitespace remain unchanged. No trimming, case folding,
text-to-number conversion or fuzzy matching occurs. A blank row is still a data row.

## Sorting and filtering

`sort` is a list of unique column specifications, highest priority first. `descending`
is a JSON boolean (default false). `blanks` is first or last (default last), independently
of direction. Both missing cells and empty text form the blank partition, retaining
their stable tie order. Multi-column sorting is stable; equal full keys preserve source
order. Nonblank ascending type order is boolean, number, ISO date text, ordinary text;
descending reverses that order and values. Numbers compare as Decimal, strings by Unicode
code points, not locale collation. ISO dates sort as stored text, not calendar instants.

Filters use the same typed AND semantics as [data-analysis.md](data-analysis.md).
eq/ne distinguish types; gt/ge/lt/le require a JSON numeric threshold and select numeric
cells only. No Python expressions are evaluated. Unknown keys, duplicate JSON keys,
unknown/out-of-range columns and invalid options fail rather than being ignored.

## Formula, date and precision boundaries

Formulas in the whole selected rectangle are rejected unless `--formulas cached` is
chosen. That option uses existing values, reports unverified caches, and refuses missing
caches or Excel error cells. No formula calculation occurs, even if a filtered row would
later be removed. Empty results are valid and return zero output rows.

Stored date serials remain numbers without format inference. ISO date cells keep the
`date_text` tag. No arithmetic or rounding is performed by this command. Decimal values
serialize as JSON strings with explicit type tags; text numbers retain their text tag.
Hidden/filtered worksheet rows are included, and merged cells are not expanded.

The selected rectangle is limited to 200,000 cells including blanks; the query is limited
to 64 KiB by a bounded read. Existing workbook package limits also apply. Processing is
in memory. Exit 0 and `status: selected` mean successful extraction; exit 1 returns JSON
`status: failed`. Check counts, typed keys and representative original row numbers before
using the result. This does not establish real Excel/WPS compatibility.
