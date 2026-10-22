# Match two tables by typed composite keys

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检), requiring exit 0 and
`status: ready`. Confirm the task directory and both input files are readable. Resolve
`<skill_dir>` from the actual SKILL.md path before execution; use task-relative input
paths. Python 3.9+, standard library only. Neither source is modified; JSON goes to stdout.

Task-local `match.json`:

```json
{
  "left":{"sheet":"Sales","keys":["A","B"]},
  "right":{"sheet":"Targets","keys":["A","B"]},
  "relationship":"one_to_one",
  "how":"left",
  "nulls":"never",
  "max_rows":10000
}
```

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_match.py" sales.xlsx targets.xlsx --query match.json
```

Two sheets in one workbook can be matched by supplying the same file twice. Each side
accepts `sheet`, `keys`, optional `range`, `header` (default true), `header_rows`,
`header_exclude`, and `include_spanning_headings`. Explicit multi-level headers and
column path/ID selectors use [multi-level-headers.md](multi-level-headers.md). Key lists
must have the same nonzero length; component order defines the composite key.

An explicit `relationship` is required:

- `one_to_one`: both sides' eligible keys must be unique.
- `one_to_many`: left keys must be unique; repeated right keys yield multiple pairs.
- `many_to_one`: right keys must be unique; repeated left keys are allowed.

Violations return exit 1, `status: failed`, no result rows, and bounded duplicate-key
row-number diagnostics on both sides. Uniqueness is checked for all eligible input keys,
not only keys that eventually match. Many-to-many matching is deliberately unsupported.

`how` is left (default), inner or full. Full appends unmatched right rows. Matched output
follows left source order and then right source order per key. Each record retains both
original row numbers and separate typed cell arrays; a missing side is null. Reports
include unmatched-row counts and up to 20 sample row numbers. Column/header metadata
is scoped separately to each input, avoiding duplicate-label or col:A naming collisions.

Keys use exact stored types and values: numeric 1 and 1.0 match, text "1", true and
numeric 1 do not. Text case, whitespace, leading zeros and long identifiers are preserved.
No fuzzy matching, trimming, text conversion, arithmetic, rounding or date-format inference
occurs. Numeric date serials remain numbers and ISO date cells remain date_text.

Any key containing a missing cell or empty text is ineligible to match by default
(`nulls: never`), remains unmatched, and is counted in null-key diagnostics. With explicit
`nulls: match`, null-containing keys participate, but missing cells and empty text remain
different typed keys. Repeated null keys may then violate the selected relationship.

Formula policy is shared across both sides: default reject; `"formulas":"cached"` uses
existing unverified caches and reports them independently. Missing caches and Excel
errors fail. Header and whole selected-region checks occur before matching. No formulas
are calculated. Inputs and output use the existing raw-value, hidden-row and merge rules.

Queries use bounded 64 KiB reads. Existing workbook/region limits apply per input.
`max_rows` must be 1..10000; output is also limited to 200,000 combined left/right cells,
counting unmatched-side width. Exceeding either budget fails before any JSON result is
emitted; narrow input ranges rather than permitting uncontrolled multiplication.
Exit 0 plus `status: matched` indicates completed matching, including zero matches.
Inspect diagnostics before treating unmatched keys as missing business records. No
real Excel/WPS validation or native ET/ETT support is implied.
