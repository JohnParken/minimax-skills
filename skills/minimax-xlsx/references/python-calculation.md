# Standard-library calculation and file contract

All commands run from the task directory. Replace `<skill_dir>` with the confirmed
skill directory. Python 3.9+; no pip packages, external applications or network access.

## Contents

- [Commands](#commands)
- [Input JSON](#input-json)
- [Formula subset](#formula-subset)
- [Results and exit codes](#results-and-exit-codes)
- [Known business computations without Excel formulas](#known-business-computations-without-excel-formulas)

## Commands

```console
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" check
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" read input.xlsx
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" read input.csv --encoding utf-8-sig
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" create report.json report.xlsx
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" create report.json static.xlsx --values-only
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" edit original.xlsx patch.json edited.xlsx
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" calculate original.xlsx calculated.xlsx
```

Output files must not exist; original inputs are never overwritten. Preserve the
`.xlsm` extension when editing a macro-enabled workbook. Creation is XLSX only.
Redirect command stdout to a task report file if a persistent JSON report is needed.

## Input JSON

A runnable example is bundled at [calculation_example.json](../assets/calculation_example.json).

Both create and edit use this shape. Edit accepts only existing sheet names and patches
only listed cells; it does not replace the sheet's remaining data.

```json
{
  "sheets": [
    {
      "name": "Sales",
      "cells": {
        "A1": {"value": "Quantity", "style": 4},
        "B1": {"value": "Unit price", "style": 4},
        "C1": {"value": "Discount factor", "style": 4},
        "D1": {"value": "Amount", "style": 4},
        "A2": {"value": 12},
        "B2": {"value": 19.90},
        "C2": {"value": 0.95},
        "D2": {"formula": "ROUND(A2*B2*C2,2)", "style": 2}
      }
    },
    {
      "name": "Summary",
      "cells": {
        "A1": {"value": "Total"},
        "B1": {"formula": "Sales!D2", "style": 3}
      }
    }
  ]
}
```

Expected numeric results: Sales!D2 and Summary!B1 both 226.86. Styles 2/3 use General;
use or add a suitable number format separately when fixed display decimals are required.
Each cell has exactly one of `value` and `formula`, plus an optional integer `style`.
Duplicate JSON keys, duplicate case-insensitive cell/sheet names, unknown keys, invalid
addresses and invalid styles are rejected. Formula strings may start with `=`.
`value` accepts null (clear), a number, boolean or string. A string beginning `=` remains
literal text. JSON decimal numbers are loaded as Decimal, not binary floats. Dates must
be deliberately supplied as text or serial numbers; no implicit date conversion occurs.

For new files, template style indices are: 0 default, 1 input, 2 formula, 3 cross-sheet,
4 header, 5 currency input, 6 currency result, 7 percentage input, 8 percentage result,
9 integer input, 10 integer result, 11 year, 12 highlighted assumption. Currency styles
5/6 display whole dollars. For existing files, indices refer to that file's own styles.
Omitting style preserves an existing cell's style. New-workbook [appearance presets](readable-appearance.md) can append shared readable
styles and formats; existing-file edits cannot create arbitrary custom styles.

## Formula subset

- Decimal numeric literals, double-quoted strings (`"a""b"` escapes quotes), TRUE/FALSE.
- Parentheses, unary +/-; binary +, -, *, / and =, <>, <, <=, >, >=.
- A1 references, optional `$`, ranges, case-insensitive sheet lookup, quoted sheet names
  (`'Sales Data'!A1`) including doubled apostrophes.
- SUM, COUNT, AVERAGE, MIN, MAX, IF (three arguments), ROUND (two arguments).
- IF evaluates only its chosen branch. Dependency cycles in evaluated paths are errors.
- Decimal precision 28; ROUND uses ROUND_HALF_UP (ties away from zero), integer digits.
  This is an explicit Python calculation contract, not a guarantee of Excel floating
  point, locale, type-coercion or date semantics.
- Blank arithmetic values are zero; direct blank-reference formula results become zero.
  Booleans in arithmetic are 1/0. Text-to-number coercion is unsupported.
- Aggregates ignore text, blanks and booleans inside ranges. Scalar text/boolean aggregate
  arguments are unsupported to avoid ambiguous reference-vs-literal coercion. Use ranges
  for mixed data. An empty numeric AVERAGE is an error; empty SUM/MIN/MAX returns zero.
- Text comparisons are case-insensitive; mixed text/numeric and mixed boolean/non-boolean
  comparisons are unsupported. Boolean arithmetic coercion does not apply to comparisons.

Not supported: names, structured/table references, external links, whole-column ranges,
3D ranges, shared/array/data-table formulas, spill functions, exponent notation/operator,
percent/concatenation operators, omitted arguments, Excel error literals and functions
outside the list. Unsupported expressions preserve formulas without fabricating results.
Workbooks using precision-as-displayed calculation report their formulas as unsupported.
Malformed expressions and existing literal error cells block output. Formula error caches
are recomputed rather than used as inputs. ISO-date arithmetic is unsupported; date serial
arithmetic remains numeric and does not imply calendar semantics.

Limits: formula 8192 characters, one range 100,000 cells, package 128 MiB uncompressed,
10,000 ZIP members, 200,000 stored cells, dependency depth 100, bounded syntax nesting.
This is an in-memory tool, not a streaming engine for arbitrarily large workbooks.

## Results and exit codes

| Exit | Status | Meaning |
|---|---|---|
| 0 | `ready` | Runtime self-check succeeded |
| 0 | `calculated_subset` | All formula cells calculated under this engine's subset; may be zero formulas |
| 0 | read result | Inspection succeeded; old formula caches remain unverified |
| 2 | `incomplete` | Unsupported formulas exist; output written with their caches removed |
| 1 | `failed` | Input/runtime/calculation failure, or incomplete values-only request; no new output |

Reports include formula_count, calculated_count, issues (sheet/cell/code/message),
output_written and excel_equivalence_verified=false. Check both exit and JSON; a file
existing does not mean every formula was calculated. Copy calculated results only when
business expectations are checked. Do not claim a full Excel recalculation.

An unsupported dependency propagates to evaluated dependents. All formula cells are
visited, so a separately stored erroneous formula still fails even if another IF skips
referencing it. Known errors/cycles prevent writing. On incomplete output, formulas
outside the subset remain intact but have no cached result. Values-only mode is available
only for creation and refuses any incomplete formula set.

Recalculation flags request a future spreadsheet application to recalculate; they do not
perform calculation. Formula caches may change on every worksheet. Chart/pivot and
external-link caches are not refreshed: their retained contents may display old data.

## Known business computations without Excel formulas

Use Decimal from strings for currency; explicitly choose rounding timing and policy.
Use `csv`, `collections`, `statistics` and ordinary Python loops for analysis. Write
computed static results through `value`, and explain their non-updating nature.
Do not wrap an arbitrary formula string in Python eval/exec.

The Python API is `Workbook.create(spec, appearance=None)`, `Workbook.load(path)`, `patch(spec)`,
`calculate(values_only=False)`, `save(path)`, `inspect()`. Import from the resolved scripts
directory. `calculate` returns a report; do not call save if status is failed. It rewrites
formula caches but does not refresh the in-memory inspection index; reopen the saved
file to inspect results. Caller-supplied caches are never accepted through the JSON API.

JSON input exceeding the parser nesting depth returns a structured failure. Bounded query
readers also enforce byte limits on actual reads, not only filesystem size metadata.
