# Explicit monthly period comparisons

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检), require exit 0 and
`status: ready`, and confirm task directory/input readability. Resolve `<skill_dir>` from
the actual SKILL.md path and replace it before running task-relative input paths.
Python 3.9+, standard library only. Inputs are unchanged; JSON is written to stdout.

Task-local `periods.json`:

```json
{"entity_keys":["A"],"period":"B","metrics":["C"],
 "period_format":"YYYY-MM","comparisons":["mom","yoy"],
 "non_numeric":"reject","precision":34}
```

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_advanced.py" compare history.xlsx --sheet Sales --query periods.json
```

Entity, period and metric selectors accept column letters, col: IDs or exact paths under
[explicit multi-level headers](multi-level-headers.md). Supply --header-rows where needed.
Entity keys are an explicit list; [] deliberately treats the table as one entity. Key
components use exact stored types, preserving leading-zero and long text identifiers.
Missing/empty entity keys fail. Metric selectors must be unique; all selected records
participate, including hidden rows. Aggregation, filtering and duplicate suppression are
not automatic. An entity/month duplicate fails with both source row numbers, even if
its values agree; aggregate deliberately upstream if business rules require that.

Only explicit `period_format: YYYY-MM` is supported. Period cells must be literal text
with four-digit year 0001..9999 and two-digit month 01..12. Numeric YYYYMM, Excel serials,
ISO date cells, quarters and fiscal calendars are rejected rather than guessed. Period
keys are indexed independently of source order. Entity keys remain exact typed keys.

`mom` compares each observed month with exactly one calendar month earlier; `yoy` uses
exactly twelve months earlier. Missing months return `missing_period`; the previous
available row is never substituted. A month before year 0001 has null base_period and
is missing. The tool does not synthesize missing current months or compare year-to-date
aggregates. Output follows original row order, then requested comparison order, with
both source row numbers and resolved current/base periods for audit.

For numeric values, difference=current-base; growth_rate=(current-base)/base is a
fraction, not a percentage. A zero base keeps the numeric difference but returns null
growth_rate and `zero_base`, including 0/0. Negative bases are allowed, flagged as
negative_base, and retain a signed denominator. For -100 to -50 the difference is +50
and growth_rate is -0.5; that sign alone must not be interpreted as business decline.
No absolute-denominator adjustment or unrequested business interpretation is performed.

Missing current/base metrics produce current_missing/base_missing with null calculations.
`non_numeric: reject` (default) fails for any nonblank text, boolean or ISO-date metric;
explicit diagnose returns current_non_numeric/base_non_numeric per comparison. No text
numbers are coerced and no missing values are filled. When several issues apply, current
missing/type issues precede missing-period and base-value issues. All records are checked
before output. Invalid month/entity/duplicate errors always fail the entire operation.

Differences use exact Decimal arithmetic up to 1000 digits; exceeding that limit fails.
Growth uses precision 16..100 significant digits (default 34), 16 guard digits and
ROUND_HALF_EVEN. Each growth result has value/rounded; rounded includes inexact
intermediate division. Decimal numbers are JSON strings. Date serials in metric columns
remain numbers, so choose columns deliberately. This is arithmetic comparison, not a
seasonality model, significance test, forecast or Excel-equivalence guarantee.

Formula policy defaults to reject for the selected rectangle, including headers.
--formulas cached uses explicitly unverified caches; missing caches and Excel errors
fail. Period cache values must still be YYYY-MM text. No recalculation occurs. Query
reads are bounded to 64 KiB; existing package/region limits apply. At most 20,000 metric
results (observed rows * metric columns * comparison kinds) are allowed.
Exit 0/status compared means processing completed, not that all comparisons are valid;
inspect per-metric status. Exit 1/status failed reports rejected input without result rows.
