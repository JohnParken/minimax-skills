# Descriptive statistics with explicit Decimal semantics

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检) and require exit 0 plus
`status: ready`. Confirm task directory/input readability. Resolve `<skill_dir>` from the
actual SKILL.md path and replace it before execution; use task-relative input paths.
Python 3.9+, standard library only. No application/network is used; inputs are not modified.

Create task-local `stats.json`:

```json
{"columns":["C",{"path":["2024年","销售","金额"]}],
 "missing":"omit","non_numeric":"reject","variance":"both",
 "quantiles":[0,0.25,0.5,0.75,1],"precision":34}
```

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_advanced.py" describe report.xlsx --sheet Sales --header-rows 3:5 --query stats.json
```

Columns accept letters, col: IDs or exact typed paths; selectors must resolve uniquely.
Multi-level paths require explicit header rows as in [multi-level-headers.md](multi-level-headers.md).
Use ordinary letters without --header-rows for the legacy one-header-row layout. Region,
header exclusions, spanning-heading choices and --no-header follow existing semantics.

Reports give numeric count, mean, median, requested quantiles, variance and standard
deviation. `missing` is omit (default) or reject; both missing cells and empty strings
are missing. Omission counts are reported. `non_numeric` is reject (default) or omit;
text numbers, whitespace, booleans and ISO-date text are never silently converted.
Date serials remain numeric because display formats are not inferred. Repeated values
are observations, not deduplicated records. No grouping/filter stage is implicitly applied.

Quantiles use **Hyndman-Fan type 7**: sort ascending, set h=(n-1)*p using zero-based ranks,
and linearly interpolate adjacent observations. p=0/1 returns min/max. Median is the same
algorithm at p=0.5, independently of requested quantiles. Probabilities must be unique
JSON numbers in [0,1], at most 21. An empty quantile list is valid; median still appears.
This is one declared algorithm, not a promise that every spreadsheet percentile variant
uses identical rules. No quantile extrapolation or approximate streaming quantiles occur.

`variance` is sample, population or both (default). Population divides squared deviations
by n (ddof=0); sample uses n-1 (ddof=1). At least one numeric observation is needed for
mean/median/quantiles/population variance, and two for sample variance. Zero observations
produce `insufficient_numeric_samples` with null metrics; one observation has population
variance/stddev zero and sample metrics null. Null is never silently replaced by zero.

Exact sums, squared sums, variance numerator and quantile interpolation use a 1000-digit
Decimal context with inexact arithmetic trapped. Large inputs that exceed that bound fail
rather than silently round. Variance uses the exact numerator n*sum(x²)-sum(x)²; its
intermediate exact arithmetic avoids cancellation caused by rounded means. Variance
normalization uses denominator n² or n*(n-1), then Decimal division and square root.

`precision` is 16..100 significant digits (default 34). Mean, quantiles, variance and
standard deviation use ROUND_HALF_EVEN; division/square-root use 16 guard digits before
final rounding. Each computed metric carries `{value,rounded}`; rounded is true when
intermediate or final arithmetic was inexact. Decimal values serialize as JSON strings.
Precision/guard/rounding metadata is included. This bounded Decimal contract is not an
assertion of Excel floating-point equivalence or exact irrational standard deviations.

Default formula policy is reject for the whole selected rectangle. `--formulas cached`
explicitly uses and counts unverified stored caches; missing caches and Excel errors fail.
No recalculation occurs. Existing workbook/region limits apply, with at most 20,000
requested metric outputs (columns times quantile count plus six summary slots). Input
query reads are bounded to 64 KiB. JSON stdout is the deliverable: exit 0/status described
means completed computation; per-column sample insufficiency still requires inspection.
Exit 1/status failed indicates rejected input, policy violation or arithmetic limits.
