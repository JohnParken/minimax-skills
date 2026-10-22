# Daily snapshots and controlled flow accumulation

Run the [SKILL self-check](../SKILL.md#执行环境约束与首步自检), require exit 0/status ready,
and confirm the task directory, input readability and XLSX write probe. Resolve
<skill_dir> from the actual SKILL.md path. Commands use relative task data paths and
Python 3.9+ standard libraries; no application, installation or upload is involved.

A configuration declares entity_keys, source columns (metric/date/value/unit/currency/
version/basis), source metrics, definition_version, start/end and date_format=YYYY-MM-DD.
Averages require stock/daily_snapshot metrics and matching source basis. Accumulation
requires flow/daily_flow; YTD, month-end and endpoint-only snapshots cannot substitute.
Units must match the declared metric. Dates are exact ISO date text (including date_text),
not guessed numeric Excel serials, timestamps or local-language dates. Entity IDs are
nonblank literal text. Currency and definition version are independently retained/checked.

## Reproducible fictional examples

Copy [average config](../assets/fictional_daily_average_config.json) and
[average facts spec](../assets/fictional_daily_average_facts.json) to the task directory.
Create the source with stdlib_xlsx.py create, then run:

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_daily.py" average daily.xlsx --sheet DailyFacts --query daily.json --xlsx average_report.xlsx
```

[Daily-flow config](../assets/fictional_daily_flow_config.json) and
[flow facts spec](../assets/fictional_daily_flow_facts.json) illustrate accumulation:

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_daily.py" accumulate flow.xlsx --sheet DailyFacts --query flow.json --xlsx flow_report.xlsx
```

All examples are fictional and contain no real customer or bank records. The ordinary
range/header/path-selection/cache options apply. Without --xlsx, JSON daily_summary_v1
is emitted and may be exported separately with xlsx_result_export.py. Source files are
never changed. No native pivot cache or charts are written.

## Exact coverage contract

start/end are inclusive Gregorian calendar dates spanning 1..366 days. Weekends and
holidays are included; no business-day calendar is guessed. Every entity/metric/currency/
version/date key must be unique. Duplicate dates fail with both source row numbers, even
when values agree or one is blank. Dates outside the requested interval fail by default;
explicit outside=exclude reports their count. Filters are applied before this date scope.

Reported days count any dated record; observed days count numeric nonmissing values.
Blank values are missing days, while numeric zero is a valid observation. coverage_details
lists expected/reported/observed counts, every missing/blank date, and coverage rounding.
coverage defaults complete: incomplete intervals have null primary results, never zero.
Observed mean/sum is separately available. coverage=observed explicitly uses the observed
value as primary and marks it observed-only; it must not be described as a full-period
mean/total. Groups/entities entirely absent from the source are not invented. An empty
result is flagged; no complete coverage of an unknown entity universe is claimed.

Daily balance mean=sum of numeric daily snapshots / numeric observed days. Under complete
coverage, observed days equal all calendar days in the interval. There is no zero fill,
forward fill, interpolation, endpoint averaging, day-weight estimation or month-end
balance averaging. Source units remain per row; coverage itself is a fraction.

Daily flow accumulation emits explicitly requested periods=['month','ytd'] by default.
Month windows intersect the requested interval and retain actual start/end; partial
month windows are labelled, not claimed as full calendar months. YTD requires the overall
interval to begin January 1; each year uses only that year's daily flows through the
requested end. Missing dates make the complete sum null; observed subtotal is explicit.
Monthly and YTD rows are alternative reporting windows and must not be added together.
Source YTD snapshots are always rejected rather than repeatedly accumulated or differenced.

Exact sums have a 1000-digit bound; exceeding it fails. Means and coverage ratios use
16..100 significant digits (default 34), 16 guard digits, ROUND_HALF_EVEN and explicit
inexact flags. JSON Decimal values are strings; XLSX preserves >15-digit results as text
with warnings. Formula caches require explicit --formulas cached and remain unverified.
No recalculation occurs; missing caches/errors fail before filtering.

The Daily XLSX sheet is accompanied by Columns/Metadata/Warnings/Definitions/Validation/
Lineage. Every window retains original file/sheet/row/metric references. A source day
may legitimately appear in both a month and YTD audit window. Bounds: 20 source metrics,
1000 groups, 20000 lineage records, 200000 output cells plus the full audit/export budget.
Exit 0 means computation completed, not that coverage is complete. Exit 1 gives structured
failure. Actual Excel/WPS rendering and regulator-specific averaging definitions are not
validated. Weighted averages, customer distinct counts and incomplete-calendar repairs
remain unsupported.
