# Configured bank management long-to-wide reports (first release)

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检); require exit 0 and
`status: ready`. Confirm task directory/input readability and a write probe for XLSX output.
Resolve `<skill_dir>` from the actual SKILL.md path before executing task-relative inputs.
Python 3.9+, standard library only; no installation, upload or spreadsheet application.

This module targets institution/product management aggregates, not individual credit,
eligibility or employee decisions. It does not implement regulatory reporting standards,
regulatory thresholds or official bank-specific metric definitions. The supplied example
is entirely fictional and its ratio/targets illustrate arithmetic, not bank policy.

## Reproducible fictional example

Copy the reusable [fact specification](../assets/fictional_management_facts.json) and
[configuration](../assets/fictional_management_config.json) into the confirmed task directory:

```console
cp "<skill_dir>/assets/fictional_management_facts.json" ./facts.json
cp "<skill_dir>/assets/fictional_management_config.json" ./management.json
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" create facts.json fictional_facts.xlsx
python3 -I -S -B "<skill_dir>/scripts/xlsx_management.py" fictional_facts.xlsx --sheet Facts --query management.json --xlsx fictional_management.xlsx
```

Without --xlsx the command emits native management_wide_v1 JSON only. That saved JSON
can also be passed to xlsx_result_export.py. --range, --header-rows, header exclusions,
--no-header and explicit --formulas cached reuse existing semantics. Formula caches
remain unverified; defaults reject formulas/errors before filtering. Existing sources
and output files are never overwritten. No native pivot tables/caches are generated.

## Fact grain and safe aggregation

Input is long-form: one metric/value per row, with explicit dimension columns, metric ID,
unit, record-scope marker and time-basis columns. Every dimension must be nonblank literal
text; long institution IDs remain text. Numeric identifiers are rejected rather than
recovered by guesswork. Dimension selectors accept letters, col: IDs or unique exact
multi-level paths. Supported dimensions are period, institution, line, currency, version,
product, region and channel. Period/institution/currency/version are mandatory.

fact_grain must list every configured input dimension plus metric_id exactly once.
Duplicate facts at that grain always fail, even when their values agree. row_dimensions
specifies output grain and must retain period, currency and version. period_format is
explicit YYYY-MM. Different months/currencies are separate rows; no FX conversion,
YTD accumulation, cross-period stock addition or automatic missing-month creation occurs.
Only the configured definition_version is accepted, and it remains a dimension.

A mandatory source record_scope column must equal one explicitly configured marker
(e.g. leaf) on every selected row. Mixed summary/detail markers fail. Apply an explicit
filter if the input deliberately contains multiple scopes. Marker truth is an upstream
data contract. Optional [validated institution rollup](institution-hierarchy.md) adds an
explicit static tree, deepest-leaf checks and ancestor aggregation; no ancestry is inferred.
Only flat, declared leaf-management data may be grouped. A report_level label is required
and retained; it does not manufacture a hierarchy or prove leaf exclusivity.

Filters use existing typed AND semantics before grain checks/aggregation. Selected and
excluded row counts plus the exact requested query are retained. Fact values must be
stored numbers; no text-number conversion occurs. missing=propagate (default) returns
null for an incomplete or absent metric with quality flags; reject fails instead. No
zero fill or silent omission is provided. SUM across dropped dimensions requires explicit
additive permissions and upstream non-overlapping partitions. See [dictionary rules](management-dictionary.md).

## Wide output and audit

Wide has configured row dimensions, report_level, metric actuals, requested target /
achievement / gap fields, and a quality-flags column. Column IDs are JSON-encoded arrays
such as ["metric","balance","actual"], not ambiguous joined strings. Parse the ID as JSON
to recover its parts. Human labels and structured paths are retained separately.

Targets must match full output grain, metric ID, unit, kind, time basis, definition version
and report level. Duplicate targets fail. No target is guessed or implicitly aggregated.
At-least completion is actual/positive target; gap=max(target-actual,0). At-most completion
is positive target/positive actual; gap=max(actual-target,0). Completion is a fraction
labelled in multiples, not a pass/fail decision. Zero/negative targets yield null completion
and diagnostics; at-most with nonpositive actual also yields null completion. Gaps remain
numeric when actual/target exist. Missing target/actual is flagged. Targets with no observed
actual group are listed in validation; no synthetic zero-valued groups are created.

XLSX includes Wide, Columns, Metadata, Warnings, Definitions, Validation and Lineage.
Lineage maps report row IDs and Wide worksheet rows to source filename/sheet/row/metric.
Dictionary, filters, target definitions, units, version and arithmetic policy are retained.
Validation pass means configured input checks passed, not regulatory/business certification.
Readable headers, fixed widths, banding and visible textual quality flags are used.
Precision-limited numeric values become exact text with warnings; [export rules](result-export.md)
apply. Entire source JSON is retained, but source-file original formatting is not recreated.

Queries are bounded to 64 KiB; at most 50 metrics, 2,000 targets and 10,000 report groups.
Existing 200,000-cell region/output and 16 MiB export-snapshot limits apply. All audit
sheets count toward export's 200,000 total cells; narrow grain/range if a report is too large.
Exit 0 reports completed computation/export; inspect quality flags and validation warnings.
Exit 1 returns structured failure without a new output. Real Excel/WPS visual acceptance,
customer distinct-count analysis and daily/weighted-average measures remain unverified or
unsupported; no missing customer IDs/day coverage/weights are invented.
