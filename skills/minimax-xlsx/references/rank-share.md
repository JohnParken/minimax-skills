# Controlled ranking and within-group shares

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检), require exit 0 and
`status: ready`, and confirm task directory/input readability. Resolve `<skill_dir>` from
the actual SKILL.md path before using task-relative inputs. Python 3.9+, standard library
only. The command reads inputs and returns JSON without altering or sorting the source.

Task-local `ranking.json`:

```json
{"group_by":["A"],"measure":"C","rank":"competition","descending":true,
 "filters":[],"missing":"omit","non_numeric":"reject",
 "denominator":"group_numeric_sum","precision":34}
```

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_rank_share.py" source.xlsx --sheet Sales --query ranking.json > ranked.json
python3 -I -S -B "<skill_dir>/scripts/xlsx_result_export.py" ranked.json ranked.xlsx
```

group_by and measure accept letters, col: IDs or exact multi-level paths. Pass explicit
--header-rows for path selection. The usual --range/--no-header/header-exclusion options
apply. Filters use typed AND predicates from [data-analysis.md](data-analysis.md) and are
applied before grouping, ranks and denominators. Omitted group_by means one global group.
Blank/empty group keys are distinct typed keys, following existing grouping semantics.

Ranking is per observed row within its group, not automatically per institution/entity
or aggregated metric. Aggregate deliberately first if business rules require that grain.
Rows are ranked by the numeric measure, descending by default; ascending is optional.
Exact Decimal equality defines ties; no rounding-to-display or tolerance grouping occurs.

- competition: ties share rank and leave gaps: 1,1,3.
- dense: ties share rank without gaps: 1,1,2.
- ordinal: each row has a distinct position: 1,2,3; ties use original source row order.

Result rows retain original order and source row IDs. missing is omit (default) or reject;
non_numeric is reject (default) or omit. Missing cells/empty strings are missing, while
text numbers and booleans are nonnumeric. Omitted observations remain visible in output
with null rank/share and an explicit missing_omitted/non_numeric_omitted status; they do
not enter ranks or denominators. Group-level counts report both types of omission.

The only denominator is the signed sum of all valid numeric observations in the same
group after filters. Share=value/group_numeric_sum, expressed as a fraction. Negative
values and negative totals are allowed; shares can be negative or exceed one (10 and -2
produce 1.25 and -0.25). Negative counts and totals are reported. No absolute-value
normalization, percentage-of-change contribution or additive time-series assumption is
introduced. With a zero total, ranks remain valid but shares are null with zero_total.
Groups without any numeric observation have null totals and insufficient_numeric_samples.

Group sums use exact Decimal arithmetic within 1000 digits; exceeding it fails. Shares
use 16..100 significant digits (default 34), 16 guard digits and ROUND_HALF_EVEN. A derived
rounded flag marks inexact division/final rounding. Numeric values serialize as JSON
strings with type tags. ISO-date text is not numeric; stored date serials remain numeric.

Output reuses selected schema with analysis=rank_share. Four explicit derived IDs are
calc:rank, calc:share, calc:share_status and calc:share_rounded. generated_columns records
label paths and derived provenance. requested_query, ranking_definition, share_definition,
arithmetic and group_statistics retain calculation scope for export/audit. Exported exact
shares with >15 digits are text with warnings, as described in [result-export.md](result-export.md).

The whole selected rectangle is subject to the existing formula policy before filtering:
reject by default, or explicit --formulas cached with unverified caches; missing caches
and Excel errors fail. No recalculation occurs. Query reads are bounded to 64 KiB;
output is limited to 200,000 typed cells including four derived columns, 16,384 columns
and 10,000 groups. Exit 0/status selected indicates completed analysis, not absence of
row/group diagnostics. Exit 1/status failed reports invalid configuration/input/budget.
No machine learning, forecasting, regulated banking metric or personnel decision is implied.
