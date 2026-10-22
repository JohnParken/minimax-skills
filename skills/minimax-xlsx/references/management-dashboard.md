# Static management dashboard

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检), requiring exit 0 and
`status: ready`; confirm task input readability and output-directory write/read/cleanup
probe. Resolve `<skill_dir>` from the actual SKILL.md before executing these commands
from the task directory. Python 3.9+, standard library only; no installation or network.

## Fictional example

Copy the bundled [fictional facts](../assets/fictional_hierarchy_facts.json),
[hierarchy configuration](../assets/fictional_hierarchy_config.json) and
[dashboard configuration](../assets/fictional_dashboard_config.json) into the confirmed
task directory as facts.json, hierarchy.json and dashboard.json. These are invented bank
figures, not customer data, regulatory definitions or recommended thresholds.

```console
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" create facts.json facts.xlsx
python3 -I -S -B "<skill_dir>/scripts/xlsx_management.py" facts.xlsx --sheet Facts --query hierarchy.json > wide.json
python3 -I -S -B "<skill_dir>/scripts/xlsx_dashboard.py" wide.json --query dashboard.json --xlsx dashboard.xlsx > dashboard-result.json
```

Input is native management_wide_v1 JSON, not arbitrary spreadsheets or daily-summary
JSON. Select 1–20 metric IDs and an explicit overview_period (YYYY-MM). trend_metrics
is an optional selected subset, default all selected metrics. The output JSON retains
the complete source report, metric dictionary, original source rows, filters, definition
version, target policy, arithmetic metadata and reversible column IDs. Dashboard records
reference wide_row_id; source_report.lineage resolves it to file/sheet/source rows.

## Comparison and warnings

Overview contains actual, target, achievement (a multiple), gap and quality flags without
inventing missing targets. Target meaning is inherited from the wide report. There are
no dashboard totals: ancestor and descendant observations remain separate. Ranks require
explicit method (competition, dense or ordinal) and boolean descending for each selected
metric. Equal values use standard competition gaps, dense consecutive ranks, or stable
source order for ordinal ties. Peers retain every dimension except institution and its
name, including period, report level, currency, version and business line. Missing actuals
and rows with any quality diagnostic are excluded with an explicit status; direction is
not interpreted as business performance. peer_count counts eligible observations only.

Thresholds require metric_id, field (actual/achievement/gap), op (gt/ge/lt/le/eq/ne), a
finite JSON numeric value, unit, severity text, and scope. Scope must explicitly retain
currency, version and report_level; other retained dimensions can narrow it. Unit must
match the metric, or `multiple` for achievement. Rules evaluate only the overview period.
Missing values are not_evaluated_missing; quality-flagged observations are
not_evaluated_quality. No matching rule means no evaluation, never an implicit normal
classification. within_configured_threshold means only that this configured predicate did
not trigger; it is not a regulatory or overall health judgement. Unmatched rules warn.

Trends is an observed-period data table, not a chart. It does not fill missing periods,
interpolate, sum balances/YTD, compare currencies or create a native chart/web dashboard.
Maximums: input/snapshot 16 MiB, query 64 KiB, 5000 source rows, 20000 metric records and
20000 rule evaluations; safe export additionally enforces 200000 cells including audits.

## XLSX and acceptance

Sheets: Overview, Alerts, Trends, Definitions, Metadata and Warnings. Metadata stores the
complete structured dashboard/source JSON in ordered chunks. Decimal outputs follow the
existing [safe result export](result-export.md) policy: over 15 significant digits or Excel
range become exact text with warnings; IDs remain text, including long leading-zero IDs.
All cells are literal values, never executable input formulas. Existing outputs are refused.
The readable preset provides pale-blue headers, frozen header rows and alternating rows;
Alerts additionally colors status categories while retaining status text. No automatic
filter controls are added. Shared styles are reused. No actual Excel/WPS visual acceptance
or rendering has been performed; generated OOXML and readback tests do not establish it.
