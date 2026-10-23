# Deposit observations, daily means and concentration

## Contents

- [Fictional example and inputs](#fictional-example-and-inputs)
- [Change decomposition and targets](#change-decomposition-and-targets)
- [Structure and concentration](#structure-and-concentration)
- [Output and limits](#output-and-limits)

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检), requiring exit 0 and
status ready. Confirm task input readability and output write/read/cleanup probe. Resolve
`<skill_dir>` from the actual SKILL.md and replace it before executing commands from the
task directory. Python 3.9+, standard library only. No private files are discovered.

## Fictional example and inputs

Copy these invented assets to the task directory: [left](../assets/fictional_deposit_left.json)
as left.json, [right](../assets/fictional_deposit_right.json) as right.json,
[change config](../assets/fictional_deposit_config.json) as deposits.json, and
[structure config](../assets/fictional_deposit_structure_config.json) as structure.json.

```console
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" create left.json left.xlsx
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" create right.json right.xlsx
python3 -I -S -B "<skill_dir>/scripts/xlsx_deposits.py" left.xlsx right.xlsx --query deposits.json --xlsx deposits.xlsx > deposits-result.json
python3 -I -S -B "<skill_dir>/scripts/xlsx_deposit_structure.py" right.xlsx --query structure.json --xlsx structure.xlsx > structure-result.json
```

Snapshot columns are entity, institution, product, balance, currency, unit, version,
date and level. Selectors accept column IDs or unique structured multi-level paths.
Require explicit literal ISO dates, one currency/unit/version/institution level, nonnegative
numeric balances and one unique stable anonymized entity key per snapshot. Declare account
or customer grain; preaggregate explicitly to one row per key. Customer rows spanning
multiple products/institutions cannot be represented by a guessed single category; prepare
an appropriate account-grain source instead. Blank/numeric-text balances and duplicates
are refused. Zero is a valid observed balance. Formula inputs are refused, not cached.

## Change decomposition and targets

For matched keys, delta = current balance - previous balance. Separate positive continuing
changes, negative continuing changes and unchanged records. Keys observed only currently
contribute their current balance; keys observed only previously contribute minus their
previous balance. Sum of these four components equals total current - total previous;
residual must be exactly zero. Absent side values stay null: absence is not evidence of
new customers, account opening, closure or intentional withdrawal. No lifecycle inference
is implemented. Inputs must describe comparable complete observation universes; the tool
cannot prove extraction completeness or detect changed anonymization keys.

Stable keys align across institution/product changes. Both old/new classifications and
change flags are retained; these changes are not whole-bank entry/exit. No institution
or product transfer-flow allocation is inferred from two snapshots. Totals cover the
provided observation universe and must not be treated as audited bank totals without
source completeness checks.

Optional `daily` config requires sheet, query, metric_id and matching grain; pass --daily
daily.xlsx. Its query is the existing [daily mean query](daily-analysis.md), with one
stable entity selector, matching version, selected metric/unit/currency, coverage complete
and end equal to right_date. It computes sum(daily balances)/inclusive calendar days only
with full numeric coverage, preserving missing dates and observed-only mean separately.
Negative selected daily balances are refused; explicitly filtered-out records are not
rechecked as part of the selected deposit universe. No zero/forward fill or endpoint-average substitute.
`end_minus_daily_average` = current balance - complete daily mean; missing inputs stay
null. Entity means are not institution/product-attributed because intra-period movements
are unknown. This difference does not establish month-end manipulation or customer intent.

Targets are entity-specific at_least only, with entity/metric/period/unit/currency/version/
level/grain/value. Metrics are balance (period = right_date), net_change (left..right),
or daily_average (daily start..end). They cannot share a target merely because amounts
have the same unit. Achievement = actual/positive target; gap = max(target-actual,0).
Missing actual gives null achievement/gap; zero/negative targets have an explicit diagnostic
and no achievement classification. No bank/institution target rollup is inferred.

## Structure and concentration

Structure input is account grain, unique account keys. Each classification has a unique
id, field (product/entity/customer) and exhaustive scalar mapping to one category per
source value. For example explicitly map products to 活期/定期 and customers to 对公/零售.
Customer field requires a separate stable customer_column and customer concentration.
Axes are independent: each axis reconciles to the same balance total, but axes must never
be summed together. Unmapped or list-valued overlapping classifications are refused.

Category share = category balance / total balance. For concentration_grain account, Top N
means accounts; for customer, require separate text customer IDs and aggregate all account
balances by customer before sorting. Repeated customer IDs across distinct accounts are
expected; repeated account rows are refused to prevent double counting. No inferred
customer ID is allowed. Top N balance / total uses exact N, stable first-observed order
for ties; effective N is min(N,population). Missing IDs fail. Zero total gives null shares
and zero_total_balance, never 0% or normal. Negative balances are outside supported scope.

## Output and limits

Both commands produce typed JSON and optional readable XLSX; --plain disables appearance.
Deposit output is an entity wide table plus Checks, Quality, Lineage, Metadata and Warnings.
Structure output includes Structure, Concentration, Entities, Lineage, Metadata and Warnings.
Metadata retains query, periods, units, scope, formulas/policies and source coordinates;
daily row IDs resolve through the retained complete daily report. Text IDs remain exact;
all output is literal, with over-15-digit numeric values exported as text and warned.
Styles are shared, pale-blue frozen headers and banding; change category text remains.

At most 5000 entities per snapshot, 10000 union keys, 10000 targets, 10 classification
axes and 10 requested N values (each 1..1000). Query 64 KiB, export snapshot 16 MiB and
200000 total cells including audit sheets. Exact sums up to 1000 digits; divided outputs
use 16–100 significant digits (default 34), ROUND_HALF_EVEN, with rounding diagnostics.
No real customer samples or Excel/WPS visual acceptance were used. Generated XML/readback
is not rendering acceptance. No regulatory thresholds or personal decisions are produced.

Actual funding cost requires interest expense, consistent average interest-bearing balance,
period/day-count and annualization convention; averaging nominal rates is not actual cost.
Maturity/renewal needs maturity date, principal, actual renewal linkage/window and treatment
of partial renewal. These computations are not implemented without those fields/rules.
