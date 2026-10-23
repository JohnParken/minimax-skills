# Explicit deposit maturity distribution

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检), requiring exit 0 and
status ready. Confirm task inputs and output write/read/cleanup probe. Resolve
`<skill_dir>` from the actual SKILL.md; replace it before commands and use task-relative
data files. Python 3.9+, standard library only; no private data scan or installation.

## Contents

- [Fictional CLI](#fictional-cli)
- [Input and status rules](#input-and-status-rules)
- [Date partitions and cumulative windows](#date-partitions-and-cumulative-windows)
- [Output and tested integration](#output-and-tested-integration)
- [Limits and remaining requirements](#limits-and-remaining-requirements)

## Fictional CLI

Copy [invented contracts](../assets/fictional_maturity_facts.json) to contracts.json and
[explicit config](../assets/fictional_maturity_config.json) to maturity.json in the task
directory. These contain no real customer records. Example horizons 7/30/90 days are
user-configurable management examples, not bank policy or regulatory defaults.

```console
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" create contracts.json contracts.xlsx
python3 -I -S -B "<skill_dir>/scripts/xlsx_maturity.py" contracts.xlsx --query maturity.json --xlsx maturity.xlsx > maturity-result.json
```

--plain disables appearance. Outputs must not exist; original source is unchanged.

## Input and status rules

Require unique explicit contract/certificate keys, numeric nonnegative principal, maturity
date, observation date, institution, currency, unit, version, level and status columns.
Selectors accept letters/IDs or unique structured header paths; range/header_rows can be
explicit. ISO YYYY-MM-DD dates are required, not guessed serial dates. Blank maturity
alone means unknown date. Invalid calendar dates, missing principal, negative values,
numeric-text amounts and duplicate contract observations fail. Zero principal is an
observed contract with zero amount, not an inferred closed contract.

Scope fixes one currency/unit/version/definition version/institution level and an explicit
institution universe. Mixed rows fail, rather than silently converting or dropping data.
Source observation dates must equal the configured date. status_mapping must explicitly
map every observed status to include or exclude; the example excludes closed records.
Unknown status fails. Duplicate IDs, invalid dates and negative amounts still fail for
excluded records. Excluded status/source-row diagnostics are retained without contract IDs.
completeness is declared_complete or unknown: input extraction completeness cannot be
proved. Unknown completeness remains a quality flag.

## Date partitions and cumulative windows

Day distance = maturity date - observation date in Gregorian calendar days, including
weekends/holidays. Negative days are past_date, zero is today, blank is unknown_date.
Past-date presence in a snapshot does not establish nonredemption, overdue status or
payment failure. Future maturity does not establish outflow or customer loss.

Future bucket bounds are inclusive and must form contiguous nonoverlapping intervals
from day 1 through an unbounded final bucket. The sample is [1,7], [8,30], [31,90],
[91,infinity]. Past/today/unknown are always separate, reserved categories. Gaps, overlaps,
reversed bounds, duplicate IDs and bounded final tails fail. Bucket principal totals must
exactly reconcile to all included principal, including unknown-date records.

Cumulative windows are separately labelled [1,N], such as 1–7, 1–30 and 1–90. They
exclude day 0, past and unknown dates and overlap each other; never add their totals.
Each partition/window outputs contract observation count, principal and principal share.
Share denominator is explicitly eligible_total_principal, including past/today/unknown,
within that institution scope. Zero principal denominator gives null shares and a quality
flag; no 0% is invented. Rounded fractions retain diagnostics in details.

## Output and tested integration

Output is native management_wide_v1 JSON with reversible metric IDs and a new readable
XLSX through [safe export](result-export.md). Sheets are Wide, Columns, Metadata, Warnings,
Definitions, Validation and Lineage. Metadata preserves status exclusions, bucket/window
definitions, exact observation date, coverage quality and source rows. No contract ID list
is exported. Shared pale-blue frozen headers/banding are used; all cells are literal.

Each scope emits a bank_scope slice plus separate same-level institution slices. Do not
add the bank slice to institutions. Automatic hierarchy rollup and parent/child source
mixtures are unsupported. Exact observation_date is a retained dimension even though
period uses YYYY-MM; a monthly period label does not change the snapshot date.

Actual integration regression verifies maturity wide -> dashboard -> XLSX and maturity
wide -> management policy -> XLSX, preserving source/quality; policy context must explicitly
fix observation_date. Unknown maturity quality prevents a clean ranking/comparison by
default. Complete natural-month daily results -> policy -> XLSX is also verified.
Customer wide -> dashboard is verified and now retains peer_group, original institution
level and observation scope so unlike peers cannot pool. Raw deposit-analysis and policy
result schemas remain explicitly refused by dashboard; their own XLSX exports work.
They require an explicit future native adapter, not an assumed direct connection.

## Limits and remaining requirements

At most 10000 contract observations, 100 same-level institutions, 20 future buckets and
20 cumulative windows (each 1..36500 days). Query 64 KiB; inherited source ZIP/XML limits,
export snapshot 16 MiB and 200000 cells including audits apply. Exact sums up to 1000
digits; ratios use 16–100 significant digits (default 34), ROUND_HALF_EVEN. Excel-unsafe
numbers become exact text with warnings. No real Excel/WPS visual acceptance was done.

Actual renewal rate remains unsupported. It needs linked maturing and actual renewal
contracts/transactions, amounts/dates, eligibility denominator, renewal window, treatment
of early/partial renewal, currency/unit policy and corrections/cancellations. Maturity
snapshots alone cannot provide those relationships. No intent or causal inference is made.
