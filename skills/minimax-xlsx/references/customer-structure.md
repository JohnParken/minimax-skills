# Observed customer aggregation and product structure

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检), requiring exit 0 and
status ready. Confirm readable task inputs and an output write/read/cleanup probe. Resolve
`<skill_dir>` from the actual SKILL.md before commands, using task-relative data files.
Python 3.9+, standard library only; no private file discovery, installation or upload.

## Contents

- [Fictional input and CLI](#fictional-input-and-cli)
- [Identity, institution ownership and scope](#identity-institution-ownership-and-scope)
- [Formulas and interpretation](#formulas-and-interpretation)
- [Outputs, privacy and limits](#outputs-privacy-and-limits)
- [Next-stage maturity data](#next-stage-maturity-data)

## Fictional input and CLI

Copy invented [left observations](../assets/fictional_customer_left.json) as left.json,
[right observations](../assets/fictional_customer_right.json) as right.json and
[configuration](../assets/fictional_customer_config.json) as customers.json into the task
directory. These examples contain fictional identifiers and are not bank customer data.

```console
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" create left.json left.xlsx
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" create right.json right.xlsx
python3 -I -S -B "<skill_dir>/scripts/xlsx_customers.py" left.xlsx right.xlsx --query customers.json --xlsx customers.xlsx > customer-result.json
```

Use --plain to disable the readable appearance preset. All outputs must be new files.
Each side declares sheet, distinct field selectors and optional range/header_rows. Fields
are account, customer, institution, product, holding, period, version, level and scope;
designated ownership additionally requires owner. Selectors accept letters, IDs or unique
multi-level paths. Metadata and identifiers must be nonblank literal text; holding must
be boolean. Formula inputs and numeric identifiers are refused, not coerced.

## Identity, institution ownership and scope

Require fact_grain account_product: one account/product observation per period. Exact
duplicates fail; one account with conflicting customer/institution/owner metadata fails.
Multiple accounts for a customer are expected and deduplicated. A shared stable account
changing customer between periods fails until identity history is explicitly resolved.
Declare two ordered YYYY-MM periods, one source version, definition version, institution
level and observation_scope label. Rows with mismatched periods/version/level/scope fail.
No filters are silently inferred. Completeness is declared_complete or unknown, as a
user assertion; extraction completeness and stable cross-period anonymization cannot be
proven by this module. Unknown completeness remains a quality flag.

Institutions must all have the configured level and an explicit peer_group. Unknown
institutions fail. ownership exclusive rejects customers observed in multiple institutions.
ownership designated requires all records for a customer to specify the same owner,
and that owner must have an observed account for the customer. All selected product
holdings are attributed to that owner. There is no guessed primary bank or proportional
allocation. Bank distinct counts are computed directly; institution and peer-group slices
are disjoint under this explicit ownership rule. Never sum across report levels.

## Formulas and interpretation

Product codes map explicitly to nonempty lists of unique category IDs. Multiple categories
may overlap; missing product mappings, duplicate category IDs or duplicate labels within
a product mapping fail. Category labels are display only; IDs preserve distinction.
A customer holds a category if any observed account/product record has holding=true and
maps to it. Holding is not inferred from balance, account presence or zero balance.
Observed customers with holding=false everywhere remain in the observed population.

Category coverage = distinct holding customers / configured observed_customers or
holding_customers denominator. A holding customer holds at least one mapped category.
Zero denominator produces null with undefined_zero_denominator, never an invented 0%.
Overlapping category coverage ratios must not be added. Product-count distribution counts
distinct mapped categories held, not accounts or raw product codes; it includes zero.
Its shares always use all observed customers as the explicitly named denominator.

Two-period changes report set intersection (continuing observed), current-only and
previous-only observations. Current count - previous count = current-only - previous-only;
residual is retained. Owner migrations among continuing bank customers are reported
separately and are not bank entry/exit. Institution membership changes can include such
migrations. No records are called newly opened/acquired, lost or closed; lifecycle evidence
is absent. Empty extracts describe zero observed records, not proof of no actual customers.

## Outputs, privacy and limits

JSON includes summary, coverage, product_count_distribution, changes, source coordinates,
quality and definitions. management_wide is a native management_wide_v1 subreport compatible
with [safe result export](result-export.md); its counts use customers and coverage uses
fraction with retained numerator/denominator definitions. Currency is not_applicable;
these are not monetary indicators. Peer group, original institution level and observation
scope remain explicit dimensions so downstream ranks cannot pool unlike groups. Scope IDs are reversible JSON [kind,id] arrays.
Category counts/ratios retain stable JSON-based metric IDs; no cross-period count sum or
ratio average is implied. Full-report sheets are Wide, Customers, Coverage, ProductCounts,
Changes, Lineage, Definitions, Quality, Metadata and Warnings. There are no customer lists.

Default JSON/XLSX omits input account/customer identifier values and all per-customer
records. Lineage retains file/sheet/row coordinates for scoped aggregates; source access
is still required for investigation and these reports are not certified anonymous.
Configuration contains only institution/product metadata, not a customer ownership list.
The tool does not perform small-cell suppression, access control or disclosure assessment.
All output is literal and source files are unchanged; numbers unsafe for Excel precision
become exact text with warnings. Shared pale-blue headers/frozen rows and banding are used.
No real customer samples or Excel/WPS visual acceptance were used.

Limits per period: 5000 customers, 20000 account-product observations; configuration has
100 institutions, 32 categories, 1000 product codes. Aggregate lineage is capped at 60000
rows. Query 64 KiB, output snapshot 16 MiB, XLSX 200000 cells including audits. Decimal
ratios use 16–100 significant digits (default 34), ROUND_HALF_EVEN with rounding flags.
No individual credit, eligibility, pricing or personnel decision is produced.

## Next-stage maturity data

Maturity distribution needs explicit account/contract key, customer key if customer-level,
remaining principal, currency/unit, observation date, maturity date, early-termination
status and user-defined nonoverlapping buckets. Renewal additionally needs actual linked
renewal contract/transactions, amount and date, renewal window and partial/early renewal
rules. Due-date balance alone cannot establish renewal, attrition or customer intention.
These methods are not implemented in this module without those definitions and fields.
