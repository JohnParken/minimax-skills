# Static institution hierarchy rollup

Use the self-check/task directory/input/output rules in [management-wide.md](management-wide.md).
Python 3.9+, standard library only. This is an explicit management mapping, not an inferred
organizational chart, regulatory consolidation scope or historical-restatement engine.

Add hierarchy to a management configuration:

```json
{"mapping_version":"fictional-tree-v1","valid_from":"2025-01","valid_to":"2025-12",
 "levels":["bank","region","branch"],"report_levels":["bank","region","branch"],
 "coverage":"complete",
 "nodes":[{"id":"ROOT","parent":null,"level":"bank","name":"虚构总行"},
          {"id":"R","parent":"ROOT","level":"region","name":"虚构区域"},
          {"id":"0001","parent":"R","level":"branch","name":"虚构甲支行"}]}
```

The object is the hierarchy value, not the complete management configuration. Input
institution IDs remain literal text. Mapping_version and inclusive validity months are
mandatory. Each node has one parent (or null), an explicit name and a level from the
ordered list. Roots use the first level, children the immediately following level, and
all leaves the deepest level. Multiple roots are permitted. Duplicate IDs, multi-parent
lists, cycles, unknown parents, invalid/skipped levels and unmapped facts fail.

Source facts must contain deepest leaf institutions only, including when source rows
claim a leaf scope marker. Any mapped parent institution fact is refused, so combining
existing parent totals with children cannot silently double-count. Each original leaf
fact contributes once to each selected ancestor. Ancestors are not re-fed as new facts.
Source fact-grain uniqueness remains checked before expansion. Institution must remain a
row dimension. Record_scope.report_level declares the deepest source level; output rows
carry their actual mapped report_level plus institution name and ID.

When any parent level is requested, each additive base metric must permit institution
in additive_dimensions (MIN/MAX use reduction_dimensions). Existing permissions for other
dropped dimensions still apply. Period/currency/definition version remain separate. At
all levels, ratios are recomputed from aggregated base components, not averaged from
branch ratios. Parent targets must explicitly name the ancestor ID and its correct level.
No automatic target allocation or summation occurs.

coverage defaults complete: for each emitted ancestor/context/metric, every mapped
leaf descendant is expected to have a fact. Missing leaf metrics propagate null with
incomplete_leaf_metric flags (or fail under missing=reject). A missing fact value also
remains incomplete. This expectation includes the selected line/product context; provide
a scope-specific map or deliberate observed coverage when some leaves are inapplicable.
coverage=observed explicitly returns sums of observed leaves, flagged as partial; it does
not prove complete coverage. No absent leaf/group is created as a zero-valued record.

Mapping validity is checked against facts and target periods. A single static tree is
used for the run; institution changes/splits/mergers do not automatically rewrite prior
periods. Split runs at known changes and supply the appropriate mapping version. The
mapping's truth and additive partitions remain upstream responsibilities.

Limits: 1..10 levels, 1..5000 nodes, source facts times requested report levels <=20000
lineage rows, and existing report/audit/export budgets. All original file/sheet/row/metric
references survive each emitted ancestor row. Same-level comparisons must include actual
report_level and comparable period/currency/version; do not rank parent and child totals
together. No individual customer/employee decision is made.

A reusable [fictional hierarchy config](../assets/fictional_hierarchy_config.json) and
[fictional source spec](../assets/fictional_hierarchy_facts.json) can be copied into the
confirmed task directory. Create the source with stdlib_xlsx.py create, then:

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_management.py" fictional_hierarchy.xlsx --sheet Facts --query hierarchy.json --xlsx hierarchy_report.xlsx
```

The generated XLSX reuses Wide/Definitions/Validation/Lineage and readable appearance.
Self-check/read-back tests are not real Excel/WPS visual or regulatory acceptance.
