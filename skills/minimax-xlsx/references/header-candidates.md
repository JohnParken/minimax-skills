# Suggest header ranges for review

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检) first, require exit 0
and `status: ready`, and confirm task directory/input readability. Resolve `<skill_dir>`
from the actual SKILL.md path, replace it in commands and use task-relative input paths.
Python 3.9+, standard library only. Inputs are not changed; JSON is written to stdout.

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_header_candidates.py" report.xlsx --sheet Sales
python3 -I -S -B "<skill_dir>/scripts/xlsx_header_candidates.py" report.xlsx --sheet Sales --scan-rows 64 --scan-columns 32 --max-depth 4
```

The default scan is A1 through at most row 32/column 32. Allowed bounds are 2..64 rows,
1..64 columns, and 1..4 header rows per candidate. The existing workbook loader's package
and cell limits still apply; only candidate evaluation, not whole-file loading, is limited
to the scan rectangle. `scan_clipped` warns about stored cells outside it. At most 256
row-range combinations are examined and the top eight successful candidates are returned.

Evidence includes complete horizontal/vertical merge structures, text-rich rows and a
following increase in numeric-value share. Full-width single-row merged headings are
excluded from candidates to avoid using report titles as all-column parents. Unmerged
notes and titles cannot be reliably identified: they may appear among weak alternatives.
Blank rows split candidate blocks; a candidate never crosses a blank row. Multiple blocks
might be separate tables, notes or intentional gaps, so the result reports ambiguity.
Repeated paths at different row starts are also reported. Headerless numeric tables,
numeric labels and all-text data may be indistinguishable and remain ambiguous.

Scores are heuristic ordering rules, not calibrated probabilities or accuracy guarantees.
`strong_evidence` means one of the documented structural/type rules fired. It is not proof
that the suggested rows are headers. Formulas and Excel errors are not interpreted as
header evidence; reports list uncertain coordinates. No cache freshness is assumed.
Candidates whose merge geometry is malformed, clipped, overlapping or crosses the
candidate boundary are rejected using the explicit-header safety checks.

`status: suggested` means one range has a sufficiently separated rule score with no
additional ambiguity. `status: ambiguous` means review is needed or evidence is lacking.
Both always set `requires_confirmation: true` and `automatic_boundary_applied: false`.
The command never sends a suggestion into profile, grouping or row processing. After
review, explicitly pass `--header-rows first:last` to the normal workflow; manual ranges
remain available regardless of candidate ranking. See [multi-level-headers.md](multi-level-headers.md).

Exit 0 means diagnosis completed, including ambiguous/empty results. Exit 1 returns JSON
`status: failed` for input/runtime errors. No real Excel/WPS application testing or
reliable automatic interpretation of arbitrary layouts is implied by these suggestions.

Merge structure is parsed once and reused across candidate ranges; all per-candidate
geometry checks remain active. No suggested boundary is applied automatically.
