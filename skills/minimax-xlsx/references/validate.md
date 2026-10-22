# Validate a standard-library output

Keep these findings separate:

1. **Readable structure:** the ZIP and XML parse, expected sheets/cells exist, values,
   types and style references are appropriate. Parsing is not full OOXML XSD validation.
2. **Change integrity:** only intended cells and required calculation metadata changed.
   Compare untouched ZIP member contents; source file and unrelated data must remain.
3. **Subset calculation:** inspect `status`, `issues`, `calculated_count`, `formula_count`
   and `output_written` from `stdlib_xlsx.py`. See [python-calculation.md](python-calculation.md).
4. **Business accuracy:** independently check representative expected totals, ranges,
   rounding and units. A formula can evaluate successfully and still reference wrong data.
5. **Unverified behavior:** report unknown formulas, unrefreshed chart/pivot caches and
   lack of Excel rendering/recalculation verification. Never claim these checks passed.

Reopen the output using `read` to verify stored formulas and caches. Reading old caches
marks them unverified because an XLSX alone does not prove their freshness. Retain the
calculation report as the evidence of this run. No third-party application is launched.

Exit 0 / calculated_subset means this engine completed its supported calculations; it
is not full Excel compatibility. Exit 2 / incomplete produces an output with unsupported
formula caches removed and must be disclosed as `部分公式未计算`. Exit 1 / failed must
not be delivered as a successfully calculated workbook. Static values-only reports must
be identified as such; inputs will not drive automatic updates.

The legacy formula_check.py is heuristic static inspection. Its clean exit is not
calculation evidence, it does not prove business correctness, and it is not required by
the Python 3.9+ standard-library path. No LibreOffice installation or runtime is required.
