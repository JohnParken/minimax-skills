# Read and analyze with the standard library

Use `stdlib_xlsx.py read input.xlsx` as described in
[python-calculation.md](python-calculation.md). The JSON lists sheets and coordinate-keyed
cells containing type, raw value, style index, formula, formula attributes and cache status.
Decimal values serialize as strings in the report to avoid binary-float precision loss.
The read command does not calculate formulas or modify files. Existing formula caches
are unverified; a missing cache is not zero. Dates remain serial numbers or ISO strings,
with the workbook date system reported separately. Display formatting is not rendered.

CSV/TSV reads return text rows, including headers and empty fields. Explicitly choose an
encoding (`--encoding gbk` where appropriate); there is no silent encoding fallback.
Preserve leading zeros and convert only known numeric columns to Decimal. Strings that
begin with `=` are plain data unless deliberately supplied as formula in a write spec.

For custom analyses use `csv`, `collections`, `statistics`, `decimal` and Python loops.
Compute sums and means from the actual selected data, not reconstructed approximations.
Make missing-data and rounding policies explicit. Do not mix old formula caches with
freshly computed values without disclosing their status. If fresh results are needed,
run `calculate` into a separate workbook and inspect its report before using them.

The old `xlsx_reader.py`/pandas pipeline is legacy and is not used by this workflow.
