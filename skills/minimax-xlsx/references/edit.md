# Edit existing cells

Use `read` to inspect the actual sheet names, addresses, formulas, raw values and style
indices. Follow [python-calculation.md](python-calculation.md) to build a JSON patch and
run `edit input.xlsx patch.json output.xlsx`. The output must have a new filename.

Only list intended cells. Existing styles remain when style is omitted. Writing null
clears a cell's value/formula but retains its style. Write a merged region's top-left
anchor only. A patch cannot add/rename sheets or insert/delete/move rows or columns.
Sheets containing shared/array/specially attributed formulas cannot be edited by this API.

Changing input values can invalidate formula caches on other sheets. The tool calculates
all formula cells using the supported subset, clears unresolved caches, and reports issues.
Known errors or cycles block output. Do not suppress a business error with arbitrary zeros.
The original file is never overwritten.

Untouched ZIP members retain their content bytes. Changed worksheet XML is reserialized;
this is not a promise of identical ZIP bytes or complete Excel feature compatibility.
Calculation-chain parts may be removed. Macro, chart, image and pivot parts are retained,
not executed or refreshed. Preserved chart/pivot caches can be stale after an input edit.
Signed/encrypted workbooks and non-worksheet sheet types are outside the workflow.

Do not use the legacy unpack/shift/insert scripts for this mode. In particular their
row-reference rewriting lacks the sheet-aware parsing needed for safe structural edits.
Verify output sheet names, unrelated data, changed cell values and recalculated results;
see [validate.md](validate.md).
