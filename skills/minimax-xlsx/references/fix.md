# Repair a known formula defect

Inspect the original formula and its intended business behavior. Use the standard-library
`read` command; if calculation is needed, use `calculate` to a separate output and inspect
its structured issues. A failed calculation produces no new output.

Correct only formulas whose intended meaning is known, using a JSON patch and `edit`.
Do not replace unsupported functions with guessed results. Do not hide division by zero
with an arbitrary default or use IFERROR merely to obtain a clean report. If a default is
part of the agreed business rule, express it explicitly using supported IF conditions.

Recheck inputs, changed formulas, dependent results and unaffected content. See
[python-calculation.md](python-calculation.md) and [validate.md](validate.md). Structural
edits and shared/array formulas require a separate implementation; do not use legacy
regex row-shifting as a repair shortcut.
