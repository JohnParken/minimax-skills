# Create a spreadsheet

Read [python-calculation.md](python-calculation.md) for the JSON schema and CLI.
Plan sheet names, input values, calculations and formatting first. Use `create` with
an explicit JSON cell specification. The tool builds from the bundled XML template.

Choose the output behavior:
- Editable model: use supported formulas; the engine writes formulas and calculated caches.
- Static report: calculate business results in Python and use `value`, or use
  `create --values-only` to evaluate supported formulas and replace them with values.

Do not invent data or formula caches. Label assumptions and units. Preserve required
column/sheet names. For financial styling consult [format.md](format.md); select template
style indices deliberately. The CLI does not set custom column widths, merges or charts.
Such XML extensions require separate implementation and validation, not a claim that
all Excel features are supported.

Check the calculation report, reopen the output with `read`, verify expected sheet names,
input values and independently known results. Explain static outputs and unresolved
formulas. See [validate.md](validate.md). No Office runtime or third-party packages are needed.
