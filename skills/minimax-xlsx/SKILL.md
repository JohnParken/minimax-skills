---
name: minimax-xlsx
version: "2.0"
description: "Create, inspect and edit XLSX cells, calculate supported formulas, and import/export CSV or TSV. Use for reports and spreadsheet data exchange."
license: MIT
---

# MiniMax XLSX — Python Standard Library

Category: productivity. Format reference: ECMA-376 Office Open XML File Formats.

Use Python 3.9+ without installing packages, launching Office/LibreOffice, or using plugins.
The maintained entrypoint is `scripts/stdlib_xlsx.py`; it imports the bundled
`formula_engine.py`. Keep both files and `assets/minimal_xlsx/` together.

## 执行环境约束与首步自检

Confirm the task working directory and input/output paths before operating. Resolve
`<skill_dir>` from the actual location of this SKILL.md; replace that placeholder in
commands with the real directory. Do not generate task files in the skill source tree.
Use relative task data paths from the confirmed task directory. No installation is needed.
Do not use parent paths, absolute data paths or mount changes to bypass task boundaries.
If the current directory is the skill source, have the harness confirm a separate task
directory first. Resolve Markdown resources relative to the containing file; never infer
the skill location from environment variables, another session, or masked host paths.

Before the first operation, run:

```console
python3 -I -S -B "<skill_dir>/scripts/stdlib_xlsx.py" check
```

Require exit 0 and JSON `status: ready`. This verifies imports, the template and a
limited calculation; it does not verify task data, directory permissions, or Excel
compatibility. Check input readability and, for writes, output-directory writability
with a uniquely named temporary probe that is read back and removed. Stop dependent
operations if a check fails. Re-read this entrypoint if its constraints are lost after
context compaction. Verify the actual output and available delivery mechanism separately.

## Route the task

| Task | Action | Reference |
|---|---|---|
| Compute median, quantiles and variance/stddev | `xlsx_advanced.py describe` | [advanced-statistics.md](references/advanced-statistics.md) |
| Compare explicit monthly periods | `xlsx_advanced.py compare` | [period-comparison.md](references/period-comparison.md) |
| Match tables by typed composite keys | `xlsx_match.py` | [table-matching.md](references/table-matching.md) |
| Deduplicate and stably sort typed rows | `xlsx_rows.py` | [row-processing.md](references/row-processing.md) |
| Suggest possible header ranges for review | `xlsx_header_candidates.py` | [header-candidates.md](references/header-candidates.md) |
| Reconstruct or select multi-level headers | `xlsx_headers.py`; `--header-rows` in data operations | [multi-level-headers.md](references/multi-level-headers.md) |
| Profile a region, filter or group data | `xlsx_analysis.py profile` / `summarize` | [data-analysis.md](references/data-analysis.md) |
| Diagnose file contents or ET/ETT candidates | `xlsx_identify.py` | [data-analysis.md](references/data-analysis.md) |
| Import CSV/TSV or export a worksheet/range | `xlsx_tabular.py import` / `export` | [tabular-interchange.md](references/tabular-interchange.md) |
| Inspect XLSX/XLSM or CSV/TSV | `read` | [read-analyze.md](references/read-analyze.md) |
| Create a report | JSON specification → `create` | [create.md](references/create.md) |
| Fill or correct existing cells | JSON patch → `edit` | [edit.md](references/edit.md) |
| Compute supported formulas in an existing file | `calculate` to a new file | [python-calculation.md](references/python-calculation.md) |
| Repair a known formula defect | Diagnose intent, then `edit` | [fix.md](references/fix.md) |
| Validate output | Inspect calculation report and preserved content | [validate.md](references/validate.md) |

Read [python-calculation.md](references/python-calculation.md) before any calculation
or writing task. It defines the JSON contract, commands, formula subset and exit codes.
Use [format.md](references/format.md) for financial formatting and XML style details;
its formatting conventions do not override static-result mode or calculation limits.

## Calculation policy

- For known business rules, calculate in Python using `Decimal` with an explicit
  rounding policy. Write those results as `value` when a static report is intended.
- For editable reports, use supported `formula` expressions. The engine evaluates
  them and writes both formula and cache. Never accept caller-supplied formula caches.
- `create --values-only` evaluates all formulas first, then writes static values. It
  refuses partial results; explain that these values will not update when inputs change.
- Existing formulas are recalculated only within the documented subset. Unsupported
  formulas and their dependents keep their formulas and lose stale caches. Exit 2
  means incomplete calculation, not a fully verified workbook.
- A known calculation error or circular dependency blocks output. Do not hide errors
  with default zeros or change business rules just to make calculation pass.
- This is Decimal subset calculation, not Excel-equivalent recalculation. Report the
  engine and limits. For unresolved formulas say `部分公式未计算` with affected cells.
- The opened workbook's old caches are unverified; missing caches are not zero.

## Editing and preservation

Never reconstruct an existing workbook from extracted values. `edit` patches the
original package and preserves unmodified ZIP member contents byte-for-byte. Modified
worksheet XML is serialized again; full-file byte identity or arbitrary feature
compatibility is not promised. Formula caches across sheets may change when an input
changes. Calculation-chain metadata may be removed and recalculation requested.

Only cell values, formulas and existing style indices are editable through the CLI.
Do not insert/delete/move rows, columns or sheets. Editing a sheet containing shared,
array or other specially attributed formulas is refused. Macro parts can be retained
in XLSM files, but macros are never run. Charts and pivots are retained, not refreshed.
Dates are exposed as stored serials or ISO strings, without display-format conversion.
Signed/encrypted files, legacy XLS/XLSB, chart sheets and strict OOXML are unsupported.

## Legacy utilities

The older scripts remain for compatibility and reference. They are not the standard
library workflow: `xlsx_reader.py` needs pandas, `libreoffice_recalc.py` needs external
software, and older helpers may require Python 3.10+. Do not invoke them for this mode.
In particular, do not use `xlsx_insert_row.py`/`xlsx_shift_rows.py`: their global row
shifting and regex reference rewriting are not a safe structural-edit implementation.
`formula_check.py` is heuristic static inspection only; exit 0 does not prove calculation.

## Development verification

From the repository root:

```console
python3 -I -S -B -m unittest discover -s skills/minimax-xlsx/scripts -p 'test_*.py' -v
```
