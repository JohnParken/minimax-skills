# Strict BIFF8 read-only record subset

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检), requiring exit 0 and
status ready. Confirm the authorized task input is readable. This reader writes no files;
no write probe is required for inspection. Resolve `<skill_dir>` from the actual SKILL.md,
replace it before commands, and use task-relative inputs. Python 3.9+, standard library
only; no macro/formula execution, network requests, Office tools or dependency installs.

## Contents

- [Evidence and input](#evidence-and-input)
- [Supported records and retained values](#supported-records-and-retained-values)
- [Dates and formula caches](#dates-and-formula-caches)
- [Headers and output](#headers-and-output)
- [Rejection rules and validation limits](#rejection-rules-and-validation-limits)

## Evidence and input

Implementation follows the Microsoft MS-XLS online record/structure sections checked
2026-10-22: Record, BOF, BoundSheet8, SST/Unicode strings, Number/RK/MulRk, Blank,
LabelSst, BoolErr, MergeCells, Date1904, XF/Format, Formula/FormulaValue and String.
Exact official source URLs are bibliographic entries in
[et-format-evidence.json](et-format-evidence.json), not runtime dependencies. No third-party
parser code or binary samples were incorporated. All tests are synthetic MS-XLS record
fixtures, with manual typed expectations and container fixtures; they are not WPS files.

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_biff8.py" input.xls
```

The file must pass the [CFB container reader](et-ett-readonly.md), contain exactly one
root Workbook/Book stream, and start with the 16-byte BOF payload identifying 0x0600
BIFF8 globals. A stream name or .et/.ett suffix does not establish its encoding. Worksheet
BOFs, BoundSheet8 offsets, EOF boundaries and the non-dialog WsBool must agree. Chart,
macro/VBA and dialog sheets, recognized macro storages and encryption are refused.
ET/ETT names may be inspected by content but native_et_ett_parsed and
wps_compatibility_verified always remain false, even after successful BIFF extraction.

## Supported records and retained values

Read Number, RK, MulRk, BoolErr, Blank, MulBlank and LabelSst cells. Shared strings support
unsegmented compressed-Unicode or UTF-16 text only; rich/phonetic and Continue encodings
are refused. Text such as long leading-zero identifiers stays text. Number/RK values use
an exact Decimal representation of the stored IEEE754 value (which may differ from the
intended decimal); raw numeric bytes are retained. No numeric text coercion occurs.

Preserve sheet order/names/hidden state, source record offsets, raw XF indices/bytes,
custom format strings and merged-range coordinates. Merges are validated for overlap and
bounds but never expanded; non-anchor data is retained, not silently erased. Built-in
format IDs remain IDs; fonts/styles/layout are not rendered or reconstructed.

Known metadata records are explicitly retained as offset/type/raw payload hex: global
CodePage (Unicode 1200 only), Window1, Font, Style; worksheet Dimensions, Row, Index,
DBCell and Window2. Their internal semantics are not validated or applied. Other records
fail with ID and offset. This is a strict data-record subset, not full MS-XLS grammar or
workbook validity certification; many genuine XLS/ET files will be outside this subset.

## Dates and formula caches

Require explicit Date1904, preserving the 1900/1904 system. A numeric date serial stays
numeric with raw XF/number-format ID and any explicit custom format text. No date-format
heuristic, locale resolution, effective-style inheritance or ISO conversion is performed.
Serial 60 in the 1900 system is not invented as a real calendar date; fractional serials
and the original bytes remain available for later, separately verified interpretation.

Formula records retain opaque token bytes, ancillary bytes, raw flags and cache bytes.
Tokens are not decompiled, semantically validated or executed; unknown token semantics do
not trigger any formula evaluation. Shared-formula flags and array/table/shared records
are refused. Supported caches are numeric, boolean, text (immediate String record), error
and explicit empty string. Every present cache is unverified and may be stale, regardless
of the source recalculation flag. Error caches remain errors, never zero. Missing expected
String cache produces null, cache_status missing, an issue and a diagnostic result status;
it is distinct from a legitimate empty-string cache. Other truncated records fail.

## Headers and output

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_biff8.py" input.xls --sheet Data --range A1:F20 --header-rows 1:3
```

Explicit header reconstruction reuses [multi-level header rules](multi-level-headers.md),
including merge-only propagation, stable column IDs, clipping/cross-boundary rejection and
non-anchor-data checks. Formula headers default to rejection; --header-formulas cached
is an explicit opt-in with a warning, and missing/error caches still fail. Dates remain
numeric serials in paths. No implicit spreadsheet-to-OOXML reconstruction is performed.

JSON schema biff8_readonly_v1 contains typed sheet cells, merges, date/format metadata,
source file/stream digest, unparsed metadata and diagnostics. biff_parsed true means this
limited record subset passed; it is separate from the container-only report and native
ET/ETT acceptance. Unsupported inputs exit 1 with no partial cell report; subset extraction
exits 0, but read_biff8_subset_with_diagnostics requires review. Missing formula caches are
explicit diagnostics, not evidence of complete calculated data. No file is written.
Python APIs: parse_stream(bytes), read_file(path), reconstruct_headers(result,...).
Generic OOXML region/profile and editing APIs still do not accept this separate schema.

## Rejection rules and validation limits

Reject unknown/private records, FilePass, unsupported BOF versions, offset aliases/gaps,
truncated/oversized payloads, nonfinite numbers, duplicate cells, bad SST/XF references,
malformed strings/merges, orphan String records and trailing unbound data. No data-bearing
record is silently skipped. Present but opaque formulas remain explicitly unverified.

Limits: stream 8 MiB, records 100000, cells 200000, sheets 128, XFs 4096, SST strings
10000, merges 256, uninterpreted metadata 1 MiB/4096 records, diagnostics 1000, JSON output
16 MiB. Payloads over 8224 bytes and all Continue records are outside the subset. Header
regions are capped at 200000 cells and BIFF8's 65536 rows/256 columns. Inherited CFB chain,
allocation and memory budgets apply before record decoding. Trailing producer padding is
not guessed or discarded, so some noncanonical streams are deliberately refused.

No real ET/ETT samples or corresponding XLSX were available. Application interoperability,
WPS private extensions and real Excel acceptance are unverified. Next samples must include
permitted provenance/license or user authorization, WPS version/platform, untouched ET/ETT
and corresponding XLSX, and expectations for typed cells/merges/date system/formula caches.
Comparisons must distinguish stored IEEE754 from rounded decimal display and unverified
cache values. Do not rename XLS/XLSX fixtures and present them as genuine WPS validation.
