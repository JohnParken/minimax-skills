# ET/ETT investigation and bounded compound-container reading

Run the [entrypoint self-check](../SKILL.md#执行环境约束与首步自检) and require exit 0,
status ready. Confirm the input is authorized and readable in the task directory. This
reader writes no files; no write probe is needed for inspection. Resolve `<skill_dir>`
from the actual SKILL.md and replace it before running commands with task-relative input.
Python 3.9+, standard library only; no Office launch, network calls or installations at runtime.

## Contents

- [Evidence and provenance](#evidence-and-provenance)
- [CLI and support levels](#cli-and-support-levels)
- [Container subset and safety](#container-subset-and-safety)
- [Tests and remaining work](#tests-and-remaining-work)

## Evidence and provenance

Research checked 2026-10-22. Sources were read for specifications/support/licensing; no
third-party parser code or test files were copied, installed or bundled. Exact source URLs
are recorded in the [evidence manifest](et-format-evidence.json) as bibliographic provenance,
not runtime resources or downloadable dependencies.

| Source | Finding and implementation use |
|---|---|
| Microsoft MS-CFB, published revision 12.0, 2024-04-23 | Documented general-purpose compound container; basis for this independent implementation, not an ET record specification |
| CFB header and directory entry | Sector sizes, header fields, directory offsets and UTF-16 names |
| SheetJS format documentation | Lists ET read support and describes proprietary extensions over BIFF8; this is evidence for a future investigation, not proof that every ET/ETT variant is BIFF8 |
| SheetJS repository/license | CE Apache-2.0; no code incorporated; documentation does not establish all ETT variants or WPS versions |
| olefile API and license | Mature generic OLE stream reader, not native ET semantics; BSD-style terms plus inherited PIL notice; not installed/copied |
| libwps upstream | Microsoft Works import library; its name does not establish Kingsoft WPS ET support |
| WPS format list | Vendor service lists ET/ETT support, but that page is not a binary layout specification |

No real ET/ETT samples were present in the maintained skill. No authorized user sample
was supplied. The public SheetJS/test_files repository was disabled when checked, so it
was not usable as a licensed/versioned ET corpus. No externally sourced sample is claimed.
Current tests generate synthetic CFB structures in memory under this repository's license;
creator is the test builder, CFB versions are 3/4, WPS producer/version is not applicable.
Renamed XLS/XLSX or simulated containers are never counted as real ET/ETT acceptance.

## CLI and support levels

```console
python3 -I -S -B "<skill_dir>/scripts/xlsx_identify.py" input.et
python3 -I -S -B "<skill_dir>/scripts/xlsx_cfb.py" input.et
```

Content, not extension, selects the container probe. Non-CFB data is rejected by xlsx_cfb;
xlsx_identify still distinguishes OOXML/ZIP/unknown/damaged input without guessing ET.
The CFB report lists component-array stream paths, lengths and SHA-256 digests, not cells
or full stream contents. Workbook/Book names are only workbook_stream_candidates.

Three separate claims must remain separate: container_readable / BIFF readable / native
ET/ETT accepted. The CFB reader only emits support_level cfb_container_only; biff_parsed,
native_et_ett_parsed and cells_decoded are false. A consistent header alone no longer
passes the compound probe. xlsx_identify returns conversion_required after a readable
container, while parser_supported remains false. Known EncryptionInfo/EncryptedPackage
streams yield encrypted_unsupported; generic payload/private encryption is not decoded.
A corrupt or outside-subset container returns invalid_or_unsupported, not a fake empty sheet.

The standalone CFB CLI exits 0 after container diagnosis, including encrypted_unsupported;
invalid/unsupported structure or I/O returns 1. The general inspector retains exit 0 for
completed diagnoses; inspect JSON status and support flags, not exit alone.
Python API: CompoundFile.load(path).report() and read_stream([component,...]) -> bytes.
read_stream refuses recognized encrypted payloads; it never extracts to a filesystem path.
Raw stream bytes are not a spreadsheet model and cannot be fed directly to region/profile.

## Container subset and safety

Support version 3 (512-byte sectors) and version 4 (4096-byte sectors), header and extended
DIFAT, FAT, directory graphs, normal streams and 64-byte mini streams/MiniFAT. Check bounds,
chain termination/length, cycles, cross-linked allocation, UTF-16 names, duplicate names,
unreachable active directory entries and unreferenced allocated sectors. Never follow
OLE links, URLs or external objects, execute macros, decrypt or decompress nested streams.
Directory graph checks do not constitute full red-black ordering/conformance validation.
Strict restrictions can reject otherwise readable noncanonical CFB files: exact stream
chain lengths, no orphan allocated sectors, empty streams must use ENDOFCHAIN, casefold
name collision checks, and aligned root mini-stream sizes. Report outside-subset rather
than silently dropping payload. The CFB entrypoint interprets no spreadsheet types, dates, merges, formulas or styles.
A separate [strict BIFF8 record reader](biff8-readonly.md) now provides a bounded record
subset and explicit header adapter, without changing CFB-only or ET compatibility claims.

Limits: input 32 MiB (+ one limit-detection byte), individual user stream 8 MiB, aggregate
user streams 32 MiB, directory 4096 entries, graph depth 32, 512 FAT and 16 DIFAT sectors.
Header-declared sizes/counts are bounded before materialization. CFB has no compression
in this reader; existing OOXML ZIP member/uncompressed limits remain separate. No changes
were made to permit unbounded compression or to trust extension-based dispatch.

## Tests and remaining work

Synthetic tests cover v3/v4, Chinese names/raw bytes, normal/mini/fragmented chains,
extended DIFAT, nested storages, empty streams, cycles, cross-links, overlong declarations,
malformed pointers/names, orphan allocations, encrypted stream rejection and CLI no-write.
These verify container behavior only, not WPS or Excel interoperability.

Next obtain permitted, versioned WPS-generated ET/ETT samples plus corresponding XLSX and
producer/platform metadata. Establish stream signatures and private record rules, then
extend the documented BIFF8 subset only after verifying ET private records and comparing
real sheet/type/date/merge data. The separate BIFF8 tests are still synthetic. Unknown records must not be silently skipped to claim full data
recovery. Formula caches would remain unverified with raw representation, never computed.
Until then use a trusted local WPS copy to Save As a new XLSX and verify it; renaming is
not conversion. No native ET/ETT or application-level acceptance has been performed.
