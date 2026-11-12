# Changelog

All notable changes to this project are documented in this file. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses calendar
versioning aligned with its module major version (`v26`): `26.11` is the November 2026
release.

Releases before 26.11 predate this file; see the
[commit history](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/commits/main)
for earlier changes.

## [26.11] — 2026-11-12

### Added

- **Number formats and date serials.** `Style.NumberFormat` accepts built-in and custom
  format codes (written to `<numFmts>` in styles.xml) and round-trips through save/load.
  `TimeToSerial` / `SerialToTime` convert between Go `time.Time` values and Excel date
  serials using the 1900 date system (`number_format.go`).
- **Named worksheets.** `Workbook.AddWorksheet(name)` creates a sheet with Excel's naming
  rules enforced (length, invalid characters, duplicate names).
- **Worksheet layout.** Merged cells, column widths, row heights, freeze panes, and
  auto-filters are now first-class worksheet features that round-trip through save/load
  (`layout.go`).
- **Cell hyperlinks.** External hyperlinks (via worksheet relationships) and internal
  hyperlinks (via the `location` attribute) with save/load support (`hyperlink.go`).
- **Formula engine expansion.** `CalculateFormula` grows from 4 to 22 functions: `SUM`,
  `AVERAGE`, `MAX`, `MIN`, `COUNT`, `COUNTIF`, `VLOOKUP`, `IF`, `AND`, `OR`, `NOT`,
  `ROUND`, `ABS`, `POWER`, `SQRT`, `CONCAT`, `LEN`, `LEFT`, `RIGHT`, `MID`, `UPPER`,
  `LOWER`. Expressions now support the arithmetic operators `+ - * / ^ &`, unary minus,
  comparisons, parentheses, literal arguments, and nested calls, plus cross-sheet
  references (`Sheet2!A1`, `'My Sheet'!B2:C3`).
- **Excel error values.** New `ErrorValue` type with the seven Excel error constants
  (`#DIV/0!`, `#N/A`, `#VALUE!`, …). Error values propagate through arithmetic,
  aggregation ranges, and cell references the way Excel does (`error_value.go`).
- **Charts: scatter and area types, axes, and legends.** `ChartTypeScatter` and
  `ChartTypeArea` join bar/line/pie; new `ChartAxis` (titles, hidden axes) and
  `ChartLegend` (position, overlay) models with spec-compliant chart XML (required axis-id
  pairs, absolute range references, literal series names).
- **Conditional formatting rule types.** `colorScale`, `dataBar`, `iconSet` (17 built-in
  icon sets), and `top10` rules join the existing cellIs/text/expression rules and
  round-trip through save/load; rule ranges are validated when added.
- **Data validation operators.** New `DataValidation.Operator` field with the eight
  `DataValidationOp*` constants covering the full OOXML `ST_DataValidationOperator`
  enumeration; invalid operators are rejected by `AddDataValidation`.
- **Row insert/delete.** `Worksheet.InsertRow` / `DeleteRow` shift cells, row heights,
  merged ranges, hyperlinks, auto-filter, data-validation / conditional-formatting /
  table / pivot ranges, chart data ranges, and picture and chart anchors. Ranges expand,
  shrink, or are dropped following Excel semantics; formula text is never rewritten
  (`rowops.go`).
- **Stream I/O.** `Workbook.SaveTo(io.Writer)`, `LoadFrom(io.Reader)`, and
  `LoadFromWithPassword(io.Reader, password)` read and write the same packages over
  arbitrary streams — HTTP responses, zip entries, network connections — without touching
  the filesystem.
- **Excel-compatible encryption.** `Workbook.SetPassword` output is rewritten to a
  standard OLE/CFB container carrying ECMA-376 Agile Encryption per MS-OFFCRYPTO
  (SHA-512 key derivation, AES-256-CBC segmented payload, HMAC-SHA512 integrity
  verification); the files open directly in Microsoft Excel. Legacy `ECRX` files written
  by earlier versions can still be loaded.
- **Ten new runnable examples** (21 total): `add_worksheet`, `chart`,
  `conditional_format`, `encryption`, `layout`, `macro`, `number_format`, `pivot_table`,
  `row_operations`, `stream_io`.

### Fixed

- **Shared strings:** loaded cells preserve their shared-string index on re-save instead
  of corrupting string values and duplicating `sharedStrings.xml` entries.
- **Dirty tracking:** `Cell.SetFormula` and `Cell.SetStyle` now mark the sheet (and
  workbook) modified, so those changes are actually persisted on save.
- **Element order:** worksheet XML emits `<f>` before `<v>` and otherwise follows ECMA-376
  ordering, fixing Excel repair prompts on generated files.
- **Package declarations:** `[Content_Types].xml` entries and relationship items are
  declared only for parts that are actually present in the package.
- **Round-trip fidelity:** unmodified sheets reuse their loaded XML and relationships
  verbatim; only sheets touched since load are regenerated, so untouched content
  round-trips byte-identically even when other sheets change.
- **Differential formats:** `dxfId` is written only when a rule has an assigned DXF
  style, and the loaded `<dxfs>` block is preserved when styles.xml is regenerated.
- **Save caching:** the in-memory XML cache is refreshed after every successful save, so
  a subsequent save without intervening edits stays fast and byte-identical (including
  retries after a failed write).
- **Part naming:** chart, pivot-table, and media part names are globally unique across
  multi-sheet workbooks; table and pivot-table display names are deduplicated; pivot-cache
  relationships are wired correctly.
- **Encryption container:** DIFAT chain walking, bounds checks that return errors instead
  of panicking or emitting corrupt data, and Agile Encryption parameters brought into
  spec compliance.
- **Inline strings:** cells stored as `t="inlineStr"` are read by both the workbook loader
  and the streaming reader.
- **Formula evaluation:** literal function arguments, quote-aware argument splitting,
  rune-based string functions (`LEFT`/`RIGHT`/`MID`/`LEN`) for non-ASCII text, and
  out-of-range `MID` returns an error instead of panicking.
- **CSV:** import infers boolean and numeric types like Excel does instead of storing
  every field as a string; export uses a sparse-safe bounding box (no more giant grids of
  empty cells) and writes a UTF-8 BOM with CRLF line endings so Excel detects encoding
  and layout correctly.
- Conditional formatting ranges are written with the spec-compliant `sqref` attribute
  (files written by older versions using `ref` still load).

### Changed

- Save-path XML generation failures now return errors instead of panicking
  (`marshalXML` / `buildEncryptionInfoXML` across all part generators).
- Previously ignored load errors (styles, feature parts, conditional formatting,
  streaming reader) are propagated to the caller.
- `verify/check_open_xlsx.go` rewritten into a thorough package validator: content types,
  relationships, and per-part structure checks.
- Dead code removed (`writeEncryptedFile`, unused cache fields); a side effect was removed
  from the load path; `splitRef` hardened against malformed range strings.
- Documentation overhauled for this release: new `docs/usage.md` chapters (stream I/O,
  row operations, number formats, layout, hyperlinks, validation operators), `AGENTS.md`
  synced with the current feature set, README updated, and this changelog added.

### Deprecated

No new deprecations in this release. Previously deprecated APIs remain deprecated:

- `Cell.Value` — use the type-safe `Cell.AsString` / `AsFloat64` / `AsInt` accessors.
- `Workbook.VerifyPassword` — use `Workbook.CheckPassword` (constant-time, and returns
  `false` when no password is set).
- `test_helpers.go` exports (`WriteTestXLSX`, `ReadTestZipEntry`, `MinimalPNG`) — test
  scaffolding, not supported API.

### Known limitations

- Row operations shift references but never rewrite formula text; inserting is refused
  when content would be pushed off the last row (1048576).
- The 1900 date system only: serials below 61 (ambiguous because of Excel's 1900
  leap-year bug) and dates beyond 9999-12-31 are rejected.
- Charts: no 3D/combo types, secondary axes, data labels, or gridline styling; pivot
  tables: no calculated fields, grouping, or custom aggregations.
- Conditional formatting: Excel 2010+ x14 extensions (gradient data bars, negative-value
  formats) are not supported.
- Hidden rows/columns, outline grouping, print setup, and sheet protection are not
  modeled; unmodeled parts (themes, docProps, comments, calcChain) are preserved
  byte-for-byte instead.
- VBA support is storage-only: `vbaProject.bin` bytes round-trip, but VBA cannot be
  created, edited, or executed.
- Stream-loaded workbooks have an empty `FilePath`.

[26.11]: https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/releases/tag/v26.11
