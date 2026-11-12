# aspose.Cells for Go Development Guide for AI Agents

You are a senior Go engineer working on a pure-Go Excel library. Prioritize Excel compatibility, round-trip fidelity, and small, reviewable diffs.

## Do

- Treat `aspose/cells_foss/` as the source of truth for library behavior and public API
- Treat `examples/` as executable usage coverage; keep examples aligned with the current API
- Use A1-style string references for cell access: `ws.Cells().Set("A1", v)` / `ws.Cells().Get("A1")`
- Preserve loaded workbook content when it was not modified, especially XML parts cached on objects such as `sourceXML`
- Add new workbook features through the existing loader/saver split: `xml_feature_loader.go` and `xml_feature_saver.go`
- Keep worksheet XML in ECMA-376-compatible element order when adding new nodes
- Use file extensions consistently when saving workbooks
- Add or update example coverage when changing user-facing behavior in `aspose/cells_foss/`
- Prefer stdlib packages already used by the repo (`encoding/xml`, `archive/zip`) over new dependencies
- Run targeted tests for the area you changed before finishing

## Never

- Never use tuple/array-index cell addressing like `ws.Cells[0][0]` — the collection is keyed by A1 strings
- Never regenerate XML for loaded objects that were not changed if preserved source XML is available
- Never change public exports in `aspose/cells_foss/` package without verifying the API impact
- Never add third-party dependencies without approval
- Never commit generated `.xlsx` files or contents from `outputfiles/`
- Never hard-code behavior in examples that contradicts the implementation in `aspose/cells_foss/`
- Never add comments that only restate the code

## PR Size Guidelines

Keep changes focused and easy to review.

- **Lines changed**: Prefer under 500 lines of code
- **Files changed**: Prefer under 10 code files
- **Single responsibility**: One feature, bugfix, or refactor per change

When the task is larger, split it by:
1. Core model/API changes in `aspose/cells_foss/`
2. XML load/save wiring
3. Example coverage in `examples/`

## Commands

Key commands:
```bash
go test ./tests/ -v          # library test suite (run from the repo root)
go test ./...                # includes tests/; examples/ is a separate module

go run ./verify/check_open_xlsx.go <file.xlsx>   # validate .xlsx package structure

cd examples && go run ./basic/      # run a single example
cd examples && go run ./load_modify_save/   # needs basic/ to have run first
```

`examples/` is its own module holding `main` packages, so it is exercised with
`go run ./<name>/`, not `go test`. Examples write to `examples/outputfiles/`.

## Boundaries

### Always do
- Read the relevant modules in `aspose/cells_foss/` before changing behavior
- Update matching examples in `examples/` when you change a public workflow
- Run relevant tests for the changed area
- Preserve backward-compatible API behavior unless the task explicitly requires a change

### Ask first
- Adding dependencies
- Changing public struct names, method signatures, or exported symbols
- Deleting or renaming source files
- Large cross-cutting refactors across multiple workbook features

### Never do
- Commit secrets, credentials, or local machine paths
- Commit generated files from `outputfiles/`
- Replace preserved source XML with regenerated XML unless the object was actually modified
- Use tuple-style cell addressing

## Project Structure

```
doc.go                          # Root package doc (pkg.go.dev indexing only)
aspose/cells_foss/              # Library source code (canonical location)
  workbook.go                   # Workbook entry point and save/load dispatch
  worksheet.go                  # Worksheet model and Workbook.AddWorksheet (name validation per Excel rules)
  cell.go / cells.go            # Cell model and A1-keyed collection
  style.go                      # Font, fill, border, alignment, number format
  number_format.go              # Number format codes (built-in ids + custom <numFmts>) and Excel date serial helpers (1900 date system)
  picture.go                    # Drawing objects
  chart.go                      # Chart support (bar, line, pie, scatter, area) with axes, legend, and save/load
  table.go                      # Excel table support
  pivot_table.go                # Pivot table support with save/load
  layout.go                     # Worksheet layout: merged cells, column widths, row heights, freeze panes, autoFilter
  rowops.go                     # Worksheet.InsertRow/DeleteRow: shift cells and every range-bearing feature (formulas not rewritten)
  hyperlink.go                  # Cell hyperlinks (external via relationships, internal via location attribute)
  datavalidation.go             # Validation models and enums (types + comparison operators, save/load)
  conditional_format.go         # Conditional formatting: cellIs/text/expression with DXF styles, plus colorScale, dataBar, iconSet, top10
  macro.go                      # VBA project storage (read/write binary data)
  errors.go                     # Custom error types
  csv_handler.go                # CSV import/export
  formula_engine.go             # Formula evaluation: functions (SUM, AVERAGE, ROUND, ABS, IF, COUNTIF, VLOOKUP, etc.), arithmetic expressions (+ - * / ^ & with literals and nested calls), cross-sheet refs (Sheet2!A1, 'My Sheet'!B2:C3)
  error_value.go                # Excel error values (#DIV/0! etc.) as cell values with Excel-style propagation through formulas
  streaming_reader.go           # Row-by-row reader for large files with filtering
  crypto.go                     # Encrypted workbooks (Agile Encryption)
  olecfb.go                     # OLE/CFB container for Excel-compatible encryption
  xmlloader.go                  # Workbook/sheet/rels XML loading
  xmlsaver.go                   # Workbook and sheet XML saving
  xml_feature_loader.go         # styles.xml, table, drawing, conditional formatting loading
  xml_feature_saver.go          # styles.xml, tables, drawings, validations, conditional formatting, charts
  xml_sharedstrings_loader.go   # sharedStrings.xml loading
  test_helpers.go               # Helpers exported for tests/ (not stable API)
examples/                       # Executable usage coverage (separate module, main packages)
  outputfiles/                  # Output from examples/ (gitignored, never commit)
tests/                          # Library test suite (external test package)
verify/check_open_xlsx.go       # Standalone .xlsx structure validator
docs/usage.md                   # Usage guide
```

## Tech Stack

- **Language**: Go 1.24.5 (see `go.mod`)
- **Workbook format**: .xlsx / ECMA-376 Open XML
- **XML**: `encoding/xml`
- **Archives**: `archive/zip`
- **Testing**: `testing` package only — no third-party assertion library
- **Excel verification**: via `verify/check_open_xlsx.go`

## Code Examples

### Good cell access
```go
wb := NewWorkbook()
ws := wb.Worksheets[0]
ws.Cells().Set("A1", "Revenue")
ws.Cells().Set("B2", 42)

cell, err := ws.Cells().Get("A1")   // errors when the cell does not exist
fmt.Println(cell.AsString())         // type-safe accessor (cell.Value is deprecated)
```

### Bad cell access
```go
ws.Cells()[0][0].Value = "Revenue"  // Don't use tuple/array indices
```

### Good feature extension pattern
```go
// Add model behavior in aspose/cells_foss/<feature>.go
// Load existing files in aspose/cells_foss/xml_feature_loader.go
// Save new or modified content in aspose/cells_foss/xml_feature_saver.go
```

### Example alignment
```go
package main

import (
    "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
    wb := cells_foss.NewWorkbook()
    ws := wb.Worksheets[0]
    dv := &cells_foss.DataValidation{
        Type:     cells_foss.DataValidationTypeList,
        Formula1: `"Yes,No"`,
    }
    ws.AddDataValidation("A1:A10", dv)
    wb.Save("outputfiles/example.xlsx")
}
```

## PR Checklist

- [ ] Change is focused and reviewable
- [ ] Relevant tests pass
- [ ] `examples/` still reflects the current API
- [ ] Public exports in `aspose/cells_foss/` package are correct
- [ ] No generated `.xlsx` files or `outputfiles/` artifacts are included

## Known Limitations

Do not assume these already work; they are documented so that decisions are not
made from misleading descriptions.

- **Encrypted files use OLE/CFB container.** `Workbook.SetPassword` produces
  standard OLE/CFB containers with ECMA-376 Agile Encryption, compatible with
  Microsoft Excel. The implementation is in `aspose/cells_foss/olecfb.go`.
  Legacy `ECRX`-format files (from earlier versions) can still be loaded.
- **Passwords are held in plaintext.** The old `VerifyPassword` returns `true`
  for any input when no password is set and uses non-constant-time comparison.
  Use `CheckPassword` instead — it returns `false` when no password is set and
  uses `crypto/subtle.ConstantTimeCompare`.
- **`Cell.Value` type varies.** After loading, numeric cells contain `float64`
  (not `string`), so direct type assertions like `cell.Value.(string)` may
  panic. Use `Cell.AsFloat64()`, `Cell.AsInt()`, or `Cell.AsString()` for
  type-safe access. `Cell.Value` is deprecated in favor of these methods.
- **Per-sheet `Modified` tracking.** Each `Worksheet` has a `Modified` bool
  field set to true when its content changes. The saver uses this to decide
  whether to regenerate a sheet's XML or reuse its cached `sourceXML`. The
  workbook-level `Workbook.Modified` flag is also maintained for backward
  compatibility. Unmodified sheets preserve byte-identical round-trip even
  when other sheets are modified.
- **Pictures support read/write.** Pictures can be loaded from existing .xlsx
  files and saved to new files. The implementation parses drawing XML and
  relationships to reconstruct `Picture` objects with position, size, and
  image data. Multi-sheet workbooks use correct per-sheet drawing and media
  part names.
- **Charts support save/load with axes and legend.** Charts can be created, saved
  to .xlsx files, and loaded back. Supported types: bar, line, pie, scatter, area.
  Cartesian charts write spec-required axis ids plus optional axis titles, hidden
  axes, and legend position/overlay. Advanced chart types (3D, combo, etc.),
  secondary axes, data labels, and gridline styling are not supported. Chart data
  is referenced by cell range, not embedded.
- **Pivot tables support basic save/load.** Pivot tables can be created, saved to
  .xlsx files, and loaded back with row fields, column fields, and data fields.
  Advanced features like calculated fields, grouping, and custom aggregations are
  not supported.
- **Conditional formatting supports DXF styles and gradient rules.** Rules can
  include Style objects with font, fill, border, alignment, and number format
  properties; these are saved as Differential Formatting (DXF) elements in
  styles.xml and referenced by the rules. colorScale, dataBar, iconSet, and
  top10 rule types round-trip through save/load (ranges are addressed with the
  spec-compliant `sqref` attribute; files written by older versions using `ref`
  still load). Excel 2010+ extensions (x14 data bar gradients, negative-value
  formats) are not supported.
- **Layout covers the visual basics.** Merged cells, column widths, row heights,
  freeze panes, autoFilter, and cell hyperlinks round-trip through save/load
  (`layout.go`, `hyperlink.go`). Hidden rows/columns, outline grouping, print
  setup, and sheet protection are not supported.
- **Row operations shift references but never formulas.** `InsertRow` /
  `DeleteRow` (`rowops.go`) move cells, merged ranges, row heights, hyperlinks,
  autoFilter, data-validation / conditional-formatting / table / pivot ranges,
  chart data ranges, and picture and chart anchors; features whose whole range
  vanishes on delete are dropped. Formula text (cell formulas, validation and
  conditional-format formulas) is not rewritten, comment anchors in preserved
  parts do not move, and inserting is refused when content would be pushed off
  the last row (1048576).
- **Formula error values propagate with limits.** `CalculateFormula` returns
  Excel error values (`ErrDiv0` etc.) as results, and they propagate through
  arithmetic, aggregation ranges, and cell references. Comparisons and text
  functions treat an error operand as its text form instead of propagating,
  and external-workbook references (`[Book.xlsx]Sheet!A1`) are not supported —
  sheet-qualified references resolve within the same workbook only.
- **Number formats use the 1900 date system only.** `Style.NumberFormat` accepts
  built-in and custom codes and round-trips through save/load. `TimeToSerial` /
  `SerialToTime` reject serials below 61 (1900-03-01), because Excel's 1900
  leap-year bug makes that range ambiguous, and reject dates beyond 9999-12-31.
  The 1904 date system is not supported.
- **VBA macro support is storage-only.** `Workbook.SetVBAProject` accepts raw
  `vbaProject.bin` binary data and stores it for round-trip preservation. The
  library cannot create, edit, or execute VBA code. Use .xlsm extension when
  saving macro-enabled workbooks.
- **`test_helpers.go` exports are deprecated.** `WriteTestXLSX`,
  `ReadTestZipEntry`, and `MinimalPNG` exist for `tests/`; they are not part of
  the supported API and are marked deprecated.

## When Stuck

1. Compare generated workbook XML with a known-good Excel file
2. Trace save behavior through `aspose/cells_foss/workbook.go`, `xmlsaver.go`, and the relevant `xml_feature_saver.go`
3. Trace load behavior through `aspose/cells_foss/xmlloader.go` and the relevant `xml_feature_loader.go`
4. Check `examples/` for the intended user-facing workflow before changing API behavior
5. Write or update the smallest example or test that reproduces the problem