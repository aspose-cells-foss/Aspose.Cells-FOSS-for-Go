# aspose.Cells for Go Development Guide for AI Agents

You are a senior Go engineer working on a pure-Go Excel library. Prioritize Excel compatibility, round-trip fidelity, and small, reviewable diffs.

## Do

- Treat `aspose/cells_foss/` as the source of truth for library behavior and public API
- Treat `examples/` as executable usage coverage; keep examples aligned with the current API
- Use A1-style string references for cell access: `ws.Cells().Set("A1", v)` / `ws.Cells().Get("A1")`
- Preserve loaded workbook content when it was not modified, especially XML parts cached on objects such as `sourceXML`
- Add new workbook features through the existing loader/saver split: `xml_feature_loader.go` and `xml_feature_saver.go`
- Keep worksheet XML in ECMA-376-compatible element order when adding new nodes
- Use `SaveFormat` or file extensions consistently when saving workbooks
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
  worksheet.go                  # Worksheet model
  cell.go / cells.go            # Cell model and A1-keyed collection
  style.go                      # Font, fill, border, alignment, number format
  picture.go                    # Drawing objects
  table.go                      # Excel table support
  datavalidation.go             # Validation models and enums
  csv_handler.go                # CSV import/export
  formula_engine.go             # SUM/AVERAGE/MAX/MIN evaluation
  streaming_reader.go           # Row-by-row reader for large files
  crypto.go                     # Encrypted workbooks (Agile Encryption)
  xmlloader.go                  # Workbook/sheet/rels XML loading
  xmlsaver.go                   # Workbook and sheet XML saving
  xml_feature_loader.go         # styles.xml and table XML loading
  xml_feature_saver.go          # styles.xml, tables, drawings, validations
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
fmt.Println(cell.Value)
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

- **Encrypted files are not Excel-readable.** `Workbook.SetPassword` wraps the
  package in a custom `ECRX`-magic container (`aspose/cells_foss/crypto.go`),
  not the OLE/CFB container Excel uses. Only this library can read the result.
  The key derivation itself is standard ECMA-376 Agile Encryption.
  `SetPassword` is deprecated for this reason.
- **Passwords are held in plaintext.** The old `VerifyPassword` returns `true`
  for any input when no password is set and uses non-constant-time comparison.
  Use `CheckPassword` instead — it returns `false` when no password is set and
  uses `crypto/subtle.ConstantTimeCompare`.
- **`Cell.Value` type varies.** After loading, numeric cells contain `float64`
  (not `string`), so direct type assertions like `cell.Value.(string)` may
  panic. Use `Cell.AsFloat64()`, `Cell.AsInt()`, or `Cell.AsString()` for
  type-safe access. `Cell.Value` is deprecated in favor of these methods.
- **The `Modified` flag is workbook-level.** Modifying any sheet marks the
  entire workbook as modified, causing all sheets to be regenerated on save.
  Per-sheet source XML caching (`Worksheet.sourceXML`) enables byte-identical
  round-trip only when no sheet is modified.
- **Pictures are write-only.** There is no drawing/picture loading. Multi-sheet
  workbooks with pictures use correct per-sheet drawing and media part names.
- **`test_helpers.go` exports are deprecated.** `WriteTestXLSX`,
  `ReadTestZipEntry`, and `MinimalPNG` exist for `tests/`; they are not part of
  the supported API and are marked deprecated.

## When Stuck

1. Compare generated workbook XML with a known-good Excel file
2. Trace save behavior through `aspose/cells_foss/workbook.go`, `xmlsaver.go`, and the relevant `xml_feature_saver.go`
3. Trace load behavior through `aspose/cells_foss/xmlloader.go` and the relevant `xml_feature_loader.go`
4. Check `examples/` for the intended user-facing workflow before changing API behavior
5. Write or update the smallest example or test that reproduces the problem