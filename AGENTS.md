# aspose.Cells for Go Development Guide for AI Agents

You are a senior Go engineer working on a pure-Go Excel library. Prioritize Excel compatibility, round-trip fidelity, and small, reviewable diffs.

## Do

- Treat `aspose/cells_foss/` as the source of truth for library behavior and public API
- Treat `examples/` as executable usage coverage; keep examples aligned with the current API
- Use A1-style string references for cell access: `ws.Cells["A1"]`
- Preserve loaded workbook content when it was not modified, especially XML parts cached on objects such as `sourceXML`
- Add new workbook features through the existing loader/saver split: `xml_feature_loader.go` and `xml_feature_saver.go`
- Keep worksheet XML in ECMA-376-compatible element order when adding new nodes
- Use `SaveFormat` or file extensions consistently when saving workbooks
- Add or update example coverage when changing user-facing behavior in `aspose/cells_foss/`
- Prefer stdlib packages already used by the repo (`encoding/xml`, `archive/zip`) over new dependencies
- Run targeted tests for the area you changed before finishing

## Never

- Never use tuple-based cell keys like `ws.Cells[0, 0]`
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
go test ./examples -v
go test ./examples/test_feature_test.go -v
```

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
aspose/cells_foss/              # Library source code (canonical location)
  workbook.go                   # Workbook entry point and save/load dispatch
  worksheet.go                  # Worksheet model
  cell.go / cells.go            # Cell model and A1-keyed collection
  style.go                      # Font, fill, border, alignment, number format
  chart.go                      # Chart models and enums
  picture.go / shape.go         # Drawing objects
  table.go                      # Excel table support
  sparkline.go                  # Sparkline support
  datavalidation.go             # Validation models and enums
  autofilter.go                 # Filter models
  documentproperties.go         # Core and extended document properties
  workbookproperties.go         # Workbook-level settings and protection
  csvhandler.go                 # CSV import/export
  markdownhandler.go            # Markdown export
  jsonhandler.go                # JSON export
  xmlloader.go                  # Workbook XML loading
  xmlsaver.go                   # Workbook XML saving
  xml_feature_loader.go         # Feature-specific XML loaders
  xml_feature_saver.go          # Feature-specific XML savers
examples/                       # Executable example tests for library features
  outputfiles/                  # Output from examples/ tests
```

## Tech Stack

- **Language**: Go 1.18+
- **Workbook format**: .xlsx / ECMA-376 Open XML
- **XML**: `encoding/xml`
- **Archives**: `archive/zip`
- **Testing**: `testing` package, `testify` for assertions
- **Excel verification**: via `verify/check_open_xlsx.go`

## Code Examples

### Good cell access
```go
wb := NewWorkbook()
ws := wb.Worksheets[0]
ws.Cells["A1"].Value = "Revenue"
ws.Cells["B2"].Value = 42
```

### Bad cell access
```go
ws.Cells[0][0].Value = "Revenue"  // Don't use tuple/array indices
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
    "github.com/aspose/cells_foss"
)

func main() {
    wb := cells_foss.NewWorkbook()
    ws := wb.Worksheets[0]
    dv := ws.DataValidations.Add("A1:A10")
    dv.Type = cells_foss.DataValidationTypeList
    dv.Formula1 = `"Yes,No"`
    wb.Save("outputfiles/example.xlsx")
}
```

## PR Checklist

- [ ] Change is focused and reviewable
- [ ] Relevant tests pass
- [ ] `examples/` still reflects the current API
- [ ] Public exports in `aspose/cells_foss/` package are correct
- [ ] No generated `.xlsx` files or `outputfiles/` artifacts are included

## When Stuck

1. Compare generated workbook XML with a known-good Excel file
2. Trace save behavior through `aspose/cells_foss/workbook.go`, `xmlsaver.go`, and the relevant `xml_feature_saver.go`
3. Trace load behavior through `aspose/cells_foss/xmlloader.go` and the relevant `xml_feature_loader.go`
4. Check `examples/` for the intended user-facing workflow before changing API behavior
5. Write or update the smallest example or test that reproduces the problem