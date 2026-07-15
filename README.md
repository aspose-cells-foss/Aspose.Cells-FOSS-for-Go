# Aspose.Cells for Go

A pure-Go Excel (.xlsx) library — create, read, and write ECMA-376 Office Open XML spreadsheets.

[![Go Version](https://img.shields.io/badge/Go-%3E%3D1.18-blue)](https://go.dev)
[![Tests](https://img.shields.io/badge/tests-passing-green)](#)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

## Features

- **Cells** — A1-style references, read/write/remove cells, supports `string`, `float64`, `int`, `bool`
- **Styles** — Font (bold/italic/size/color), fill, alignment, borders with automatic deduplication
- **Formulas** — Write Excel formulas, built-in `SUM` / `AVERAGE` / `MAX` / `MIN` calculation engine
- **Data Validation** — Dropdown lists, whole/decimal ranges, custom error messages and styles
- **Tables** — Structured data regions with header rows and auto-filters
- **Pictures** — PNG / JPEG embedding with anchor positioning
- **CSV Import/Export** — Custom delimiters, automatic type conversion
- **Streaming Reader** — Row-by-row processing, O(row-width) memory, handles 100K+ row files
- **Encryption** — ECMA-376 Agile Encryption (SHA-512 + AES-256-CBC) with password protection

## Installation

```bash
go get github.com/aspose/cells_foss
```

Requires **Go 1.18+**. Only depends on the standard library + `golang.org/x/crypto`.

## Quick Start

```go
package main

import cells_foss "github.com/aspose/cells_foss/aspose/cells_foss"

func main() {
    // Create.
    wb := cells_foss.NewWorkbook()
    ws := wb.Worksheets[0]

    // Write.
    ws.Cells().Set("A1", "Hello, World!")
    ws.Cells().Set("B1", 42)

    // Save.
    wb.Save("hello.xlsx")
}
```

Run: `go run main.go` → produces `hello.xlsx`

## Common Use Cases

### Open and modify an existing file

```go
wb, _ := cells_foss.LoadWorkbook("input.xlsx")
ws := wb.Worksheets[0]

cell, _ := ws.Cells().Get("A1")
fmt.Println("Current value:", cell.Value)

ws.Cells().Set("A1", "Updated value")
wb.Save("output.xlsx")
```

### Add styles to a report

```go
header := cells_foss.NewStyle()
header.Font.Bold = true
header.Font.Size = 14

cell, _ := ws.Cells().Get("A1")
cell.SetStyle(header)
```

### Use formulas

```go
ws.Cells().Set("A1", float64(100))
ws.Cells().Set("A2", float64(200))

ws.Cells().Set("A3", nil)
cell, _ := ws.Cells().Get("A3")
cell.SetFormula("SUM(A1:A2)")

result, _ := cells_foss.CalculateFormula("SUM(A1:A2)", ws)
fmt.Println("Result:", result) // 300
```

### Process large files (streaming)

```go
sr := cells_foss.NewStreamingReader("huge.xlsx")
sr.ProcessRows("Sheet1", func(rowIdx int, cells map[string]string) error {
    fmt.Printf("Row %d: %v\n", rowIdx, cells)
    return nil
})
```

## Project Structure

```
├── aspose/cells_foss/        # Library source code
│   ├── workbook.go           #   Workbook entry point
│   ├── worksheet.go          #   Worksheet model
│   ├── cell.go / cells.go    #   Cell model and collection
│   ├── style.go              #   Styles (font, fill, alignment, border)
│   ├── formula_engine.go     #   Formula calculation engine
│   ├── datavalidation.go     #   Data validation
│   ├── table.go              #   Tables
│   ├── picture.go            #   Picture embedding
│   ├── csv_handler.go        #   CSV import/export
│   ├── streaming_reader.go   #   Streaming row reader
│   ├── crypto.go             #   Encryption (SHA-512 + AES-256-CBC)
│   ├── xmlloader.go          #   XML loading
│   ├── xmlsaver.go           #   XML saving
│   ├── xml_feature_loader.go #   Style/table/picture loading
│   ├── xml_feature_saver.go  #   Style/table/picture/validation saving
│   └── xml_sharedstrings_loader.go  # Shared strings
├── tests/                    # Integration tests (public API)
├── examples/                 # Runnable example programs
│   ├── basic/                #   Create and save
│   ├── load_modify_save/     #   Load → modify → save
│   ├── style/                #   Style application
│   ├── formula/              #   Formulas and calculation
│   ├── table/                #   Structured tables
│   ├── data_validation/      #   Data validation
│   ├── picture/              #   Picture embedding
│   ├── csv_export/           #   CSV export
│   ├── csv_import/           #   CSV import
│   └── streaming/            #   Streaming large data
├── docs/                     # Documentation
│   └── usage.md              #   Detailed usage guide
├── verify/                   # Verification tool
│   └── check_open_xlsx.go    #   .xlsx structure validator
├── go.mod
├── go.sum
└── README.md
```

## Testing

```bash
# All tests (integration tests)
go test ./...

# Integration tests only (public API)
go test ./tests/ -v

# Run all examples
cd examples
for d in */; do go run ./$d; done
```

## Running Examples

```bash
cd examples

go run ./basic/             # Create a basic workbook
go run ./load_modify_save/  # Load → modify → save
go run ./style/             # Bold/center/color/borders
go run ./formula/           # SUM/AVERAGE/MAX/MIN
go run ./table/             # Structured tables
go run ./data_validation/   # Dropdown lists and range validation
go run ./picture/           # PNG picture embedding
go run ./csv_export/        # Export to CSV
go run ./csv_import/        # Import from CSV
go run ./streaming/         # Stream large datasets
```

## Detailed Documentation

For the complete API reference and usage guide, see **[docs/usage.md](docs/usage.md)**.

## Constraints

- **A1-style references** (e.g. `"A1"`, `"B2"`) — tuple/array indices (e.g. `[0, 0]`) are not supported
- Modified content regenerates XML; unmodified content reuses original XML
- ECMA-376-compatible element ordering
- No third-party dependencies beyond `golang.org/x/crypto`
- Do not commit generated `.xlsx` files or contents of `outputfiles/`

## License

MIT License
