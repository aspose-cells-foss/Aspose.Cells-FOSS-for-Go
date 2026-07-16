# Aspose.Cells for Go — Usage Guide

## Table of Contents

- [Quick Start](#quick-start)
- [Workbook Operations](#workbook-operations)
- [Cell Operations](#cell-operations)
- [Styles](#styles)
- [Formulas](#formulas)
- [Data Validation](#data-validation)
- [Tables](#tables)
- [Pictures](#pictures)
- [CSV Import/Export](#csv-importexport)
- [Streaming Reader](#streaming-reader)
- [Encryption & Password Protection](#encryption--password-protection)
- [API Reference](#api-reference)

---

## Quick Start

```go
package main

import cells_foss "github.com/aspose/cells_foss/v26/aspose/cells_foss"

func main() {
    // Create a workbook.
    wb := cells_foss.NewWorkbook()
    ws := wb.Worksheets[0]

    // Write data.
    ws.Cells().Set("A1", "Hello")
    ws.Cells().Set("B1", "World")
    ws.Cells().Set("A2", 42)

    // Save.
    wb.Save("output.xlsx")
}
```

Run:
```bash
go run main.go
# → produces output.xlsx
```

---

## Workbook Operations

### Creating and Loading

```go
// Create an empty workbook (with one blank worksheet "Sheet1").
wb := cells_foss.NewWorkbook()

// Load from a file.
wb, err := cells_foss.LoadWorkbook("existing.xlsx")

// Load from an encrypted file.
wb, err := cells_foss.LoadWithPassword("encrypted.xlsx", "password")
```

### Saving

```go
// Plain save.
wb.Save("output.xlsx")

// Set a password before saving (encrypts the file).
wb.SetPassword("secret123")
wb.Save("encrypted.xlsx")

// Remove password protection.
wb.SetPassword("")
wb.Save("plain.xlsx")
```

### Multiple Worksheets

```go
// Import a CSV as a new worksheet.
wb.ImportFromCSV("data.csv", "Employees", ',')

// Manually create a worksheet (ensure Cells are initialised).
ws := &cells_foss.Worksheet{
    Name:  "Report",
    Index: len(wb.Worksheets),
}
wb.Worksheets = append(wb.Worksheets, ws)

// Access worksheets.
firstSheet := wb.Worksheets[0]
secondSheet := wb.Worksheets[1]
```

---

## Cell Operations

### Reading and Writing Cells

```go
ws := wb.Worksheets[0]
cells := ws.Cells()

// Write (type is inferred automatically).
cells.Set("A1", "text")
cells.Set("B1", 42)          // int
cells.Set("C1", 3.14)        // float64
cells.Set("D1", true)        // bool

// Read.
cell, err := cells.Get("A1")
if err == nil {
    fmt.Println(cell.Value) // "text"
}

// Remove.
cells.Remove("Z99")

// Iterate all cells.
for ref, cell := range cells.All() {
    fmt.Printf("%s = %v\n", ref, cell.Value)
}
```

### A1 Reference Convention

All cell references use **A1-style strings**: column letter(s) + row number (e.g. `"A1"`, `"B2"`, `"AA10"`). Both row and column numbering start at 1. **Tuple/array indices (e.g. `[0, 0]`) are not supported.**

---

## Styles

### Basic Styling

```go
style := cells_foss.NewStyle()

// Font.
style.Font.Name = "Arial"
style.Font.Size = 14
style.Font.Bold = true
style.Font.Italic = true
style.Font.Color = "FFFF0000" // ARGB red

// Fill.
style.Fill = &cells_foss.Fill{
    Type:  cells_foss.FillTypeSolid,
    Color: "FFFFFF00", // yellow
}

// Alignment.
style.Alignment = &cells_foss.Alignment{
    Horizontal: cells_foss.AlignCenter,
    Vertical:   cells_foss.AlignMiddle,
    WrapText:   true,
}

// Border.
style.Border = &cells_foss.Border{
    Top: true, Bottom: true, Left: true, Right: true,
    Color: "FF000000",
}

// Apply to a cell.
cell, _ := ws.Cells().Get("A1")
cell.SetStyle(style)
```

### Alignment Constants

| Constant | Value | Description |
|---|---|---|
| `AlignLeft` | `"left"` | Horizontal left |
| `AlignCenter` | `"center"` | Horizontal center |
| `AlignRight` | `"right"` | Horizontal right |
| `AlignTop` | `"top"` | Vertical top |
| `AlignMiddle` | `"center"` | Vertical middle |
| `AlignBottom` | `"bottom"` | Vertical bottom |

### Style Deduplication

Identical style values automatically share the same ID — calling `SetStyle` multiple times with the same style does not produce duplicate entries in the output.

---

## Formulas

### Setting Formulas

```go
ws.Cells().Set("A1", float64(10))
ws.Cells().Set("A2", float64(20))
ws.Cells().Set("A3", float64(30))

// Create a formula cell.
ws.Cells().Set("A4", nil)
cell, _ := ws.Cells().Get("A4")
cell.SetFormula("SUM(A1:A3)")
```

### Formula Calculation Engine

A lightweight built-in engine supports aggregate functions:

```go
// SUM — total.
result, _ := cells_foss.CalculateFormula("SUM(A1:A10)", ws)

// AVERAGE — mean.
result, _ = cells_foss.CalculateFormula("AVERAGE(B1:B5)", ws)

// MAX / MIN — extremes.
max, _ := cells_foss.CalculateFormula("MAX(C1:C20)", ws)
min, _ := cells_foss.CalculateFormula("MIN(C1:C20)", ws)
```

Supported reference formats:
- Single cell: `A1`
- Vertical range: `A1:A10`
- Horizontal range: `A1:C1`
- Rectangular range: `A1:C10`

> Formulas are written to the file on save. When opened in Excel, Excel will recalculate automatically.

---

## Data Validation

### Adding Validation Rules

```go
// Dropdown list.
dv := &cells_foss.DataValidation{
    Type:             cells_foss.DataValidationTypeList,
    Formula1:         `"Apple,Banana,Cherry"`,
    AllowBlank:       true,
    ShowErrorMessage: true,
    ErrorTitle:       "Invalid Fruit",
    ErrorMessage:     "Please pick a fruit from the list.",
    ErrorStyle:       cells_foss.ErrorStyleStop,
}
ws.AddDataValidation("A2:A10", dv)

// Whole-number range.
ws.AddDataValidation("B2:B10", &cells_foss.DataValidation{
    Type:             cells_foss.DataValidationTypeWhole,
    Formula1:         "1",
    Formula2:         "100",
    ShowErrorMessage: true,
    ErrorTitle:       "Out of Range",
    ErrorMessage:     "Enter a whole number between 1 and 100.",
})

// Remove.
ws.RemoveDataValidation("A2:A10")
```

### Validation Type Constants

| Constant | Value |
|---|---|
| `DataValidationTypeList` | `"list"` |
| `DataValidationTypeWhole` | `"whole"` |
| `DataValidationTypeDecimal` | `"decimal"` |
| `DataValidationTypeDate` | `"date"` |
| `DataValidationTypeTime` | `"time"` |
| `DataValidationTypeTextLength` | `"textLength"` |

### Error Styles

| Constant | Effect |
|---|---|
| `ErrorStyleStop` | Red cross — blocks invalid input |
| `ErrorStyleWarning` | Yellow warning — allows override |
| `ErrorStyleInformation` | Blue info — allows override |

---

## Tables

### Creating Tables

```go
// Populate data.
ws.Cells().Set("A1", "Name")
ws.Cells().Set("B1", "Score")
ws.Cells().Set("A2", "Alice")
ws.Cells().Set("B2", float64(95))

// Create a table over the data range.
tbl := ws.AddTable("A1:B2")
tbl.HasHeaderRow = true
tbl.StyleName = "TableStyleMedium6"

// Look up a table by name.
found := ws.GetTable("Table1")
fmt.Println(found.Range) // "A1:B2"
```

When opened in Excel, tables display filter buttons and built-in styling.

---

## Pictures

### Embedding Pictures

```go
// Read image bytes.
data, _ := os.ReadFile("logo.png")

// Create a Picture.
pic := cells_foss.NewPicture(data, "png")
pic.Width = 200
pic.Height = 150
pic.SetAnchor(2, 1) // row 2, column B

ws.AddPicture(pic)
```

Supported formats: **PNG**, **JPEG**.

---

## CSV Import/Export

### Exporting to CSV

```go
// Export the first worksheet.
wb.ExportToCSV(0, "output.csv", ',')   // comma-separated
wb.ExportToCSV(0, "output.tsv", '\t')  // tab-separated

// Get a 2D string slice.
data, _ := ws.ToCSV(',')
fmt.Println(data[0][0]) // first row, first column
```

### Importing from CSV

```go
// Import as a new worksheet.
wb.ImportFromCSV("input.csv", "Data", ',')

// Populate from existing data.
ws.FromCSV(csvRows, ',')
```

---

## Streaming Reader

`StreamingReader` processes data one row at a time with constant memory usage (`O(row-width)`), independent of file size.

```go
sr := cells_foss.NewStreamingReader("huge.xlsx")

sr.ProcessRows("Sheet1", func(rowIdx int, cells map[string]string) error {
    fmt.Printf("Row %d: %v\n", rowIdx, cells)

    // Return an error to stop processing early.
    if rowIdx >= 1000 {
        return fmt.Errorf("stop")
    }
    return nil
})
```

---

## Encryption & Password Protection

```go
// Save an encrypted file.
wb := cells_foss.NewWorkbook()
wb.Worksheets[0].Cells().Set("A1", "Confidential data")
wb.SetPassword("mypassword")
wb.Save("protected.xlsx")

// Open an encrypted file.
wb, err := cells_foss.LoadWithPassword("protected.xlsx", "mypassword")
if err != nil {
    // Wrong password.
}

// Remove password protection.
wb.SetPassword("")
wb.Save("unprotected.xlsx")
```

Encryption uses **ECMA-376 Agile Encryption** (SHA-512 iterative hashing + AES-256-CBC).

---

## API Reference

### Workbook

| Method | Description |
|---|---|
| `NewWorkbook() *Workbook` | Create an empty workbook |
| `LoadWorkbook(path) (*Workbook, error)` | Load from an .xlsx file |
| `LoadWithPassword(path, pw) (*Workbook, error)` | Load from an encrypted file |
| `Save(path) error` | Save (encrypts if a password is set) |
| `SetPassword(pw) error` | Set or change the password |
| `ExportToCSV(idx, path, delim) error` | Export a worksheet to CSV |
| `ImportFromCSV(path, name, delim) error` | Import a CSV as a new worksheet |

### Cells

| Method | Description |
|---|---|
| `Get(ref) (*Cell, error)` | Get a cell by A1 reference |
| `Set(ref, value) error` | Set a cell value |
| `Remove(ref) error` | Remove a cell |
| `All() map[string]*Cell` | Return all cells |

### Cell

| Method | Description |
|---|---|
| `SetStyle(*Style) error` | Apply a style |
| `GetStyle() *Style` | Get the current style |
| `SetFormula(f string)` | Set a formula |
| `GetFormula() string` | Get the formula |
| `Value interface{}` | Cell value |
| `Ref string` | A1 reference address |

### Worksheet

| Method | Description |
|---|---|
| `Cells() *Cells` | Get the cell collection |
| `AddTable(ref) *Table` | Create a table |
| `GetTable(name) *Table` | Look up a table by name |
| `AddDataValidation(ref, *DataValidation) error` | Add a data validation rule |
| `RemoveDataValidation(ref) error` | Remove a data validation rule |
| `AddPicture(*Picture) error` | Embed a picture |
| `ToCSV(delim) ([][]string, error)` | Export as a 2D string slice |
| `FromCSV(data, delim) error` | Populate from a 2D string slice |
