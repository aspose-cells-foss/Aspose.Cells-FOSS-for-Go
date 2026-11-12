# Aspose.Cells for Go — Usage Guide

## Table of Contents

- [Quick Start](#quick-start)
- [Workbook Operations](#workbook-operations)
- [Cell Operations](#cell-operations)
- [Styles](#styles)
- [Worksheet Layout](#worksheet-layout)
- [Formulas](#formulas)
- [Data Validation](#data-validation)
- [Tables](#tables)
- [Pictures](#pictures)
- [Charts](#charts)
- [Pivot Tables](#pivot-tables)
- [Conditional Formatting](#conditional-formatting)
- [CSV Import/Export](#csv-importexport)
- [Streaming Reader](#streaming-reader)
- [VBA Macros](#vba-macros)
- [Encryption & Password Protection](#encryption--password-protection)
- [API Reference](#api-reference)

---

## Quick Start

```go
package main

import cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"

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

### Stream I/O

`SaveTo` writes the exact same .xlsx bytes `Save` would — verbatim round-trip of unmodified packages and password encryption included — to any `io.Writer`. `LoadFrom` reads a package from any `io.Reader`. Neither touches `FilePath`, since a stream has no file location.

```go
// Save to an in-memory buffer (an http.ResponseWriter works the same).
var buf bytes.Buffer
err := wb.SaveTo(&buf)

// Load from a stream. Encrypted packages are rejected — use the
// password variant for those.
wb, err := cells_foss.LoadFrom(bytes.NewReader(buf.Bytes()))
wb, err = cells_foss.LoadFromWithPassword(r, "secret123")
```

A workbook loaded from an encrypted stream keeps its password, so a later `SaveTo` re-encrypts. See `examples/stream_io/`.

### Multiple Worksheets

```go
// Add a named worksheet. The returned sheet is ready for cell access.
ws, err := wb.AddWorksheet("Report")
ws.Cells().Set("A1", "Quarterly figures")

// Import a CSV as a new worksheet (also validates the name).
err = wb.ImportFromCSV("data.csv", "Employees", ',')

// Access worksheets.
firstSheet := wb.Worksheets[0]
secondSheet := wb.Worksheets[1]
```

`AddWorksheet` enforces Excel's sheet-naming rules and returns an error for
names that would produce a file Excel rejects: empty or longer than 31
characters, containing any of `[]:*?/\`, starting or ending with an
apostrophe, equal to the reserved name `History`, or duplicating an existing
sheet name (case-insensitive). Renaming the default sheet is done by
assigning `wb.Worksheets[0].Name` before saving.

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

Identical style values automatically share the same ID — calling `SetStyle` multiple times with the same style does not produce duplicate entries in the output. `NumberFormat` is part of style identity: two styles differing only in their format code are distinct.

### Number Formats

`Style.NumberFormat` takes any Excel number format code. Codes Excel knows as built-ins (`"0.00"`, `"0%"`, `"mm-dd-yy"`, ...) are stored under their standard `numFmtId`; any other code is written as a custom `<numFmt>` entry in styles.xml with an id starting at 164. Empty (or `"General"`) means the General format.

```go
price := cells_foss.NewStyle()
price.NumberFormat = `"$"#,##0.00`         // custom code
share := cells_foss.NewStyle()
share.NumberFormat = cells_foss.NumFmtPercent2 // "0.00%" — built-in id 10

cell, _ := ws.Cells().Get("B2")
cell.SetStyle(price)
```

Common codes are available as constants: `NumFmtGeneral`, `NumFmtInt` (`"0"`), `NumFmtFloat2` (`"0.00"`), `NumFmtThousand` (`"#,##0"`), `NumFmtMoney` (`"#,##0.00"`), `NumFmtPercent`, `NumFmtPercent2`, `NumFmtDate` (`"yyyy-mm-dd"`), `NumFmtDateTime`, `NumFmtTime`, `NumFmtText` (`"@"`). Formats round-trip through save/load, including on workbooks loaded from existing files. Conditional-formatting rule styles support `NumberFormat` too; it is emitted inside the rule's `<dxf>` entry.

### Dates and Times

Excel stores dates as serial numbers: whole days since 1899-12-30 plus a fraction for the time of day (1900 date system). `TimeToSerial` and `SerialToTime` convert between Go times and serials:

```go
serial, err := cells_foss.TimeToSerial(time.Date(2024, 3, 15, 13, 45, 0, 0, time.UTC))
// serial == 45366.572916...

ws.Cells().Set("D2", serial)
cell, _ := ws.Cells().Get("D2")
dateStyle := cells_foss.NewStyle()
dateStyle.NumberFormat = cells_foss.NumFmtDateTime // makes Excel render it as a date
cell.SetStyle(dateStyle)

back, err := cells_foss.SerialToTime(serial) // 2024-03-15 13:45:00 UTC
```

Conversion uses wall-clock fields only — the input's time zone is ignored and `SerialToTime` returns UTC. Excel's historical leap-year bug (it treats 1900 as a leap year) makes low serials ambiguous, so both functions reject the range before 1900-03-01 (serial 61, including the nonexistent 1900-02-29); serials above 2958465 (9999-12-31) are rejected as well. Only the 1900 date system is supported — the 1904 system is not.

---

## Worksheet Layout

All layout properties round-trip through save/load: they are parsed back into the `Worksheet` so regenerated sheet XML keeps them. A runnable demonstration lives in `examples/layout/`.

### Inserting and Deleting Rows

```go
err := ws.InsertRow(3)   // blank row at 3; old row 3 and below move down
err = ws.DeleteRow(5)    // row 5 is removed; rows below move up
```

Both take a 1-based row number and move everything row-aware along with the data: cells (values, formulas, styles), merged ranges, row heights, hyperlinks, the autofilter range, data-validation and conditional-formatting ranges, table and pivot-table ranges, chart data ranges, and picture and chart anchors. Ranges spanning an inserted row expand; ranges spanning a deleted row shrink, and features whose whole range vanishes (single-row merges, validations on the deleted row only) are dropped. Inserting fails when sheet content already reaches the last row (1048576).

Formula text is **not** rewritten: `=SUM(B2:B4)` keeps its references even when those rows move — update formulas yourself if the shift changes their meaning. See `examples/row_operations/`.

### Merged Cells

```go
ws.Cells().Set("A1", "Quarterly Sales Report")
err := ws.AddMergedCells("A1:D1")  // range must cover at least two cells
err = ws.RemoveMergedCells("A1:D1")
```

References are normalized (upper-case, corners ordered). Excel rejects files whose merged ranges overlap, so duplicates and overlapping ranges are refused. The current merges are readable from `ws.MergedCells` (`[]string` of normalized A1 ranges).

### Column Widths and Row Heights

```go
err := ws.SetColumnWidth("B", 20)   // character units, (0, 255]; default is 8.43
err = ws.SetRowHeight(1, 26)        // points, (0, 409.5]; default is 15
```

Widths are keyed by column letter in `ws.ColumnWidths` (`map[string]float64`) and heights by 1-based row in `ws.RowHeights` (`map[int]float64`). Rows carrying only a height — no cells — are still written to the file.

### Frozen Panes

```go
err := ws.FreezePanes(2, 1)  // keep the top 2 rows and leftmost column visible
err = ws.FreezePanes(0, 0)  // remove the freeze
```

The counts are readable from `ws.FrozenRow` / `ws.FrozenCol`.

### AutoFilter

```go
err := ws.SetAutoFilter("A2:D5")  // filter buttons over the range; normalized
err = ws.SetAutoFilter("")        // remove the filter
```

The current range is `ws.AutoFilterRange`.

### Hyperlinks

```go
ws.Cells().Set("E3", "Q1 report")  // the cell value is the visible label
err := ws.AddHyperlink("E3", "https://example.com/reports/q1")
err = ws.AddInternalHyperlink("E4", "Details!A1")  // jump inside the workbook
err = ws.RemoveHyperlink("E3")
```

External links are stored as sheet relationships, internal links as a `location` attribute. Excel allows one link per cell — a second link on the same reference is refused. Links are readable from `ws.Hyperlinks` (`[]*Hyperlink` with `Ref`, `Target`, `Location`, `Tooltip`, `Display`; exactly one of `Target`/`Location` is set).

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

A lightweight built-in engine evaluates formulas against a worksheet:

```go
// Aggregation.
result, _ := cells_foss.CalculateFormula("SUM(A1:A10)", ws)
result, _ = cells_foss.CalculateFormula("AVERAGE(B1:B5)", ws)
result, _ = cells_foss.CalculateFormula("MAX(C1:C20)", ws)
result, _ = cells_foss.CalculateFormula("MIN(C1:C20)", ws)
result, _ = cells_foss.CalculateFormula("COUNT(A1:A10, 5)", ws)

// Logic, lookup, and conditional counting.
result, _ = cells_foss.CalculateFormula(`IF(A1>10,"high","low")`, ws)
result, _ = cells_foss.CalculateFormula("AND(A1>0,B1>0)", ws)
result, _ = cells_foss.CalculateFormula("COUNTIF(A1:A10,\">15\")", ws)
result, _ = cells_foss.CalculateFormula("VLOOKUP(\"Banana\",A1:B10,2)", ws)

// Math and text.
result, _ = cells_foss.CalculateFormula("ROUND(3.14159,2)", ws)   // 3.14
result, _ = cells_foss.CalculateFormula("POWER(2,10)", ws)          // 1024
result, _ = cells_foss.CalculateFormula("LEN(A1)", ws)              // counts runes, not bytes
result, _ = cells_foss.CalculateFormula(`CONCAT(A1," / ",B1)`, ws)
result, _ = cells_foss.CalculateFormula("UPPER(LEFT(A1,3))", ws)

// Arithmetic expressions and literals work anywhere, including as
// arguments and nested calls.
result, _ = cells_foss.CalculateFormula("1+1", ws)                  // 2
result, _ = cells_foss.CalculateFormula("SUM(A1:A3)*2+10", ws)
result, _ = cells_foss.CalculateFormula("IF(A1>0,SUM(B1:B5),0)", ws)
result, _ = cells_foss.CalculateFormula(`A1&"!"`, ws)               // & concatenates
```

Supported functions: `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNT`, `IF`, `AND`,
`OR`, `NOT`, `COUNTIF`, `VLOOKUP`, `ROUND`, `ABS`, `POWER`, `SQRT`, `CONCAT`,
`LEN`, `LEFT`, `RIGHT`, `MID`, `UPPER`, `LOWER`.

Operators: `+` `-` `*` `/` `^` (power) and `&` (string concatenation), with
parentheses and unary minus.

Supported reference formats:
- Single cell: `A1`
- Vertical range: `A1:A10`
- Horizontal range: `A1:C1`
- Rectangular range: `A1:C10`
- Sheet-qualified cell: `Sheet2!A1`
- Sheet-qualified range: `Sheet2!A1:C10`
- Quoted sheet name (for names with spaces): `'My Sheet'!B2`

### Cross-Sheet References

A reference may be qualified with a sheet name; the sheet is looked up (case-insensitively, as Excel does) in the workbook the passed worksheet belongs to:

```go
result, _ := cells_foss.CalculateFormula("SUM(Sheet2!B2:B6)", ws)
result, _ = cells_foss.CalculateFormula("'Q1 Data'!C3*2", ws)
```

A reference to a sheet that does not exist returns a `*SheetNotFoundError` (checkable with `errors.As`). Cross-sheet references require the worksheet to be attached to a `Workbook` — the sheets created by `NewWorkbook` and `AddWorksheet`, and those produced by loading a file, all are.

### Error Values

The engine follows Excel's error-value semantics: an invalid result is a *value*, not a Go error, and it propagates through subsequent operations.

```go
v, err := cells_foss.CalculateFormula("1/0", ws)   // v == cells_foss.ErrDiv0, err == nil
v, _ = cells_foss.CalculateFormula(`"text"+1`, ws) // cells_foss.ErrValue
v, _ = cells_foss.CalculateFormula("SQRT(-4)", ws) // cells_foss.ErrNum
```

The seven Excel error values are exported as `ErrDiv0`, `ErrNA`, `ErrName`, `ErrNull`, `ErrNum`, `ErrRef`, `ErrValue` (all of type `ErrorValue`). Store one in a cell like any other value — it is written as a `t="e"` cell and reads back as an `ErrorValue`:

```go
ws.Cells().Set("B11", cells_foss.ErrDiv0)
c, _ := ws.Cells().Get("B11")
c.AsString()          // "#DIV/0!"
_, ok := c.AsFloat64() // ok == false
```

An error inside an aggregation range propagates out (`SUM` over a range containing `#DIV/0!` yields `#DIV/0!`), matching Excel.

Text functions are Unicode-aware: `LEN("日本語")` is 3, and `LEFT`/`RIGHT`/`MID`
never split UTF-8 sequences.

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

// Whole-number range. Operator picks the comparison; empty means "between".
ws.AddDataValidation("B2:B10", &cells_foss.DataValidation{
    Type:             cells_foss.DataValidationTypeWhole,
    Operator:         cells_foss.DataValidationOpBetween,
    Formula1:         "1",
    Formula2:         "100",
    ShowErrorMessage: true,
    ErrorTitle:       "Out of Range",
    ErrorMessage:     "Enter a whole number between 1 and 100.",
})

// Single-bound comparisons need only Formula1.
ws.AddDataValidation("C2:C10", &cells_foss.DataValidation{
    Type:     cells_foss.DataValidationTypeDecimal,
    Operator: cells_foss.DataValidationOpGreaterThanOrEqual,
    Formula1: "0",
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

### Operator Constants

`Operator` applies to the value-comparing types (`whole`, `decimal`, `date`, `time`, `textLength`); list and custom validations ignore it. Empty means `"between"`, the schema default, and is omitted from the saved XML just like Excel does. `AddDataValidation` rejects unknown operators.

| Constant | Value |
|---|---|
| `DataValidationOpBetween` | `"between"` |
| `DataValidationOpNotBetween` | `"notBetween"` |
| `DataValidationOpEqual` | `"equal"` |
| `DataValidationOpNotEqual` | `"notEqual"` |
| `DataValidationOpLessThan` | `"lessThan"` |
| `DataValidationOpLessThanOrEqual` | `"lessThanOrEqual"` |
| `DataValidationOpGreaterThan` | `"greaterThan"` |
| `DataValidationOpGreaterThanOrEqual` | `"greaterThanOrEqual"` |

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

## Charts

Bar, line, pie, scatter, and area charts can be created, saved, and loaded
back. Chart data is referenced by cell range, not embedded. See
`examples/chart/`.

```go
ch := &cells_foss.Chart{
    Type:  cells_foss.ChartTypeBar, // ChartTypeLine, ChartTypePie, ChartTypeScatter, ChartTypeArea
    Title: "Revenue by Region",
    Series: []*cells_foss.ChartSeries{
        {Name: "Revenue", Categories: "A2:A5", Values: "B2:B5"},
    },
    Anchor: cells_foss.ChartAnchor{Row: 0, Col: 5}, // 0-based position
    Width:  480,
    Height: 320,
}
ws.AddChart(ch)
```

### Axes and Legend

Cartesian charts (everything but pie) accept axis configuration and a
legend. Both are optional; nil keeps Excel's defaults.

```go
ch.XAxis = &cells_foss.ChartAxis{Title: "Region"}
ch.YAxis = &cells_foss.ChartAxis{Title: "USD", Hidden: true} // Hidden suppresses the axis
ch.Legend = &cells_foss.ChartLegend{
    Position: cells_foss.LegendBottom, // LegendLeft, LegendTop, LegendRight
}
```

For scatter charts, each series' `Categories` range supplies the numeric X
values and `Values` the Y values; both axes are value axes.

Advanced chart types (3D, combo, etc.), secondary axes, data labels, and
gridline styling are not supported.

---

## Pivot Tables

Pivot tables with row fields, column fields, and aggregated data fields can
be created, saved, and loaded back. See `examples/pivot_table/`.

```go
pt := &cells_foss.PivotTable{
    Name:      "SalesPivot",
    Ref:       "E1",          // where the pivot is placed
    SourceRef: "A1:C7",       // source data incl. header row
    RowFields: []string{"Region"},
    ColFields: []string{"Product"},
    DataFields: []*cells_foss.PivotDataField{
        {Name: "Revenue", Aggregation: cells_foss.PivotAggregationSum},
    },
}
ws.AddPivotTable(pt)
```

Aggregations: `PivotAggregationSum`, `PivotAggregationCount`,
`PivotAggregationAverage`, `PivotAggregationMax`, `PivotAggregationMin`.
Calculated fields, grouping, and custom aggregations are not supported.

---

## Conditional Formatting

Rules highlight cells by value, text, or expression, and gradient rules
(color scales, data bars, icon sets) plus top-N rules are supported. A
rule's `Style` is saved as a DXF (differential format) in styles.xml. See
`examples/conditional_format/`.

```go
cf := &cells_foss.ConditionalFormatting{
    Ref: "B2:B6",
    Rules: []*cells_foss.ConditionalFormattingRule{
        {
            Type:     cells_foss.CondTypeCellIs,
            Operator: cells_foss.CondOpGreaterThan,
            Formula:  "80",
            Priority: 1,
            Style: &cells_foss.Style{
                Font: &cells_foss.Font{Bold: true, Color: "FF006100"},
                Fill: &cells_foss.Fill{Type: "solid", Color: "FFC6EFCE"},
            },
        },
    },
}
ws.AddConditionalFormatting(cf)
```

Rule types: `CondTypeCellIs`, `CondTypeContainsText`, `CondTypeBeginsWith`,
`CondTypeEndsWith`, `CondTypeExpression`, `CondTypeColorScale`,
`CondTypeDataBar`, `CondTypeIconSet`, `CondTypeTop10`. Operators for
`cellIs`: `greaterThan`, `lessThan`, `between`, `equal`, `notEqual`,
`greaterThanOrEqual`, `lessThanOrEqual` (as `CondOp*` constants).

### Gradient and Top-N Rules

Color scales, data bars, and icon sets carry their configuration on the
rule; thresholds use `CondThreshold` with a type of `CondValueMin`,
`CondValueMax`, `CondValueNum`, `CondValuePercent`, or `CondValueFormula`
(the last three require `Val`). Colors are ARGB strings.

```go
// Red → yellow → green gradient across the range.
{Type: cells_foss.CondTypeColorScale, Priority: 1,
    ColorScale: &cells_foss.ColorScale{
        Thresholds: []cells_foss.CondThreshold{
            {Type: cells_foss.CondValueMin},
            {Type: cells_foss.CondValuePercent, Val: "50"},
            {Type: cells_foss.CondValueMax},
        },
        Colors: []string{"FFF8696B", "FFFFEB84", "FF63BE7B"},
    }}

// In-cell bars; HideValue draws the bar without the number.
{Type: cells_foss.CondTypeDataBar, Priority: 2,
    DataBar: &cells_foss.DataBar{
        Min:   cells_foss.CondThreshold{Type: cells_foss.CondValueNum, Val: "0"},
        Max:   cells_foss.CondThreshold{Type: cells_foss.CondValueNum, Val: "100"},
        Color: "FF638EC6",
    }}

// Three arrows split by percent rank (IconSet3Arrows, IconSet5Quarters, …).
{Type: cells_foss.CondTypeIconSet, Priority: 3,
    IconSet: &cells_foss.IconSet{
        IconSet: cells_foss.IconSet3Arrows,
        Thresholds: []cells_foss.CondThreshold{
            {Type: cells_foss.CondValuePercent, Val: "0"},
            {Type: cells_foss.CondValuePercent, Val: "33"},
            {Type: cells_foss.CondValuePercent, Val: "67"},
        },
    }}

// Gold font on the top 2 cells (Bottom: true takes the lowest;
// Percent: true reads Rank as a percentage).
{Type: cells_foss.CondTypeTop10, Rank: 2, Priority: 4,
    Style: &cells_foss.Style{Font: &cells_foss.Font{Bold: true, Color: "FF9C6500"}}}
```

The icon set name's leading digit is how many thresholds are required;
`AddConditionalFormatting` rejects mismatched counts, unknown icon names,
and empty colors. Excel 2010+ extensions (gradient stops for data bar
fills, negative-value formats) are not supported.

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

Exports are written with a UTF-8 BOM so Excel detects the encoding, and with
CRLF line endings. The grid is trimmed to the bounding box of the populated
cells; extremely sparse sheets (e.g. data only in `A1` and `XFD1048576`)
return an error instead of attempting a huge allocation.

### Importing from CSV

```go
// Import as a new worksheet.
wb.ImportFromCSV("input.csv", "Data", ',')

// Populate from existing data.
ws.FromCSV(csvRows, ',')
```

Imports strip a leading UTF-8 BOM if present. Values are typed the way Excel
types them: `42` and `3.14` become `float64`, `TRUE`/`FALSE` (any case)
become `bool`. Conversion is conservative — text whose canonical form does
not round-trip exactly (`007`, `1.50`, `1e5`) stays a string.

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

## VBA Macros

VBA support is **storage-only**: the library preserves the raw
`vbaProject.bin` bytes of macro-enabled workbooks across load/save but cannot
create, edit, or execute VBA code. Use the `.xlsm` extension for
macro-enabled workbooks. See `examples/macro/`.

```go
// Attach macro bytes taken from an existing .xlsm workbook.
data, _ := os.ReadFile("vbaProject.bin")
wb.SetVBAProject(data)
wb.Save("workbook.xlsm")

// Inspect and remove.
if wb.HasVBAProject() {
    fmt.Println(len(wb.GetVBAProject().Data), "bytes")
}
wb.ClearVBAProject()
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

// Verify a password without decrypting (constant-time; false when unset).
ok := wb.CheckPassword("mypassword")

// Remove password protection.
wb.SetPassword("")
wb.Save("unprotected.xlsx")
```

Encryption uses **ECMA-376 Agile Encryption** (SHA-512 iterative key
derivation + AES-256-CBC, 4096-byte segments, HMAC-SHA512 data integrity)
inside a standard **OLE/CFB container**, interoperable with Microsoft Excel
and other spec-compliant readers. Files written by older library versions
(legacy scheme and ECRX containers) can still be loaded.

> `VerifyPassword` is deprecated: it returns `true` for any input when no
> password is set. Use `CheckPassword` instead. See `examples/encryption/`.

---

## API Reference

### Workbook

| Method | Description |
|---|---|
| `NewWorkbook() *Workbook` | Create an empty workbook |
| `LoadWorkbook(path) (*Workbook, error)` | Load from an .xlsx file |
| `AddWorksheet(name) (*Worksheet, error)` | Append a named worksheet (validates Excel naming rules) |
| `LoadWithPassword(path, pw) (*Workbook, error)` | Load from an encrypted file |
| `LoadFrom(r io.Reader) (*Workbook, error)` | Load from a stream |
| `LoadFromWithPassword(r io.Reader, pw) (*Workbook, error)` | Load a possibly encrypted stream |
| `Save(path) error` | Save (encrypts if a password is set) |
| `SaveTo(w io.Writer) error` | Save to a stream (same bytes as Save) |
| `SetPassword(pw) error` | Set, change, or remove (empty) the password |
| `CheckPassword(pw) bool` | Constant-time password check; false when unset |
| `SetVBAProject(data) error` | Attach raw vbaProject.bin bytes |
| `GetVBAProject() *VBAProject` | Get the attached VBA project (or nil) |
| `HasVBAProject() bool` | Whether a VBA project is attached |
| `ClearVBAProject()` | Detach the VBA project |
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
| `AsString() string` | Value as a string (type-safe) |
| `AsFloat64() (float64, bool)` | Value as float64; ok=false when not numeric |
| `AsInt() (int, bool)` | Value as int; ok=false when not integral |
| `Value interface{}` | Cell value (deprecated — use the As* accessors) |
| `Ref string` | A1 reference address |

### Number Formats & Dates

| Symbol | Description |
|---|---|
| `Style.NumberFormat string` | Number format code ("" or "General" = General) |
| `NumFmtGeneral` ... `NumFmtText` | Constants for common format codes |
| `TimeToSerial(t) (float64, error)` | Go time to Excel serial (1900 date system) |
| `SerialToTime(serial) (time.Time, error)` | Excel serial to UTC time |

### Formulas

| Symbol | Description |
|---|---|
| `CalculateFormula(formula, ws) (interface{}, error)` | Evaluate a formula (functions, arithmetic, sheet-qualified refs) |
| `ErrorValue` type | An Excel error value; storable in cells, saved as `t="e"` |
| `ErrDiv0` ... `ErrValue` | The seven Excel error value constants |
| `IsErrorValue(s) bool` | Reports whether s is one of the seven error codes |
| `SheetNotFoundError` | Returned for a reference to a sheet that does not exist |

### Worksheet

| Method | Description |
|---|---|
| `Cells() *Cells` | Get the cell collection |
| `AddTable(ref) *Table` | Create a table |
| `GetTable(name) *Table` | Look up a table by name |
| `AddDataValidation(ref, *DataValidation) error` | Add a data validation rule |
| `RemoveDataValidation(ref) error` | Remove a data validation rule |
| `AddPicture(*Picture) error` | Embed a picture |
| `AddChart(*Chart) error` | Embed a bar/line/pie/scatter/area chart |
| `RemoveChart(index) error` | Remove a chart by index |
| `AddPivotTable(*PivotTable) error` | Add a pivot table |
| `RemovePivotTable(name) error` | Remove a pivot table by name |
| `AddConditionalFormatting(*ConditionalFormatting) error` | Add conditional formatting rules |
| `RemoveConditionalFormatting(ref) error` | Remove rules by range |
| `InsertRow(row) error` | Insert a blank row, shifting content down (formulas not rewritten) |
| `DeleteRow(row) error` | Delete a row, shifting content up (formulas not rewritten) |
| `AddMergedCells(ref) error` | Merge a range of cells (refuses overlaps) |
| `RemoveMergedCells(ref) error` | Un-merge a previously merged range |
| `SetColumnWidth(col, width) error` | Set a column's width in character units |
| `SetRowHeight(row, height) error` | Set a row's height in points |
| `FreezePanes(row, col) error` | Freeze the top rows / leftmost columns |
| `SetAutoFilter(ref) error` | Enable filter buttons over a range ("" clears) |
| `AddHyperlink(ref, target) error` | Attach an external link to a cell |
| `AddInternalHyperlink(ref, location) error` | Attach an in-workbook link to a cell |
| `RemoveHyperlink(ref) error` | Detach the link from a cell |
| `ToCSV(delim) ([][]string, error)` | Export as a 2D string slice |
| `FromCSV(data, delim) error` | Populate from a 2D string slice |
