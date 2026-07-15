package cells_foss

import "fmt"

// Table represents a structured range of data (a "table" in Excel terminology)
// with optional header row and built-in auto-filter.
type Table struct {
	// Name is the internal table name (e.g. "Table1").  It must be unique
	// within the worksheet and is automatically assigned by AddTable.
	Name string

	// Range is the A1-style reference that defines the table area,
	// e.g. "A1:D10".  The first row is treated as a header when
	// HasHeaderRow is true.
	Range string

	// HasHeaderRow controls whether the first row of the table is formatted
	// as a header row with auto-filter dropdowns.
	HasHeaderRow bool

	// StyleName is the name of the built-in table style to apply
	// (e.g. "TableStyleMedium9").  Excel ships with many predefined styles;
	// an empty string means the default style.
	StyleName string
}

// NewTable creates a Table value initialised with sensible defaults.
// The caller should use Worksheet.AddTable() to attach the table to a sheet.
func NewTable(name, rangeRef string) *Table {
	return &Table{
		Name:         name,
		Range:        rangeRef,
		HasHeaderRow: true,
		StyleName:    "TableStyleMedium9",
	}
}

// AddTable creates a new Table covering the given range, assigns it a unique
// auto-generated name ("Table1", "Table2", …), appends it to the worksheet,
// and returns it for further configuration.
func (ws *Worksheet) AddTable(rangeRef string) *Table {
	name := fmt.Sprintf("Table%d", len(ws.Tables)+1)
	t := NewTable(name, rangeRef)
	ws.Tables = append(ws.Tables, t)
	if ws.cells != nil && ws.cells.wb != nil {
		ws.cells.wb.Modified = true
	}
	return t
}

// GetTable returns the Table with the given name, or nil when no match is found.
func (ws *Worksheet) GetTable(name string) *Table {
	for _, t := range ws.Tables {
		if t.Name == name {
			return t
		}
	}
	return nil
}
