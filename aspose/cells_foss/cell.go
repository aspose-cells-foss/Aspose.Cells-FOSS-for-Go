// Package cells_foss provides a pure-Go library for reading, creating,
// and writing Excel (.xlsx) workbooks compatible with ECMA-376 Office Open XML.
package cells_foss

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

// Cell represents a single cell in a worksheet grid.
// Its Ref field holds the A1-style address (e.g. "A1", "B2") that uniquely
// identifies the cell within its parent worksheet.
type Cell struct {
	XMLName xml.Name `xml:"c"`
	Ref     string   `xml:"r,attr"`
	StyleID int      `xml:"s,attr,omitempty"`
	// Value holds the cell's data. After loading, numeric cells contain
	// float64, boolean cells contain bool, and string cells contain string.
	//
	// Deprecated: Use AsFloat64, AsInt, or AsString for type-safe access.
	// The concrete type of Value may vary (e.g. float64 for numeric cells
	// loaded from .xlsx), so direct type assertions like Value.(string)
	// may panic.
	Value   interface{} `xml:"v,omitempty"`
	Formula string      `xml:"f,omitempty"`

	// cells is a back-reference to the owning Cells collection, used by
	// SetStyle / GetStyle to access the Workbook-level style registry.
	cells *Cells
}

// AsFloat64 returns the cell value as float64.
// It returns (0, false) if the cell is empty or the value cannot be
// converted to a number.
func (c *Cell) AsFloat64() (float64, bool) {
	switch v := c.Value.(type) {
	case nil:
		return 0, false
	case float64:
		return v, true
	case float32:
		return float64(v), true
	case int:
		return float64(v), true
	case int64:
		return float64(v), true
	case string:
		f, err := strconv.ParseFloat(v, 64)
		if err != nil {
			return 0, false
		}
		return f, true
	case bool:
		if v {
			return 1, true
		}
		return 0, true
	default:
		f, err := strconv.ParseFloat(fmt.Sprint(v), 64)
		if err != nil {
			return 0, false
		}
		return f, true
	}
}

// AsInt returns the cell value as int.
// It returns (0, false) if the cell is empty or the value cannot be
// converted to an integer.
func (c *Cell) AsInt() (int, bool) {
	f, ok := c.AsFloat64()
	if !ok {
		return 0, false
	}
	return int(f), true
}

// AsString returns the cell value as a string.
// It returns "" for nil (empty) cells and formats numeric/boolean values
// using the same conventions as CellToString.
func (c *Cell) AsString() string {
	return CellToString(c.Value)
}

// SetStyle assigns the given Style to this cell.  When the owning Workbook is
// available the style is automatically registered (or deduplicated) and the
// cell's StyleID is updated.  If the cell has no parent Workbook (e.g. it was
// created outside of a workbook context) the call returns an error.
func (c *Cell) SetStyle(style *Style) error {
	if c.cells == nil || c.cells.wb == nil {
		return fmt.Errorf("cells_foss: cannot set style on a cell that is not part of a Workbook")
	}
	c.StyleID = c.cells.wb.registerStyle(style)
	return nil
}

// GetStyle returns the Style currently applied to this cell, or nil when the
// cell has no parent Workbook or the StyleID cannot be resolved.
func (c *Cell) GetStyle() *Style {
	if c.cells == nil || c.cells.wb == nil {
		return nil
	}
	return c.cells.wb.getStyle(c.StyleID)
}

// SetFormula stores a formula expression in this cell and marks the owning
// Workbook as modified.  The formula is written as-is into the <f> element
// during save; no parsing or validation is performed at write time.
func (c *Cell) SetFormula(formula string) {
	c.Formula = formula
	if c.cells != nil && c.cells.wb != nil {
		c.cells.wb.Modified = true
	}
}

// GetFormula returns the formula expression stored in this cell, or an empty
// string when the cell contains no formula.
func (c *Cell) GetFormula() string {
	return c.Formula
}
