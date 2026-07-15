// Package cells_foss provides a pure-Go library for reading, creating,
// and writing Excel (.xlsx) workbooks compatible with ECMA-376 Office Open XML.
package cells_foss

import (
	"encoding/xml"
	"fmt"
)

// Cell represents a single cell in a worksheet grid.
// Its Ref field holds the A1-style address (e.g. "A1", "B2") that uniquely
// identifies the cell within its parent worksheet.
type Cell struct {
	XMLName xml.Name    `xml:"c"`
	Ref     string      `xml:"r,attr"`
	StyleID int         `xml:"s,attr,omitempty"`
	Value   interface{} `xml:"v,omitempty"`
	Formula string      `xml:"f,omitempty"`

	// cells is a back-reference to the owning Cells collection, used by
	// SetStyle / GetStyle to access the Workbook-level style registry.
	cells *Cells
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
