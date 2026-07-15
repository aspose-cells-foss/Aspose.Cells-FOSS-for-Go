package cells_foss

import "fmt"

// Cells is a collection of Cell values indexed by A1-style string references
// (e.g. "A1", "B2", "Z100"). It is the primary API for reading and writing
// cell data within a worksheet.
type Cells struct {
	cells map[string]*Cell

	// wb is a back-reference to the owning Workbook, used to mark the
	// workbook as modified when cell data changes.
	wb *Workbook
}

// setParent wires the back-reference to the owning Workbook so that
// mutations (Set, Remove) can automatically flag the workbook as modified.
func (c *Cells) setParent(wb *Workbook) {
	c.wb = wb
}

// Get returns the Cell at the given A1 reference. An error is returned when
// no cell exists at that reference.
func (c *Cells) Get(ref string) (*Cell, error) {
	if c.cells == nil {
		return nil, fmt.Errorf("cells: cell %q not found", ref)
	}
	cell, ok := c.cells[ref]
	if !ok {
		return nil, fmt.Errorf("cells: cell %q not found", ref)
	}
	return cell, nil
}

// Set stores a value at the given A1 reference. If a cell already exists at
// ref its Value is updated; otherwise a new Cell is created with its Ref
// field set to ref.
//
// Every call to Set marks the owning Workbook as modified so that subsequent
// calls to Save will regenerate (rather than reuse cached) XML.
func (c *Cells) Set(ref string, value interface{}) error {
	if c.cells == nil {
		c.cells = make(map[string]*Cell)
	}
	if c.wb != nil {
		c.wb.Modified = true
	}
	if cell, ok := c.cells[ref]; ok {
		cell.Value = value
		return nil
	}
	c.cells[ref] = &Cell{Ref: ref, Value: value, cells: c}
	return nil
}

// Remove deletes the cell at the given A1 reference. An error is returned
// when no cell exists at that reference.
//
// A successful removal marks the owning Workbook as modified.
func (c *Cells) Remove(ref string) error {
	if c.cells == nil {
		return fmt.Errorf("cells: cell %q not found", ref)
	}
	if _, ok := c.cells[ref]; !ok {
		return fmt.Errorf("cells: cell %q not found", ref)
	}
	delete(c.cells, ref)
	if c.wb != nil {
		c.wb.Modified = true
	}
	return nil
}

// All returns the underlying map of all cells keyed by A1 reference.
// Callers may safely range over the returned map.
func (c *Cells) All() map[string]*Cell {
	if c.cells == nil {
		return make(map[string]*Cell)
	}
	return c.cells
}
