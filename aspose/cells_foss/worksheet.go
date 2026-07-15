package cells_foss

// Worksheet represents a single sheet within a workbook. It holds the sheet
// metadata (name and positional index) together with the collection of cells
// and data validations that make up the sheet content.
type Worksheet struct {
	Name  string
	Index int

	// cells holds the cell collection for this worksheet. Use the Cells()
	// accessor to obtain the *Cells value.
	cells *Cells

	// DataValidations holds the data-validation rules applied to this sheet.
	DataValidations []*DataValidation

	// Tables holds the structured tables defined on this worksheet.
	Tables []*Table

	// Pictures holds the images embedded on this worksheet.
	Pictures []*Picture
}

// Cells returns the Cells collection for this worksheet, enabling cell-level
// read and write operations via A1-style references.
func (ws *Worksheet) Cells() *Cells {
	return ws.cells
}
