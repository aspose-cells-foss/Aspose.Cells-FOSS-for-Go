package cells_foss

import "fmt"

// CellNotFoundError is returned when attempting to access a cell that does
// not exist in the worksheet.
type CellNotFoundError struct {
	Ref string
}

func (e *CellNotFoundError) Error() string {
	return fmt.Sprintf("cell %q not found", e.Ref)
}

// InvalidRefError is returned when a cell reference string is malformed or
// cannot be parsed.
type InvalidRefError struct {
	Ref string
}

func (e *InvalidRefError) Error() string {
	return fmt.Sprintf("invalid cell reference %q", e.Ref)
}

// SheetNotFoundError is returned when attempting to access a worksheet by
// name that does not exist in the workbook.
type SheetNotFoundError struct {
	Name string
}

func (e *SheetNotFoundError) Error() string {
	return fmt.Sprintf("sheet %q not found", e.Name)
}

// FormulaError is returned when a formula cannot be parsed or evaluated.
type FormulaError struct {
	Formula string
	Reason  string
}

func (e *FormulaError) Error() string {
	return fmt.Sprintf("formula error in %q: %s", e.Formula, e.Reason)
}
