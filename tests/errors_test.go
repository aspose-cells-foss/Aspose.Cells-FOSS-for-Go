package cells_foss_test

import (
	"errors"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestCellNotFoundError(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	_, err := ws.Cells().Get("X99")
	if err == nil {
		t.Fatal("expected error for non-existent cell")
	}

	var cnfErr *cells_foss.CellNotFoundError
	if !errors.As(err, &cnfErr) {
		t.Errorf("expected CellNotFoundError, got %T: %v", err, err)
	}
	if cnfErr.Ref != "X99" {
		t.Errorf("Ref = %q, want X99", cnfErr.Ref)
	}
	if cnfErr.Error() != `cell "X99" not found` {
		t.Errorf("Error() = %q", cnfErr.Error())
	}
}

func TestInvalidRefError(t *testing.T) {
	err := &cells_foss.InvalidRefError{Ref: "!!!"}
	if err.Error() != `invalid cell reference "!!!"` {
		t.Errorf("Error() = %q", err.Error())
	}
}

func TestSheetNotFoundError(t *testing.T) {
	err := &cells_foss.SheetNotFoundError{Name: "NonExistent"}
	if err.Error() != `sheet "NonExistent" not found` {
		t.Errorf("Error() = %q", err.Error())
	}
}

func TestFormulaError(t *testing.T) {
	err := &cells_foss.FormulaError{
		Formula: "=SUM(",
		Reason:  "unterminated parenthesis",
	}
	expected := `formula error in "=SUM(": unterminated parenthesis`
	if err.Error() != expected {
		t.Errorf("Error() = %q, want %q", err.Error(), expected)
	}
}

func TestErrorTypes(t *testing.T) {
	// Verify that error types implement the error interface.
	var _ error = &cells_foss.CellNotFoundError{}
	var _ error = &cells_foss.InvalidRefError{}
	var _ error = &cells_foss.SheetNotFoundError{}
	var _ error = &cells_foss.FormulaError{}
}
