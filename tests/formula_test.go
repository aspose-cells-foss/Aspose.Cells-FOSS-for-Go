package cells_foss_test

import (
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestFormula_SetGetAndRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()

	c.Set("A1", float64(10))
	c.Set("A2", float64(20))
	c.Set("A3", float64(30))

	c.Set("A4", nil)
	ca4, _ := c.Get("A4")
	ca4.SetFormula("SUM(A1:A3)")

	if ca4.GetFormula() != "SUM(A1:A3)" {
		t.Errorf("GetFormula = %q", ca4.GetFormula())
	}

	dir := t.TempDir()
	p := filepath.Join(dir, "formula.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	cell, _ := loaded.Worksheets[0].Cells().Get("A4")
	if cell.GetFormula() != "SUM(A1:A3)" {
		t.Errorf("formula after reload = %q", cell.GetFormula())
	}
}

func TestCalculateFormula_AllFunctions(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	c := ws.Cells()

	// Row range.
	c.Set("A1", float64(10))
	c.Set("A2", float64(20))
	c.Set("A3", float64(30))

	tests := []struct {
		formula string
		want    float64
	}{
		{"SUM(A1:A3)", 60},
		{"AVERAGE(A1:A3)", 20},
		{"MAX(A1:A3)", 30},
		{"MIN(A1:A3)", 10},
	}
	for _, tc := range tests {
		r, err := cells_foss.CalculateFormula(tc.formula, ws)
		if err != nil {
			t.Errorf("%s: %v", tc.formula, err)
		} else if r != tc.want {
			t.Errorf("%s = %v, want %v", tc.formula, r, tc.want)
		}
	}

	// Multi-column range (A1=10, A2=20, B1=1, B2=2).
	c.Set("B1", float64(1))
	c.Set("B2", float64(2))
	r, _ := cells_foss.CalculateFormula("SUM(A1:B2)", ws)
	if r != float64(33) {
		t.Errorf("SUM(A1:B2) = %v, want 33", r)
	}

	// Skips non-numeric.
	c.Set("A1", "text")
	r, _ = cells_foss.CalculateFormula("SUM(A1:A3)", ws)
	if r != float64(50) {
		t.Errorf("SUM with text = %v, want 50", r)
	}

	// Boolean values.
	c.Set("A1", true)
	c.Set("A2", false)
	r, _ = cells_foss.CalculateFormula("SUM(A1:A2)", ws)
	if r != float64(1) {
		t.Errorf("SUM bool = %v, want 1", r)
	}

	// Errors.
	if _, err := cells_foss.CalculateFormula("COUNT(A1:A5)", ws); err == nil {
		t.Error("unsupported function should error")
	}
	if _, err := cells_foss.CalculateFormula("", ws); err == nil {
		t.Error("empty formula should error")
	}
}
