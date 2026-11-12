package cells_foss_test

import (
	"errors"
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// TestFormula_ErrorValues covers Excel error-value semantics: bad results
// are values that propagate, not evaluation failures.
func TestFormula_ErrorValues(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("C1", "hello")
	ws.Cells().Set("B1", float64(1))
	ws.Cells().Set("B2", cells_foss.ErrDiv0) // an error stored in a cell
	ws.Cells().Set("B3", float64(3))

	cases := []struct {
		formula string
		want    cells_foss.ErrorValue
	}{
		{"1/0", cells_foss.ErrDiv0},
		{"1/0+5", cells_foss.ErrDiv0},        // propagates through +
		{"5*(1/0)", cells_foss.ErrDiv0},      // through *
		{"-(1/0)", cells_foss.ErrDiv0},       // through unary minus
		{`"x"&(1/0)`, cells_foss.ErrDiv0},    // through &
		{"C1+1", cells_foss.ErrValue},        // arithmetic on text
		{"SQRT(-4)", cells_foss.ErrNum},      // invalid numeric result
		{"POWER(-8,0.5)", cells_foss.ErrNum}, // NaN result
		{"(-8)^0.5", cells_foss.ErrNum},      // via the ^ operator
		{"B2", cells_foss.ErrDiv0},           // error cell reference
		{"B2+1", cells_foss.ErrDiv0},         // error cell propagates
		{"SUM(B1:B3)", cells_foss.ErrDiv0},   // errors propagate through aggregation
	}
	for _, tc := range cases {
		got, err := cells_foss.CalculateFormula(tc.formula, ws)
		if err != nil {
			t.Errorf("%s: unexpected Go error: %v", tc.formula, err)
			continue
		}
		if got != interface{}(tc.want) {
			t.Errorf("%s = %v (%T), want %v", tc.formula, got, got, tc.want)
		}
	}

	// A missing sheet is still a Go-level failure, typed for errors.As.
	_, err := cells_foss.CalculateFormula("Nope!A1", ws)
	var snf *cells_foss.SheetNotFoundError
	if !errors.As(err, &snf) {
		t.Errorf("missing sheet: got %v, want *SheetNotFoundError", err)
	}

	if !cells_foss.IsErrorValue("#DIV/0!") || cells_foss.IsErrorValue("#DIV/0") {
		t.Error("IsErrorValue misclassified")
	}
}

// TestFormula_ErrorValuePersistence covers the t="e" save/load round trip.
func TestFormula_ErrorValuePersistence(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", cells_foss.ErrNA)
	ws.Cells().Set("A2", "text")

	path := filepath.Join(t.TempDir(), "err.xlsx")
	if err := wb.Save(path); err != nil {
		t.Fatalf("Save: %v", err)
	}
	raw, err := cells_foss.ReadTestZipEntry(path, "xl/worksheets/sheet1.xml")
	if err != nil {
		t.Fatalf("read sheet: %v", err)
	}
	if !strings.Contains(string(raw), `<c r="A1" t="e">`) {
		t.Errorf("sheet XML missing t=\"e\" cell:\n%s", raw)
	}
	if !strings.Contains(string(raw), `<v>#N/A</v>`) {
		t.Errorf("sheet XML missing error code:\n%s", raw)
	}
	// Error codes must not leak into the shared-string table.
	sst, err := cells_foss.ReadTestZipEntry(path, "xl/sharedStrings.xml")
	if err == nil && strings.Contains(string(sst), "#N/A") {
		t.Errorf("error value stored as shared string:\n%s", sst)
	}

	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("Load: %v", err)
	}
	c, err := wb2.Worksheets[0].Cells().Get("A1")
	if err != nil {
		t.Fatalf("Get A1: %v", err)
	}
	if ev, ok := c.Value.(cells_foss.ErrorValue); !ok || ev != cells_foss.ErrNA {
		t.Errorf("reloaded A1 = %v (%T), want ErrorValue #N/A", c.Value, c.Value)
	}
	if c.AsString() != "#N/A" {
		t.Errorf("AsString = %q, want #N/A", c.AsString())
	}
	if _, ok := c.AsFloat64(); ok {
		t.Error("AsFloat64 on an error value should report false")
	}
}

// TestFormula_CrossSheetRefs covers sheet-qualified references in formulas.
func TestFormula_CrossSheetRefs(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	data, err := wb.AddWorksheet("Data")
	if err != nil {
		t.Fatalf("AddWorksheet: %v", err)
	}
	spaced, err := wb.AddWorksheet("My Data")
	if err != nil {
		t.Fatalf("AddWorksheet: %v", err)
	}

	ws.Cells().Set("A1", float64(99)) // same ref, different value than Data!A1
	data.Cells().Set("A1", float64(5))
	data.Cells().Set("A2", float64(7))
	spaced.Cells().Set("B2", float64(4))

	cases := []struct {
		formula string
		want    interface{}
	}{
		{"Data!A1", float64(5)},          // resolves against Data, not the calling sheet
		{"Data!A1+1", float64(6)},        // in arithmetic
		{"SUM(Data!A1:A2)", float64(12)}, // sheet-qualified range
		{"data!A1*2", float64(10)},       // sheet names are case-insensitive
		{"'My Data'!B2", float64(4)},     // quoted name with a space
		{"'My Data'!B2^2", float64(16)},
		{"SUM(Data!A1:A2)+'My Data'!B2", float64(16)},
	}
	for _, tc := range cases {
		got, err := cells_foss.CalculateFormula(tc.formula, ws)
		if err != nil {
			t.Errorf("%s: unexpected error: %v", tc.formula, err)
			continue
		}
		if got != tc.want {
			t.Errorf("%s = %v (%T), want %v", tc.formula, got, got, tc.want)
		}
	}

	// Cross-sheet refs also work on a loaded workbook.
	path := filepath.Join(t.TempDir(), "cross.xlsx")
	if err := wb.Save(path); err != nil {
		t.Fatalf("Save: %v", err)
	}
	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("Load: %v", err)
	}
	got, err := cells_foss.CalculateFormula("SUM(Data!A1:A2)", wb2.Worksheets[0])
	if err != nil || got != float64(12) {
		t.Errorf("loaded cross-sheet SUM = %v (%v), want 12", got, err)
	}
	got, err = cells_foss.CalculateFormula("'My Data'!B2", wb2.Worksheets[0])
	if err != nil || got != float64(4) {
		t.Errorf("loaded quoted ref = %v (%v), want 4", got, err)
	}
}
