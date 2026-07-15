package cells_foss_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose/cells_foss/aspose/cells_foss"
)

func TestCSV_ExportImport(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()

	c.Set("A1", "Name")
	c.Set("B1", "Age")
	c.Set("A2", "Alice")
	c.Set("B2", float64(30))
	c.Set("A3", "Bob")
	c.Set("B3", float64(25))

	dir := t.TempDir()
	csvPath := filepath.Join(dir, "export.csv")

	// Export.
	if err := wb.ExportToCSV(0, csvPath, ','); err != nil {
		t.Fatalf("ExportToCSV: %v", err)
	}
	raw, _ := os.ReadFile(csvPath)
	if !strings.Contains(string(raw), "Name,Age") {
		t.Errorf("CSV content: %s", string(raw))
	}

	// Import.
	wb2 := cells_foss.NewWorkbook()
	if err := wb2.ImportFromCSV(csvPath, "Imported", ','); err != nil {
		t.Fatalf("ImportFromCSV: %v", err)
	}
	if len(wb2.Worksheets) != 2 {
		t.Fatalf("expected 2 sheets, got %d", len(wb2.Worksheets))
	}

	ws := wb2.Worksheets[1]
	if ws.Name != "Imported" {
		t.Errorf("name = %q", ws.Name)
	}
	c1, _ := ws.Cells().Get("A1")
	if c1.Value != "Name" {
		t.Errorf("A1 = %v", c1.Value)
	}
	c2, _ := ws.Cells().Get("B3")
	if cells_foss.CellToString(c2.Value) != "25" {
		t.Errorf("B3 = %v", c2.Value)
	}
}

func TestCSV_TabDelimiter(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "tabs.csv")
	os.WriteFile(p, []byte("Col1\tCol2\nA\tB\n"), 0644)

	wb := cells_foss.NewWorkbook()
	wb.ImportFromCSV(p, "Tabs", '\t')
	c, _ := wb.Worksheets[1].Cells().Get("A1")
	if c.Value != "Col1" {
		t.Errorf("A1 = %v", c.Value)
	}
}

func TestCSV_ToFromCSV(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("B2", "center")
	ws.Cells().Set("D5", "corner")

	data, _ := ws.ToCSV(',')
	if len(data) < 4 {
		t.Errorf("sparse: expected >= 4 rows, got %d", len(data))
	}

	// FromCSV.
	wb2 := cells_foss.NewWorkbook()
	wb2.Worksheets[0].FromCSV([][]string{{"H1", "H2"}, {"v1", "v2"}}, ',')
	c, _ := wb2.Worksheets[0].Cells().Get("A1")
	if c.Value != "H1" {
		t.Errorf("A1 = %v", c.Value)
	}
}

func TestCSV_Errors(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	if err := wb.ExportToCSV(99, "/tmp/x.csv", ','); err == nil {
		t.Error("bad index should error")
	}
	var nilWB *cells_foss.Workbook
	if err := nilWB.ExportToCSV(0, "/tmp/x.csv", ','); err == nil {
		t.Error("nil workbook should error")
	}
}

func TestCellToString(t *testing.T) {
	tests := []struct {
		val  interface{}
		want string
	}{
		{nil, ""}, {"hello", "hello"}, {true, "TRUE"}, {false, "FALSE"},
		{float64(3.14), "3.14"}, {int(42), "42"},
	}
	for _, tc := range tests {
		got := cells_foss.CellToString(tc.val)
		if got != tc.want {
			t.Errorf("CellToString(%v) = %q, want %q", tc.val, got, tc.want)
		}
	}
}
