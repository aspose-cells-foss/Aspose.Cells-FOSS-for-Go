package cells_foss_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
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

func TestCSV_TypeInference(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	data := [][]string{{"42", "007", "3.14", "TRUE", "true", "1e5", "1.50", "hello", "-7", ""}}
	if err := ws.FromCSV(data, ','); err != nil {
		t.Fatalf("FromCSV: %v", err)
	}

	get := func(ref string) interface{} {
		c, err := ws.Cells().Get(ref)
		if err != nil {
			t.Fatalf("Get(%s): %v", ref, err)
		}
		return c.Value
	}

	if v := get("A1"); v != float64(42) {
		t.Errorf("A1 = %#v, want float64(42)", v)
	}
	if v := get("B1"); v != "007" {
		t.Errorf("B1 = %#v, want string \"007\" (leading zeros preserved)", v)
	}
	if v := get("C1"); v != float64(3.14) {
		t.Errorf("C1 = %#v, want float64(3.14)", v)
	}
	if v := get("D1"); v != true {
		t.Errorf("D1 = %#v, want bool true", v)
	}
	if v := get("E1"); v != true {
		t.Errorf("E1 = %#v, want bool true (case-insensitive)", v)
	}
	if v := get("F1"); v != "1e5" {
		t.Errorf("F1 = %#v, want string \"1e5\" (exponent spelling preserved)", v)
	}
	if v := get("G1"); v != "1.50" {
		t.Errorf("G1 = %#v, want string \"1.50\" (trailing zero preserved)", v)
	}
	if v := get("H1"); v != "hello" {
		t.Errorf("H1 = %#v, want string \"hello\"", v)
	}
	if v := get("I1"); v != float64(-7) {
		t.Errorf("I1 = %#v, want float64(-7)", v)
	}
}

func TestCSV_BOMRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", "Name")
	c.Set("B1", "Stadt")
	c.Set("A2", "Müller")

	dir := t.TempDir()
	p := filepath.Join(dir, "bom.csv")
	if err := wb.ExportToCSV(0, p, ','); err != nil {
		t.Fatalf("ExportToCSV: %v", err)
	}

	raw, err := os.ReadFile(p)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}
	if len(raw) < 3 || raw[0] != 0xEF || raw[1] != 0xBB || raw[2] != 0xBF {
		t.Errorf("export should start with UTF-8 BOM, got % X", raw[:min(3, len(raw))])
	}
	if !strings.Contains(string(raw), "Müller") {
		t.Error("export should contain UTF-8 text")
	}

	// Import must strip the BOM from the first field.
	wb2 := cells_foss.NewWorkbook()
	if err := wb2.ImportFromCSV(p, "Import", ','); err != nil {
		t.Fatalf("ImportFromCSV: %v", err)
	}
	ws := wb2.Worksheets[len(wb2.Worksheets)-1]
	a1, err := ws.Cells().Get("A1")
	if err != nil {
		t.Fatalf("Get A1: %v", err)
	}
	if a1.Value != "Name" {
		t.Errorf("A1 = %#v, want \"Name\" (BOM must be stripped)", a1.Value)
	}
	a2, _ := ws.Cells().Get("A2")
	if a2.Value != "Müller" {
		t.Errorf("A2 = %#v, want \"Müller\"", a2.Value)
	}

	// Importing a file without a BOM must also work.
	p2 := filepath.Join(dir, "nobom.csv")
	if err := os.WriteFile(p2, []byte("X,Y\n1,2\n"), 0o644); err != nil {
		t.Fatalf("WriteFile: %v", err)
	}
	wb3 := cells_foss.NewWorkbook()
	if err := wb3.ImportFromCSV(p2, "NoBom", ','); err != nil {
		t.Fatalf("ImportFromCSV without BOM: %v", err)
	}
	ws3 := wb3.Worksheets[len(wb3.Worksheets)-1]
	if x1, err := ws3.Cells().Get("A1"); err != nil || x1.Value != "X" {
		t.Errorf("A1 = %v (%v), want \"X\"", x1.Value, err)
	}
	if a2, err := ws3.Cells().Get("B2"); err != nil || a2.Value != float64(2) {
		t.Errorf("B2 = %v (%v), want float64(2)", a2.Value, err)
	}
}

func TestCSV_ToCSVTrimsColumns(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("B2", "b")
	ws.Cells().Set("D5", "d")

	rows, err := ws.ToCSV(',')
	if err != nil {
		t.Fatalf("ToCSV: %v", err)
	}
	if len(rows) != 4 {
		t.Fatalf("got %d rows, want 4 (rows 2..5)", len(rows))
	}
	for _, r := range rows {
		if len(r) != 3 {
			t.Fatalf("got %d columns, want 3 (cols B..D)", len(r))
		}
	}
	if rows[0][0] != "b" || rows[3][2] != "d" {
		t.Errorf("grid = %v, want b at [0][0] and d at [3][2]", rows)
	}
}

func TestCSV_SparseSheetGuard(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "corner")
	// A far corner cell: the bounding box spans the whole sheet area.
	if err := ws.Cells().Set("XFD1048576", "far"); err != nil {
		t.Skipf("library rejects extreme reference: %v", err)
	}

	_, err := ws.ToCSV(',')
	if err == nil {
		t.Error("ToCSV on extremely sparse sheet should return an error, not allocate the full grid")
	} else if !strings.Contains(err.Error(), "sparse") {
		t.Errorf("unexpected error: %v", err)
	}
}
