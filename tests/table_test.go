package cells_foss_test

import (
	"bytes"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose/cells_foss/aspose/cells_foss"
)

func TestTable_AddGetRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	c := ws.Cells()
	c.Set("A1", "Name")
	c.Set("B1", "Score")
	c.Set("A2", "Alice")
	c.Set("B2", float64(95))

	tbl := ws.AddTable("A1:B2")
	tbl.HasHeaderRow = true

	if tbl.Name != "Table1" {
		t.Errorf("name = %q", tbl.Name)
	}
	if ws.GetTable("Table1") == nil {
		t.Error("GetTable returned nil")
	}
	if ws.GetTable("Nonexistent") != nil {
		t.Error("GetTable should return nil")
	}

	dir := t.TempDir()
	p := filepath.Join(dir, "table.xlsx")
	wb.Save(p)

	// Verify table files in ZIP.
	if raw := readZipEntry(t, p, "xl/tables/table1.xml"); !bytes.Contains(raw, []byte("Table1")) {
		t.Error("table XML missing")
	}
	if raw := readZipEntry(t, p, "xl/worksheets/_rels/sheet1.xml.rels"); !bytes.Contains(raw, []byte("table1.xml")) {
		t.Error("sheet rels missing table ref")
	}

	// Reload.
	loaded, _ := cells_foss.LoadWorkbook(p)
	lws := loaded.Worksheets[0]
	if len(lws.Tables) != 1 {
		t.Fatalf("tables = %d, want 1", len(lws.Tables))
	}
	if lws.Tables[0].Name != "Table1" {
		t.Errorf("reloaded name = %q", lws.Tables[0].Name)
	}
}

func TestTable_NoHeaderRow(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", float64(1))
	ws.Cells().Set("A2", float64(2))

	tbl := ws.AddTable("A1:A2")
	tbl.HasHeaderRow = false

	dir := t.TempDir()
	p := filepath.Join(dir, "noheader.xlsx")
	wb.Save(p)

	raw := readZipEntry(t, p, "xl/tables/table1.xml")
	if bytes.Contains(raw, []byte(`headerRowCount="1"`)) {
		t.Error("should not have headerRowCount=1")
	}
}

func TestTable_MultipleTables(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "T1")
	ws.Cells().Set("C1", "T2")
	ws.AddTable("A1:A5")
	ws.AddTable("C1:D10")

	dir := t.TempDir()
	p := filepath.Join(dir, "multitable.xlsx")
	wb.Save(p)

	if raw := readZipEntry(t, p, "xl/tables/table1.xml"); len(raw) == 0 {
		t.Error("table1 missing")
	}
	if raw := readZipEntry(t, p, "xl/tables/table2.xml"); len(raw) == 0 {
		t.Error("table2 missing")
	}

	loaded, _ := cells_foss.LoadWorkbook(p)
	if len(loaded.Worksheets[0].Tables) != 2 {
		t.Errorf("reloaded tables = %d", len(loaded.Worksheets[0].Tables))
	}
}

func TestTable_MarksModified(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Modified = false
	wb.Worksheets[0].AddTable("A1:B10")
	if !wb.Modified {
		t.Error("AddTable should mark Modified")
	}
}
