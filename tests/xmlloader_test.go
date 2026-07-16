package cells_foss_test

import (
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose/cells_foss/v26/aspose/cells_foss"
)

func TestLoadWorkbook_NumericCells(t *testing.T) {
	dir := t.TempDir()
	testFile := filepath.Join(dir, "numeric.xlsx")

	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>3.14</v></c><c r="C1"><v>-7</v></c></row>
    <row r="2"><c r="A2"><v>Hello</v></c></row>
  </sheetData>
</worksheet>`

	if err := cells_foss.WriteTestXLSX(testFile, sheetXML, ""); err != nil {
		t.Fatalf("WriteTestXLSX: %v", err)
	}

	wb, err := cells_foss.LoadWorkbook(testFile)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	if wb.FilePath != testFile {
		t.Errorf("FilePath = %q, want %q", wb.FilePath, testFile)
	}
	if len(wb.Worksheets) != 1 {
		t.Fatalf("got %d worksheets, want 1", len(wb.Worksheets))
	}
	if len(wb.SourceXML) == 0 {
		t.Error("SourceXML should not be empty after load")
	}

	cells := wb.Worksheets[0].Cells()
	if c, _ := cells.Get("A1"); cells_foss.CellToString(c.Value) != "42" {
		t.Errorf("A1 = %v", c.Value)
	}
	if c, _ := cells.Get("B1"); cells_foss.CellToString(c.Value) != "3.14" {
		t.Errorf("B1 = %v", c.Value)
	}
	if _, err := cells.Get("Z99"); err == nil {
		t.Error("expected error for missing Z99")
	}
}

func TestLoadWorkbook_SharedStrings(t *testing.T) {
	dir := t.TempDir()
	testFile := filepath.Join(dir, "ss.xlsx")

	ssXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Apple</t></si><si><t>Banana</t></si>
</sst>`
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>
</worksheet>`

	cells_foss.WriteTestXLSX(testFile, sheetXML, ssXML)
	wb, _ := cells_foss.LoadWorkbook(testFile)
	cells := wb.Worksheets[0].Cells()

	c, _ := cells.Get("A1")
	if c.Value != "Apple" {
		t.Errorf("A1 = %v", c.Value)
	}
	c, _ = cells.Get("B1")
	if c.Value != "Banana" {
		t.Errorf("B1 = %v", c.Value)
	}
}

func TestLoadWorkbook_BooleanAndEmpty(t *testing.T) {
	dir := t.TempDir()
	testFile := filepath.Join(dir, "bool.xlsx")
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c><c r="C1"></c></row></sheetData>
</worksheet>`

	cells_foss.WriteTestXLSX(testFile, sheetXML, "")
	wb, _ := cells_foss.LoadWorkbook(testFile)
	cells := wb.Worksheets[0].Cells()

	c, _ := cells.Get("A1")
	if c.Value != true {
		t.Errorf("A1 = %v, want true", c.Value)
	}
	c, _ = cells.Get("B1")
	if c.Value != false {
		t.Errorf("B1 = %v, want false", c.Value)
	}
	c, _ = cells.Get("C1")
	if c.Value != nil {
		t.Errorf("C1 = %v, want nil", c.Value)
	}
}

func TestLoadWorkbook_ErrorCases(t *testing.T) {
	if _, err := cells_foss.LoadWorkbook("nonexistent.xlsx"); err == nil {
		t.Error("expected error")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "bad.xlsx")
	if err := writeStringFile(p, "not a zip"); err != nil {
		t.Fatal(err)
	}
	if _, err := cells_foss.LoadWorkbook(p); err == nil {
		t.Error("expected error for invalid zip")
	}
}

func TestLoadWorkbook_AllMethod(t *testing.T) {
	dir := t.TempDir()
	testFile := filepath.Join(dir, "all.xlsx")
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row></sheetData>
</worksheet>`

	cells_foss.WriteTestXLSX(testFile, sheetXML, "")
	wb, _ := cells_foss.LoadWorkbook(testFile)
	all := wb.Worksheets[0].Cells().All()
	if len(all) != 3 {
		t.Errorf("All() = %d, want 3", len(all))
	}
}
