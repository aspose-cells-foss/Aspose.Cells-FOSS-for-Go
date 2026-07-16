package cells_foss_test

import (
	"bytes"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose/cells_foss/v26/aspose/cells_foss"
)

func TestSaveNewWorkbook_RoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", "Name")
	c.Set("B1", "Score")
	c.Set("A2", "Alice")
	c.Set("B2", float64(95.5))
	c.Set("A3", "Bob")
	c.Set("B3", float64(87))

	dir := t.TempDir()
	p := filepath.Join(dir, "rt.xlsx")
	if err := wb.Save(p); err != nil {
		t.Fatalf("Save: %v", err)
	}

	loaded, _ := cells_foss.LoadWorkbook(p)
	lc := loaded.Worksheets[0].Cells()
	checks := map[string]string{"A1": "Name", "B1": "Score", "A2": "Alice", "B2": "95.5", "A3": "Bob", "B3": "87"}
	for ref, want := range checks {
		c, _ := lc.Get(ref)
		if got := cells_foss.CellToString(c.Value); got != want {
			t.Errorf("%s = %q, want %q", ref, got, want)
		}
	}
}

func TestSaveUnmodified_RawSheetBytesIdentical(t *testing.T) {
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>Hello</v></c></row></sheetData>
</worksheet>`

	inPath := makeNumericXLSX(t, sheetXML)
	originalSheet := readRawSheetXML(t, inPath)

	wb, _ := cells_foss.LoadWorkbook(inPath)
	dir := t.TempDir()
	outPath := filepath.Join(dir, "unmodified.xlsx")
	wb.Save(outPath)

	savedSheet := readRawSheetXML(t, outPath)
	if !bytes.Equal(originalSheet, savedSheet) {
		t.Errorf("unmodified sheet XML changed:\noriginal: %s\nsaved:   %s", originalSheet, savedSheet)
	}
}

func TestSaveModified_RegeneratesSheetXML(t *testing.T) {
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>
</worksheet>`

	inPath := makeNumericXLSX(t, sheetXML)
	originalSheet := readRawSheetXML(t, inPath)

	wb, _ := cells_foss.LoadWorkbook(inPath)
	wb.Worksheets[0].Cells().Set("A2", float64(99))

	dir := t.TempDir()
	outPath := filepath.Join(dir, "modified.xlsx")
	wb.Save(outPath)

	savedSheet := readRawSheetXML(t, outPath)
	if bytes.Equal(originalSheet, savedSheet) {
		t.Error("modified sheet should differ from original")
	}
}

func TestSave_BooleanRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", true)
	c.Set("A2", false)

	dir := t.TempDir()
	p := filepath.Join(dir, "bool.xlsx")
	wb.Save(p)
	loaded, _ := cells_foss.LoadWorkbook(p)
	lc := loaded.Worksheets[0].Cells()

	c1, _ := lc.Get("A1")
	if c1.Value != true {
		t.Errorf("A1 = %v", c1.Value)
	}
	c2, _ := lc.Get("A2")
	if c2.Value != false {
		t.Errorf("A2 = %v", c2.Value)
	}
}

func TestSave_UpdatesFilePath(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	dir := t.TempDir()
	p := filepath.Join(dir, "fp.xlsx")
	wb.Save(p)
	if wb.FilePath != p {
		t.Errorf("FilePath = %q, want %q", wb.FilePath, p)
	}
}

func TestSave_ResetsModifiedFlag(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	if !wb.Modified {
		t.Error("new workbook should be modified")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "mf.xlsx")
	wb.Save(p)
	if wb.Modified {
		t.Error("modified flag should be false after save")
	}
}

func TestSave_NilAndRepeated(t *testing.T) {
	var nilWB *cells_foss.Workbook
	if err := nilWB.Save("/tmp/x.xlsx"); err == nil {
		t.Error("nil save should error")
	}

	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "persistent")
	dir := t.TempDir()
	wb.Save(filepath.Join(dir, "s1.xlsx"))
	wb.Save(filepath.Join(dir, "s2.xlsx")) // repeated save without modifications
}
