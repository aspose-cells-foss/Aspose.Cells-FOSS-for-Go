package cells_foss_test

import (
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestStreamingReader_Basic(t *testing.T) {
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>Hello</v></c></row>
    <row r="2"><c r="A2"><v>99</v></c></row>
  </sheetData>
</worksheet>`

	dir := t.TempDir()
	p := filepath.Join(dir, "stream.xlsx")
	cells_foss.WriteTestXLSX(p, sheetXML, "")

	sr := cells_foss.NewStreamingReader(p)
	var rows []map[string]string
	sr.ProcessRows("", func(rowIdx int, cells map[string]string) error {
		rows = append(rows, cells)
		return nil
	})

	if len(rows) != 2 {
		t.Fatalf("got %d rows, want 2", len(rows))
	}
	if rows[0]["A1"] != "42" || rows[0]["B1"] != "Hello" {
		t.Errorf("row1: %v", rows[0])
	}
	if rows[1]["A2"] != "99" {
		t.Errorf("row2: %v", rows[1])
	}
}

func TestStreamingReader_SharedStrings(t *testing.T) {
	ssXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Apple</t></si><si><t>Banana</t></si>
</sst>`
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>
</worksheet>`

	dir := t.TempDir()
	p := filepath.Join(dir, "ss.xlsx")
	cells_foss.WriteTestXLSX(p, sheetXML, ssXML)

	sr := cells_foss.NewStreamingReader(p)
	var rows []map[string]string
	sr.ProcessRows("", func(_ int, cells map[string]string) error {
		rows = append(rows, cells)
		return nil
	})
	if rows[0]["A1"] != "Apple" || rows[0]["B1"] != "Banana" {
		t.Errorf("shared strings: %v", rows[0])
	}
}

func TestStreamingReader_LargeData(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	numRows := 500
	for r := 1; r <= numRows; r++ {
		c.Set(fmt.Sprintf("A%d", r), r*10)
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "large.xlsx")
	wb.Save(p)

	sr := cells_foss.NewStreamingReader(p)
	count := 0
	sr.ProcessRows("", func(_ int, _ map[string]string) error { count++; return nil })
	if count != numRows {
		t.Errorf("streamed %d rows, want %d", count, numRows)
	}
}

func TestStreamingReader_EarlyStop(t *testing.T) {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>`)
	for r := 1; r <= 20; r++ {
		sb.WriteString(fmt.Sprintf(`<row r="%d"><c r="A%d"><v>%d</v></c></row>`, r, r, r))
	}
	sb.WriteString(`</sheetData></worksheet>`)

	dir := t.TempDir()
	p := filepath.Join(dir, "early.xlsx")
	cells_foss.WriteTestXLSX(p, sb.String(), "")

	sr := cells_foss.NewStreamingReader(p)
	count := 0
	err := sr.ProcessRows("", func(_ int, _ map[string]string) error {
		count++
		if count >= 5 {
			return fmt.Errorf("stop")
		}
		return nil
	})
	if err == nil || err.Error() != "stop" {
		t.Errorf("expected 'stop' error, got %v", err)
	}
	if count != 5 {
		t.Errorf("count = %d, want 5", count)
	}
}

func TestStreamingReader_Errors(t *testing.T) {
	sr := cells_foss.NewStreamingReader("nonexistent.xlsx")
	if err := sr.ProcessRows("", func(_ int, _ map[string]string) error { return nil }); err == nil {
		t.Error("nonexistent file should error")
	}
	if err := cells_foss.NewStreamingReader("test.xlsx").ProcessRows("", nil); err == nil {
		t.Error("nil callback should error")
	}
}

func TestStreamingReader_Boolean(t *testing.T) {
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c></row></sheetData>
</worksheet>`

	dir := t.TempDir()
	p := filepath.Join(dir, "bool.xlsx")
	cells_foss.WriteTestXLSX(p, sheetXML, "")

	sr := cells_foss.NewStreamingReader(p)
	var rows []map[string]string
	sr.ProcessRows("", func(_ int, c map[string]string) error { rows = append(rows, c); return nil })
	if rows[0]["A1"] != "TRUE" || rows[0]["B1"] != "FALSE" {
		t.Errorf("bool: %v", rows[0])
	}
}

func TestStreamingReader_RowRange(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	for r := 1; r <= 20; r++ {
		c.Set(fmt.Sprintf("A%d", r), r*10)
		c.Set(fmt.Sprintf("B%d", r), fmt.Sprintf("Row%d", r))
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "range.xlsx")
	wb.Save(p)

	sr := cells_foss.NewStreamingReader(p)
	var rows []int
	err := sr.ProcessRowsWithRange("", 5, 10, func(rowIdx int, _ map[string]string) error {
		rows = append(rows, rowIdx)
		return nil
	})
	if err != nil {
		t.Fatalf("ProcessRowsWithRange: %v", err)
	}
	if len(rows) != 6 {
		t.Errorf("got %d rows, want 6 (rows 5-10)", len(rows))
	}
	if rows[0] != 5 || rows[len(rows)-1] != 10 {
		t.Errorf("rows = %v, want [5..10]", rows)
	}
}

func TestStreamingReader_ColumnFilter(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", "ColA")
	c.Set("B1", "ColB")
	c.Set("C1", "ColC")
	c.Set("D1", "ColD")
	dir := t.TempDir()
	p := filepath.Join(dir, "cols.xlsx")
	wb.Save(p)

	sr := cells_foss.NewStreamingReader(p)
	var cells map[string]string
	err := sr.ProcessRowsWithColumns("", []string{"A", "C"}, func(_ int, c map[string]string) error {
		cells = c
		return nil
	})
	if err != nil {
		t.Fatalf("ProcessRowsWithColumns: %v", err)
	}
	if len(cells) != 2 {
		t.Errorf("got %d cells, want 2", len(cells))
	}
	if cells["A1"] != "ColA" || cells["C1"] != "ColC" {
		t.Errorf("cells = %v, want A1=ColA, C1=ColC", cells)
	}
	if _, ok := cells["B1"]; ok {
		t.Error("B1 should be filtered out")
	}
}

func TestStreamingReader_CombinedFilter(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	for r := 1; r <= 10; r++ {
		c.Set(fmt.Sprintf("A%d", r), r)
		c.Set(fmt.Sprintf("B%d", r), r*10)
		c.Set(fmt.Sprintf("C%d", r), r*100)
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "combined.xlsx")
	wb.Save(p)

	sr := cells_foss.NewStreamingReader(p)
	opts := &cells_foss.StreamOptions{
		StartRow: 3,
		EndRow:   7,
		Columns:  []string{"A", "C"},
	}
	var rows []map[string]string
	err := sr.ProcessRowsWithFilter("", opts, func(_ int, c map[string]string) error {
		rows = append(rows, c)
		return nil
	})
	if err != nil {
		t.Fatalf("ProcessRowsWithFilter: %v", err)
	}
	if len(rows) != 5 {
		t.Errorf("got %d rows, want 5 (rows 3-7)", len(rows))
	}
	// Check first row (row 3)
	if len(rows[0]) != 2 {
		t.Errorf("first row has %d cells, want 2", len(rows[0]))
	}
	if rows[0]["A3"] != "3" || rows[0]["C3"] != "300" {
		t.Errorf("first row = %v, want A3=3, C3=300", rows[0])
	}
}
