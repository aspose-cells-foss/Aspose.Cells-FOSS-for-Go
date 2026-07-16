package cells_foss_test

import (
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose/cells_foss/v26/aspose/cells_foss"
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
