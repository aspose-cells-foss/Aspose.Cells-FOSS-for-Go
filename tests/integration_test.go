package cells_foss_test

import (
	"os"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose/cells_foss/aspose/cells_foss"
)

func TestIntegration_FullWorkflow(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	c := ws.Cells()

	// Data + styles + formulas + table + DV + picture + CSV + streaming + encrypt.
	c.Set("A1", "Product")
	c.Set("B1", "Price")
	c.Set("A2", "Widget")
	c.Set("B2", float64(12.99))

	bold := cells_foss.NewStyle()
	bold.Font.Bold = true
	cell, _ := c.Get("A1")
	cell.SetStyle(bold)

	c.Set("B3", nil)
	cb3, _ := c.Get("B3")
	cb3.SetFormula("SUM(B2)")

	ws.AddTable("A1:B2")
	ws.AddDataValidation("A2:A10", &cells_foss.DataValidation{
		Type: cells_foss.DataValidationTypeList, Formula1: `"A,B"`,
	})
	ws.AddPicture(cells_foss.NewPicture(cells_foss.MinimalPNG(), "png"))

	// Formula calc.
	r, _ := cells_foss.CalculateFormula("SUM(B2)", ws)
	if r != float64(12.99) {
		t.Errorf("SUM = %v", r)
	}

	// CSV + save + encrypt.
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "export.csv")
	wb.ExportToCSV(0, csvPath, ',')

	wb.SetPassword("pw")
	p := filepath.Join(dir, "enc.xlsx")
	wb.Save(p)

	loaded, err := cells_foss.LoadWithPassword(p, "pw")
	if err != nil {
		t.Fatalf("LoadWithPassword: %v", err)
	}

	lws := loaded.Worksheets[0]
	lc := lws.Cells()
	if c, _ := lc.Get("A1"); c.Value != "Product" {
		t.Errorf("A1 = %v", c.Value)
	}
	if c, _ := lc.Get("B3"); c.GetFormula() != "SUM(B2)" {
		t.Errorf("formula lost")
	}
	if len(lws.Tables) != 1 {
		t.Error("table lost")
	}
	if len(lws.DataValidations) != 1 {
		t.Error("DV lost")
	}

	// Streaming.
	sr := cells_foss.NewStreamingReader(filepath.Join(dir, "plain.xlsx"))
	// First save a plain copy.
	wb.SetPassword("")
	plainPath := filepath.Join(dir, "plain.xlsx")
	wb.Save(plainPath)

	sr = cells_foss.NewStreamingReader(plainPath)
	rows := 0
	sr.ProcessRows("", func(_ int, _ map[string]string) error { rows++; return nil })
	if rows < 2 {
		t.Errorf("streaming: got %d rows", rows)
	}

	// CSV import.
	wb2 := cells_foss.NewWorkbook()
	wb2.ImportFromCSV(csvPath, "Data", ',')
	if len(wb2.Worksheets) != 2 {
		t.Error("CSV import failed")
	}
}

func TestIntegration_ConcurrentSheets(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "First")

	dir := t.TempDir()
	csvPath := filepath.Join(dir, "s2.csv")
	os.WriteFile(csvPath, []byte("Col1,Col2\na,b\n"), 0644)
	wb.ImportFromCSV(csvPath, "Second", ',')

	p := filepath.Join(dir, "multi.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	if len(loaded.Worksheets) != 2 {
		t.Fatalf("sheets = %d, want 2", len(loaded.Worksheets))
	}
	for i, want := range []string{"Sheet1", "Second"} {
		if loaded.Worksheets[i].Name != want {
			t.Errorf("sheet %d = %q, want %q", i, loaded.Worksheets[i].Name, want)
		}
	}
}

func TestIntegration_WorkbookLifecycle(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "v1")

	dir := t.TempDir()
	p1 := filepath.Join(dir, "lifecycle1.xlsx")
	wb.Save(p1)

	loaded, _ := cells_foss.LoadWorkbook(p1)
	loaded.Worksheets[0].Cells().Set("A2", "v2")
	loaded.Modified = true

	p2 := filepath.Join(dir, "lifecycle2.xlsx")
	loaded.Save(p2)

	reloaded, _ := cells_foss.LoadWorkbook(p2)
	c1, _ := reloaded.Worksheets[0].Cells().Get("A1")
	if c1.Value != "v1" {
		t.Errorf("A1 = %v", c1.Value)
	}
	c2, _ := reloaded.Worksheets[0].Cells().Get("A2")
	if c2.Value != "v2" {
		t.Errorf("A2 = %v", c2.Value)
	}
}

func TestIntegration_ErrorCases(t *testing.T) {
	var nilWB *cells_foss.Workbook
	if err := nilWB.Save("/tmp/x.xlsx"); err == nil {
		t.Error("nil save")
	}
	if err := nilWB.SetPassword("x"); err == nil {
		t.Error("nil SetPassword")
	}
	if err := nilWB.ExportToCSV(0, "/tmp/x.csv", ','); err == nil {
		t.Error("nil ExportToCSV")
	}
	if _, err := cells_foss.LoadWorkbook("nonexistent.xlsx"); err == nil {
		t.Error("nonexistent LoadWorkbook")
	}
	if _, err := cells_foss.LoadWithPassword("nonexistent.xlsx", "x"); err == nil {
		t.Error("nonexistent LoadWithPassword")
	}
}

func TestIntegration_CellsCRUD(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()

	c.Set("X1", "hello")
	c.Set("X2", float64(42))

	if cell, _ := c.Get("X1"); cell.Value != "hello" {
		t.Error("X1 fail")
	}
	if len(c.All()) != 2 {
		t.Error("All len")
	}
	c.Remove("X1")
	if _, err := c.Get("X1"); err == nil {
		t.Error("X1 should be removed")
	}
	c.Set("A1", true)
	c.Set("A2", false)
	if cell, _ := c.Get("A1"); cell.Value != true {
		t.Error("bool true")
	}
	if cell, _ := c.Get("A2"); cell.Value != false {
		t.Error("bool false")
	}
}

func TestIntegration_FormulaSetAndRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("D5", nil)
	cell, _ := c.Get("D5")
	cell.SetFormula("SUM(D2:D4)")

	dir := t.TempDir()
	p := filepath.Join(dir, "f.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	cell, _ = loaded.Worksheets[0].Cells().Get("D5")
	if cell.GetFormula() != "SUM(D2:D4)" {
		t.Errorf("formula = %q", cell.GetFormula())
	}
}

func TestIntegration_DataValidationTypes(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	ws.AddDataValidation("A1:A5", &cells_foss.DataValidation{
		Type: cells_foss.DataValidationTypeWhole, Formula1: "1", Formula2: "10",
	})
	dir := t.TempDir()
	p := filepath.Join(dir, "dv2.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	if len(loaded.Worksheets[0].DataValidations) != 1 {
		t.Fatal("DV lost")
	}
	ldv := loaded.Worksheets[0].DataValidations[0]
	if ldv.Formula1 != "1" || ldv.Formula2 != "10" {
		t.Errorf("formulas: %q, %q", ldv.Formula1, ldv.Formula2)
	}
}

func TestIntegration_CellSetGetStyle(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("Z1", "test")

	style := cells_foss.NewStyle()
	style.Font.Bold = true
	style.Font.Italic = true
	style.Font.Size = 16
	style.Font.Color = "FF0000FF"

	cell, _ := c.Get("Z1")
	cell.SetStyle(style)

	s := cell.GetStyle()
	if s == nil || !s.Font.Bold || !s.Font.Italic || s.Font.Size != 16 {
		t.Error("style get/set failed")
	}

	dir := t.TempDir()
	p := filepath.Join(dir, "cellstyle.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	lc := loaded.Worksheets[0].Cells()
	cell, _ = lc.Get("Z1")
	s = cell.GetStyle()
	if s == nil || !s.Font.Bold {
		t.Error("style not preserved in round-trip")
	}
}

func TestIntegration_BorderStyles(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("B2", "bordered")

	style := cells_foss.NewStyle()
	style.Border = &cells_foss.Border{Top: true, Bottom: true, Left: true, Right: true, Color: "FF000000"}
	cell, _ := c.Get("B2")
	cell.SetStyle(style)

	dir := t.TempDir()
	p := filepath.Join(dir, "border.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	cell, _ = loaded.Worksheets[0].Cells().Get("B2")
	s := cell.GetStyle()
	if s.Border == nil || !s.Border.Top || !s.Border.Bottom {
		t.Error("border lost")
	}
}

func TestIntegration_FillAndAlignment(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("C3", "styled")

	style := cells_foss.NewStyle()
	style.Fill = &cells_foss.Fill{Type: cells_foss.FillTypeSolid, Color: "FF00FF00"}
	style.Alignment = &cells_foss.Alignment{Horizontal: cells_foss.AlignCenter, WrapText: true}
	cell, _ := c.Get("C3")
	cell.SetStyle(style)

	dir := t.TempDir()
	p := filepath.Join(dir, "fillalign.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	cell, _ = loaded.Worksheets[0].Cells().Get("C3")
	s := cell.GetStyle()
	if s.Fill.Type != cells_foss.FillTypeSolid || s.Fill.Color != "FF00FF00" {
		t.Error("fill lost")
	}
	if s.Alignment.Horizontal != cells_foss.AlignCenter || !s.Alignment.WrapText {
		t.Error("alignment lost")
	}
}

func TestIntegration_CountItems(t *testing.T) {
	// Count test helpers.
	if png := cells_foss.MinimalPNG(); len(png) == 0 {
		t.Error("MinimalPNG empty")
	}

	// CellToString nil.
	if s := cells_foss.CellToString(nil); s != "" {
		t.Errorf("CellToString(nil) = %q", s)
	}
	if s := cells_foss.CellToString(float64(1.5)); s != "1.5" {
		t.Errorf("CellToString(1.5) = %q", s)
	}

	// Default style check.
	s := cells_foss.NewStyle()
	if s.Font.Name != "Calibri" || s.Font.Size != 11 {
		t.Error("default style mismatch")
	}
}

func TestIntegration_TableCount(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "H")
	ws.Cells().Set("B1", "V")
	ws.Cells().Set("A2", "d1")
	ws.Cells().Set("B2", "d2")

	tbl := ws.AddTable("A1:B2")
	if tbl.Name != "Table1" || tbl.Range != "A1:B2" {
		t.Errorf("table: name=%q range=%q", tbl.Name, tbl.Range)
	}
	if !tbl.HasHeaderRow {
		t.Error("HasHeaderRow default")
	}
	if tbl.StyleName != "TableStyleMedium9" {
		t.Errorf("StyleName = %q", tbl.StyleName)
	}

	// NewTable standalone.
	nt := cells_foss.NewTable("Test", "A1:C10")
	if nt.Name != "Test" || nt.Range != "A1:C10" {
		t.Error("NewTable failed")
	}
}

func TestIntegration_PictureAnchor(t *testing.T) {
	pic := cells_foss.NewPicture(cells_foss.MinimalPNG(), "png")
	pic.SetAnchor(3, 2)
	if pic.Row != 3 || pic.Col != 2 {
		t.Errorf("anchor: row=%d col=%d", pic.Row, pic.Col)
	}
}

func TestIntegration_RemoveDataValidation(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.AddDataValidation("A1:A5", &cells_foss.DataValidation{Type: cells_foss.DataValidationTypeList, Formula1: `"X"`})
	ws.AddDataValidation("B1:B5", &cells_foss.DataValidation{Type: cells_foss.DataValidationTypeWhole, Formula1: "1"})

	if err := ws.RemoveDataValidation("A1:A5"); err != nil {
		t.Fatal(err)
	}
	if len(ws.DataValidations) != 1 {
		t.Errorf("expected 1 after remove, got %d", len(ws.DataValidations))
	}
	if ws.DataValidations[0].Ref != "B1:B5" {
		t.Errorf("remaining ref = %q", ws.DataValidations[0].Ref)
	}
}

func TestIntegration_FileNotFound(t *testing.T) {
	if _, err := cells_foss.LoadWorkbook("no_such_file_xyz.xlsx"); err == nil {
		t.Error("should error")
	}

	sr := cells_foss.NewStreamingReader("no_such_file.xlsx")
	if err := sr.ProcessRows("", func(_ int, _ map[string]string) error { return nil }); err == nil {
		t.Error("should error")
	}
}

func TestIntegration_EncryptDecryptFlow(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", "sensitive")
	c.Set("B1", float64(999))

	wb.SetPassword("strong!")
	dir := t.TempDir()
	p := filepath.Join(dir, "sec.xlsx")
	wb.Save(p)

	// Reload, modify, re-save with password, re-load.
	loaded, err := cells_foss.LoadWithPassword(p, "strong!")
	if err != nil {
		t.Fatalf("load: %v", err)
	}
	loaded.Worksheets[0].Cells().Set("A2", "new data")
	loaded.Modified = true

	p2 := filepath.Join(dir, "sec2.xlsx")
	loaded.Save(p2)

	reloaded, _ := cells_foss.LoadWithPassword(p2, "strong!")
	ca1, _ := reloaded.Worksheets[0].Cells().Get("A1")
	if ca1.Value != "sensitive" {
		t.Error("A1 lost")
	}
	ca2, _ := reloaded.Worksheets[0].Cells().Get("A2")
	if ca2.Value != "new data" {
		t.Error("A2 lost")
	}
}
