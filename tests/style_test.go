package cells_foss_test

import (
	"bytes"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestStyle_RoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()

	c.Set("A1", "Bold")
	bold := cells_foss.NewStyle()
	bold.Font.Bold = true
	bold.Font.Size = 14
	ca1, _ := c.Get("A1")
	ca1.SetStyle(bold)

	c.Set("B2", "Yellow")
	filled := cells_foss.NewStyle()
	filled.Fill = &cells_foss.Fill{Type: cells_foss.FillTypeSolid, Color: "FFFFFF00"}
	cb2, _ := c.Get("B2")
	cb2.SetStyle(filled)

	dir := t.TempDir()
	p := filepath.Join(dir, "styled.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	lc := loaded.Worksheets[0].Cells()

	c1, _ := lc.Get("A1")
	s := c1.GetStyle()
	if s == nil || !s.Font.Bold || s.Font.Size != 14 {
		t.Error("bold style lost")
	}

	c2, _ := lc.Get("B2")
	s = c2.GetStyle()
	if s == nil || s.Fill.Type != cells_foss.FillTypeSolid || s.Fill.Color != "FFFFFF00" {
		t.Error("fill style lost")
	}
}

func TestStyle_Alignment(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("C3", "Centered")

	centered := cells_foss.NewStyle()
	centered.Alignment = &cells_foss.Alignment{Horizontal: cells_foss.AlignCenter, Vertical: cells_foss.AlignMiddle}
	cell, _ := c.Get("C3")
	cell.SetStyle(centered)

	dir := t.TempDir()
	p := filepath.Join(dir, "aligned.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	cell, _ = loaded.Worksheets[0].Cells().Get("C3")
	s := cell.GetStyle()
	if s.Alignment.Horizontal != cells_foss.AlignCenter {
		t.Error("alignment lost")
	}
}

func TestStyle_Border(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("D4", "Boxed")

	boxed := cells_foss.NewStyle()
	boxed.Border = &cells_foss.Border{Top: true, Bottom: true, Left: true, Right: true}
	cell, _ := c.Get("D4")
	cell.SetStyle(boxed)

	dir := t.TempDir()
	p := filepath.Join(dir, "bordered.xlsx")
	wb.Save(p)

	loaded, _ := cells_foss.LoadWorkbook(p)
	cell, _ = loaded.Worksheets[0].Cells().Get("D4")
	s := cell.GetStyle()
	if s.Border == nil || !s.Border.Top || !s.Border.Bottom || !s.Border.Left || !s.Border.Right {
		t.Error("border lost")
	}
}

func TestStyle_Deduplication(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()

	red := cells_foss.NewStyle()
	red.Font.Color = "FFFF0000"

	c.Set("A1", "Red1")
	c1, _ := c.Get("A1")
	c1.SetStyle(red)

	c.Set("A2", "Red2")
	c2, _ := c.Get("A2")
	c2.SetStyle(red)

	if c1.StyleID != c2.StyleID {
		t.Errorf("same style got different IDs: %d vs %d", c1.StyleID, c2.StyleID)
	}
}

func TestStyle_OrphanCell(t *testing.T) {
	cell := &cells_foss.Cell{}
	if err := cell.SetStyle(cells_foss.NewStyle()); err == nil {
		t.Error("orphan cell SetStyle should error")
	}
	if s := cell.GetStyle(); s != nil {
		t.Error("orphan cell GetStyle should return nil")
	}
}

func TestStyle_StylesXMLContent(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", "Red Bold")

	red := cells_foss.NewStyle()
	red.Font.Bold = true
	red.Font.Color = "FFFF0000"
	cell, _ := c.Get("A1")
	cell.SetStyle(red)

	dir := t.TempDir()
	p := filepath.Join(dir, "stylexml.xlsx")
	wb.Save(p)

	raw := readZipEntry(t, p, "xl/styles.xml")
	if !bytes.Contains(raw, []byte("FFFF0000")) {
		t.Error("styles.xml missing colour")
	}
}
