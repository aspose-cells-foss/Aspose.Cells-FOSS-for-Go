package cells_foss_test

import (
	"bytes"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestPicture_AddAndSave(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Hello")

	pic := cells_foss.NewPicture(cells_foss.MinimalPNG(), "png")
	pic.Width = 100
	pic.Height = 80
	pic.SetAnchor(2, 1)

	if err := ws.AddPicture(pic); err != nil {
		t.Fatalf("AddPicture: %v", err)
	}
	if pic.Name != "Picture 1" {
		t.Errorf("Name = %q", pic.Name)
	}

	// Second picture.
	pic2 := cells_foss.NewPicture(cells_foss.MinimalPNG(), "png")
	pic2.SetAnchor(5, 3)
	ws.AddPicture(pic2)
	if pic2.Name != "Picture 2" {
		t.Errorf("Name = %q", pic2.Name)
	}

	dir := t.TempDir()
	p := filepath.Join(dir, "pic.xlsx")
	wb.Save(p)

	// Verify drawing files.
	if raw := readZipEntry(t, p, "xl/drawings/drawing1.xml"); !bytes.Contains(raw, []byte("Picture 1")) {
		t.Error("drawing XML missing")
	}
	if raw := readZipEntry(t, p, "xl/drawings/_rels/drawing1.xml.rels"); !bytes.Contains(raw, []byte("image1.png")) {
		t.Error("drawing rels missing")
	}
	if raw := readZipEntry(t, p, "xl/media/image1.png"); len(raw) == 0 {
		t.Error("image file missing")
	}
	if raw := readZipEntry(t, p, "xl/worksheets/sheet1.xml"); !bytes.Contains(raw, []byte("<drawing")) {
		t.Error("sheet XML missing drawing reference")
	}
}

func TestPicture_InvalidCases(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	if err := ws.AddPicture(nil); err == nil {
		t.Error("nil picture should error")
	}
	if err := ws.AddPicture(cells_foss.NewPicture([]byte{}, "png")); err == nil {
		t.Error("empty data should error")
	}
	if err := ws.AddPicture(cells_foss.NewPicture([]byte{1, 2, 3}, "gif")); err == nil {
		t.Error("unsupported format should error")
	}
}

func TestPicture_JPGNormalized(t *testing.T) {
	pic := cells_foss.NewPicture([]byte{0xFF, 0xD8}, "JPG")
	if pic.Format != "jpeg" {
		t.Errorf("Format = %q, want jpeg", pic.Format)
	}
}

func TestPicture_MarksModified(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Modified = false
	wb.Worksheets[0].AddPicture(cells_foss.NewPicture(cells_foss.MinimalPNG(), "png"))
	if !wb.Modified {
		t.Error("AddPicture should mark Modified")
	}
}

func TestPicture_LoadRoundTrip(t *testing.T) {
	// 1. Create a workbook with pictures.
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Hello")

	pic1 := cells_foss.NewPicture(cells_foss.MinimalPNG(), "png")
	pic1.Width = 100
	pic1.Height = 80
	pic1.SetAnchor(2, 1)
	ws.AddPicture(pic1)

	pic2 := cells_foss.NewPicture(cells_foss.MinimalPNG(), "png")
	pic2.Width = 150
	pic2.Height = 120
	pic2.SetAnchor(5, 3)
	ws.AddPicture(pic2)

	dir := t.TempDir()
	path := filepath.Join(dir, "pic_load.xlsx")
	wb.Save(path)

	// 2. Load the workbook and verify pictures are restored.
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}

	loadedWS := loaded.Worksheets[0]
	if len(loadedWS.Pictures) != 2 {
		t.Fatalf("expected 2 pictures, got %d", len(loadedWS.Pictures))
	}

	// Check first picture.
	p1 := loadedWS.Pictures[0]
	if p1.Row != 2 || p1.Col != 1 {
		t.Errorf("pic1 anchor = (%d,%d), want (2,1)", p1.Row, p1.Col)
	}
	if p1.Width != 100 || p1.Height != 80 {
		t.Errorf("pic1 size = %dx%d, want 100x80", p1.Width, p1.Height)
	}
	if p1.Format != "png" {
		t.Errorf("pic1 format = %q, want png", p1.Format)
	}
	if len(p1.Data) == 0 {
		t.Error("pic1 data is empty")
	}

	// Check second picture.
	p2 := loadedWS.Pictures[1]
	if p2.Row != 5 || p2.Col != 3 {
		t.Errorf("pic2 anchor = (%d,%d), want (5,3)", p2.Row, p2.Col)
	}
	if p2.Width != 150 || p2.Height != 120 {
		t.Errorf("pic2 size = %dx%d, want 150x120", p2.Width, p2.Height)
	}
}

func TestPicture_NoPictures(t *testing.T) {
	// Create a workbook without pictures.
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "No pictures here")

	dir := t.TempDir()
	path := filepath.Join(dir, "no_pic.xlsx")
	wb.Save(path)

	// Load and verify no pictures.
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}

	if len(loaded.Worksheets[0].Pictures) != 0 {
		t.Errorf("expected 0 pictures, got %d", len(loaded.Worksheets[0].Pictures))
	}
}
