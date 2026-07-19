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
