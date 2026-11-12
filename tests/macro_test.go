package cells_foss_test

import (
	"bytes"
	"os"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestVBAProject_SetAndGet(t *testing.T) {
	wb := cells_foss.NewWorkbook()

	// Initially no VBA project.
	if wb.HasVBAProject() {
		t.Error("new workbook should not have VBA project")
	}

	// Set VBA project.
	data := []byte{0x01, 0x02, 0x03, 0x04} // Dummy binary data
	if err := wb.SetVBAProject(data); err != nil {
		t.Fatalf("SetVBAProject: %v", err)
	}

	if !wb.HasVBAProject() {
		t.Error("workbook should have VBA project after SetVBAProject")
	}

	vba := wb.GetVBAProject()
	if vba == nil {
		t.Fatal("GetVBAProject returned nil")
	}
	if len(vba.Data) != 4 {
		t.Errorf("VBA data length = %d, want 4", len(vba.Data))
	}
}

func TestVBAProject_EmptyData(t *testing.T) {
	wb := cells_foss.NewWorkbook()

	if err := wb.SetVBAProject([]byte{}); err == nil {
		t.Error("empty VBA data should error")
	}

	if err := wb.SetVBAProject(nil); err == nil {
		t.Error("nil VBA data should error")
	}
}

func TestVBAProject_Clear(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.SetVBAProject([]byte{0x01, 0x02})

	if !wb.HasVBAProject() {
		t.Error("should have VBA project")
	}

	wb.ClearVBAProject()

	if wb.HasVBAProject() {
		t.Error("should not have VBA project after ClearVBAProject")
	}

	if wb.GetVBAProject() != nil {
		t.Error("GetVBAProject should return nil after ClearVBAProject")
	}
}

func TestVBAProject_MarksModified(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Modified = false

	wb.SetVBAProject([]byte{0x01})
	if !wb.Modified {
		t.Error("SetVBAProject should mark workbook as modified")
	}

	wb.Modified = false
	wb.ClearVBAProject()
	if !wb.Modified {
		t.Error("ClearVBAProject should mark workbook as modified")
	}
}

// TestVBAProject_SaveLoadRoundTrip verifies the VBA project part survives a
// save/load cycle: the bytes must be re-read from xl/vbaProject.bin.
func TestVBAProject_SaveLoadRoundTrip(t *testing.T) {
	tmpFile := "test_vba_roundtrip.xlsm"
	defer os.Remove(tmpFile)

	data := []byte("dummy-vba-project-bytes\x00\x01\x02")
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "macro")
	if err := wb.SetVBAProject(data); err != nil {
		t.Fatalf("SetVBAProject: %v", err)
	}
	if err := wb.Save(tmpFile); err != nil {
		t.Fatalf("Save: %v", err)
	}

	loaded, err := cells_foss.Load(tmpFile)
	if err != nil {
		t.Fatalf("Load: %v", err)
	}
	if !loaded.HasVBAProject() {
		t.Fatal("loaded workbook lost its VBA project")
	}
	if !bytes.Equal(loaded.GetVBAProject().Data, data) {
		t.Errorf("VBA data mismatch: got %q, want %q", loaded.GetVBAProject().Data, data)
	}

	// The package must be macro-enabled and declare the VBA part.
	ct, _ := cells_foss.ReadTestZipEntry(tmpFile, "[Content_Types].xml")
	if !bytes.Contains(ct, []byte("application/vnd.ms-excel.sheet.macroEnabled.main+xml")) {
		t.Error("workbook.xml is not declared macro-enabled")
	}
	if !bytes.Contains(ct, []byte("/xl/vbaProject.bin")) {
		t.Error("content types do not declare xl/vbaProject.bin")
	}
	rels, _ := cells_foss.ReadTestZipEntry(tmpFile, "xl/_rels/workbook.xml.rels")
	if !bytes.Contains(rels, []byte(`Target="vbaProject.bin"`)) {
		t.Error("workbook rels do not reference vbaProject.bin")
	}
}

// TestVBAProject_ClearedNotPreserved verifies ClearVBAProject removes the
// part on save rather than preserving it from the loaded package.
func TestVBAProject_ClearedNotPreserved(t *testing.T) {
	src := "test_vba_clear_src.xlsm"
	dst := "test_vba_clear_dst.xlsx"
	defer os.Remove(src)
	defer os.Remove(dst)

	wb := cells_foss.NewWorkbook()
	wb.SetVBAProject([]byte("vba-bytes"))
	if err := wb.Save(src); err != nil {
		t.Fatalf("Save src: %v", err)
	}

	loaded, err := cells_foss.Load(src)
	if err != nil {
		t.Fatalf("Load: %v", err)
	}
	loaded.ClearVBAProject()
	loaded.Worksheets[0].Cells().Set("B1", "edited") // force the dirty path
	if err := loaded.Save(dst); err != nil {
		t.Fatalf("Save dst: %v", err)
	}

	if _, err := cells_foss.ReadTestZipEntry(dst, "xl/vbaProject.bin"); err == nil {
		t.Error("vbaProject.bin survived ClearVBAProject")
	}
	ct, _ := cells_foss.ReadTestZipEntry(dst, "[Content_Types].xml")
	if bytes.Contains(ct, []byte("macroEnabled")) {
		t.Error("package still declared macro-enabled after ClearVBAProject")
	}
}
