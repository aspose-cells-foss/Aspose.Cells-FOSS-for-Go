package cells_foss_test

import (
	"os"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestEncryption_RoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", "Secret")
	c.Set("B1", float64(42))

	wb.SetPassword("test123")
	dir := t.TempDir()
	p := filepath.Join(dir, "enc.xlsx")
	wb.Save(p)

	// Normal load should fail.
	if _, err := cells_foss.LoadWorkbook(p); err == nil {
		t.Error("LoadWorkbook on encrypted file should fail")
	}

	// Wrong password should fail.
	if _, err := cells_foss.LoadWithPassword(p, "wrong"); err == nil {
		t.Error("wrong password should fail")
	}

	// Correct password.
	loaded, err := cells_foss.LoadWithPassword(p, "test123")
	if err != nil {
		t.Fatalf("LoadWithPassword: %v", err)
	}

	lc := loaded.Worksheets[0].Cells()
	ca1, _ := lc.Get("A1")
	if ca1.Value != "Secret" {
		t.Errorf("A1 = %v", ca1.Value)
	}
	cb1, _ := lc.Get("B1")
	if cells_foss.CellToString(cb1.Value) != "42" {
		t.Errorf("B1 = %v", cb1.Value)
	}

	// Password preserved.
	if !loaded.VerifyPassword("test123") {
		t.Error("VerifyPassword failed")
	}
	if loaded.VerifyPassword("wrong") {
		t.Error("VerifyPassword should return false")
	}

	// Remove password, re-save.
	loaded.Modified = true
	loaded.SetPassword("")
	p2 := filepath.Join(dir, "plain.xlsx")
	loaded.Save(p2)

	wb2, _ := cells_foss.LoadWorkbook(p2)
	c2, _ := wb2.Worksheets[0].Cells().Get("A1")
	if c2.Value != "Secret" {
		t.Errorf("after password removal: A1 = %v", c2.Value)
	}
}

func TestEncryption_EmptyAndNil(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "plain")

	dir := t.TempDir()
	p := filepath.Join(dir, "plain.xlsx")
	wb.Save(p)

	// LoadWithPassword on plain file should work.
	loaded, err := cells_foss.LoadWithPassword(p, "any")
	if err != nil {
		t.Fatalf("LoadWithPassword on plain: %v", err)
	}
	if !loaded.VerifyPassword("any") {
		t.Error("VerifyPassword on plain should return true")
	}

	var nilWB *cells_foss.Workbook
	if err := nilWB.SetPassword("x"); err == nil {
		t.Error("nil SetPassword should error")
	}
}

func TestEncryption_LoadWithPasswordErrors(t *testing.T) {
	if _, err := cells_foss.LoadWithPassword("nonexistent.xlsx", "pw"); err == nil {
		t.Error("nonexistent file should error")
	}
}

func TestEncryption_CFBStructure(t *testing.T) {
	// Verify that encrypted files use the standard OLE/CFB format.
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "test")
	wb.SetPassword("secret")

	dir := t.TempDir()
	p := filepath.Join(dir, "enc.xlsx")
	wb.Save(p)

	// Read the file and verify it starts with OLE/CFB magic.
	data, err := os.ReadFile(p)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}

	// OLE/CFB magic: D0 CF 11 E0 A1 B1 1A E1
	magic := []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
	if len(data) < 8 {
		t.Fatal("file too short")
	}
	for i, b := range magic {
		if data[i] != b {
			t.Errorf("byte %d: got 0x%02X, want 0x%02X", i, data[i], b)
		}
	}

	// Verify the file can be loaded back.
	loaded, err := cells_foss.LoadWithPassword(p, "secret")
	if err != nil {
		t.Fatalf("LoadWithPassword: %v", err)
	}
	c, _ := loaded.Worksheets[0].Cells().Get("A1")
	if c.Value != "test" {
		t.Errorf("A1 = %v, want 'test'", c.Value)
	}
}
