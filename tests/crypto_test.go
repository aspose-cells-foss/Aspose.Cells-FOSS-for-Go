package cells_foss_test

import (
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose/cells_foss/aspose/cells_foss"
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
