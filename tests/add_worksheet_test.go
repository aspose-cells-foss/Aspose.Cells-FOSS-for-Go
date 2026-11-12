package cells_foss_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestAddWorksheet_Basics(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws, err := wb.AddWorksheet("Data")
	if err != nil {
		t.Fatalf("AddWorksheet: %v", err)
	}
	if ws.Name != "Data" {
		t.Errorf("Name = %q, want Data", ws.Name)
	}
	if ws.Index != 1 {
		t.Errorf("Index = %d, want 1", ws.Index)
	}
	if len(wb.Worksheets) != 2 {
		t.Fatalf("len(Worksheets) = %d, want 2", len(wb.Worksheets))
	}

	// The returned sheet must be wired for cell access.
	if err := ws.Cells().Set("A1", "hello"); err != nil {
		t.Fatalf("Set: %v", err)
	}
	ws.Cells().Set("B1", 42.0)

	path := filepath.Join(t.TempDir(), "addws.xlsx")
	if err := wb.Save(path); err != nil {
		t.Fatalf("Save: %v", err)
	}

	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	if len(wb2.Worksheets) != 2 {
		t.Fatalf("reloaded %d sheets, want 2", len(wb2.Worksheets))
	}
	if wb2.Worksheets[0].Name != "Sheet1" || wb2.Worksheets[1].Name != "Data" {
		t.Errorf("sheet names = %q, %q", wb2.Worksheets[0].Name, wb2.Worksheets[1].Name)
	}
	a1, err := wb2.Worksheets[1].Cells().Get("A1")
	if err != nil {
		t.Fatalf("Get A1: %v", err)
	}
	if a1.AsString() != "hello" {
		t.Errorf("A1 = %q, want hello", a1.AsString())
	}
	b1, _ := wb2.Worksheets[1].Cells().Get("B1")
	if v, ok := b1.AsFloat64(); !ok || v != 42 {
		t.Errorf("B1 = %v (ok=%v), want 42", v, ok)
	}
}

func TestAddWorksheet_NameValidation(t *testing.T) {
	bad := []string{
		"",
		strings.Repeat("x", 32),
		"a[b", "a]b", "a:b", "a*b", "a?b", "a/b", `a\b`,
		"'leading", "trailing'",
		"History", "history", // reserved, case-insensitive
		"Sheet1", "sheet1", // duplicate of the default sheet
	}
	for _, name := range bad {
		wb := cells_foss.NewWorkbook()
		if ws, err := wb.AddWorksheet(name); err == nil {
			t.Errorf("AddWorksheet(%q) succeeded (Index=%d), want error", name, ws.Index)
		}
		if len(wb.Worksheets) != 1 {
			t.Errorf("AddWorksheet(%q) modified the sheet list", name)
		}
	}

	// 31 Unicode characters is the maximum, and non-ASCII names are fine.
	wb := cells_foss.NewWorkbook()
	name := strings.Repeat("日", 31)
	if _, err := wb.AddWorksheet(name); err != nil {
		t.Errorf("AddWorksheet(31 runes) = %v, want nil", err)
	}
	// The name is now taken (case-insensitive duplicate check).
	if _, err := wb.AddWorksheet(name); err == nil {
		t.Error("duplicate AddWorksheet should fail")
	}
}

func TestAddWorksheet_LoadedWorkbook(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "original")
	path := filepath.Join(t.TempDir(), "base.xlsx")
	if err := wb.Save(path); err != nil {
		t.Fatalf("Save: %v", err)
	}

	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	ws, err := wb2.AddWorksheet("Extra")
	if err != nil {
		t.Fatalf("AddWorksheet on loaded workbook: %v", err)
	}
	ws.Cells().Set("A1", "added")

	path2 := filepath.Join(t.TempDir(), "extended.xlsx")
	if err := wb2.Save(path2); err != nil {
		t.Fatalf("Save: %v", err)
	}

	wb3, err := cells_foss.LoadWorkbook(path2)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	if len(wb3.Worksheets) != 2 {
		t.Fatalf("reloaded %d sheets, want 2", len(wb3.Worksheets))
	}
	orig, err := wb3.Worksheets[0].Cells().Get("A1")
	if err != nil {
		t.Fatalf("Get original A1: %v", err)
	}
	if orig.AsString() != "original" {
		t.Errorf("unmodified sheet A1 = %q, want original", orig.AsString())
	}
	added, err := wb3.Worksheets[1].Cells().Get("A1")
	if err != nil {
		t.Fatalf("Get added A1: %v", err)
	}
	if added.AsString() != "added" {
		t.Errorf("added sheet A1 = %q, want added", added.AsString())
	}
}

func TestImportFromCSV_RejectsDuplicateSheetName(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "in.csv")
	if err := os.WriteFile(csvPath, []byte("a,b\n1,2\n"), 0644); err != nil {
		t.Fatalf("write csv: %v", err)
	}

	wb := cells_foss.NewWorkbook()
	// "Sheet1" already exists: the import must fail and leave the
	// workbook unchanged (previously duplicate names were silently
	// accepted, producing files Excel rejects).
	if err := wb.ImportFromCSV(csvPath, "Sheet1", ','); err == nil {
		t.Fatal("ImportFromCSV with duplicate sheet name should fail")
	}
	if len(wb.Worksheets) != 1 {
		t.Errorf("failed import left %d sheets, want 1", len(wb.Worksheets))
	}

	// A fresh name still works.
	if err := wb.ImportFromCSV(csvPath, "Imported", ','); err != nil {
		t.Fatalf("ImportFromCSV: %v", err)
	}
	if len(wb.Worksheets) != 2 {
		t.Errorf("len(Worksheets) = %d, want 2", len(wb.Worksheets))
	}
}
