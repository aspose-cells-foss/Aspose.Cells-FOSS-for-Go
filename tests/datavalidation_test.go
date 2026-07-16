package cells_foss_test

import (
	"bytes"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose/cells_foss/v26/aspose/cells_foss"
)

func TestDataValidation_AddRemoveRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Pick")

	dv := &cells_foss.DataValidation{
		Type:             cells_foss.DataValidationTypeList,
		Formula1:         `"Red,Green,Blue"`,
		AllowBlank:       true,
		ShowErrorMessage: true,
		ErrorTitle:       "Colour",
		ErrorMessage:     "Pick a colour.",
		ErrorStyle:       cells_foss.ErrorStyleStop,
	}
	if err := ws.AddDataValidation("A2:A20", dv); err != nil {
		t.Fatalf("AddDataValidation: %v", err)
	}
	if dv.Ref != "A2:A20" {
		t.Errorf("Ref = %q", dv.Ref)
	}

	// Add second and remove.
	ws.AddDataValidation("B1:B5", &cells_foss.DataValidation{Type: cells_foss.DataValidationTypeWhole, Formula1: "1", Formula2: "100"})
	if err := ws.RemoveDataValidation("B1:B5"); err != nil {
		t.Fatalf("RemoveDataValidation: %v", err)
	}
	if err := ws.RemoveDataValidation("Z99"); err == nil {
		t.Error("RemoveDataValidation nonexistent should error")
	}

	dir := t.TempDir()
	p := filepath.Join(dir, "dv.xlsx")
	wb.Save(p)

	// Check XML.
	sheetRaw := readZipEntry(t, p, "xl/worksheets/sheet1.xml")
	if !bytes.Contains(sheetRaw, []byte("<dataValidations")) {
		t.Error("missing dataValidations element")
	}

	// Reload.
	loaded, _ := cells_foss.LoadWorkbook(p)
	lws := loaded.Worksheets[0]
	if len(lws.DataValidations) != 1 {
		t.Fatalf("DVs = %d, want 1", len(lws.DataValidations))
	}
	ldv := lws.DataValidations[0]
	if ldv.Type != cells_foss.DataValidationTypeList {
		t.Errorf("type = %q", ldv.Type)
	}
	if ldv.Formula1 != `"Red,Green,Blue"` {
		t.Errorf("Formula1 = %q", ldv.Formula1)
	}
	if ldv.ErrorStyle != cells_foss.ErrorStyleStop {
		t.Errorf("ErrorStyle = %q", ldv.ErrorStyle)
	}
}

func TestDataValidation_NilAndEmpty(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	if err := ws.AddDataValidation("A1", nil); err == nil {
		t.Error("nil DV should error")
	}
	if err := ws.AddDataValidation("", &cells_foss.DataValidation{Type: cells_foss.DataValidationTypeList}); err == nil {
		t.Error("empty ref should error")
	}
}

func TestDataValidation_MarksModified(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Modified = false
	ws := wb.Worksheets[0]
	ws.AddDataValidation("A1:A5", &cells_foss.DataValidation{Type: cells_foss.DataValidationTypeList, Formula1: `"X"`})
	if !wb.Modified {
		t.Error("AddDataValidation should mark Modified")
	}
	wb.Modified = false
	ws.RemoveDataValidation("A1:A5")
	if !wb.Modified {
		t.Error("RemoveDataValidation should mark Modified")
	}
}
