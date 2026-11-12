package cells_foss_test

import (
	"bytes"
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
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

func TestDataValidation_Operator(t *testing.T) {
	// All eight spec operators are accepted; unknown ones are refused.
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ops := []string{
		cells_foss.DataValidationOpBetween,
		cells_foss.DataValidationOpNotBetween,
		cells_foss.DataValidationOpEqual,
		cells_foss.DataValidationOpNotEqual,
		cells_foss.DataValidationOpLessThan,
		cells_foss.DataValidationOpLessThanOrEqual,
		cells_foss.DataValidationOpGreaterThan,
		cells_foss.DataValidationOpGreaterThanOrEqual,
	}
	for i, op := range ops {
		dv := &cells_foss.DataValidation{
			Type: cells_foss.DataValidationTypeWhole, Operator: op, Formula1: "1",
		}
		if err := ws.AddDataValidation("A"+string(rune('1'+i)), dv); err != nil {
			t.Fatalf("operator %q rejected: %v", op, err)
		}
	}
	if err := ws.AddDataValidation("B1", &cells_foss.DataValidation{
		Type: cells_foss.DataValidationTypeWhole, Operator: "sideways", Formula1: "1",
	}); err == nil {
		t.Error("invalid operator should be rejected")
	}

	// Save/load: an explicit operator is written; the empty default is
	// omitted from the XML, matching Excel's own output.
	wb = cells_foss.NewWorkbook()
	ws = wb.Worksheets[0]
	if err := ws.AddDataValidation("A2:A5", &cells_foss.DataValidation{
		Type:     cells_foss.DataValidationTypeWhole,
		Operator: cells_foss.DataValidationOpGreaterThan,
		Formula1: "10",
	}); err != nil {
		t.Fatalf("AddDataValidation (greaterThan): %v", err)
	}
	if err := ws.AddDataValidation("B2:B5", &cells_foss.DataValidation{
		Type: cells_foss.DataValidationTypeDecimal, Formula1: "0", Formula2: "1",
	}); err != nil {
		t.Fatalf("AddDataValidation (default): %v", err)
	}
	if err := ws.AddDataValidation("C2:C5", &cells_foss.DataValidation{
		Type:     cells_foss.DataValidationTypeWhole,
		Operator: cells_foss.DataValidationOpBetween,
		Formula1: "1", Formula2: "9",
	}); err != nil {
		t.Fatalf("AddDataValidation (explicit between): %v", err)
	}

	p := filepath.Join(t.TempDir(), "dvop.xlsx")
	if err := wb.Save(p); err != nil {
		t.Fatalf("Save: %v", err)
	}
	sheetRaw := string(readZipEntry(t, p, "xl/worksheets/sheet1.xml"))
	if n := strings.Count(sheetRaw, "operator="); n != 2 {
		t.Errorf("found %d operator attributes, want 2 (empty default omitted)", n)
	}
	if !strings.Contains(sheetRaw, `operator="greaterThan"`) {
		t.Error(`missing operator="greaterThan"`)
	}
	if !strings.Contains(sheetRaw, `operator="between"`) {
		t.Error(`missing explicit operator="between"`)
	}

	loaded, err := cells_foss.LoadWorkbook(p)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	dvs := loaded.Worksheets[0].DataValidations
	if len(dvs) != 3 {
		t.Fatalf("loaded %d DVs, want 3", len(dvs))
	}
	if dvs[0].Operator != cells_foss.DataValidationOpGreaterThan {
		t.Errorf("Operator = %q, want greaterThan", dvs[0].Operator)
	}
	if dvs[1].Operator != "" {
		t.Errorf("Operator = %q, want empty (omitted default)", dvs[1].Operator)
	}
	if dvs[2].Operator != cells_foss.DataValidationOpBetween {
		t.Errorf("Operator = %q, want between", dvs[2].Operator)
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
