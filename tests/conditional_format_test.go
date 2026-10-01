package cells_foss_test

import (
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestConditionalFormatting_AddAndSave(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", 50)
	ws.Cells().Set("A2", 150)
	ws.Cells().Set("A3", 250)

	cf := &cells_foss.ConditionalFormatting{
		Ref: "A1:A10",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeCellIs,
				Operator: cells_foss.CondOpGreaterThan,
				Formula:  "100",
				Priority: 1,
			},
		},
	}
	if err := ws.AddConditionalFormatting(cf); err != nil {
		t.Fatalf("AddConditionalFormatting: %v", err)
	}

	dir := t.TempDir()
	p := filepath.Join(dir, "cf.xlsx")
	wb.Save(p)

	// Load and verify.
	loaded, err := cells_foss.LoadWorkbook(p)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	if len(loaded.Worksheets[0].ConditionalFormattings) != 1 {
		t.Fatalf("expected 1 conditional formatting, got %d", len(loaded.Worksheets[0].ConditionalFormattings))
	}
	lcf := loaded.Worksheets[0].ConditionalFormattings[0]
	if lcf.Ref != "A1:A10" {
		t.Errorf("Ref = %q, want A1:A10", lcf.Ref)
	}
	if len(lcf.Rules) != 1 {
		t.Fatalf("expected 1 rule, got %d", len(lcf.Rules))
	}
	rule := lcf.Rules[0]
	if rule.Type != cells_foss.CondTypeCellIs {
		t.Errorf("Type = %q, want cellIs", rule.Type)
	}
	if rule.Operator != cells_foss.CondOpGreaterThan {
		t.Errorf("Operator = %q, want greaterThan", rule.Operator)
	}
	if rule.Formula != "100" {
		t.Errorf("Formula = %q, want 100", rule.Formula)
	}
}

func TestConditionalFormatting_MultipleRules(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	cf := &cells_foss.ConditionalFormatting{
		Ref: "B1:B20",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeCellIs,
				Operator: cells_foss.CondOpBetween,
				Formula:  "10",
				Formula2: "20",
				Priority: 1,
			},
			{
				Type:       cells_foss.CondTypeContainsText,
				Text:       "error",
				Priority:   2,
				StopIfTrue: true,
			},
		},
	}
	ws.AddConditionalFormatting(cf)

	dir := t.TempDir()
	p := filepath.Join(dir, "cf_multi.xlsx")
	wb.Save(p)

	loaded, err := cells_foss.LoadWorkbook(p)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	lcf := loaded.Worksheets[0].ConditionalFormattings[0]
	if len(lcf.Rules) != 2 {
		t.Fatalf("expected 2 rules, got %d", len(lcf.Rules))
	}
	if lcf.Rules[0].Operator != cells_foss.CondOpBetween {
		t.Errorf("rule1 operator = %q, want between", lcf.Rules[0].Operator)
	}
	if lcf.Rules[1].Text != "error" {
		t.Errorf("rule2 text = %q, want error", lcf.Rules[1].Text)
	}
	if !lcf.Rules[1].StopIfTrue {
		t.Error("rule2 StopIfTrue should be true")
	}
}

func TestConditionalFormatting_Remove(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	cf := &cells_foss.ConditionalFormatting{
		Ref: "C1:C10",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{Type: cells_foss.CondTypeExpression, Formula: "TRUE", Priority: 1},
		},
	}
	ws.AddConditionalFormatting(cf)

	if err := ws.RemoveConditionalFormatting("C1:C10"); err != nil {
		t.Fatalf("RemoveConditionalFormatting: %v", err)
	}
	if len(ws.ConditionalFormattings) != 0 {
		t.Errorf("expected 0 conditional formattings after remove, got %d", len(ws.ConditionalFormattings))
	}

	// Remove non-existent should error.
	if err := ws.RemoveConditionalFormatting("X1:X10"); err == nil {
		t.Error("removing non-existent ref should error")
	}
}

func TestConditionalFormatting_NilAndEmpty(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	if err := ws.AddConditionalFormatting(nil); err == nil {
		t.Error("nil ConditionalFormatting should error")
	}
	cf := &cells_foss.ConditionalFormatting{}
	if err := ws.AddConditionalFormatting(cf); err == nil {
		t.Error("empty ref should error")
	}
}
