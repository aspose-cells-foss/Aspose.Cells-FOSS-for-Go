package cells_foss

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestRow_InsertShiftsCells(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "one")
	ws.Cells().Set("A2", "two")
	ws.Cells().Set("A3", "three")
	ws.Cells().Set("B2", 2.0)
	ws.Modified = false

	if err := ws.InsertRow(2); err != nil {
		t.Fatalf("InsertRow: %v", err)
	}

	c, err := ws.Cells().Get("A1")
	if err != nil || c.AsString() != "one" {
		t.Fatalf("A1 = %v, %v; want one", c, err)
	}
	if _, err := ws.Cells().Get("A2"); err == nil {
		t.Fatal("A2 should be empty after inserting a row at 2")
	}
	c, err = ws.Cells().Get("A3")
	if err != nil {
		t.Fatalf("Get A3: %v", err)
	}
	if c.AsString() != "two" {
		t.Fatalf("A3 = %q, want %q (moved from A2)", c.AsString(), "two")
	}
	if c.Ref != "A3" {
		t.Fatalf("moved cell Ref = %q, want A3", c.Ref)
	}
	c, err = ws.Cells().Get("A4")
	if err != nil || c.AsString() != "three" {
		t.Fatalf("A4 = %v, %v; want three", c, err)
	}
	if _, err := ws.Cells().Get("B3"); err != nil {
		t.Fatalf("B3 should hold the moved B2 value: %v", err)
	}
	if !ws.Modified {
		t.Fatal("InsertRow should mark the sheet modified")
	}
	if !wb.Modified {
		t.Fatal("InsertRow should mark the workbook modified")
	}
}

func TestRow_InsertShiftsFeatures(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	if err := ws.SetRowHeight(2, 30); err != nil {
		t.Fatalf("SetRowHeight(2): %v", err)
	}
	if err := ws.SetRowHeight(5, 40); err != nil {
		t.Fatalf("SetRowHeight(5): %v", err)
	}
	if err := ws.AddMergedCells("B2:C3"); err != nil {
		t.Fatalf("AddMergedCells: %v", err)
	}
	if err := ws.SetAutoFilter("A1:D10"); err != nil {
		t.Fatalf("SetAutoFilter: %v", err)
	}
	if err := ws.AddHyperlink("A9", "https://example.com"); err != nil {
		t.Fatalf("AddHyperlink: %v", err)
	}
	dv := &cells_foss.DataValidation{
		Type:     cells_foss.DataValidationTypeList,
		Formula1: `"Yes,No"`,
	}
	if err := ws.AddDataValidation("E2:E8", dv); err != nil {
		t.Fatalf("AddDataValidation: %v", err)
	}
	cf := &cells_foss.ConditionalFormatting{
		Ref: "F2:F8",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{Type: cells_foss.CondTypeCellIs, Operator: cells_foss.CondOpGreaterThan, Formula: "1", Priority: 1},
		},
	}
	if err := ws.AddConditionalFormatting(cf); err != nil {
		t.Fatalf("AddConditionalFormatting: %v", err)
	}
	ws.AddTable("G2:H5")
	ws.Charts = append(ws.Charts, &cells_foss.Chart{
		Name:   "Chart 1",
		Type:   cells_foss.ChartTypeBar,
		Anchor: cells_foss.ChartAnchor{Row: 12, Col: 0},
		Series: []*cells_foss.ChartSeries{
			{Name: "s", Categories: "Sheet1!A2:A5", Values: "B2:B5"},
		},
	})
	ws.Pictures = append(ws.Pictures, &cells_foss.Picture{Row: 20, Col: 1})

	if err := ws.InsertRow(2); err != nil {
		t.Fatalf("InsertRow: %v", err)
	}

	if h := ws.RowHeights[3]; h != 30 {
		t.Fatalf("RowHeights[3] = %v, want 30", h)
	}
	if h := ws.RowHeights[6]; h != 40 {
		t.Fatalf("RowHeights[6] = %v, want 40", h)
	}
	if len(ws.RowHeights) != 2 {
		t.Fatalf("RowHeights = %v, want exactly rows 3 and 6", ws.RowHeights)
	}
	if len(ws.MergedCells) != 1 || ws.MergedCells[0] != "B3:C4" {
		t.Fatalf("MergedCells = %v, want [B3:C4]", ws.MergedCells)
	}
	if ws.AutoFilterRange != "A1:D11" {
		t.Fatalf("AutoFilterRange = %q, want A1:D11", ws.AutoFilterRange)
	}
	if len(ws.Hyperlinks) != 1 || ws.Hyperlinks[0].Ref != "A10" {
		t.Fatalf("Hyperlinks = %v, want one at A10", ws.Hyperlinks)
	}
	if ws.DataValidations[0].Ref != "E3:E9" {
		t.Fatalf("DV Ref = %q, want E3:E9", ws.DataValidations[0].Ref)
	}
	if ws.ConditionalFormattings[0].Ref != "F3:F9" {
		t.Fatalf("CF Ref = %q, want F3:F9", ws.ConditionalFormattings[0].Ref)
	}
	if ws.Tables[0].Range != "G3:H6" {
		t.Fatalf("Table Range = %q, want G3:H6", ws.Tables[0].Range)
	}
	ch := ws.Charts[0]
	if ch.Anchor.Row != 13 {
		t.Fatalf("chart Anchor.Row = %d, want 13", ch.Anchor.Row)
	}
	if ch.Series[0].Categories != "Sheet1!A3:A6" {
		t.Fatalf("Categories = %q, want Sheet1!A3:A6", ch.Series[0].Categories)
	}
	if ch.Series[0].Values != "B3:B6" {
		t.Fatalf("Values = %q, want B3:B6", ch.Series[0].Values)
	}
	if ws.Pictures[0].Row != 21 {
		t.Fatalf("picture Row = %d, want 21", ws.Pictures[0].Row)
	}
}

func TestRow_InsertExpandsSpannedRanges(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	if err := ws.AddMergedCells("B4:C6"); err != nil {
		t.Fatalf("AddMergedCells: %v", err)
	}
	if err := ws.InsertRow(5); err != nil {
		t.Fatalf("InsertRow: %v", err)
	}
	if len(ws.MergedCells) != 1 || ws.MergedCells[0] != "B4:C7" {
		t.Fatalf("MergedCells = %v, want [B4:C7] (insert inside expands)", ws.MergedCells)
	}

	// Inserting exactly at the top edge shifts the range down instead.
	if err := ws.InsertRow(4); err != nil {
		t.Fatalf("InsertRow(4): %v", err)
	}
	if ws.MergedCells[0] != "B5:C8" {
		t.Fatalf("MergedCells = %v, want [B5:C8] (insert at top edge shifts)", ws.MergedCells)
	}
}

func TestRow_InsertMultiRangeSqref(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	cf := &cells_foss.ConditionalFormatting{
		Ref: "A2:A5 C3:C8",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{Type: cells_foss.CondTypeExpression, Formula: "TRUE", Priority: 1},
		},
	}
	if err := ws.AddConditionalFormatting(cf); err != nil {
		t.Fatalf("AddConditionalFormatting: %v", err)
	}
	if err := ws.InsertRow(3); err != nil {
		t.Fatalf("InsertRow: %v", err)
	}
	if cf.Ref != "A2:A6 C4:C9" {
		t.Fatalf("CF Ref = %q, want %q", cf.Ref, "A2:A6 C4:C9")
	}
}

func TestRow_DeleteDropsAndShifts(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "keep")
	ws.Cells().Set("A2", "drop")
	ws.Cells().Set("A3", "rise")
	if err := ws.SetRowHeight(2, 30); err != nil {
		t.Fatalf("SetRowHeight(2): %v", err)
	}
	if err := ws.SetRowHeight(3, 25); err != nil {
		t.Fatalf("SetRowHeight(3): %v", err)
	}
	if err := ws.AddMergedCells("B2:C3"); err != nil {
		t.Fatalf("AddMergedCells(1): %v", err)
	}
	if err := ws.AddMergedCells("D5:E5"); err != nil {
		t.Fatalf("AddMergedCells(2): %v", err)
	}
	if err := ws.AddHyperlink("A2", "https://gone.example"); err != nil {
		t.Fatalf("AddHyperlink(A2): %v", err)
	}
	if err := ws.AddHyperlink("A3", "https://rise.example"); err != nil {
		t.Fatalf("AddHyperlink(A3): %v", err)
	}
	dv := &cells_foss.DataValidation{
		Type:     cells_foss.DataValidationTypeList,
		Formula1: `"Yes,No"`,
	}
	if err := ws.AddDataValidation("E2:E8", dv); err != nil {
		t.Fatalf("AddDataValidation: %v", err)
	}
	cf := &cells_foss.ConditionalFormatting{
		Ref: "F1:F4",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{Type: cells_foss.CondTypeExpression, Formula: "TRUE", Priority: 1},
		},
	}
	if err := ws.AddConditionalFormatting(cf); err != nil {
		t.Fatalf("AddConditionalFormatting: %v", err)
	}
	ws.AddTable("G2:H5")
	ws.Charts = append(ws.Charts, &cells_foss.Chart{
		Name:   "Chart 1",
		Type:   cells_foss.ChartTypeBar,
		Anchor: cells_foss.ChartAnchor{Row: 1, Col: 0},
		Series: []*cells_foss.ChartSeries{
			{Name: "s", Categories: "Sheet1!A2:A5", Values: "B3:B5"},
		},
	})
	ws.Pictures = append(ws.Pictures, &cells_foss.Picture{Row: 4, Col: 1})

	if err := ws.DeleteRow(2); err != nil {
		t.Fatalf("DeleteRow: %v", err)
	}

	c, err := ws.Cells().Get("A1")
	if err != nil || c.AsString() != "keep" {
		t.Fatalf("A1 = %v, %v; want keep", c, err)
	}
	c, err = ws.Cells().Get("A2")
	if err != nil {
		t.Fatalf("Get A2: %v", err)
	}
	if c.AsString() != "rise" {
		t.Fatalf("A2 = %q, want %q (moved up from A3)", c.AsString(), "rise")
	}
	if c.Ref != "A2" {
		t.Fatalf("moved cell Ref = %q, want A2", c.Ref)
	}
	if _, err := ws.Cells().Get("A3"); err == nil {
		t.Fatal("A3 should be empty after the delete")
	}
	if len(ws.RowHeights) != 1 || ws.RowHeights[2] != 25 {
		t.Fatalf("RowHeights = %v, want {2:25}", ws.RowHeights)
	}
	if len(ws.MergedCells) != 2 || ws.MergedCells[0] != "B2:C2" || ws.MergedCells[1] != "D4:E4" {
		t.Fatalf("MergedCells = %v, want [B2:C2 D4:E4]", ws.MergedCells)
	}
	if len(ws.Hyperlinks) != 1 || ws.Hyperlinks[0].Ref != "A2" {
		t.Fatalf("Hyperlinks = %v, want one at A2 (link on deleted row dropped)", ws.Hyperlinks)
	}
	if ws.DataValidations[0].Ref != "E2:E7" {
		t.Fatalf("DV Ref = %q, want E2:E7", ws.DataValidations[0].Ref)
	}
	if ws.ConditionalFormattings[0].Ref != "F1:F3" {
		t.Fatalf("CF Ref = %q, want F1:F3", ws.ConditionalFormattings[0].Ref)
	}
	if ws.Tables[0].Range != "G2:H4" {
		t.Fatalf("Table Range = %q, want G2:H4", ws.Tables[0].Range)
	}
	ch := ws.Charts[0]
	if ch.Anchor.Row != 0 {
		t.Fatalf("chart Anchor.Row = %d, want 0 (slides up, clamped)", ch.Anchor.Row)
	}
	if ch.Series[0].Categories != "Sheet1!A2:A4" {
		t.Fatalf("Categories = %q, want Sheet1!A2:A4", ch.Series[0].Categories)
	}
	if ch.Series[0].Values != "B2:B4" {
		t.Fatalf("Values = %q, want B2:B4", ch.Series[0].Values)
	}
	if ws.Pictures[0].Row != 3 {
		t.Fatalf("picture Row = %d, want 3", ws.Pictures[0].Row)
	}
}

func TestRow_DeleteRemovesVanishedFeatures(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	if err := ws.AddMergedCells("D5:E5"); err != nil {
		t.Fatalf("AddMergedCells: %v", err)
	}
	dv := &cells_foss.DataValidation{
		Type:     cells_foss.DataValidationTypeWhole,
		Formula1: "1",
		Formula2: "10",
	}
	if err := ws.AddDataValidation("A1:B1", dv); err != nil {
		t.Fatalf("AddDataValidation: %v", err)
	}
	cf := &cells_foss.ConditionalFormatting{
		Ref: "C2:C2",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{Type: cells_foss.CondTypeExpression, Formula: "TRUE", Priority: 1},
		},
	}
	if err := ws.AddConditionalFormatting(cf); err != nil {
		t.Fatalf("AddConditionalFormatting: %v", err)
	}

	if err := ws.DeleteRow(5); err != nil {
		t.Fatalf("DeleteRow(5): %v", err)
	}
	if len(ws.MergedCells) != 0 {
		t.Fatalf("MergedCells = %v, want empty (single-row merge vanishes)", ws.MergedCells)
	}
	if err := ws.DeleteRow(2); err != nil {
		t.Fatalf("DeleteRow(2): %v", err)
	}
	if len(ws.ConditionalFormattings) != 0 {
		t.Fatalf("ConditionalFormattings = %v, want dropped (range vanished)", ws.ConditionalFormattings)
	}
	if err := ws.DeleteRow(1); err != nil {
		t.Fatalf("DeleteRow(1): %v", err)
	}
	if len(ws.DataValidations) != 0 {
		t.Fatalf("DataValidations = %v, want dropped (range vanished)", ws.DataValidations)
	}
}

func TestRow_FormulasAreNotRewritten(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", 1.0)
	ws.Cells().Set("A2", 2.0)
	c, err := ws.Cells().Get("A2")
	if err != nil {
		t.Fatalf("Get: %v", err)
	}
	c.SetFormula("SUM(A1:A2)")
	before := c.Formula

	if err := ws.InsertRow(1); err != nil {
		t.Fatalf("InsertRow: %v", err)
	}
	moved, err := ws.Cells().Get("A3")
	if err != nil {
		t.Fatalf("Get A3: %v", err)
	}
	if moved.Formula != before {
		t.Fatalf("Formula = %q, want unchanged %q", moved.Formula, before)
	}
}

func TestRow_BoundsAndOverflow(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	if err := ws.InsertRow(0); err == nil || !strings.Contains(err.Error(), "outside the sheet area") {
		t.Fatalf("InsertRow(0) = %v, want out-of-range error", err)
	}
	if err := ws.InsertRow(1048577); err == nil {
		t.Fatal("InsertRow(1048577) should fail")
	}
	if err := ws.DeleteRow(0); err == nil {
		t.Fatal("DeleteRow(0) should fail")
	}
	if err := ws.DeleteRow(-3); err == nil {
		t.Fatal("DeleteRow(-3) should fail")
	}

	// A cell on the last row blocks inserts that would shift it off-sheet.
	ws.Cells().Set("A1048576", "last")
	if err := ws.InsertRow(10); err == nil || !strings.Contains(err.Error(), "last row") {
		t.Fatalf("InsertRow with content on row 1048576 = %v, want overflow error", err)
	}
	// Deleting never overflows.
	if err := ws.DeleteRow(1048576); err != nil {
		t.Fatalf("DeleteRow(1048576): %v", err)
	}
	if err := ws.InsertRow(10); err != nil {
		t.Fatalf("InsertRow after delete: %v", err)
	}
}

func TestRow_SaveLoadRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "header")
	ws.Cells().Set("A2", "first")
	ws.Cells().Set("A3", "second")
	ws.Cells().Set("A4", "third")
	if err := ws.SetRowHeight(3, 42); err != nil {
		t.Fatalf("SetRowHeight: %v", err)
	}

	if err := ws.InsertRow(3); err != nil {
		t.Fatalf("InsertRow: %v", err)
	}
	if err := ws.DeleteRow(1); err != nil {
		t.Fatalf("DeleteRow: %v", err)
	}

	path := filepath.Join(t.TempDir(), "rows.xlsx")
	if err := wb.Save(path); err != nil {
		t.Fatalf("Save: %v", err)
	}
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	lws := loaded.Worksheets[0]
	// After insert@3 then delete@1: first→A1, second→A3 (gap at A2), third→A3? no:
	// insert@3: A1 header, A2 first, A3 (blank), A4 second, A5 third
	// delete@1: A1 first, A2 (blank), A3 second, A4 third
	want := map[string]string{"A1": "first", "A3": "second", "A4": "third"}
	for ref, val := range want {
		c, err := lws.Cells().Get(ref)
		if err != nil {
			t.Fatalf("Get %s: %v", ref, err)
		}
		if c.AsString() != val {
			t.Fatalf("%s = %q, want %q", ref, c.AsString(), val)
		}
	}
	if _, err := lws.Cells().Get("A2"); err == nil {
		t.Fatal("A2 should be empty after the round trip")
	}
	if h := lws.RowHeights[3]; h != 42 {
		t.Fatalf("RowHeights[3] = %v, want 42 (moved 3→4 by insert, 4→3 by delete)", h)
	}
}
