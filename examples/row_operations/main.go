// Row operations example: insert and delete rows while cells, merged
// ranges, data validations, conditional formats and chart anchors move
// along with the data.
//
//	cd examples && go run ./row_operations/
package main

import (
	"fmt"
	"os"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	ws.Cells().Set("A1", "Quarter")
	ws.Cells().Set("B1", "Revenue")
	rows := []struct {
		q       string
		revenue float64
	}{
		{"Q1", 1200},
		{"Q2", 1500},
		{"Q3", 900},
	}
	for i, r := range rows {
		ws.Cells().Set(fmt.Sprintf("A%d", i+2), r.q)
		ws.Cells().Set(fmt.Sprintf("B%d", i+2), r.revenue)
	}

	if err := ws.AddMergedCells("D2:E3"); err != nil {
		fmt.Fprintf(os.Stderr, "AddMergedCells: %v\n", err)
		os.Exit(1)
	}
	dv := &cells_foss.DataValidation{
		Type:     cells_foss.DataValidationTypeList,
		Formula1: `"Q1,Q1b,Q2,Q3"`,
	}
	if err := ws.AddDataValidation("A2:A4", dv); err != nil {
		fmt.Fprintf(os.Stderr, "AddDataValidation: %v\n", err)
		os.Exit(1)
	}
	cf := &cells_foss.ConditionalFormatting{
		Ref: "B2:B4",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeDataBar,
				Priority: 1,
				DataBar: &cells_foss.DataBar{
					Min:   cells_foss.CondThreshold{Type: cells_foss.CondValueMin},
					Max:   cells_foss.CondThreshold{Type: cells_foss.CondValueMax},
					Color: "FF638EC6",
				},
			},
		},
	}
	if err := ws.AddConditionalFormatting(cf); err != nil {
		fmt.Fprintf(os.Stderr, "AddConditionalFormatting: %v\n", err)
		os.Exit(1)
	}
	chart := &cells_foss.Chart{
		Name:   "Revenue Chart",
		Type:   cells_foss.ChartTypeBar,
		Anchor: cells_foss.ChartAnchor{Row: 6, Col: 0},
		Width:  480,
		Height: 300,
		Series: []*cells_foss.ChartSeries{
			{Name: "Revenue", Categories: "A2:A4", Values: "B2:B4"},
		},
	}
	if err := ws.AddChart(chart); err != nil {
		fmt.Fprintf(os.Stderr, "AddChart: %v\n", err)
		os.Exit(1)
	}

	fmt.Println("Before:")
	printState(ws)

	// Insert a blank row above Q2 — everything at or below row 3 moves down,
	// and ranges spanning the insertion point expand.
	if err := ws.InsertRow(3); err != nil {
		fmt.Fprintf(os.Stderr, "InsertRow: %v\n", err)
		os.Exit(1)
	}
	ws.Cells().Set("A3", "Q1b")
	ws.Cells().Set("B3", 1350.0)
	fmt.Println("\nAfter InsertRow(3) + filling the new row:")
	printState(ws)

	// Delete the Q3 row — cells below slide up and ranges shrink. The insert
	// moved Q3 from row 4 to row 5.
	if err := ws.DeleteRow(5); err != nil {
		fmt.Fprintf(os.Stderr, "DeleteRow: %v\n", err)
		os.Exit(1)
	}
	fmt.Println("\nAfter DeleteRow(5):")
	printState(ws)

	path := "outputfiles/row_operations.xlsx"
	if err := wb.Save(path); err != nil {
		fmt.Fprintf(os.Stderr, "Save: %v\n", err)
		os.Exit(1)
	}

	// The shifted layout is what gets written: reload and confirm.
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Load: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("\nSaved %s; reloaded state:\n", path)
	printState(loaded.Worksheets[0])
	fmt.Println("Formulas are NOT rewritten by row operations — a formula")
	fmt.Println("like =SUM(B2:B4) keeps its text even when those rows move.")
}

// printState dumps the cells plus every row-aware feature of the sheet.
func printState(ws *cells_foss.Worksheet) {
	for r := 1; r <= 6; r++ {
		a, aerr := ws.Cells().Get(fmt.Sprintf("A%d", r))
		b, berr := ws.Cells().Get(fmt.Sprintf("B%d", r))
		if aerr != nil && berr != nil {
			continue
		}
		fmt.Printf("  row %d: %-10s %s\n", r, text(a, aerr), text(b, berr))
	}
	fmt.Printf("  merged:      %v\n", ws.MergedCells)
	fmt.Printf("  validation:  %s\n", ws.DataValidations[0].Ref)
	fmt.Printf("  cond format: %s\n", ws.ConditionalFormattings[0].Ref)
	fmt.Printf("  chart at 0-based row %d, series %s / %s\n",
		ws.Charts[0].Anchor.Row,
		ws.Charts[0].Series[0].Categories,
		ws.Charts[0].Series[0].Values)
}

func text(c *cells_foss.Cell, err error) string {
	if err != nil || c == nil {
		return ""
	}
	return c.AsString()
}
