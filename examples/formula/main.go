// Formula example: create a workbook with SUM / AVERAGE formulas and
// evaluate them with the built-in formula engine.
//
//	cd examples && go run ./formula/
package main

import (
	"fmt"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// ---- Populate sales data ----
	data := []float64{1200, 850, 1400, 960, 1780}
	headers := []string{"Month", "Sales"}

	ws.Cells().Set("A1", headers[0])
	ws.Cells().Set("B1", headers[1])

	months := []string{"Jan", "Feb", "Mar", "Apr", "May"}
	for i, m := range months {
		row := i + 2
		ws.Cells().Set(fmt.Sprintf("A%d", row), m)
		ws.Cells().Set(fmt.Sprintf("B%d", row), data[i])
	}

	// ---- Add formula cells ----
	lastDataRow := len(data) + 1
	totalRef := fmt.Sprintf("B2:B%d", lastDataRow)

	// SUM formula.
	ws.Cells().Set("B7", nil)
	sumCell, _ := ws.Cells().Get("B7")
	sumCell.SetFormula(fmt.Sprintf("SUM(%s)", totalRef))

	// AVERAGE formula.
	ws.Cells().Set("B8", nil)
	avgCell, _ := ws.Cells().Get("B8")
	avgCell.SetFormula(fmt.Sprintf("AVERAGE(%s)", totalRef))

	// MAX formula.
	ws.Cells().Set("B9", nil)
	maxCell, _ := ws.Cells().Get("B9")
	maxCell.SetFormula(fmt.Sprintf("MAX(%s)", totalRef))

	// MIN formula.
	ws.Cells().Set("B10", nil)
	minCell, _ := ws.Cells().Get("B10")
	minCell.SetFormula(fmt.Sprintf("MIN(%s)", totalRef))

	// Labels.
	ws.Cells().Set("A7", "TOTAL")
	ws.Cells().Set("A8", "AVERAGE")
	ws.Cells().Set("A9", "MAX")
	ws.Cells().Set("A10", "MIN")

	// ---- Evaluate formulas with the engine ----
	fmt.Println("Evaluating formulas …")
	fmt.Println()

	for _, row := range []int{7, 8, 9, 10} {
		cell, _ := ws.Cells().Get(fmt.Sprintf("B%d", row))
		formula := cell.GetFormula()
		result, err := cells_foss.CalculateFormula(formula, ws)
		if err != nil {
			fmt.Printf("  %s = ERROR: %v\n", formula, err)
		} else {
			fmt.Printf("  %s = %v\n", formula, result)
		}
	}

	// ---- Cross-sheet references ----
	summary, err := wb.AddWorksheet("Summary")
	if err != nil {
		fmt.Printf("AddWorksheet: %v\n", err)
		return
	}
	summary.Cells().Set("A1", "TOTAL (from Sheet1)")
	summary.Cells().Set("B1", nil)
	crossCell, _ := summary.Cells().Get("B1")
	crossCell.SetFormula("SUM(Sheet1!B2:B6)")
	crossResult, err := cells_foss.CalculateFormula(crossCell.GetFormula(), summary)
	if err != nil {
		fmt.Printf("  cross-sheet: %v\n", err)
	} else {
		fmt.Printf("  %s = %v (cross-sheet)\n", crossCell.GetFormula(), crossResult)
	}

	// ---- Error values follow Excel semantics ----
	// Division by zero yields the #DIV/0! error value, not a failure, and
	// error values can be stored in cells like any other value.
	divResult, err := cells_foss.CalculateFormula("1/0", ws)
	fmt.Printf("  1/0 = %v (error value, err=%v)\n", divResult, err)
	ws.Cells().Set("A11", "DIV ERROR")
	ws.Cells().Set("B11", divResult)

	outPath := "outputfiles/formula.xlsx"
	if err := wb.Save(outPath); err != nil {
		fmt.Printf("\nError saving: %v\n", err)
		return
	}
	fmt.Printf("\nWorkbook saved to %s\n", outPath)

	// Error values round-trip: they are stored as t="e" cells.
	if loaded, err := cells_foss.LoadWorkbook(outPath); err == nil {
		if c, err := loaded.Worksheets[0].Cells().Get("B11"); err == nil {
			fmt.Printf("Reloaded B11: %v (%T)\n", c.Value, c.Value)
		}
	}
	fmt.Println("Open in Excel — formulas will auto-calculate.")
}
