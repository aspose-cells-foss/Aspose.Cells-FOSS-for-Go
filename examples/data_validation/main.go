// Data validation example: add a list-type validation rule to a range.
//
//	cd examples && go run ./data_validation/
package main

import (
	"fmt"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Set up header and data cells.
	ws.Cells().Set("A1", "Fruit")
	ws.Cells().Set("A2", "Apple")
	ws.Cells().Set("A3", "Banana")

	// Create a list-type data validation.
	dv := &cells_foss.DataValidation{
		Type:             cells_foss.DataValidationTypeList,
		Formula1:         `"Apple,Banana,Cherry,Dragonfruit"`,
		AllowBlank:       true,
		ShowErrorMessage: true,
		ErrorTitle:       "Invalid Fruit",
		ErrorMessage:     "Please pick a fruit from the list.",
		ErrorStyle:       cells_foss.ErrorStyleStop,
	}

	if err := ws.AddDataValidation("A2:A10", dv); err != nil {
		fmt.Printf("Error: %v\n", err)
		return
	}
	fmt.Printf("Added %q validation on A2:A10\n", dv.Type)

	// Add a second validation — whole number range. The Operator field
	// selects the comparison; empty means "between" (Excel's default).
	dv2 := &cells_foss.DataValidation{
		Type:             cells_foss.DataValidationTypeWhole,
		Operator:         cells_foss.DataValidationOpBetween,
		Formula1:         "1",
		Formula2:         "100",
		ShowErrorMessage: true,
		ErrorTitle:       "Invalid Value",
		ErrorMessage:     "Enter a whole number between 1 and 100.",
		ErrorStyle:       cells_foss.ErrorStyleWarning,
	}
	ws.Cells().Set("B1", "Score (1-100)")
	ws.AddDataValidation("B2:B10", dv2)
	fmt.Println("Added whole-number validation on B2:B10 (between 1 and 100)")

	// Add a third validation — decimals at or above zero, using a different
	// operator from the same family (greaterThan, lessThanOrEqual, …).
	dv3 := &cells_foss.DataValidation{
		Type:             cells_foss.DataValidationTypeDecimal,
		Operator:         cells_foss.DataValidationOpGreaterThanOrEqual,
		Formula1:         "0",
		ShowErrorMessage: true,
		ErrorTitle:       "Negative Value",
		ErrorMessage:     "Enter a margin of 0 or more.",
		ErrorStyle:       cells_foss.ErrorStyleStop,
	}
	ws.Cells().Set("C1", "Margin (>= 0)")
	ws.AddDataValidation("C2:C10", dv3)
	fmt.Println("Added decimal validation on C2:C10 (>= 0)")

	outPath := "outputfiles/data_validation.xlsx"
	if err := wb.Save(outPath); err != nil {
		fmt.Printf("Error saving workbook: %v\n", err)
		return
	}
	fmt.Printf("Workbook saved to %s\n", outPath)

	// Operators round-trip through save/load.
	loaded, err := cells_foss.LoadWorkbook(outPath)
	if err != nil {
		fmt.Printf("Error loading workbook: %v\n", err)
		return
	}
	for _, ldv := range loaded.Worksheets[0].DataValidations {
		fmt.Printf("  reloaded %s validation on %s, operator %q\n", ldv.Type, ldv.Ref, ldv.Operator)
	}
	fmt.Println("Open in Excel — A2:A10 shows a dropdown; B2:B10 accepts 1-100;")
	fmt.Println("C2:C10 rejects negative decimals.")
}
