// AddWorksheet example: build a multi-sheet workbook by adding named
// worksheets.
//
//	cd examples && go run ./add_worksheet/
package main

import (
	"fmt"
	"os"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()

	// Rename the default sheet by editing Name directly, then add two more.
	wb.Worksheets[0].Name = "Data"

	summary, err := wb.AddWorksheet("Summary")
	if err != nil {
		fmt.Fprintf(os.Stderr, "AddWorksheet(Summary): %v\n", err)
		os.Exit(1)
	}
	notes, err := wb.AddWorksheet("Notes")
	if err != nil {
		fmt.Fprintf(os.Stderr, "AddWorksheet(Notes): %v\n", err)
		os.Exit(1)
	}

	data := wb.Worksheets[0]
	data.Cells().Set("A1", "Month")
	data.Cells().Set("B1", "Revenue")
	for i, m := range []struct {
		month string
		value float64
	}{{"Jan", 1200}, {"Feb", 1450}, {"Mar", 980}} {
		data.Cells().Set(fmt.Sprintf("A%d", i+2), m.month)
		data.Cells().Set(fmt.Sprintf("B%d", i+2), m.value)
	}

	summary.Cells().Set("A1", "Total")
	summary.Cells().Set("B1", 3630.0)
	notes.Cells().Set("A1", "Figures are illustrative.")

	// Names must be unique and Excel-legal; duplicates are rejected.
	if _, err := wb.AddWorksheet("data"); err == nil {
		fmt.Fprintln(os.Stderr, "unexpected: duplicate sheet name accepted")
		os.Exit(1)
	} else {
		fmt.Printf("Duplicate name rejected: %v\n", err)
	}

	path := "outputfiles/add_worksheet.xlsx"
	if err := wb.Save(path); err != nil {
		fmt.Fprintf(os.Stderr, "Save: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Saved %s with %d sheet(s)\n", path, len(wb.Worksheets))

	// All sheets and their content survive a save/load round trip.
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Load: %v\n", err)
		os.Exit(1)
	}
	for _, ws := range loaded.Worksheets {
		a1, err := ws.Cells().Get("A1")
		if err != nil {
			fmt.Printf("  reloaded sheet %q (empty)\n", ws.Name)
			continue
		}
		fmt.Printf("  reloaded sheet %q: A1=%q\n", ws.Name, a1.AsString())
	}
}
