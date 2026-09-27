// Example: typed cell accessors.
//
// After loading a workbook, Cell.Value may hold float64 (numeric cells),
// string (shared-string cells), or bool (boolean cells). Direct type
// assertions like cell.Value.(string) can panic for numeric cells.
//
// Use the typed accessors AsFloat64(), AsInt(), AsString() for safe access.
package main

import (
	"fmt"
	"log"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	c := ws.Cells()

	// Populate cells with different types.
	c.Set("A1", "Product")     // string
	c.Set("B1", "Price")      // string
	c.Set("A2", "Widget")     // string
	c.Set("B2", float64(12))  // numeric (int-valued float64)
	c.Set("A3", "Gadget")     // string
	c.Set("B3", float64(29))  // numeric
	c.Set("A4", "In stock")   // string
	c.Set("B4", true)         // boolean

	// Save and reload to demonstrate that loaded numeric cells come back
	// as float64, not string.
	outPath := "outputfiles/typed_access.xlsx"
	if err := wb.Save(outPath); err != nil {
		log.Fatal(err)
	}
	fmt.Println("Saved workbook to", outPath)

	loaded, err := cells_foss.LoadWorkbook(outPath)
	if err != nil {
		log.Fatal(err)
	}
	lc := loaded.Worksheets[0].Cells()

	// AsFloat64: converts numeric cells, numeric strings, and booleans.
	fmt.Println("\n--- AsFloat64 ---")
	for _, ref := range []string{"B2", "B3", "B4"} {
		cell, _ := lc.Get(ref)
		if v, ok := cell.AsFloat64(); ok {
			fmt.Printf("  %s = %.2f\n", ref, v)
		} else {
			fmt.Printf("  %s: not numeric\n", ref)
		}
	}

	// AsInt: truncates to integer.
	fmt.Println("\n--- AsInt ---")
	cell, _ := lc.Get("B2")
	if v, ok := cell.AsInt(); ok {
		fmt.Printf("  B2 as int = %d (type-safe, no assertion needed)\n", v)
	}

	// AsString: formats any value as its display string.
	fmt.Println("\n--- AsString ---")
	for _, ref := range []string{"A2", "B2", "B4"} {
		cell, _ := lc.Get(ref)
		fmt.Printf("  %s = %q\n", ref, cell.AsString())
	}

	// Contrast with unsafe direct access: cell.Value.(string) panics for
	// numeric cells because Value is float64 after loading.
	fmt.Println("\n--- Why typed accessors matter ---")
	cell, _ = lc.Get("B2")
	fmt.Printf("  cell.Value has concrete type %T\n", cell.Value)
	fmt.Println("  cell.Value.(string) would panic here.")
	fmt.Println("  cell.AsString() safely returns the display form.")
}
