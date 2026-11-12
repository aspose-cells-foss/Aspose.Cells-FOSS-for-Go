// VBA macro storage example: attach a vbaProject.bin to a workbook and save
// it as .xlsm. The library is storage-only — it cannot create or edit VBA
// code. Real macro bytes come from an existing .xlsm file:
//
//	data, _ := os.ReadFile("path/to/vbaProject.bin") // extracted from an .xlsm
//	wb.SetVBAProject(data)
//
// This example uses placeholder bytes to demonstrate the API and round-trip
// behavior; Excel will not execute (or open) a fabricated VBA project.
//
//	cd examples && go run ./macro/
package main

import (
	"fmt"
	"os"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Workbook with attached VBA project")

	// Placeholder vbaProject.bin bytes. Replace with the real binary from an
	// existing macro-enabled workbook for a file Excel can actually open.
	placeholder := []byte("vbaProject-placeholder-not-a-real-macro")
	if err := wb.SetVBAProject(placeholder); err != nil {
		fmt.Fprintf(os.Stderr, "SetVBAProject: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Attached VBA project: %d bytes, HasVBAProject=%v\n",
		len(wb.GetVBAProject().Data), wb.HasVBAProject())

	// Macro-enabled workbooks must use the .xlsm extension.
	path := "outputfiles/macro.xlsm"
	if err := wb.Save(path); err != nil {
		fmt.Fprintf(os.Stderr, "Save: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Saved %s\n", path)

	// The VBA binary is preserved verbatim across save/load.
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Load: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Reloaded: HasVBAProject=%v, %d bytes preserved\n",
		loaded.HasVBAProject(), len(loaded.GetVBAProject().Data))

	// ClearVBAProject detaches the macro data; a following save writes a
	// plain workbook without vbaProject.bin.
	loaded.ClearVBAProject()
	clean := "outputfiles/macro_stripped.xlsx"
	if err := loaded.Save(clean); err != nil {
		fmt.Fprintf(os.Stderr, "Save stripped: %v\n", err)
		os.Exit(1)
	}
	reloaded, err := cells_foss.LoadWorkbook(clean)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Load stripped: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("After ClearVBAProject: HasVBAProject=%v (saved to %s)\n",
		reloaded.HasVBAProject(), clean)
}
