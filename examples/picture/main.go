// Picture example: create a workbook with an embedded image.
//
//	cd examples && go run ./picture/
package main

import (
	"fmt"

	"github.com/aspose/cells_foss/aspose/cells_foss"
)

// generateSmallPNG creates a minimal valid 1×1 PNG image.
func generateSmallPNG() []byte {
	return []byte{
		0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
		0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
		0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
		0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
		0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
		0x00, 0x00, 0x03, 0x00, 0x01, 0x42, 0x0A, 0xF1,
		0x45, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
		0x44, 0xAE, 0x42, 0x60, 0x82,
	}
}

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// ---- Populate some cells ----
	ws.Cells().Set("A1", "Product Catalog")
	ws.Cells().Set("A3", "Item")
	ws.Cells().Set("B3", "Price")
	ws.Cells().Set("A4", "Widget")
	ws.Cells().Set("B4", "$9.99")

	// ---- Embed a picture ----
	pic := cells_foss.NewPicture(generateSmallPNG(), "png")
	pic.Width = 100
	pic.Height = 80
	pic.SetAnchor(5, 1) // position at row 5, column B

	if err := ws.AddPicture(pic); err != nil {
		fmt.Printf("Error adding picture: %v\n", err)
		return
	}

	fmt.Printf("Added %q at row=%d, col=%d (%d×%d px)\n",
		pic.Name, pic.Row, pic.Col, pic.Width, pic.Height)

	outPath := "outputfiles/picture.xlsx"
	if err := wb.Save(outPath); err != nil {
		fmt.Printf("Error saving: %v\n", err)
		return
	}
	fmt.Printf("Workbook saved to %s\n", outPath)
	fmt.Println("Open in Excel — the image should appear at row 5, column B.")
}
