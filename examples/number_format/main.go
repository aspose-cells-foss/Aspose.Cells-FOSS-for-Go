// Number formats example: apply built-in and custom number format codes to
// cells and store dates as Excel serial numbers.
//
//	cd examples && go run ./number_format/
package main

import (
	"fmt"
	"os"
	"time"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// setStyled writes a value and applies a style to it. The same *Style can be
// reused across cells: identical styles are deduplicated into one cellXfs
// entry when the workbook is saved.
func setStyled(ws *cells_foss.Worksheet, ref string, val interface{}, st *cells_foss.Style) {
	ws.Cells().Set(ref, val)
	c, err := ws.Cells().Get(ref)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Get(%s): %v\n", ref, err)
		os.Exit(1)
	}
	if err := c.SetStyle(st); err != nil {
		fmt.Fprintf(os.Stderr, "SetStyle(%s): %v\n", ref, err)
		os.Exit(1)
	}
}

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	ws.Cells().Set("A1", "Product")
	ws.Cells().Set("B1", "Price")
	ws.Cells().Set("C1", "Margin")
	ws.Cells().Set("D1", "Released")

	// Styles differ only in NumberFormat — which is part of style identity,
	// so each gets its own cellXfs entry.
	money := cells_foss.NewStyle()
	money.NumberFormat = `"$"#,##0.00` // custom code -> <numFmt> id 164

	percent := cells_foss.NewStyle()
	percent.NumberFormat = cells_foss.NumFmtPercent2 // "0.00%" -> built-in id 10

	date := cells_foss.NewStyle()
	date.NumberFormat = cells_foss.NumFmtDate // "yyyy-mm-dd" -> custom code

	rows := []struct {
		name     string
		price    float64
		margin   float64
		released time.Time
	}{
		{"Widget", 1299.99, 0.18, time.Date(2023, 5, 12, 0, 0, 0, 0, time.UTC)},
		{"Gadget", 24.5, 0.425, time.Date(2024, 1, 8, 0, 0, 0, 0, time.UTC)},
		{"Doohickey", 899, 0.07, time.Date(2024, 11, 30, 0, 0, 0, 0, time.UTC)},
	}
	for i, r := range rows {
		row := i + 2
		ws.Cells().Set(fmt.Sprintf("A%d", row), r.name)
		setStyled(ws, fmt.Sprintf("B%d", row), r.price, money)
		setStyled(ws, fmt.Sprintf("C%d", row), r.margin, percent)

		// Dates are stored as Excel serial numbers: whole days since
		// 1899-12-30 plus a fraction for the time of day. The number
		// format makes Excel render the serial as a date.
		serial, err := cells_foss.TimeToSerial(r.released)
		if err != nil {
			fmt.Fprintf(os.Stderr, "TimeToSerial: %v\n", err)
			os.Exit(1)
		}
		setStyled(ws, fmt.Sprintf("D%d", row), serial, date)
	}

	path := "outputfiles/number_format.xlsx"
	if err := wb.Save(path); err != nil {
		fmt.Fprintf(os.Stderr, "Save: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Saved %s\n", path)

	// Values and format codes survive a save/load round trip.
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Load: %v\n", err)
		os.Exit(1)
	}
	lws := loaded.Worksheets[0]
	for _, ref := range []string{"B2", "C2", "D2"} {
		c, err := lws.Cells().Get(ref)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Get(%s): %v\n", ref, err)
			os.Exit(1)
		}
		fmt.Printf("  %s: value=%v format=%q\n", ref, c.AsString(), c.GetStyle().NumberFormat)
	}

	// Convert a loaded serial back to a Go time.
	d2, _ := lws.Cells().Get("D2")
	serial, _ := d2.AsFloat64()
	when, err := cells_foss.SerialToTime(serial)
	if err != nil {
		fmt.Fprintf(os.Stderr, "SerialToTime: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("  D2 serial %v = %s\n", serial, when.Format(time.RFC3339))
	fmt.Println("Open in Excel — prices show $1,299.99, margins 18.00%, dates 2023-05-12.")
}
