// Layout example: merged cells, column widths, row heights, frozen panes,
// autoFilter, and hyperlinks.
//
//	cd examples && go run ./layout/
package main

import (
	"fmt"
	"os"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func must(err error) {
	if err != nil {
		fmt.Fprintf(os.Stderr, "error: %v\n", err)
		os.Exit(1)
	}
}

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Name = "Report"
	details, err := wb.AddWorksheet("Details")
	must(err)
	details.Cells().Set("A1", "Full methodology lives here.")

	// Merged, tall title row.
	ws.Cells().Set("A1", "Quarterly Sales Report")
	must(ws.AddMergedCells("A1:D1"))
	must(ws.SetRowHeight(1, 26))

	title := cells_foss.NewStyle()
	title.Font.Bold = true
	title.Font.Size = 14
	c, err := ws.Cells().Get("A1")
	must(err)
	must(c.SetStyle(title))

	// Header row, then data.
	header := cells_foss.NewStyle()
	header.Font.Bold = true
	for i, h := range []string{"Region", "Product", "Units", "Revenue"} {
		ref := fmt.Sprintf("%s2", string(rune('A'+i)))
		ws.Cells().Set(ref, h)
		hc, err := ws.Cells().Get(ref)
		must(err)
		must(hc.SetStyle(header))
	}
	rows := []struct {
		region, product string
		units           float64
		revenue         float64
	}{
		{"North", "Widget", 120, 1299.99},
		{"South", "Gadget", 54, 24.5},
		{"East", "Doohickey", 310, 899},
	}
	for i, r := range rows {
		row := i + 3
		ws.Cells().Set(fmt.Sprintf("A%d", row), r.region)
		ws.Cells().Set(fmt.Sprintf("B%d", row), r.product)
		ws.Cells().Set(fmt.Sprintf("C%d", row), r.units)
		ws.Cells().Set(fmt.Sprintf("D%d", row), r.revenue)
	}

	// Column widths in Excel character units (default is 8.43).
	must(ws.SetColumnWidth("A", 14))
	must(ws.SetColumnWidth("B", 20))
	must(ws.SetColumnWidth("D", 12))

	// Keep the title and header rows plus the Region column in view.
	must(ws.FreezePanes(2, 1))

	// Filter buttons over the table body.
	must(ws.SetAutoFilter("A2:D5"))

	// Hyperlinks: one external URL, one jump inside the workbook. The cell
	// value is the visible label; the link itself is stored separately.
	ws.Cells().Set("E3", "Q1 report")
	must(ws.AddHyperlink("E3", "https://example.com/reports/q1?a=1&b=2"))
	ws.Cells().Set("E4", "Methodology")
	must(ws.AddInternalHyperlink("E4", "Details!A1"))
	must(ws.SetColumnWidth("E", 16))

	// Excel allows one link per cell; a second is refused.
	err = ws.AddHyperlink("E3", "https://example.com/other")
	if err == nil {
		fmt.Fprintln(os.Stderr, "unexpected: second hyperlink on E3 accepted")
		os.Exit(1)
	}
	fmt.Printf("Second link on E3 rejected: %v\n", err)

	path := "outputfiles/layout.xlsx"
	must(wb.Save(path))
	fmt.Printf("Saved %s\n", path)

	// Every layout property survives the save/load round trip.
	loaded, err := cells_foss.LoadWorkbook(path)
	must(err)
	lws := loaded.Worksheets[0]
	fmt.Printf("  merged:      %v\n", lws.MergedCells)
	fmt.Printf("  widths:      A=%v B=%v D=%v E=%v\n",
		lws.ColumnWidths["A"], lws.ColumnWidths["B"], lws.ColumnWidths["D"], lws.ColumnWidths["E"])
	fmt.Printf("  row height:  row 1 = %v\n", lws.RowHeights[1])
	fmt.Printf("  frozen:      rows=%d cols=%d\n", lws.FrozenRow, lws.FrozenCol)
	fmt.Printf("  autoFilter:  %s\n", lws.AutoFilterRange)
	for _, h := range lws.Hyperlinks {
		if h.Target != "" {
			fmt.Printf("  hyperlink:   %s → %s (external)\n", h.Ref, h.Target)
		} else {
			fmt.Printf("  hyperlink:   %s → %s (internal)\n", h.Ref, h.Location)
		}
	}
	fmt.Println("Open in Excel — the title spans A1:D1, the header stays frozen,")
	fmt.Println("filter buttons sit on row 2, and E3/E4 are clickable links.")
}
