// Pivot table example: summarize transactional data with row/column fields.
//
//	cd examples && go run ./pivot_table/
package main

import (
	"fmt"
	"os"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Source data with a header row; the pivot field names must match the
	// headers exactly.
	ws.Cells().Set("A1", "Region")
	ws.Cells().Set("B1", "Product")
	ws.Cells().Set("C1", "Revenue")
	data := [][]interface{}{
		{"North", "Widget", 1200.0},
		{"South", "Widget", 900.0},
		{"North", "Gadget", 1600.0},
		{"East", "Widget", 2100.0},
		{"South", "Gadget", 750.0},
		{"East", "Gadget", 1300.0},
	}
	for i, row := range data {
		ws.Cells().Set(fmt.Sprintf("A%d", i+2), row[0])
		ws.Cells().Set(fmt.Sprintf("B%d", i+2), row[1])
		ws.Cells().Set(fmt.Sprintf("C%d", i+2), row[2])
	}

	// Pivot table placed at E1, summarizing Revenue by Region (rows) and
	// Product (columns).
	pt := &cells_foss.PivotTable{
		Name:      "SalesPivot",
		Ref:       "E1",
		SourceRef: "A1:C7",
		RowFields: []string{"Region"},
		ColFields: []string{"Product"},
		DataFields: []*cells_foss.PivotDataField{
			{Name: "Revenue", Aggregation: cells_foss.PivotAggregationSum, DisplayName: "Total Revenue"},
		},
	}
	if err := ws.AddPivotTable(pt); err != nil {
		fmt.Fprintf(os.Stderr, "AddPivotTable: %v\n", err)
		os.Exit(1)
	}

	path := "outputfiles/pivot_table.xlsx"
	if err := wb.Save(path); err != nil {
		fmt.Fprintf(os.Stderr, "Save: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Saved %s with pivot table %q (source %s)\n", path, pt.Name, pt.SourceRef)

	// Pivot tables survive a save/load round trip.
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Load: %v\n", err)
		os.Exit(1)
	}
	for _, p := range loaded.Worksheets[0].PivotTables {
		fmt.Printf("  reloaded pivot %q: ref=%s source=%s rows=%v cols=%v dataFields=%d\n",
			p.Name, p.Ref, p.SourceRef, p.RowFields, p.ColFields, len(p.DataFields))
	}
	fmt.Println("Open in Excel — the pivot table appears at E1; refresh it to see the sums.")
}
