// Chart example: create bar, pie, and scatter charts from worksheet data,
// with axis titles and a legend.
//
//	cd examples && go run ./chart/
package main

import (
	"fmt"
	"os"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Source data: quarterly revenue by region, plus numeric pairs for the
	// scatter chart.
	ws.Cells().Set("A1", "Region")
	ws.Cells().Set("B1", "Revenue")
	ws.Cells().Set("C1", "Units")
	ws.Cells().Set("D1", "Margin")
	rows := []struct {
		region  string
		revenue float64
		units   float64
		margin  float64
	}{
		{"North", 12500, 3, 12.5},
		{"South", 9800, 7, 9.8},
		{"East", 15200, 5, 15.2},
		{"West", 7300, 9, 7.3},
	}
	for i, r := range rows {
		ws.Cells().Set(fmt.Sprintf("A%d", i+2), r.region)
		ws.Cells().Set(fmt.Sprintf("B%d", i+2), r.revenue)
		ws.Cells().Set(fmt.Sprintf("C%d", i+2), r.units)
		ws.Cells().Set(fmt.Sprintf("D%d", i+2), r.margin)
	}

	// Bar chart anchored to the right of the data. Series reference the
	// worksheet ranges by A1-style addresses; the data is not embedded.
	// Axis titles and a legend are optional additions.
	bar := &cells_foss.Chart{
		Type:  cells_foss.ChartTypeBar,
		Title: "Revenue by Region",
		Series: []*cells_foss.ChartSeries{
			{Name: "Revenue", Categories: "A2:A5", Values: "B2:B5"},
		},
		Anchor: cells_foss.ChartAnchor{Row: 0, Col: 5},
		Width:  480,
		Height: 320,
		XAxis:  &cells_foss.ChartAxis{Title: "Region"},
		YAxis:  &cells_foss.ChartAxis{Title: "USD"},
		Legend: &cells_foss.ChartLegend{Position: cells_foss.LegendBottom},
	}
	if err := ws.AddChart(bar); err != nil {
		fmt.Fprintf(os.Stderr, "AddChart (bar): %v\n", err)
		os.Exit(1)
	}

	// Pie chart below the bar chart. Pie charts have no axes.
	pie := &cells_foss.Chart{
		Type:  cells_foss.ChartTypePie,
		Title: "Revenue Share",
		Series: []*cells_foss.ChartSeries{
			{Name: "Share", Categories: "A2:A5", Values: "B2:B5"},
		},
		Anchor: cells_foss.ChartAnchor{Row: 18, Col: 5},
		Width:  480,
		Height: 320,
		Legend: &cells_foss.ChartLegend{Position: cells_foss.LegendRight},
	}
	if err := ws.AddChart(pie); err != nil {
		fmt.Fprintf(os.Stderr, "AddChart (pie): %v\n", err)
		os.Exit(1)
	}

	// Scatter chart below the pie. For scatter charts, Categories supplies
	// the numeric X values and Values the Y values.
	scatter := &cells_foss.Chart{
		Type:  cells_foss.ChartTypeScatter,
		Title: "Units vs Margin",
		Series: []*cells_foss.ChartSeries{
			{Name: "Regions", Categories: "C2:C5", Values: "D2:D5"},
		},
		Anchor: cells_foss.ChartAnchor{Row: 36, Col: 5},
		Width:  480,
		Height: 320,
		XAxis:  &cells_foss.ChartAxis{Title: "Units sold"},
		YAxis:  &cells_foss.ChartAxis{Title: "Margin %"},
		Legend: &cells_foss.ChartLegend{Position: cells_foss.LegendTop},
	}
	if err := ws.AddChart(scatter); err != nil {
		fmt.Fprintf(os.Stderr, "AddChart (scatter): %v\n", err)
		os.Exit(1)
	}

	path := "outputfiles/chart.xlsx"
	if err := wb.Save(path); err != nil {
		fmt.Fprintf(os.Stderr, "Save: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Saved %s with %d chart(s)\n", path, len(ws.Charts))

	// Charts survive a save/load round trip, including axes and legend.
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Load: %v\n", err)
		os.Exit(1)
	}
	for _, ch := range loaded.Worksheets[0].Charts {
		fmt.Printf("  reloaded chart: type=%s title=%q series=%d anchor=(%d,%d) %dx%d\n",
			ch.Type, ch.Title, len(ch.Series), ch.Anchor.Row, ch.Anchor.Col, ch.Width, ch.Height)
		if ch.XAxis != nil || ch.YAxis != nil {
			xt, yt := "", ""
			if ch.XAxis != nil {
				xt = ch.XAxis.Title
			}
			if ch.YAxis != nil {
				yt = ch.YAxis.Title
			}
			fmt.Printf("    axes: x=%q y=%q\n", xt, yt)
		}
		if ch.Legend != nil {
			fmt.Printf("    legend: position=%s overlay=%v\n", ch.Legend.Position, ch.Legend.Overlay)
		}
	}
	fmt.Println("Open in Excel — bar, pie, and scatter charts appear next to the data.")
}
