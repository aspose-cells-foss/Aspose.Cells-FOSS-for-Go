package cells_foss

import "fmt"

// ChartType enumerates the supported chart types.
type ChartType string

// Supported chart types.
const (
	ChartTypeBar  ChartType = "bar"
	ChartTypeLine ChartType = "line"
	ChartTypePie  ChartType = "pie"
)

// Chart represents a chart embedded in a worksheet.
type Chart struct {
	// Type is the chart type: "bar", "line", or "pie".
	Type ChartType

	// Title is the chart title displayed above the chart.
	Title string

	// Series holds the data series to plot.
	Series []*ChartSeries

	// Anchor specifies the position of the chart on the worksheet.
	Anchor ChartAnchor

	// Width is the chart width in pixels.
	Width int

	// Height is the chart height in pixels.
	Height int

	// Name is the internal name used in drawing XML.
	Name string
}

// ChartSeries represents a single data series in a chart.
type ChartSeries struct {
	// Name is the series name displayed in the legend.
	Name string

	// Categories is the A1-style range for category labels (e.g., "A1:A10").
	Categories string

	// Values is the A1-style range for data values (e.g., "B1:B10").
	Values string
}

// ChartAnchor specifies the position of a chart on the worksheet.
type ChartAnchor struct {
	// Row is the 0-based row index of the anchor position.
	Row int

	// Col is the 0-based column index of the anchor position.
	Col int

	// RowOff is the vertical offset from the top of the anchor row, in EMUs.
	RowOff int64

	// ColOff is the horizontal offset from the left of the anchor column, in EMUs.
	ColOff int64
}

// AddChart adds a chart to the worksheet. The chart must have at least one
// series with valid data ranges.
//
//	chart := &cells_foss.Chart{
//	    Type:   cells_foss.ChartTypeBar,
//	    Title:  "Sales Data",
//	    Series: []*cells_foss.ChartSeries{
//	        {
//	            Name:       "Revenue",
//	            Categories: "A2:A10",
//	            Values:     "B2:B10",
//	        },
//	    },
//	    Anchor: cells_foss.ChartAnchor{Row: 0, Col: 2},
//	    Width:  400,
//	    Height: 300,
//	}
//	ws.AddChart(chart)
func (ws *Worksheet) AddChart(chart *Chart) error {
	if chart == nil {
		return fmt.Errorf("cells_foss: Chart is nil")
	}
	if len(chart.Series) == 0 {
		return fmt.Errorf("cells_foss: Chart must have at least one series")
	}
	if chart.Width <= 0 || chart.Height <= 0 {
		return fmt.Errorf("cells_foss: Chart width and height must be positive")
	}

	// Auto-assign a unique name.
	chart.Name = fmt.Sprintf("Chart %d", len(ws.Charts)+1)

	ws.Charts = append(ws.Charts, chart)
	ws.Modified = true
	if ws.cells != nil && ws.cells.wb != nil {
		ws.cells.wb.Modified = true
	}
	return nil
}

// RemoveChart removes the chart at the specified index.
func (ws *Worksheet) RemoveChart(index int) error {
	if index < 0 || index >= len(ws.Charts) {
		return fmt.Errorf("cells_foss: chart index %d out of range", index)
	}
	ws.Charts = append(ws.Charts[:index], ws.Charts[index+1:]...)
	ws.Modified = true
	if ws.cells != nil && ws.cells.wb != nil {
		ws.cells.wb.Modified = true
	}
	return nil
}
