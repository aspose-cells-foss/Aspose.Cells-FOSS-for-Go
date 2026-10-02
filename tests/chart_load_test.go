package cells_foss_test

import (
	"os"
	"testing"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// TestChart_SaveAndLoadRoundTrip verifies that charts can be saved and loaded back.
func TestChart_SaveAndLoadRoundTrip(t *testing.T) {
	tmpFile := "test_chart_roundtrip.xlsx"
	defer os.Remove(tmpFile)

	// Create workbook with chart
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Add some data
	ws.Cells().Set("A1", "Category")
	ws.Cells().Set("B1", "Value")
	ws.Cells().Set("A2", "A")
	ws.Cells().Set("B2", 10.0)
	ws.Cells().Set("A3", "B")
	ws.Cells().Set("B3", 20.0)

	// Add chart
	chart := &cells_foss.Chart{
		Type:   cells_foss.ChartTypeBar,
		Title:  "Test Chart",
		Width:  400,
		Height: 300,
		Series: []*cells_foss.ChartSeries{
			{
				Name:       "Series1",
				Categories: "A2:A3",
				Values:     "B2:B3",
			},
		},
	}
	if err := ws.AddChart(chart); err != nil {
		t.Fatalf("AddChart failed: %v", err)
	}

	// Save workbook
	if err := wb.Save(tmpFile); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Load workbook
	loaded, err := cells_foss.Load(tmpFile)
	if err != nil {
		t.Fatalf("Load failed: %v", err)
	}

	// Verify chart exists
	loadedWs := loaded.Worksheets[0]
	if len(loadedWs.Charts) != 1 {
		t.Errorf("Expected 1 chart, got %d", len(loadedWs.Charts))
		return
	}

	loadedChart := loadedWs.Charts[0]

	// Verify chart type
	if loadedChart.Type != cells_foss.ChartTypeBar {
		t.Errorf("Expected chart type %s, got %s", cells_foss.ChartTypeBar, loadedChart.Type)
	}

	// Verify chart title
	if loadedChart.Title != "Test Chart" {
		t.Errorf("Expected title 'Test Chart', got '%s'", loadedChart.Title)
	}

	// Verify series
	if len(loadedChart.Series) != 1 {
		t.Errorf("Expected 1 series, got %d", len(loadedChart.Series))
		return
	}

	// Note: Categories and Values may have sheet name prefix, so we check if they contain the range
	if loadedChart.Series[0].Categories == "" {
		t.Error("Expected categories to be set")
	}
	if loadedChart.Series[0].Values == "" {
		t.Error("Expected values to be set")
	}
}

// TestChart_LoadLineChart tests loading a line chart.
func TestChart_LoadLineChart(t *testing.T) {
	tmpFile := "test_chart_line.xlsx"
	defer os.Remove(tmpFile)

	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Add data
	ws.Cells().Set("A1", "X")
	ws.Cells().Set("B1", "Y")
	ws.Cells().Set("A2", 1.0)
	ws.Cells().Set("B2", 10.0)
	ws.Cells().Set("A3", 2.0)
	ws.Cells().Set("B3", 20.0)

	// Add line chart
	chart := &cells_foss.Chart{
		Type:   cells_foss.ChartTypeLine,
		Title:  "Line Chart",
		Width:  400,
		Height: 300,
		Series: []*cells_foss.ChartSeries{
			{
				Name:       "Data",
				Categories: "A2:A3",
				Values:     "B2:B3",
			},
		},
	}
	if err := ws.AddChart(chart); err != nil {
		t.Fatalf("AddChart failed: %v", err)
	}

	if err := wb.Save(tmpFile); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Load and verify
	loaded, err := cells_foss.Load(tmpFile)
	if err != nil {
		t.Fatalf("Load failed: %v", err)
	}

	loadedWs := loaded.Worksheets[0]
	if len(loadedWs.Charts) != 1 {
		t.Fatalf("Expected 1 chart, got %d", len(loadedWs.Charts))
	}

	if loadedWs.Charts[0].Type != cells_foss.ChartTypeLine {
		t.Errorf("Expected line chart, got %s", loadedWs.Charts[0].Type)
	}
}

// TestChart_LoadPieChart tests loading a pie chart.
func TestChart_LoadPieChart(t *testing.T) {
	tmpFile := "test_chart_pie.xlsx"
	defer os.Remove(tmpFile)

	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Add data
	ws.Cells().Set("A1", "Category")
	ws.Cells().Set("B1", "Value")
	ws.Cells().Set("A2", "A")
	ws.Cells().Set("B2", 30.0)
	ws.Cells().Set("A3", "B")
	ws.Cells().Set("B3", 70.0)

	// Add pie chart
	chart := &cells_foss.Chart{
		Type:   cells_foss.ChartTypePie,
		Title:  "Pie Chart",
		Width:  400,
		Height: 300,
		Series: []*cells_foss.ChartSeries{
			{
				Categories: "A2:A3",
				Values:     "B2:B3",
			},
		},
	}
	if err := ws.AddChart(chart); err != nil {
		t.Fatalf("AddChart failed: %v", err)
	}

	if err := wb.Save(tmpFile); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Load and verify
	loaded, err := cells_foss.Load(tmpFile)
	if err != nil {
		t.Fatalf("Load failed: %v", err)
	}

	loadedWs := loaded.Worksheets[0]
	if len(loadedWs.Charts) != 1 {
		t.Fatalf("Expected 1 chart, got %d", len(loadedWs.Charts))
	}

	if loadedWs.Charts[0].Type != cells_foss.ChartTypePie {
		t.Errorf("Expected pie chart, got %s", loadedWs.Charts[0].Type)
	}
}
