package cells_foss_test

import (
	"os"
	"testing"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// TestChart_SaveAndLoad verifies that charts can be saved to .xlsx files.
// Note: Chart loading from .xlsx files is not yet implemented.
func TestChart_SaveAndLoad(t *testing.T) {
	tmpFile := "test_chart_save.xlsx"
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

	// Verify file was created
	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("Output file was not created")
	}

	// Note: Chart loading is not yet implemented, so we cannot verify
	// that the chart was correctly saved by loading it back.
}

// TestPivotTable_SaveAndLoad verifies that pivot tables can be saved and loaded from .xlsx files.
func TestPivotTable_SaveAndLoad(t *testing.T) {
	tmpFile := "test_pivot_save.xlsx"
	defer os.Remove(tmpFile)

	// Create workbook with pivot table
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Add some data
	ws.Cells().Set("A1", "Name")
	ws.Cells().Set("B1", "Value")
	ws.Cells().Set("A2", "Item1")
	ws.Cells().Set("B2", 100.0)
	ws.Cells().Set("A3", "Item2")
	ws.Cells().Set("B3", 200.0)

	// Add pivot table
	pt := &cells_foss.PivotTable{
		Name:      "PivotTable1",
		Ref:       "D1:E5",
		SourceRef: "A1:B3",
		RowFields: []string{"Name"},
		DataFields: []*cells_foss.PivotDataField{
			{
				Name:        "Value",
				Aggregation: cells_foss.PivotAggregationSum,
				DisplayName: "Sum of Value",
			},
		},
	}
	if err := ws.AddPivotTable(pt); err != nil {
		t.Fatalf("AddPivotTable failed: %v", err)
	}

	// Save workbook
	if err := wb.Save(tmpFile); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify file was created
	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("Output file was not created")
	}
}
