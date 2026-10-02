package cells_foss_test

import (
	"os"
	"testing"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// TestPivotTable_SaveAndLoadRoundTrip verifies that pivot tables can be saved and loaded back.
func TestPivotTable_SaveAndLoadRoundTrip(t *testing.T) {
	tmpFile := "test_pivot_roundtrip.xlsx"
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

	// Load workbook
	loaded, err := cells_foss.Load(tmpFile)
	if err != nil {
		t.Fatalf("Load failed: %v", err)
	}

	// Verify pivot table exists
	loadedWs := loaded.Worksheets[0]
	if len(loadedWs.PivotTables) != 1 {
		t.Errorf("Expected 1 pivot table, got %d", len(loadedWs.PivotTables))
		return
	}

	loadedPT := loadedWs.PivotTables[0]

	// Verify pivot table name
	if loadedPT.Name != "PivotTable1" {
		t.Errorf("Expected name 'PivotTable1', got '%s'", loadedPT.Name)
	}

	// Verify source ref
	if loadedPT.SourceRef != "A1:B3" {
		t.Errorf("Expected source ref 'A1:B3', got '%s'", loadedPT.SourceRef)
	}

	// Verify row fields
	if len(loadedPT.RowFields) != 1 {
		t.Errorf("Expected 1 row field, got %d", len(loadedPT.RowFields))
	}

	// Verify data fields
	if len(loadedPT.DataFields) != 1 {
		t.Errorf("Expected 1 data field, got %d", len(loadedPT.DataFields))
	}
}
