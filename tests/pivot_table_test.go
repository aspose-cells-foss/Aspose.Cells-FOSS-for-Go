package cells_foss_test

import (
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestPivotTable_AddAndRemove(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	pt := &cells_foss.PivotTable{
		Name:      "SalesPivot",
		Ref:       "F1",
		SourceRef: "A1:D100",
		RowFields: []string{"Region"},
		DataFields: []*cells_foss.PivotDataField{
			{Name: "Revenue", Aggregation: cells_foss.PivotAggregationSum},
		},
	}

	if err := ws.AddPivotTable(pt); err != nil {
		t.Fatalf("AddPivotTable: %v", err)
	}

	if len(ws.PivotTables) != 1 {
		t.Errorf("len(PivotTables) = %d, want 1", len(ws.PivotTables))
	}

	if ws.PivotTables[0].Name != "SalesPivot" {
		t.Errorf("Name = %q, want SalesPivot", ws.PivotTables[0].Name)
	}

	if err := ws.RemovePivotTable("SalesPivot"); err != nil {
		t.Fatalf("RemovePivotTable: %v", err)
	}

	if len(ws.PivotTables) != 0 {
		t.Errorf("len(PivotTables) after remove = %d, want 0", len(ws.PivotTables))
	}
}

func TestPivotTable_Validation(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Nil pivot table.
	if err := ws.AddPivotTable(nil); err == nil {
		t.Error("nil pivot table should error")
	}

	// Empty ref.
	pt := &cells_foss.PivotTable{SourceRef: "A1:D100"}
	if err := ws.AddPivotTable(pt); err == nil {
		t.Error("empty ref should error")
	}

	// Empty source ref.
	pt = &cells_foss.PivotTable{Ref: "F1"}
	if err := ws.AddPivotTable(pt); err == nil {
		t.Error("empty source ref should error")
	}

	// No data fields.
	pt = &cells_foss.PivotTable{Ref: "F1", SourceRef: "A1:D100"}
	if err := ws.AddPivotTable(pt); err == nil {
		t.Error("no data fields should error")
	}
}

func TestPivotTable_AutoName(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	pt := &cells_foss.PivotTable{
		Ref:       "F1",
		SourceRef: "A1:D100",
		DataFields: []*cells_foss.PivotDataField{
			{Name: "Revenue", Aggregation: cells_foss.PivotAggregationSum},
		},
	}
	ws.AddPivotTable(pt)

	if pt.Name != "PivotTable 1" {
		t.Errorf("auto-generated name = %q, want 'PivotTable 1'", pt.Name)
	}
}

func TestPivotTable_RemoveNotFound(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	if err := ws.RemovePivotTable("NonExistent"); err == nil {
		t.Error("removing non-existent pivot table should error")
	}
}

func TestPivotTable_AggregationTypes(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	aggregations := []string{
		cells_foss.PivotAggregationSum,
		cells_foss.PivotAggregationCount,
		cells_foss.PivotAggregationAverage,
		cells_foss.PivotAggregationMax,
		cells_foss.PivotAggregationMin,
	}

	for _, agg := range aggregations {
		pt := &cells_foss.PivotTable{
			Ref:       "F1",
			SourceRef: "A1:D100",
			DataFields: []*cells_foss.PivotDataField{
				{Name: "Value", Aggregation: agg},
			},
		}
		if err := ws.AddPivotTable(pt); err != nil {
			t.Errorf("AddPivotTable with aggregation %q: %v", agg, err)
		}
	}
}
