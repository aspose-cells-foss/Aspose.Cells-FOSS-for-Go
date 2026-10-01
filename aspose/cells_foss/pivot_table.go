package cells_foss

import "fmt"

// PivotTable represents a pivot table on a worksheet.
type PivotTable struct {
	// Name is the display name of the pivot table.
	Name string

	// Ref is the A1-style range where the pivot table is placed (e.g., "F1").
	Ref string

	// SourceRef is the A1-style range of the source data (e.g., "A1:D100").
	SourceRef string

	// RowFields lists the field names to use as row labels.
	RowFields []string

	// ColFields lists the field names to use as column labels.
	ColFields []string

	// DataFields lists the value fields with aggregation settings.
	DataFields []*PivotDataField

	// internalID is assigned during save for relationship wiring.
	internalID int
}

// PivotDataField represents a value field in a pivot table with aggregation.
type PivotDataField struct {
	// Name is the source field name.
	Name string

	// Aggregation is the aggregation function: "sum", "count", "average",
	// "max", "min". Default is "sum".
	Aggregation string

	// DisplayName is the custom name shown in the pivot table. If empty,
	// Name is used.
	DisplayName string
}

// Aggregation function constants for pivot table data fields.
const (
	PivotAggregationSum     = "sum"
	PivotAggregationCount   = "count"
	PivotAggregationAverage = "average"
	PivotAggregationMax     = "max"
	PivotAggregationMin     = "min"
)

// AddPivotTable adds a pivot table to the worksheet. The pivot table must
// have a valid source range and at least one data field.
//
//	pt := &cells_foss.PivotTable{
//	    Name:      "SalesPivot",
//	    Ref:       "F1",
//	    SourceRef: "A1:D100",
//	    RowFields: []string{"Region"},
//	    DataFields: []*cells_foss.PivotDataField{
//	        {Name: "Revenue", Aggregation: cells_foss.PivotAggregationSum},
//	    },
//	}
//	ws.AddPivotTable(pt)
func (ws *Worksheet) AddPivotTable(pt *PivotTable) error {
	if pt == nil {
		return fmt.Errorf("cells_foss: PivotTable is nil")
	}
	if pt.Ref == "" {
		return fmt.Errorf("cells_foss: PivotTable ref is empty")
	}
	if pt.SourceRef == "" {
		return fmt.Errorf("cells_foss: PivotTable source ref is empty")
	}
	if len(pt.DataFields) == 0 {
		return fmt.Errorf("cells_foss: PivotTable must have at least one data field")
	}

	if pt.Name == "" {
		pt.Name = fmt.Sprintf("PivotTable %d", len(ws.PivotTables)+1)
	}

	ws.PivotTables = append(ws.PivotTables, pt)
	ws.Modified = true
	if ws.cells != nil && ws.cells.wb != nil {
		ws.cells.wb.Modified = true
	}
	return nil
}

// RemovePivotTable removes the pivot table with the given name.
func (ws *Worksheet) RemovePivotTable(name string) error {
	for i, pt := range ws.PivotTables {
		if pt.Name == name {
			ws.PivotTables = append(ws.PivotTables[:i], ws.PivotTables[i+1:]...)
			ws.Modified = true
			if ws.cells != nil && ws.cells.wb != nil {
				ws.cells.wb.Modified = true
			}
			return nil
		}
	}
	return fmt.Errorf("cells_foss: no PivotTable found with name %q", name)
}
