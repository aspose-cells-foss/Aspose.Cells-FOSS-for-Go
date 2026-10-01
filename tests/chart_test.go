package cells_foss_test

import (
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestChart_AddAndRemove(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	chart := &cells_foss.Chart{
		Type:  cells_foss.ChartTypeBar,
		Title: "Test Chart",
		Series: []*cells_foss.ChartSeries{
			{
				Name:       "Series1",
				Categories: "A1:A5",
				Values:     "B1:B5",
			},
		},
		Width:  400,
		Height: 300,
	}

	if err := ws.AddChart(chart); err != nil {
		t.Fatalf("AddChart: %v", err)
	}

	if chart.Name != "Chart 1" {
		t.Errorf("Name = %q, want 'Chart 1'", chart.Name)
	}

	if len(ws.Charts) != 1 {
		t.Errorf("len(Charts) = %d, want 1", len(ws.Charts))
	}

	if err := ws.RemoveChart(0); err != nil {
		t.Fatalf("RemoveChart: %v", err)
	}

	if len(ws.Charts) != 0 {
		t.Errorf("len(Charts) after remove = %d, want 0", len(ws.Charts))
	}
}

func TestChart_Validation(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Nil chart.
	if err := ws.AddChart(nil); err == nil {
		t.Error("nil chart should error")
	}

	// No series.
	chart := &cells_foss.Chart{Type: cells_foss.ChartTypeBar}
	if err := ws.AddChart(chart); err == nil {
		t.Error("chart with no series should error")
	}

	// Invalid dimensions.
	chart = &cells_foss.Chart{
		Type:   cells_foss.ChartTypeBar,
		Series: []*cells_foss.ChartSeries{{Name: "S1"}},
		Width:  0,
		Height: 300,
	}
	if err := ws.AddChart(chart); err == nil {
		t.Error("chart with zero width should error")
	}
}

func TestChart_Types(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	types := []cells_foss.ChartType{
		cells_foss.ChartTypeBar,
		cells_foss.ChartTypeLine,
		cells_foss.ChartTypePie,
	}

	for _, ct := range types {
		chart := &cells_foss.Chart{
			Type:   ct,
			Series: []*cells_foss.ChartSeries{{Name: "S1", Categories: "A1:A5", Values: "B1:B5"}},
			Width:  400,
			Height: 300,
		}
		if err := ws.AddChart(chart); err != nil {
			t.Errorf("AddChart with type %q: %v", ct, err)
		}
	}

	if len(ws.Charts) != 3 {
		t.Errorf("len(Charts) = %d, want 3", len(ws.Charts))
	}
}

func TestChart_RemoveOutOfRange(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	if err := ws.RemoveChart(0); err == nil {
		t.Error("removing from empty charts should error")
	}

	chart := &cells_foss.Chart{
		Type:   cells_foss.ChartTypeBar,
		Series: []*cells_foss.ChartSeries{{Name: "S1"}},
		Width:  400,
		Height: 300,
	}
	ws.AddChart(chart)

	if err := ws.RemoveChart(5); err == nil {
		t.Error("removing out-of-range index should error")
	}
}
