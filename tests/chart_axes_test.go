package cells_foss_test

import (
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// saveChartWorkbook saves a workbook holding one chart and returns the file
// path plus the raw xl/charts/chart1.xml part.
func saveChartWorkbook(t *testing.T, wb *cells_foss.Workbook, chart *cells_foss.Chart) (string, string) {
	t.Helper()
	if err := wb.Worksheets[0].AddChart(chart); err != nil {
		t.Fatalf("AddChart: %v", err)
	}
	path := filepath.Join(t.TempDir(), "chart.xlsx")
	if err := wb.Save(path); err != nil {
		t.Fatalf("Save: %v", err)
	}
	raw, err := cells_foss.ReadTestZipEntry(path, "xl/charts/chart1.xml")
	if err != nil {
		t.Fatalf("read chart part: %v", err)
	}
	return path, string(raw)
}

func loadFirstChart(t *testing.T, path string) *cells_foss.Chart {
	t.Helper()
	wb, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("Load: %v", err)
	}
	charts := wb.Worksheets[0].Charts
	if len(charts) != 1 {
		t.Fatalf("loaded %d charts, want 1", len(charts))
	}
	return charts[0]
}

func assertContains(t *testing.T, xmlStr, want, what string) {
	t.Helper()
	if !strings.Contains(xmlStr, want) {
		t.Errorf("%s: chart XML missing %s\n%s", what, want, xmlStr)
	}
}

// TestChart_AxesAndLegend covers axis titles, a hidden axis, legend
// placement, literal series names, and absolute references on a bar chart.
func TestChart_AxesAndLegend(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Quarter")
	ws.Cells().Set("B1", "Revenue")
	for i, v := range []float64{10, 20, 30, 40} {
		ws.Cells().Set("A"+string(rune('2'+i)), "Q")
		ws.Cells().Set("B"+string(rune('2'+i)), v)
	}

	chart := &cells_foss.Chart{
		Type:   cells_foss.ChartTypeBar,
		Title:  "Sales",
		Series: []*cells_foss.ChartSeries{{Name: "Revenue", Categories: "A2:A5", Values: "B2:B5"}},
		Width:  400,
		Height: 300,
		XAxis:  &cells_foss.ChartAxis{Title: "Quarter"},
		YAxis:  &cells_foss.ChartAxis{Title: "USD", Hidden: true},
		Legend: &cells_foss.ChartLegend{Position: cells_foss.LegendBottom},
	}
	path, xmlStr := saveChartWorkbook(t, wb, chart)

	// The bar group must carry the axId pair its axes reference.
	assertContains(t, xmlStr, `<c:axId val="111111111"/>`, "axId cat")
	assertContains(t, xmlStr, `<c:axId val="222222222"/>`, "axId val")
	assertContains(t, xmlStr, `<c:catAx>`, "catAx")
	assertContains(t, xmlStr, `<c:axPos val="b"/><c:title>`, "catAx title position")
	assertContains(t, xmlStr, `<a:t>Quarter</a:t>`, "catAx title text")
	// The hidden Y axis writes delete=1; the visible X axis writes delete=0.
	assertContains(t, xmlStr, `<c:delete val="1"/>`, "hidden valAx")
	assertContains(t, xmlStr, `<c:delete val="0"/>`, "visible catAx")
	assertContains(t, xmlStr, `<a:t>USD</a:t>`, "valAx title text")
	assertContains(t, xmlStr, `<c:legendPos val="b"/>`, "legend position")
	assertContains(t, xmlStr, `<c:overlay val="0"/></c:legend>`, "legend overlay")
	// Literal series names go to c:v, and data refs are absolute.
	assertContains(t, xmlStr, `<c:tx><c:v>Revenue</c:v></c:tx>`, "series name literal")
	assertContains(t, xmlStr, `'Sheet1'!$A$2:$A$5`, "absolute cat ref")
	assertContains(t, xmlStr, `'Sheet1'!$B$2:$B$5`, "absolute val ref")
	if strings.Contains(xmlStr, `<c:strRef><c:f>'Sheet1'!Revenue`) {
		t.Error("series name written as a bogus cell reference")
	}

	ch := loadFirstChart(t, path)
	if ch.Type != cells_foss.ChartTypeBar {
		t.Errorf("Type = %q, want bar", ch.Type)
	}
	if ch.Title != "Sales" {
		t.Errorf("Title = %q, want Sales", ch.Title)
	}
	if len(ch.Series) != 1 || ch.Series[0].Name != "Revenue" ||
		ch.Series[0].Categories != "A2:A5" || ch.Series[0].Values != "B2:B5" {
		t.Errorf("series round trip: %+v", ch.Series)
	}
	if ch.XAxis == nil || ch.XAxis.Title != "Quarter" || ch.XAxis.Hidden {
		t.Errorf("XAxis round trip: %+v", ch.XAxis)
	}
	if ch.YAxis == nil || ch.YAxis.Title != "USD" || !ch.YAxis.Hidden {
		t.Errorf("YAxis round trip: %+v", ch.YAxis)
	}
	if ch.Legend == nil || ch.Legend.Position != cells_foss.LegendBottom || ch.Legend.Overlay {
		t.Errorf("Legend round trip: %+v", ch.Legend)
	}
}

// TestChart_Scatter covers scatter charts: xVal/yVal references and two
// value axes distinguished by axPos.
func TestChart_Scatter(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	for i := 0; i < 4; i++ {
		ws.Cells().Set("A"+string(rune('2'+i)), float64(i+1))
		ws.Cells().Set("B"+string(rune('2'+i)), float64((i+1)*(i+1)))
	}

	chart := &cells_foss.Chart{
		Type:   cells_foss.ChartTypeScatter,
		Title:  "Trajectory",
		Series: []*cells_foss.ChartSeries{{Name: "Path", Categories: "A2:A5", Values: "B2:B5"}},
		Width:  400,
		Height: 300,
		XAxis:  &cells_foss.ChartAxis{Title: "Time"},
		YAxis:  &cells_foss.ChartAxis{Title: "Distance"},
		Legend: &cells_foss.ChartLegend{Position: cells_foss.LegendLeft, Overlay: true},
	}
	path, xmlStr := saveChartWorkbook(t, wb, chart)

	assertContains(t, xmlStr, `<c:scatterChart>`, "scatterChart")
	assertContains(t, xmlStr, `<c:scatterStyle val="lineMarker"/>`, "scatterStyle")
	assertContains(t, xmlStr, `<c:xVal><c:numRef><c:f>'Sheet1'!$A$2:$A$5</c:f>`, "xVal ref")
	assertContains(t, xmlStr, `<c:yVal><c:numRef><c:f>'Sheet1'!$B$2:$B$5</c:f>`, "yVal ref")
	assertContains(t, xmlStr, `<c:axPos val="b"/><c:title>`, "x axis position")
	assertContains(t, xmlStr, `<a:t>Time</a:t>`, "x axis title")
	assertContains(t, xmlStr, `<c:axPos val="l"/><c:title>`, "y axis position")
	assertContains(t, xmlStr, `<a:t>Distance</a:t>`, "y axis title")
	assertContains(t, xmlStr, `<c:legendPos val="l"/><c:overlay val="1"/>`, "legend")
	if strings.Contains(xmlStr, `<c:catAx>`) {
		t.Error("scatter chart must not write a category axis")
	}

	ch := loadFirstChart(t, path)
	if ch.Type != cells_foss.ChartTypeScatter {
		t.Errorf("Type = %q, want scatter", ch.Type)
	}
	if len(ch.Series) != 1 || ch.Series[0].Name != "Path" ||
		ch.Series[0].Categories != "A2:A5" || ch.Series[0].Values != "B2:B5" {
		t.Errorf("series round trip: %+v", ch.Series)
	}
	if ch.XAxis == nil || ch.XAxis.Title != "Time" {
		t.Errorf("XAxis round trip: %+v", ch.XAxis)
	}
	if ch.YAxis == nil || ch.YAxis.Title != "Distance" {
		t.Errorf("YAxis round trip: %+v", ch.YAxis)
	}
	if ch.Legend == nil || ch.Legend.Position != cells_foss.LegendLeft || !ch.Legend.Overlay {
		t.Errorf("Legend round trip: %+v", ch.Legend)
	}
}

// TestChart_Area covers area charts, including multiple series.
func TestChart_Area(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	chart := &cells_foss.Chart{
		Type: cells_foss.ChartTypeArea,
		Series: []*cells_foss.ChartSeries{
			{Name: "S1", Categories: "A2:A5", Values: "B2:B5"},
			{Name: "S2", Categories: "A2:A5", Values: "C2:C5"},
		},
		Width:  400,
		Height: 300,
	}
	path, xmlStr := saveChartWorkbook(t, wb, chart)

	assertContains(t, xmlStr, `<c:areaChart>`, "areaChart")
	assertContains(t, xmlStr, `<c:grouping val="standard"/>`, "area grouping")
	assertContains(t, xmlStr, `<c:catAx>`, "default catAx")
	assertContains(t, xmlStr, `<c:valAx>`, "default valAx")
	if strings.Count(xmlStr, `<c:ser>`) != 2 {
		t.Errorf("want 2 series elements:\n%s", xmlStr)
	}

	ch := loadFirstChart(t, path)
	if ch.Type != cells_foss.ChartTypeArea {
		t.Errorf("Type = %q, want area", ch.Type)
	}
	if len(ch.Series) != 2 || ch.Series[1].Name != "S2" || ch.Series[1].Values != "C2:C5" {
		t.Errorf("series round trip: %+v", ch.Series)
	}
	// Default axes are recovered as present but unconfigured.
	if ch.XAxis == nil || ch.XAxis.Title != "" || ch.XAxis.Hidden {
		t.Errorf("XAxis round trip: %+v", ch.XAxis)
	}
	if ch.YAxis == nil || ch.YAxis.Title != "" || ch.YAxis.Hidden {
		t.Errorf("YAxis round trip: %+v", ch.YAxis)
	}
	if ch.Legend != nil {
		t.Errorf("Legend should stay nil when unset: %+v", ch.Legend)
	}
}

// TestChart_PieHasNoAxes verifies pie charts carry no axis elements and
// still keep only their first series.
func TestChart_PieHasNoAxes(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	chart := &cells_foss.Chart{
		Type: cells_foss.ChartTypePie,
		Series: []*cells_foss.ChartSeries{
			{Categories: "A2:A5", Values: "B2:B5"},
			{Categories: "A2:A5", Values: "C2:C5"},
		},
		Width:  400,
		Height: 300,
	}
	path, xmlStr := saveChartWorkbook(t, wb, chart)

	assertContains(t, xmlStr, `<c:pieChart>`, "pieChart")
	for _, banned := range []string{`<c:axId`, `<c:catAx>`, `<c:valAx>`} {
		if strings.Contains(xmlStr, banned) {
			t.Errorf("pie chart must not contain %s\n%s", banned, xmlStr)
		}
	}
	if strings.Count(xmlStr, `<c:ser>`) != 1 {
		t.Errorf("pie should keep one series:\n%s", xmlStr)
	}

	ch := loadFirstChart(t, path)
	if ch.Type != cells_foss.ChartTypePie {
		t.Errorf("Type = %q, want pie", ch.Type)
	}
	if ch.XAxis != nil || ch.YAxis != nil {
		t.Errorf("pie axes should be nil: %+v / %+v", ch.XAxis, ch.YAxis)
	}
}

// TestChart_LegendDefaults covers the empty-legend case and normalization
// of unknown positions.
func TestChart_LegendDefaults(t *testing.T) {
	for _, tc := range []struct {
		pos  cells_foss.LegendPosition
		want cells_foss.LegendPosition
	}{
		{"", cells_foss.LegendRight},
		{"corner", cells_foss.LegendRight}, // unknown values normalize
		{cells_foss.LegendTop, cells_foss.LegendTop},
	} {
		wb := cells_foss.NewWorkbook()
		chart := &cells_foss.Chart{
			Type:   cells_foss.ChartTypeLine,
			Series: []*cells_foss.ChartSeries{{Categories: "A2:A5", Values: "B2:B5"}},
			Width:  400,
			Height: 300,
			Legend: &cells_foss.ChartLegend{Position: tc.pos},
		}
		path, xmlStr := saveChartWorkbook(t, wb, chart)
		assertContains(t, xmlStr, `<c:legendPos val="`+string(tc.want)+`"/>`, "legendPos")

		ch := loadFirstChart(t, path)
		if ch.Legend == nil || ch.Legend.Position != tc.want {
			t.Errorf("pos %q: Legend round trip %+v, want %q", tc.pos, ch.Legend, tc.want)
		}
	}
}
