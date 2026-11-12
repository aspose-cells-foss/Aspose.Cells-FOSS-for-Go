package cells_foss_test

import (
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// saveSheet saves the workbook and returns the path together with the raw
// sheet1 XML and its .rels part (empty string when absent).
func saveSheet(t *testing.T, wb *cells_foss.Workbook) (path, sheetXML, relsXML string) {
	t.Helper()
	path = filepath.Join(t.TempDir(), "layout.xlsx")
	if err := wb.Save(path); err != nil {
		t.Fatalf("Save: %v", err)
	}
	raw, err := cells_foss.ReadTestZipEntry(path, "xl/worksheets/sheet1.xml")
	if err != nil {
		t.Fatalf("read sheet1.xml: %v", err)
	}
	sheetXML = string(raw)
	if rels, err := cells_foss.ReadTestZipEntry(path, "xl/worksheets/_rels/sheet1.xml.rels"); err == nil {
		relsXML = string(rels)
	}
	return path, sheetXML, relsXML
}

func TestLayout_MergedCells(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Title")

	// Lower-case and reversed corners are normalized.
	if err := ws.AddMergedCells("b2:a1"); err != nil {
		t.Fatalf("AddMergedCells: %v", err)
	}
	if len(ws.MergedCells) != 1 || ws.MergedCells[0] != "A1:B2" {
		t.Fatalf("MergedCells = %v, want [A1:B2]", ws.MergedCells)
	}
	// Duplicates and overlaps are refused.
	if err := ws.AddMergedCells("A1:B2"); err == nil {
		t.Error("duplicate merge accepted")
	}
	if err := ws.AddMergedCells("B2:C3"); err == nil {
		t.Error("overlapping merge accepted")
	}
	// A single cell cannot be merged.
	if err := ws.AddMergedCells("D4"); err == nil {
		t.Error("single-cell merge accepted")
	}
	if err := ws.AddMergedCells("D4:D4"); err == nil {
		t.Error("degenerate merge accepted")
	}
	// Disjoint ranges are fine.
	if err := ws.AddMergedCells("D4:E5"); err != nil {
		t.Fatalf("disjoint merge: %v", err)
	}

	_, sheetXML, _ := saveSheet(t, wb)
	if !strings.Contains(sheetXML, `<mergeCells count="2">`) {
		t.Errorf("sheet XML missing mergeCells block:\n%s", sheetXML)
	}
	if !strings.Contains(sheetXML, `<mergeCell ref="A1:B2"`) || !strings.Contains(sheetXML, `<mergeCell ref="D4:E5"`) {
		t.Errorf("sheet XML missing mergeCell refs:\n%s", sheetXML)
	}
	// mergeCells follows sheetData and precedes conditionalFormatting per
	// the CT_Worksheet sequence (no CF here, so just check after sheetData).
	if strings.Index(sheetXML, "<mergeCells") < strings.Index(sheetXML, "<sheetData") {
		t.Errorf("mergeCells must follow sheetData:\n%s", sheetXML)
	}
}

func TestLayout_MergedCellsRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Title")
	if err := ws.AddMergedCells("A1:C1"); err != nil {
		t.Fatalf("AddMergedCells: %v", err)
	}
	path, _, _ := saveSheet(t, wb)

	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	ws2 := wb2.Worksheets[0]
	if len(ws2.MergedCells) != 1 || ws2.MergedCells[0] != "A1:C1" {
		t.Fatalf("reloaded MergedCells = %v, want [A1:C1]", ws2.MergedCells)
	}

	// Modify the sheet so its XML is regenerated: the merge must survive.
	ws2.Cells().Set("A2", "changed")
	path2 := filepath.Join(t.TempDir(), "layout2.xlsx")
	if err := wb2.Save(path2); err != nil {
		t.Fatalf("Save modified: %v", err)
	}
	wb3, err := cells_foss.LoadWorkbook(path2)
	if err != nil {
		t.Fatalf("LoadWorkbook modified: %v", err)
	}
	if len(wb3.Worksheets[0].MergedCells) != 1 || wb3.Worksheets[0].MergedCells[0] != "A1:C1" {
		t.Errorf("merge lost after modification: %v", wb3.Worksheets[0].MergedCells)
	}

	// RemoveMergedCells works and rejects unknown ranges.
	if err := wb3.Worksheets[0].RemoveMergedCells("a1:c1"); err != nil {
		t.Errorf("RemoveMergedCells: %v", err)
	}
	if err := wb3.Worksheets[0].RemoveMergedCells("A1:C1"); err == nil {
		t.Error("removing twice should fail")
	}
}

func TestLayout_ColumnWidthAndRowHeight(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "x")

	if err := ws.SetColumnWidth("b", 20); err != nil {
		t.Fatalf("SetColumnWidth: %v", err)
	}
	if err := ws.SetColumnWidth("A", 12.5); err != nil {
		t.Fatalf("SetColumnWidth A: %v", err)
	}
	// Row 3 has no cells but an explicit height.
	if err := ws.SetRowHeight(3, 30); err != nil {
		t.Fatalf("SetRowHeight: %v", err)
	}
	if err := ws.SetRowHeight(1, 22.5); err != nil {
		t.Fatalf("SetRowHeight 1: %v", err)
	}

	// Validation.
	if err := ws.SetColumnWidth("1A", 10); err == nil {
		t.Error("invalid column ref accepted")
	}
	if err := ws.SetColumnWidth("ABCD", 10); err == nil {
		t.Error("over-long column ref accepted")
	}
	if err := ws.SetColumnWidth("A", 0); err == nil {
		t.Error("zero width accepted")
	}
	if err := ws.SetColumnWidth("A", 300); err == nil {
		t.Error("width above 255 accepted")
	}
	if err := ws.SetRowHeight(0, 10); err == nil {
		t.Error("row 0 accepted")
	}
	if err := ws.SetRowHeight(2, 500); err == nil {
		t.Error("height above 409.5 accepted")
	}

	path, sheetXML, _ := saveSheet(t, wb)
	if !strings.Contains(sheetXML, `<col min="1" max="1" width="12.5" customWidth="1"`) {
		t.Errorf("sheet XML missing col A width:\n%s", sheetXML)
	}
	if !strings.Contains(sheetXML, `<col min="2" max="2" width="20" customWidth="1"`) {
		t.Errorf("sheet XML missing col B width:\n%s", sheetXML)
	}
	if !strings.Contains(sheetXML, `<row r="3" ht="30" customHeight="1"`) {
		t.Errorf("sheet XML missing cell-less row with height:\n%s", sheetXML)
	}
	if strings.Index(sheetXML, "<cols") > strings.Index(sheetXML, "<sheetData") {
		t.Errorf("cols must precede sheetData:\n%s", sheetXML)
	}

	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	ws2 := wb2.Worksheets[0]
	if w := ws2.ColumnWidths["A"]; w != 12.5 {
		t.Errorf("reloaded ColumnWidths[A] = %v, want 12.5", w)
	}
	if w := ws2.ColumnWidths["B"]; w != 20 {
		t.Errorf("reloaded ColumnWidths[B] = %v, want 20", w)
	}
	if h := ws2.RowHeights[3]; h != 30 {
		t.Errorf("reloaded RowHeights[3] = %v, want 30", h)
	}
	if h := ws2.RowHeights[1]; h != 22.5 {
		t.Errorf("reloaded RowHeights[1] = %v, want 22.5", h)
	}
}

func TestLayout_FreezePanes(t *testing.T) {
	cases := []struct {
		name     string
		row, col int
		wantPane string // "" means no <pane> expected
		wantRow  int
		wantCol  int
	}{
		{"none", 0, 0, "", 0, 0},
		{"rows only", 2, 0, `<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"`, 2, 0},
		{"cols only", 0, 1, `<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"`, 0, 1},
		{"both", 1, 1, `<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"`, 1, 1},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			wb := cells_foss.NewWorkbook()
			ws := wb.Worksheets[0]
			ws.Cells().Set("A1", "x")
			if err := ws.FreezePanes(tc.row, tc.col); err != nil {
				t.Fatalf("FreezePanes: %v", err)
			}
			path, sheetXML, _ := saveSheet(t, wb)

			if tc.wantPane == "" {
				if strings.Contains(sheetXML, "<pane") || strings.Contains(sheetXML, "<sheetViews") {
					t.Errorf("unexpected pane block:\n%s", sheetXML)
				}
			} else {
				if !strings.Contains(sheetXML, tc.wantPane) {
					t.Errorf("sheet XML missing %s:\n%s", tc.wantPane, sheetXML)
				}
				if strings.Index(sheetXML, "<sheetViews") > strings.Index(sheetXML, "<sheetData") {
					t.Errorf("sheetViews must precede sheetData:\n%s", sheetXML)
				}
			}

			wb2, err := cells_foss.LoadWorkbook(path)
			if err != nil {
				t.Fatalf("LoadWorkbook: %v", err)
			}
			ws2 := wb2.Worksheets[0]
			if ws2.FrozenRow != tc.wantRow || ws2.FrozenCol != tc.wantCol {
				t.Errorf("reloaded freeze = (%d,%d), want (%d,%d)", ws2.FrozenRow, ws2.FrozenCol, tc.wantRow, tc.wantCol)
			}
		})
	}

	wb := cells_foss.NewWorkbook()
	if err := wb.Worksheets[0].FreezePanes(-1, 0); err == nil {
		t.Error("negative freeze accepted")
	}
	if err := wb.Worksheets[0].FreezePanes(0, 16384); err == nil {
		t.Error("freeze of all columns accepted")
	}
}

func TestLayout_AutoFilter(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "h")
	ws.Cells().Set("C10", "v")

	if err := ws.SetAutoFilter("a1:c10"); err != nil {
		t.Fatalf("SetAutoFilter: %v", err)
	}
	if ws.AutoFilterRange != "A1:C10" {
		t.Fatalf("AutoFilterRange = %q, want A1:C10", ws.AutoFilterRange)
	}
	if err := ws.SetAutoFilter("A1:C10:X20"); err == nil {
		t.Error("invalid range accepted")
	}

	path, sheetXML, _ := saveSheet(t, wb)
	if !strings.Contains(sheetXML, `<autoFilter ref="A1:C10"`) {
		t.Errorf("sheet XML missing autoFilter:\n%s", sheetXML)
	}
	if strings.Index(sheetXML, "<autoFilter") < strings.Index(sheetXML, "</sheetData>") {
		t.Errorf("autoFilter must follow sheetData:\n%s", sheetXML)
	}

	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	if got := wb2.Worksheets[0].AutoFilterRange; got != "A1:C10" {
		t.Errorf("reloaded AutoFilterRange = %q", got)
	}

	// Clearing works.
	if err := wb2.Worksheets[0].SetAutoFilter(""); err != nil {
		t.Fatalf("SetAutoFilter clear: %v", err)
	}
	if wb2.Worksheets[0].AutoFilterRange != "" {
		t.Error("autoFilter not cleared")
	}
}

func TestLayout_Hyperlinks(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "docs")
	ws.Cells().Set("A2", "go to B5")

	const url = `https://example.com/report?a=1&b=2#frag`
	if err := ws.AddHyperlink("a1", url); err != nil {
		t.Fatalf("AddHyperlink: %v", err)
	}
	if err := ws.AddInternalHyperlink("A2", "Sheet1!B5"); err != nil {
		t.Fatalf("AddInternalHyperlink: %v", err)
	}
	// One link per cell.
	if err := ws.AddHyperlink("A1", "https://other.example.com"); err == nil {
		t.Error("second hyperlink on A1 accepted")
	}

	path, sheetXML, relsXML := saveSheet(t, wb)

	if !strings.Contains(sheetXML, `location="Sheet1!B5"`) {
		t.Errorf("sheet XML missing internal location:\n%s", sheetXML)
	}
	// encoding/xml binds the relationships namespace with its own prefix.
	if !strings.Contains(sheetXML, `relationships:id="rId1"`) {
		t.Errorf("sheet XML missing external r:id:\n%s", sheetXML)
	}
	// The URL's & must be escaped in the rels part, with TargetMode External.
	if !strings.Contains(relsXML, `Target="https://example.com/report?a=1&amp;b=2#frag"`) ||
		!strings.Contains(relsXML, `TargetMode="External"`) {
		t.Errorf("rels XML missing escaped external target:\n%s", relsXML)
	}
	if strings.Index(sheetXML, "<hyperlinks") < strings.Index(sheetXML, "</sheetData>") {
		t.Errorf("hyperlinks must follow sheetData:\n%s", sheetXML)
	}

	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	links := wb2.Worksheets[0].Hyperlinks
	if len(links) != 2 {
		t.Fatalf("reloaded %d hyperlinks, want 2", len(links))
	}
	if links[0].Ref != "A1" || links[0].Target != url {
		t.Errorf("external link = %+v, want A1 → %s", links[0], url)
	}
	if links[1].Ref != "A2" || links[1].Location != "Sheet1!B5" {
		t.Errorf("internal link = %+v, want A2 → Sheet1!B5", links[1])
	}

	// Removal.
	if err := wb2.Worksheets[0].RemoveHyperlink("A1"); err != nil {
		t.Errorf("RemoveHyperlink: %v", err)
	}
	if err := wb2.Worksheets[0].RemoveHyperlink("A1"); err == nil {
		t.Error("removing twice should fail")
	}
}

func TestLayout_HyperlinkRIDAfterDrawing(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", 1.0)
	ws.Cells().Set("B1", 2.0)

	// A chart forces a drawing part, which occupies rId1 in the sheet rels;
	// the hyperlink must then take rId2 in both the rels and the sheet XML.
	ch := &cells_foss.Chart{
		Type:   cells_foss.ChartTypeBar,
		Series: []*cells_foss.ChartSeries{{Name: "s", Categories: "A1:A1", Values: "B1:B1"}},
		Anchor: cells_foss.ChartAnchor{Row: 0, Col: 3},
		Width:  300, Height: 200,
	}
	if err := ws.AddChart(ch); err != nil {
		t.Fatalf("AddChart: %v", err)
	}
	if err := ws.AddHyperlink("A1", "https://example.com"); err != nil {
		t.Fatalf("AddHyperlink: %v", err)
	}

	_, sheetXML, relsXML := saveSheet(t, wb)
	if !strings.Contains(relsXML, `Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"`) {
		t.Errorf("drawing should hold rId1:\n%s", relsXML)
	}
	if !strings.Contains(relsXML, `Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"`) {
		t.Errorf("hyperlink should hold rId2:\n%s", relsXML)
	}
	if !strings.Contains(sheetXML, `relationships:id="rId2"`) {
		t.Errorf("sheet hyperlink should reference rId2:\n%s", sheetXML)
	}
}
