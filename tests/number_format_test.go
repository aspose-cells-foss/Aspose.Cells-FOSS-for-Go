package cells_foss_test

import (
	"math"
	"path/filepath"
	"strings"
	"testing"
	"time"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// saveStyles saves the workbook to a fresh temp file and returns the path
// together with the raw xl/styles.xml content.
func saveStyles(t *testing.T, wb *cells_foss.Workbook) (string, string) {
	t.Helper()
	p := filepath.Join(t.TempDir(), "nf.xlsx")
	if err := wb.Save(p); err != nil {
		t.Fatalf("Save: %v", err)
	}
	raw, err := cells_foss.ReadTestZipEntry(p, "xl/styles.xml")
	if err != nil {
		t.Fatalf("read styles.xml: %v", err)
	}
	return p, string(raw)
}

func styledCell(t *testing.T, ws *cells_foss.Worksheet, ref string, val interface{}, code string) {
	t.Helper()
	ws.Cells().Set(ref, val)
	c, err := ws.Cells().Get(ref)
	if err != nil {
		t.Fatalf("Get(%s): %v", ref, err)
	}
	st := cells_foss.NewStyle()
	st.NumberFormat = code
	if err := c.SetStyle(st); err != nil {
		t.Fatalf("SetStyle(%s): %v", ref, err)
	}
}

func TestNumberFormat_BuiltinRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	styledCell(t, ws, "A1", 3.14159, cells_foss.NumFmtFloat2) // "0.00" -> id 2
	styledCell(t, ws, "B1", 0.25, cells_foss.NumFmtPercent)   // "0%" -> id 9
	styledCell(t, ws, "C1", 45000.0, "mm-dd-yy")              // builtin id 14

	path, styles := saveStyles(t, wb)

	for _, want := range []string{`numFmtId="2"`, `numFmtId="9"`, `numFmtId="14"`, `applyNumberFormat="1"`} {
		if !strings.Contains(styles, want) {
			t.Errorf("styles.xml missing %s:\n%s", want, styles)
		}
	}
	if strings.Contains(styles, "<numFmts") {
		t.Errorf("builtin codes must not create custom <numFmts> entries:\n%s", styles)
	}

	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	ws2 := wb2.Worksheets[0]
	for ref, want := range map[string]string{"A1": "0.00", "B1": "0%", "C1": "mm-dd-yy"} {
		c, err := ws2.Cells().Get(ref)
		if err != nil {
			t.Fatalf("Get(%s): %v", ref, err)
		}
		st := c.GetStyle()
		if st == nil {
			t.Fatalf("%s: GetStyle() = nil", ref)
		}
		if st.NumberFormat != want {
			t.Errorf("%s: NumberFormat = %q, want %q", ref, st.NumberFormat, want)
		}
	}
	if v, ok := mustGet(t, ws2, "A1").AsFloat64(); !ok || v != 3.14159 {
		t.Errorf("A1 value = %v (ok=%v), want 3.14159", v, ok)
	}
}

func TestNumberFormat_CustomRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	styledCell(t, ws, "A1", 45366.5, cells_foss.NumFmtDateTime) // custom "yyyy-mm-dd hh:mm:ss"
	styledCell(t, ws, "A2", 1234.5, `"$"#,##0.00`)              // second custom code

	path, styles := saveStyles(t, wb)

	if !strings.Contains(styles, `<numFmts count="2">`) {
		t.Errorf("styles.xml missing <numFmts count=\"2\">:\n%s", styles)
	}
	if !strings.Contains(styles, `numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"`) {
		t.Errorf("styles.xml missing custom numFmt 164:\n%s", styles)
	}
	if !strings.Contains(styles, `numFmtId="165"`) {
		t.Errorf("styles.xml missing custom numFmt 165:\n%s", styles)
	}
	// ECMA-376 orders numFmts ahead of fonts in the styleSheet.
	if strings.Index(styles, "<numFmts") > strings.Index(styles, "<fonts") {
		t.Errorf("numFmts must precede fonts in styleSheet:\n%s", styles)
	}

	// Both codes survive the round trip, including the one with embedded
	// quotes (escaped as &#34; in the XML attribute).
	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	for ref, want := range map[string]string{"A1": cells_foss.NumFmtDateTime, "A2": `"$"#,##0.00`} {
		st := mustGet(t, wb2.Worksheets[0], ref).GetStyle()
		if st == nil {
			t.Errorf("%s: GetStyle() = nil", ref)
			continue
		}
		if st.NumberFormat != want {
			t.Errorf("%s: NumberFormat = %q, want %q", ref, st.NumberFormat, want)
		}
	}
}

func TestNumberFormat_DedupAndDistinct(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	styledCell(t, ws, "A1", 1.0, cells_foss.NumFmtPercent2)
	styledCell(t, ws, "A2", 2.0, cells_foss.NumFmtPercent2) // identical style
	styledCell(t, ws, "A3", 3.0, cells_foss.NumFmtFloat2)   // differs only in NumberFormat

	a1, _ := ws.Cells().Get("A1")
	a2, _ := ws.Cells().Get("A2")
	a3, _ := ws.Cells().Get("A3")
	if a1.StyleID != a2.StyleID {
		t.Errorf("identical styles not deduplicated: A1 s=%d, A2 s=%d", a1.StyleID, a2.StyleID)
	}
	if a1.StyleID == a3.StyleID {
		t.Errorf("styles differing only in NumberFormat share StyleID %d", a1.StyleID)
	}
}

func TestNumberFormat_SurvivesLaterModification(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	styledCell(t, ws, "A1", 0.125, cells_foss.NumFmtPercent2)
	ws.Cells().Set("B1", "text")

	p := filepath.Join(t.TempDir(), "nf.xlsx")
	if err := wb.Save(p); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// Load, modify an unrelated cell (forces sheet XML regeneration and
	// styles.xml rebuild), save again. The number format must survive —
	// previously loadStyles dropped numFmtId and it was silently lost.
	wb2, err := cells_foss.LoadWorkbook(p)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	wb2.Worksheets[0].Cells().Set("B2", 99)
	p2 := filepath.Join(t.TempDir(), "nf2.xlsx")
	if err := wb2.Save(p2); err != nil {
		t.Fatalf("Save modified: %v", err)
	}

	wb3, err := cells_foss.LoadWorkbook(p2)
	if err != nil {
		t.Fatalf("LoadWorkbook modified: %v", err)
	}
	st := mustGet(t, wb3.Worksheets[0], "A1").GetStyle()
	if st == nil || st.NumberFormat != cells_foss.NumFmtPercent2 {
		t.Errorf("number format lost after modification: %+v", st)
	}
	if b2, ok := mustGet(t, wb3.Worksheets[0], "B2").AsInt(); !ok || b2 != 99 {
		t.Errorf("B2 = %v (ok=%v), want 99", b2, ok)
	}
}

func TestNumberFormat_ConditionalFormatDxf(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("B1", 90.0)
	ws.Cells().Set("B2", 10.0)

	// Rules carry styles with number formats but no cell ever calls
	// SetStyle: the styles part must still be written (regression: it used
	// to be skipped, leaving dangling dxfId refs in the sheet XML).
	cf := &cells_foss.ConditionalFormatting{
		Ref: "B1:B2",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeCellIs,
				Operator: cells_foss.CondOpGreaterThan,
				Formula:  "50",
				Priority: 1,
				Style:    &cells_foss.Style{NumberFormat: cells_foss.NumFmtFloat2},
			},
			{
				Type:     cells_foss.CondTypeCellIs,
				Operator: cells_foss.CondOpLessThan,
				Formula:  "50",
				Priority: 2,
				Style:    &cells_foss.Style{NumberFormat: cells_foss.NumFmtDate},
			},
		},
	}
	if err := ws.AddConditionalFormatting(cf); err != nil {
		t.Fatalf("AddConditionalFormatting: %v", err)
	}

	path, styles := saveStyles(t, wb)

	// Built-in code goes straight into the <dxf>; custom code also lands
	// in the workbook-wide <numFmts> table.
	if !strings.Contains(styles, `<numFmt numFmtId="2" formatCode="0.00"`) {
		t.Errorf("styles.xml missing dxf numFmt for builtin code:\n%s", styles)
	}
	if !strings.Contains(styles, `<numFmts count="1">`) || !strings.Contains(styles, `formatCode="yyyy-mm-dd"`) {
		t.Errorf("styles.xml missing custom numFmts entry:\n%s", styles)
	}
	if !strings.Contains(styles, "<dxfs count=\"2\">") {
		t.Errorf("styles.xml missing merged dxfs block:\n%s", styles)
	}

	// The two rules must reference distinct dxf ids in the sheet XML.
	sheetRaw, err := cells_foss.ReadTestZipEntry(path, "xl/worksheets/sheet1.xml")
	if err != nil {
		t.Fatalf("read sheet1.xml: %v", err)
	}
	sheet := string(sheetRaw)
	if !strings.Contains(sheet, `dxfId="0"`) || !strings.Contains(sheet, `dxfId="1"`) {
		t.Errorf("sheet XML missing distinct dxfId refs:\n%s", sheet)
	}

	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	cfs := wb2.Worksheets[0].ConditionalFormattings
	if len(cfs) != 1 || len(cfs[0].Rules) != 2 {
		t.Fatalf("reloaded %d conditional formatting(s)", len(cfs))
	}
}

func TestDateSerial_Conversions(t *testing.T) {
	cases := []struct {
		date   time.Time
		serial float64
	}{
		{time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC), 61},
		{time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC), 45292},
		{time.Date(2024, 1, 1, 12, 0, 0, 0, time.UTC), 45292.5},
		{time.Date(9999, 12, 31, 0, 0, 0, 0, time.UTC), 2958465},
	}
	for _, tc := range cases {
		got, err := cells_foss.TimeToSerial(tc.date)
		if err != nil {
			t.Fatalf("TimeToSerial(%s): %v", tc.date, err)
		}
		if got != tc.serial {
			t.Errorf("TimeToSerial(%s) = %v, want %v", tc.date.Format("2006-01-02 15:04"), got, tc.serial)
		}
		back, err := cells_foss.SerialToTime(tc.serial)
		if err != nil {
			t.Fatalf("SerialToTime(%v): %v", tc.serial, err)
		}
		if !back.Equal(tc.date) {
			t.Errorf("SerialToTime(%v) = %s, want %s", tc.serial, back, tc.date)
		}
	}

	// Wall-clock semantics: the location of the input is ignored.
	loc := time.FixedZone("X", 5*3600)
	got, err := cells_foss.TimeToSerial(time.Date(2024, 1, 1, 12, 0, 0, 0, loc))
	if err != nil || got != 45292.5 {
		t.Errorf("TimeToSerial with non-UTC zone = %v, %v; want 45292.5", got, err)
	}

	// The ambiguous low range (including the nonexistent 1900-02-29) is rejected.
	if _, err := cells_foss.TimeToSerial(time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC)); err == nil {
		t.Error("TimeToSerial(1900-02-28) should fail")
	}
	if _, err := cells_foss.TimeToSerial(time.Date(10000, 1, 1, 0, 0, 0, 0, time.UTC)); err == nil {
		t.Error("TimeToSerial(10000-01-01) should fail")
	}
	for _, bad := range []float64{60, 60.5, 59, 0, -1, math.NaN(), 2958466, 3000000} {
		if _, err := cells_foss.SerialToTime(bad); err == nil {
			t.Errorf("SerialToTime(%v) should fail", bad)
		}
	}
}

func TestDateSerial_CellRoundTrip(t *testing.T) {
	d := time.Date(2024, 3, 15, 13, 45, 0, 0, time.UTC)
	serial, err := cells_foss.TimeToSerial(d)
	if err != nil {
		t.Fatalf("TimeToSerial: %v", err)
	}

	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	styledCell(t, ws, "A1", serial, cells_foss.NumFmtDateTime)
	ws.Cells().Set("B1", "shipped")

	path, _ := saveStyles(t, wb)
	wb2, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	c := mustGet(t, wb2.Worksheets[0], "A1")
	v, ok := c.AsFloat64()
	if !ok || math.Abs(v-serial) > 1e-9 {
		t.Fatalf("A1 = %v (ok=%v), want %v", v, ok, serial)
	}
	back, err := cells_foss.SerialToTime(v)
	if err != nil {
		t.Fatalf("SerialToTime: %v", err)
	}
	if !back.Equal(d) {
		t.Errorf("round-tripped date = %s, want %s", back, d)
	}
	if st := c.GetStyle(); st == nil || st.NumberFormat != cells_foss.NumFmtDateTime {
		t.Errorf("date cell format = %+v, want %q", st, cells_foss.NumFmtDateTime)
	}
}

func mustGet(t *testing.T, ws *cells_foss.Worksheet, ref string) *cells_foss.Cell {
	t.Helper()
	c, err := ws.Cells().Get(ref)
	if err != nil {
		t.Fatalf("Get(%s): %v", ref, err)
	}
	return c
}
