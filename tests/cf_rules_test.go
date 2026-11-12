package cells_foss_test

import (
	"path/filepath"
	"strings"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// saveCF saves a workbook holding one conditional formatting and returns
// the file path plus the raw sheet1.xml.
func saveCF(t *testing.T, cf *cells_foss.ConditionalFormatting) (string, string) {
	t.Helper()
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	if err := ws.AddConditionalFormatting(cf); err != nil {
		t.Fatalf("AddConditionalFormatting: %v", err)
	}
	path := filepath.Join(t.TempDir(), "cf.xlsx")
	if err := wb.Save(path); err != nil {
		t.Fatalf("Save: %v", err)
	}
	raw, err := cells_foss.ReadTestZipEntry(path, "xl/worksheets/sheet1.xml")
	if err != nil {
		t.Fatalf("read sheet: %v", err)
	}
	return path, string(raw)
}

func loadCF(t *testing.T, path string) *cells_foss.ConditionalFormatting {
	t.Helper()
	wb, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("Load: %v", err)
	}
	cfs := wb.Worksheets[0].ConditionalFormattings
	if len(cfs) != 1 {
		t.Fatalf("loaded %d conditional formattings, want 1", len(cfs))
	}
	return cfs[0]
}

// TestCF_SqrefAttribute pins the spec-compliant sqref attribute; Excel
// addresses conditional formatting ranges through sqref, not ref.
func TestCF_SqrefAttribute(t *testing.T) {
	cf := &cells_foss.ConditionalFormatting{
		Ref: "A1:A10",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{Type: cells_foss.CondTypeCellIs, Operator: cells_foss.CondOpGreaterThan, Formula: "3", Priority: 1},
		},
	}
	_, sheet := saveCF(t, cf)
	if !strings.Contains(sheet, `<conditionalFormatting sqref="A1:A10">`) {
		t.Errorf("missing sqref attribute:\n%s", sheet)
	}
	if strings.Contains(sheet, `<conditionalFormatting ref=`) {
		t.Error("legacy ref attribute still written")
	}
}

func TestCF_ColorScale(t *testing.T) {
	cf := &cells_foss.ConditionalFormatting{
		Ref: "A1:A10",
		Rules: []*cells_foss.ConditionalFormattingRule{{
			Type:     cells_foss.CondTypeColorScale,
			Priority: 1,
			ColorScale: &cells_foss.ColorScale{
				Thresholds: []cells_foss.CondThreshold{
					{Type: cells_foss.CondValueMin},
					{Type: cells_foss.CondValuePercent, Val: "50"},
					{Type: cells_foss.CondValueMax},
				},
				Colors: []string{"FFF8696B", "FFFFEB84", "FF63BE7B"},
			},
		}},
	}
	path, sheet := saveCF(t, cf)

	for _, want := range []string{
		`<cfRule type="colorScale" priority="1">`,
		`<colorScale>`,
		`<cfvo type="min">`,
		`<cfvo type="percent" val="50">`,
		`<cfvo type="max">`,
		`<color rgb="FFF8696B">`,
		`<color rgb="FFFFEB84">`,
		`<color rgb="FF63BE7B">`,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("sheet XML missing %s\n%s", want, sheet)
		}
	}
	// Schema order: every cfvo precedes every color.
	if strings.Index(sheet, `<cfvo type="max">`) > strings.Index(sheet, `<color rgb=`) {
		t.Error("cfvo elements must precede color elements")
	}

	lcf := loadCF(t, path)
	if len(lcf.Rules) != 1 {
		t.Fatalf("loaded %d rules", len(lcf.Rules))
	}
	r := lcf.Rules[0]
	if r.Type != cells_foss.CondTypeColorScale || r.ColorScale == nil {
		t.Fatalf("rule round trip: %+v", r)
	}
	wantT := []cells_foss.CondThreshold{
		{Type: cells_foss.CondValueMin},
		{Type: cells_foss.CondValuePercent, Val: "50"},
		{Type: cells_foss.CondValueMax},
	}
	if len(r.ColorScale.Thresholds) != 3 || len(r.ColorScale.Colors) != 3 {
		t.Fatalf("colorScale round trip: %+v", r.ColorScale)
	}
	for i := range wantT {
		if r.ColorScale.Thresholds[i] != wantT[i] {
			t.Errorf("threshold %d = %+v, want %+v", i, r.ColorScale.Thresholds[i], wantT[i])
		}
	}
	if r.ColorScale.Colors[0] != "FFF8696B" || r.ColorScale.Colors[2] != "FF63BE7B" {
		t.Errorf("colors round trip: %v", r.ColorScale.Colors)
	}
}

func TestCF_DataBar(t *testing.T) {
	cf := &cells_foss.ConditionalFormatting{
		Ref: "B1:B10",
		Rules: []*cells_foss.ConditionalFormattingRule{{
			Type:     cells_foss.CondTypeDataBar,
			Priority: 1,
			DataBar: &cells_foss.DataBar{
				Min:       cells_foss.CondThreshold{Type: cells_foss.CondValueNum, Val: "0"},
				Max:       cells_foss.CondThreshold{Type: cells_foss.CondValueNum, Val: "100"},
				Color:     "FF638EC6",
				HideValue: true,
				MinLength: 20,
				MaxLength: 80,
			},
		}},
	}
	path, sheet := saveCF(t, cf)

	for _, want := range []string{
		`<dataBar minLength="20" maxLength="80" showValue="0">`,
		`<cfvo type="num" val="0">`,
		`<cfvo type="num" val="100">`,
		`<color rgb="FF638EC6">`,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("sheet XML missing %s\n%s", want, sheet)
		}
	}

	r := loadCF(t, path).Rules[0]
	if r.Type != cells_foss.CondTypeDataBar || r.DataBar == nil {
		t.Fatalf("rule round trip: %+v", r)
	}
	db := r.DataBar
	if db.Min.Type != cells_foss.CondValueNum || db.Min.Val != "0" ||
		db.Max.Type != cells_foss.CondValueNum || db.Max.Val != "100" {
		t.Errorf("endpoints round trip: %+v / %+v", db.Min, db.Max)
	}
	if db.Color != "FF638EC6" || !db.HideValue || db.MinLength != 20 || db.MaxLength != 80 {
		t.Errorf("dataBar round trip: %+v", db)
	}

	// Zero-value defaults: shown value, no explicit lengths.
	cf2 := &cells_foss.ConditionalFormatting{
		Ref: "C1:C10",
		Rules: []*cells_foss.ConditionalFormattingRule{{
			Type:     cells_foss.CondTypeDataBar,
			Priority: 1,
			DataBar: &cells_foss.DataBar{
				Min:   cells_foss.CondThreshold{Type: cells_foss.CondValueMin},
				Max:   cells_foss.CondThreshold{Type: cells_foss.CondValueMax},
				Color: "FF638EC6",
			},
		}},
	}
	path2, sheet2 := saveCF(t, cf2)
	if !strings.Contains(sheet2, `<dataBar showValue="1">`) {
		t.Errorf("default dataBar wrong:\n%s", sheet2)
	}
	if strings.Contains(sheet2, "minLength") {
		t.Error("zero MinLength must omit the attribute")
	}
	db2 := loadCF(t, path2).Rules[0].DataBar
	if db2 == nil || db2.HideValue || db2.MinLength != 0 || db2.MaxLength != 0 {
		t.Errorf("default dataBar round trip: %+v", db2)
	}
}

func TestCF_IconSet(t *testing.T) {
	cf := &cells_foss.ConditionalFormatting{
		Ref: "D1:D10",
		Rules: []*cells_foss.ConditionalFormattingRule{{
			Type:     cells_foss.CondTypeIconSet,
			Priority: 1,
			IconSet: &cells_foss.IconSet{
				IconSet: cells_foss.IconSet3Arrows,
				Thresholds: []cells_foss.CondThreshold{
					{Type: cells_foss.CondValuePercent, Val: "0"},
					{Type: cells_foss.CondValuePercent, Val: "33"},
					{Type: cells_foss.CondValuePercent, Val: "67"},
				},
			},
		}},
	}
	path, sheet := saveCF(t, cf)

	for _, want := range []string{
		`<iconSet iconSet="3Arrows" showValue="1" percent="1">`,
		`<cfvo type="percent" val="33">`,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("sheet XML missing %s\n%s", want, sheet)
		}
	}

	is := loadCF(t, path).Rules[0].IconSet
	if is == nil || is.IconSet != cells_foss.IconSet3Arrows || len(is.Thresholds) != 3 {
		t.Fatalf("iconSet round trip: %+v", is)
	}
	if is.Thresholds[2].Val != "67" || is.HideValue || is.Reverse {
		t.Errorf("iconSet fields round trip: %+v", is)
	}

	// Number thresholds flip percent off; HideValue/Reverse write 0/1.
	cf2 := &cells_foss.ConditionalFormatting{
		Ref: "E1:E10",
		Rules: []*cells_foss.ConditionalFormattingRule{{
			Type:     cells_foss.CondTypeIconSet,
			Priority: 1,
			IconSet: &cells_foss.IconSet{
				IconSet: cells_foss.IconSet4Rating,
				Thresholds: []cells_foss.CondThreshold{
					{Type: cells_foss.CondValueNum, Val: "0"},
					{Type: cells_foss.CondValueNum, Val: "2"},
					{Type: cells_foss.CondValueNum, Val: "4"},
					{Type: cells_foss.CondValueNum, Val: "6"},
				},
				HideValue: true,
				Reverse:   true,
			},
		}},
	}
	path2, sheet2 := saveCF(t, cf2)
	if !strings.Contains(sheet2, `<iconSet iconSet="4Rating" showValue="0" percent="0" reverse="1">`) {
		t.Errorf("numeric iconSet wrong:\n%s", sheet2)
	}
	is2 := loadCF(t, path2).Rules[0].IconSet
	if is2 == nil || !is2.HideValue || !is2.Reverse || len(is2.Thresholds) != 4 {
		t.Errorf("numeric iconSet round trip: %+v", is2)
	}
}

func TestCF_Top10(t *testing.T) {
	gold := &cells_foss.Style{Font: &cells_foss.Font{Bold: true, Color: "FF9C6500"}}
	cf := &cells_foss.ConditionalFormatting{
		Ref: "B2:B6",
		Rules: []*cells_foss.ConditionalFormattingRule{{
			Type:     cells_foss.CondTypeTop10,
			Priority: 1,
			Rank:     5,
			Bottom:   true,
			Percent:  true,
			Style:    gold,
		}},
	}
	path, sheet := saveCF(t, cf)

	if !strings.Contains(sheet, `<cfRule type="top10" priority="1" rank="5" bottom="1" percent="1" dxfId="0">`) {
		t.Errorf("top10 rule wrong:\n%s", sheet)
	}

	r := loadCF(t, path).Rules[0]
	if r.Type != cells_foss.CondTypeTop10 || r.Rank != 5 || !r.Bottom || !r.Percent {
		t.Errorf("top10 round trip: %+v", r)
	}
	if r.StyleID != 0 {
		t.Errorf("StyleID = %d, want restored dxfId 0", r.StyleID)
	}

	// Unset rank saves Excel's default of 10; no style means no dxfId.
	cf2 := &cells_foss.ConditionalFormatting{
		Ref: "B2:B6",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{Type: cells_foss.CondTypeTop10, Priority: 1},
		},
	}
	path2, sheet2 := saveCF(t, cf2)
	if !strings.Contains(sheet2, `<cfRule type="top10" priority="1" rank="10">`) {
		t.Errorf("default top10 wrong:\n%s", sheet2)
	}
	if strings.Contains(sheet2, "dxfId") {
		t.Error("style-less rule must not carry dxfId")
	}
	if r2 := loadCF(t, path2).Rules[0]; r2.Rank != 10 || r2.Bottom || r2.Percent {
		t.Errorf("default top10 round trip: %+v", r2)
	}
}

// TestCF_Validation covers the checks AddConditionalFormatting applies to
// the new rule types while leaving legacy types permissive.
func TestCF_Validation(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	valid3 := []cells_foss.CondThreshold{
		{Type: cells_foss.CondValueMin},
		{Type: cells_foss.CondValuePercent, Val: "50"},
		{Type: cells_foss.CondValueMax},
	}
	invalid := []struct {
		name string
		rule *cells_foss.ConditionalFormattingRule
	}{
		{"colorScale without config", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeColorScale}},
		{"colorScale one threshold", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeColorScale,
			ColorScale: &cells_foss.ColorScale{
				Thresholds: []cells_foss.CondThreshold{{Type: cells_foss.CondValueMin}},
				Colors:     []string{"FFF8696B"},
			}}},
		{"colorScale color mismatch", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeColorScale,
			ColorScale: &cells_foss.ColorScale{Thresholds: valid3, Colors: []string{"FFF8696B", "FF63BE7B"}}}},
		{"colorScale empty color", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeColorScale,
			ColorScale: &cells_foss.ColorScale{
				Thresholds: []cells_foss.CondThreshold{{Type: cells_foss.CondValueMin}, {Type: cells_foss.CondValueMax}},
				Colors:     []string{"FFF8696B", ""},
			}}},
		{"dataBar without config", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeDataBar}},
		{"dataBar without color", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeDataBar,
			DataBar: &cells_foss.DataBar{
				Min: cells_foss.CondThreshold{Type: cells_foss.CondValueMin},
				Max: cells_foss.CondThreshold{Type: cells_foss.CondValueMax},
			}}},
		{"dataBar length out of range", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeDataBar,
			DataBar: &cells_foss.DataBar{
				Min:   cells_foss.CondThreshold{Type: cells_foss.CondValueMin},
				Max:   cells_foss.CondThreshold{Type: cells_foss.CondValueMax},
				Color: "FF638EC6", MaxLength: 150,
			}}},
		{"threshold unknown type", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeDataBar,
			DataBar: &cells_foss.DataBar{
				Min:   cells_foss.CondThreshold{Type: "bogus"},
				Max:   cells_foss.CondThreshold{Type: cells_foss.CondValueMax},
				Color: "FF638EC6",
			}}},
		{"percent threshold without val", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeDataBar,
			DataBar: &cells_foss.DataBar{
				Min:   cells_foss.CondThreshold{Type: cells_foss.CondValuePercent},
				Max:   cells_foss.CondThreshold{Type: cells_foss.CondValueMax},
				Color: "FF638EC6",
			}}},
		{"iconSet without config", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeIconSet}},
		{"iconSet unknown name", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeIconSet,
			IconSet: &cells_foss.IconSet{IconSet: "6Stars", Thresholds: valid3}}},
		{"iconSet wrong threshold count", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeIconSet,
			IconSet: &cells_foss.IconSet{IconSet: cells_foss.IconSet3Arrows,
				Thresholds: valid3[:2]}}},
		{"top10 negative rank", &cells_foss.ConditionalFormattingRule{Type: cells_foss.CondTypeTop10, Rank: -1}},
	}
	for _, tc := range invalid {
		cf := &cells_foss.ConditionalFormatting{Ref: "A1:A10", Rules: []*cells_foss.ConditionalFormattingRule{tc.rule}}
		if err := ws.AddConditionalFormatting(cf); err == nil {
			t.Errorf("%s: expected an error", tc.name)
		}
	}

	valid := []*cells_foss.ConditionalFormattingRule{
		{Type: cells_foss.CondTypeColorScale, ColorScale: &cells_foss.ColorScale{
			Thresholds: []cells_foss.CondThreshold{{Type: cells_foss.CondValueMin}, {Type: cells_foss.CondValueMax}},
			Colors:     []string{"FFF8696B", "FF63BE7B"},
		}},
		{Type: cells_foss.CondTypeIconSet, IconSet: &cells_foss.IconSet{
			IconSet: cells_foss.IconSet5Quarters,
			Thresholds: []cells_foss.CondThreshold{
				{Type: cells_foss.CondValuePercent, Val: "0"},
				{Type: cells_foss.CondValuePercent, Val: "20"},
				{Type: cells_foss.CondValuePercent, Val: "40"},
				{Type: cells_foss.CondValuePercent, Val: "60"},
				{Type: cells_foss.CondValuePercent, Val: "80"},
			},
		}},
		{Type: cells_foss.CondTypeTop10, Rank: 3},
		{Type: cells_foss.CondTypeCellIs, Operator: cells_foss.CondOpEqual, Formula: "1"}, // legacy stays permissive
	}
	for _, rule := range valid {
		cf := &cells_foss.ConditionalFormatting{Ref: "A1:A10", Rules: []*cells_foss.ConditionalFormattingRule{rule}}
		if err := ws.AddConditionalFormatting(cf); err != nil {
			t.Errorf("valid %s rule rejected: %v", rule.Type, err)
		}
	}
}

// TestCF_LoadExcelStyleRules decodes rules the way Excel writes them:
// sqref, attribute defaults omitted, extra attributes we do not model.
func TestCF_LoadExcelStyleRules(t *testing.T) {
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="C1"><v>5</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="C2:C6">
    <cfRule type="dataBar" priority="1">
      <dataBar>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FF638EC6" tint="-0.25"/>
      </dataBar>
    </cfRule>
    <cfRule type="iconSet" priority="2">
      <iconSet>
        <cfvo type="percent" val="0"/>
        <cfvo type="percent" val="33"/>
        <cfvo type="percent" val="67"/>
      </iconSet>
    </cfRule>
    <cfRule type="top10" dxfId="0" priority="3" rank="2" bottom="1"/>
  </conditionalFormatting>
</worksheet>`
	path := filepath.Join(t.TempDir(), "excel_cf.xlsx")
	if err := cells_foss.WriteTestXLSX(path, sheetXML, ""); err != nil {
		t.Fatalf("WriteTestXLSX: %v", err)
	}
	wb, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("Load: %v", err)
	}
	cfs := wb.Worksheets[0].ConditionalFormattings
	if len(cfs) != 1 || cfs[0].Ref != "C2:C6" {
		t.Fatalf("loaded cfs: %+v", cfs)
	}
	if len(cfs[0].Rules) != 3 {
		t.Fatalf("loaded %d rules", len(cfs[0].Rules))
	}

	db := cfs[0].Rules[0]
	if db.Type != cells_foss.CondTypeDataBar || db.DataBar == nil {
		t.Fatalf("dataBar rule: %+v", db)
	}
	// showValue omitted means Excel's default: value shown.
	if db.DataBar.HideValue || db.DataBar.Color != "FF638EC6" {
		t.Errorf("dataBar decode: %+v", db.DataBar)
	}

	is := cfs[0].Rules[1].IconSet
	if is == nil || is.IconSet != "3TrafficLights1" || len(is.Thresholds) != 3 || is.HideValue {
		t.Errorf("iconSet decode (schema default name): %+v", is)
	}

	t10 := cfs[0].Rules[2]
	if t10.Type != cells_foss.CondTypeTop10 || t10.Rank != 2 || !t10.Bottom || t10.StyleID != 0 {
		t.Errorf("top10 decode: %+v", t10)
	}
}

// TestCF_LoadLegacyRefAttribute covers files written by older library
// versions, which used ref= instead of sqref=.
func TestCF_LoadLegacyRefAttribute(t *testing.T) {
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>5</v></c></row>
  </sheetData>
  <conditionalFormatting ref="A1:A10">
    <cfRule type="cellIs" operator="greaterThan" priority="1">
      <formula>3</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>`
	path := filepath.Join(t.TempDir(), "legacy_cf.xlsx")
	if err := cells_foss.WriteTestXLSX(path, sheetXML, ""); err != nil {
		t.Fatalf("WriteTestXLSX: %v", err)
	}
	wb, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("Load: %v", err)
	}
	cfs := wb.Worksheets[0].ConditionalFormattings
	if len(cfs) != 1 || cfs[0].Ref != "A1:A10" {
		t.Fatalf("legacy ref not recovered: %+v", cfs)
	}
	r := cfs[0].Rules[0]
	if r.Type != cells_foss.CondTypeCellIs || r.Operator != cells_foss.CondOpGreaterThan || r.Formula != "3" {
		t.Errorf("legacy rule decode: %+v", r)
	}
}
