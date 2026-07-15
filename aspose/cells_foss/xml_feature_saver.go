package cells_foss

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strconv"
	"strings"
)

// ======================================================================
// Output XML types for styles.xml
// ======================================================================

type outStyleSheet struct {
	XMLName      xml.Name       `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	Fonts        outFonts       `xml:"fonts"`
	Fills        outFills       `xml:"fills"`
	Borders      outBorders     `xml:"borders"`
	CellStyleXfs outCellXfs     `xml:"cellStyleXfs"`
	CellXfs      outCellXfs     `xml:"cellXfs"`
}

type outFonts struct {
	Count int       `xml:"count,attr"`
	Fonts []outFont `xml:"font"`
}

type outFont struct {
	Sz    outVal    `xml:"sz"`
	Name  outVal    `xml:"name"`
	Color *outColor `xml:"color,omitempty"`
	B     *outEmpty `xml:"b,omitempty"`
	I     *outEmpty `xml:"i,omitempty"`
}

type outFills struct {
	Count int       `xml:"count,attr"`
	Fills []outFill `xml:"fill"`
}

type outFill struct {
	PatternFill *outPatternFill `xml:"patternFill"`
}

type outPatternFill struct {
	PatternType string    `xml:"patternType,attr"`
	FgColor     *outColor `xml:"fgColor,omitempty"`
}

type outBorders struct {
	Count   int         `xml:"count,attr"`
	Borders []outBorder `xml:"border"`
}

type outBorder struct {
	Left   *outBorderSide `xml:"left"`
	Right  *outBorderSide `xml:"right"`
	Top    *outBorderSide `xml:"top"`
	Bottom *outBorderSide `xml:"bottom"`
}

type outBorderSide struct {
	Style string `xml:"style,attr,omitempty"`
}

type outCellXfs struct {
	Count int     `xml:"count,attr"`
	Xfs   []outXf `xml:"xf"`
}

type outXf struct {
	NumFmtId       int       `xml:"numFmtId,attr"`
	FontId         int       `xml:"fontId,attr"`
	FillId         int       `xml:"fillId,attr"`
	BorderId       int       `xml:"borderId,attr"`
	XfId           int       `xml:"xfId,attr"`
	ApplyFont      int       `xml:"applyFont,attr,omitempty"`
	ApplyFill      int       `xml:"applyFill,attr,omitempty"`
	ApplyBorder    int       `xml:"applyBorder,attr,omitempty"`
	ApplyAlignment int       `xml:"applyAlignment,attr,omitempty"`
	Alignment      *outAlign `xml:"alignment,omitempty"`
}

type outAlign struct {
	Horizontal string `xml:"horizontal,attr,omitempty"`
	Vertical   string `xml:"vertical,attr,omitempty"`
	WrapText   string `xml:"wrapText,attr,omitempty"`
}

type outVal struct {
	Val string `xml:"val,attr"`
}

type outColor struct {
	RGB string `xml:"rgb,attr,omitempty"`
}

type outEmpty struct{}

// ======================================================================
// generateStylesXML
// ======================================================================

// generateStylesXML produces the content of xl/styles.xml from the
// Workbook's style registry.  Every Style is decomposed into its
// constituent parts (font, fill, border, alignment), the parts are
// deduplicated into the OOXML tables, and a <cellXfs> entry is emitted
// that references the correct table indices.
func generateStylesXML(wb *Workbook) string {
	if len(wb.styles) == 0 {
		return ""
	}

	// ---- Font table ----
	fontTable, fontIndex := buildFontTable(wb.styles)

	// ---- Fill table ----
	fillTable, fillIndex := buildFillTable(wb.styles)

	// ---- Border table ----
	borderTable, borderIndex := buildBorderTable(wb.styles)

	// ---- cellXfs ----
	xfCount := len(wb.styles)
	cellXfs := make([]outXf, xfCount)
	for i, st := range wb.styles {
		fID := fontIndex[fontKey(st.Font)]
		fiID := fillIndex[fillKey(st.Fill)]
		bID := borderIndex[borderKey(st.Border)]

		xf := outXf{
			NumFmtId: 0,
			FontId:   fID,
			FillId:   fiID,
			BorderId: bID,
			XfId:     0,
		}

		// Apply flags.
		if fID > 0 {
			xf.ApplyFont = 1
		}
		if fiID > 1 { // >1 because indices 0 and 1 are reserved
			xf.ApplyFill = 1
		}
		if bID > 0 {
			xf.ApplyBorder = 1
		}

		// Alignment.
		if st.Alignment != nil && (st.Alignment.Horizontal != "" || st.Alignment.Vertical != "" || st.Alignment.WrapText) {
			xf.ApplyAlignment = 1
			align := &outAlign{
				Horizontal: st.Alignment.Horizontal,
				Vertical:   st.Alignment.Vertical,
			}
			if st.Alignment.WrapText {
				align.WrapText = "1"
			}
			xf.Alignment = align
		}

		cellXfs[i] = xf
	}

	ss := outStyleSheet{
		Fonts:   outFonts{Count: len(fontTable), Fonts: fontTable},
		Fills:   outFills{Count: len(fillTable), Fills: fillTable},
		Borders: outBorders{Count: len(borderTable), Borders: borderTable},
		CellStyleXfs: outCellXfs{
			Count: 1,
			Xfs:   []outXf{{NumFmtId: 0, FontId: 0, FillId: 0, BorderId: 0, XfId: 0}},
		},
		CellXfs: outCellXfs{Count: xfCount, Xfs: cellXfs},
	}

	return marshalXML(ss)
}

// ======================================================================
// Table builders
// ======================================================================

func buildFontTable(styles []*Style) ([]outFont, map[string]int) {
	seen := make(map[string]int)
	var table []outFont

	for _, st := range styles {
		key := fontKey(st.Font)
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = len(table)

		f := outFont{
			Sz:   outVal{Val: strconv.FormatFloat(st.Font.Size, 'f', -1, 64)},
			Name: outVal{Val: st.Font.Name},
		}
		if st.Font.Bold {
			f.B = &outEmpty{}
		}
		if st.Font.Italic {
			f.I = &outEmpty{}
		}
		if st.Font.Color != "" && st.Font.Color != "FF000000" {
			f.Color = &outColor{RGB: st.Font.Color}
		}
		table = append(table, f)
	}

	if len(table) == 0 {
		// Default Calibri 11.
		table = append(table, outFont{
			Sz:   outVal{Val: "11"},
			Name: outVal{Val: "Calibri"},
		})
		seen[fontKey(DefaultStyle().Font)] = 0
	}

	return table, seen
}

func buildFillTable(styles []*Style) ([]outFill, map[string]int) {
	// Indices 0 and 1 are reserved in OOXML.
	table := []outFill{
		{PatternFill: &outPatternFill{PatternType: "none"}},
		{PatternFill: &outPatternFill{PatternType: "gray125"}},
	}
	seen := map[string]int{
		fillKey(&Fill{Type: FillTypeNone}):    0,
		fillKey(&Fill{Type: FillTypeGray125}): 1,
	}

	for _, st := range styles {
		key := fillKey(st.Fill)
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = len(table)

		fl := outFill{PatternFill: &outPatternFill{PatternType: st.Fill.Type}}
		if st.Fill.Color != "" {
			fl.PatternFill.FgColor = &outColor{RGB: st.Fill.Color}
		}
		table = append(table, fl)
	}
	return table, seen
}

func buildBorderTable(styles []*Style) ([]outBorder, map[string]int) {
	// Index 0: no borders.
	table := []outBorder{{
		Left:   &outBorderSide{},
		Right:  &outBorderSide{},
		Top:    &outBorderSide{},
		Bottom: &outBorderSide{},
	}}
	seen := map[string]int{
		borderKey(&Border{}): 0,
	}

	for _, st := range styles {
		key := borderKey(st.Border)
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = len(table)

		b := outBorder{}
		if st.Border.Left {
			b.Left = &outBorderSide{Style: "thin"}
		} else {
			b.Left = &outBorderSide{}
		}
		if st.Border.Right {
			b.Right = &outBorderSide{Style: "thin"}
		} else {
			b.Right = &outBorderSide{}
		}
		if st.Border.Top {
			b.Top = &outBorderSide{Style: "thin"}
		} else {
			b.Top = &outBorderSide{}
		}
		if st.Border.Bottom {
			b.Bottom = &outBorderSide{Style: "thin"}
		} else {
			b.Bottom = &outBorderSide{}
		}
		table = append(table, b)
	}
	return table, seen
}

// ======================================================================
// Stable key helpers for deduplication
// ======================================================================

func fontKey(f *Font) string {
	if f == nil {
		return ""
	}
	return "F:" + f.Name + "|" +
		strconv.FormatFloat(f.Size, 'f', -1, 64) + "|" +
		strconv.FormatBool(f.Bold) + "|" +
		strconv.FormatBool(f.Italic) + "|" +
		f.Color
}

func fillKey(f *Fill) string {
	if f == nil {
		return ""
	}
	return "L:" + f.Type + "|" + f.Color
}

func borderKey(b *Border) string {
	if b == nil {
		return ""
	}
	parts := []string{"B"}
	if b.Top {
		parts = append(parts, "T")
	}
	if b.Bottom {
		parts = append(parts, "B")
	}
	if b.Left {
		parts = append(parts, "L")
	}
	if b.Right {
		parts = append(parts, "R")
	}
	sort.Strings(parts)
	return parts[0] + ":" + joinSorted(parts[1:]) + "|" + b.Color
}

func joinSorted(ss []string) string {
	s := ""
	for i, v := range ss {
		if i > 0 {
			s += ","
		}
		s += v
	}
	return s
}

// ======================================================================
// Table XML generation
// ======================================================================

type outTable struct {
	XMLName        xml.Name        `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main table"`
	ID             string          `xml:"id,attr"`
	Name           string          `xml:"name,attr"`
	DisplayName    string          `xml:"displayName,attr"`
	Ref            string          `xml:"ref,attr"`
	HeaderRowCount int             `xml:"headerRowCount,attr,omitempty"`
	AutoFilter     *outAutoFilter  `xml:"autoFilter"`
	TableColumns   outTableColumns `xml:"tableColumns"`
	TableStyleInfo outTableStyle   `xml:"tableStyleInfo"`
}

type outAutoFilter struct {
	Ref string `xml:"ref,attr"`
}

type outTableColumns struct {
	Count   int              `xml:"count,attr"`
	Columns []outTableColumn `xml:"tableColumn"`
}

type outTableColumn struct {
	ID   int    `xml:"id,attr"`
	Name string `xml:"name,attr"`
}

type outTableStyle struct {
	Name              string `xml:"name,attr"`
	ShowFirstColumn   int    `xml:"showFirstColumn,attr"`
	ShowLastColumn    int    `xml:"showLastColumn,attr"`
	ShowRowStripes    int    `xml:"showRowStripes,attr"`
	ShowColumnStripes int    `xml:"showColumnStripes,attr"`
}

// generateTableXML produces the content of xl/tables/tableN.xml.
func generateTableXML(t *Table, tableID int) string {
	hdrCount := 0
	if t.HasHeaderRow {
		hdrCount = 1
	}

	colCount := rangeColumnCount(t.Range)
	columns := make([]outTableColumn, colCount)
	for i := 0; i < colCount; i++ {
		columns[i] = outTableColumn{
			ID:   i + 1,
			Name: fmt.Sprintf("Column%d", i+1),
		}
	}

	table := outTable{
		ID:             strconv.Itoa(tableID),
		Name:           t.Name,
		DisplayName:    t.Name,
		Ref:            t.Range,
		HeaderRowCount: hdrCount,
		AutoFilter:     &outAutoFilter{Ref: t.Range},
		TableColumns:   outTableColumns{Count: colCount, Columns: columns},
		TableStyleInfo: outTableStyle{
			Name:              t.StyleName,
			ShowFirstColumn:   0,
			ShowLastColumn:    0,
			ShowRowStripes:    1,
			ShowColumnStripes: 0,
		},
	}

	return marshalXML(table)
}

// rangeColumnCount returns the number of columns spanned by a range like "A1:D10".
func rangeColumnCount(rangeRef string) int {
	parts := splitRange(rangeRef)
	if len(parts) != 2 {
		return 1
	}
	sc, _ := splitRef(parts[0])
	ec, _ := splitRef(parts[1])
	return colToNum(ec) - colToNum(sc) + 1
}

// splitRange splits "A1:D10" into ("A1", "D10").
func splitRange(ref string) []string {
	parts := make([]string, 0, 2)
	for _, p := range strings.SplitN(ref, ":", 2) {
		parts = append(parts, strings.TrimSpace(p))
	}
	return parts
}

// generateSheetRelsXML produces xl/worksheets/_rels/sheetN.xml.rels with
// table relationships when the sheet contains tables.
func generateSheetRelsXML(tables []*Table) string {
	if len(tables) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")
	for i, t := range tables {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/%s.xml"/>`+"\n",
			i+1, strings.ToLower(t.Name))
	}
	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// generateContentTypesForTables returns additional <Override> elements for
// each table part that should appear in [Content_Types].xml.
func generateContentTypesForTables(worksheets []*Worksheet) string {
	var b strings.Builder
	for _, ws := range worksheets {
		for _, t := range ws.Tables {
			fmt.Fprintf(&b, `  <Override PartName="/xl/tables/%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`+"\n",
				strings.ToLower(t.Name))
		}
	}
	return b.String()
}

// outTableParts is added to the worksheet XML when the sheet has tables.
type outTableParts struct {
	Count      int            `xml:"count,attr"`
	TableParts []outTablePart `xml:"tablePart"`
}

type outTablePart struct {
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// buildTableParts returns nil when there are no tables; otherwise a populated
// outTableParts that references the table relationships in the sheet rels.
func buildTableParts(tables []*Table) *outTableParts {
	if len(tables) == 0 {
		return nil
	}
	parts := make([]outTablePart, len(tables))
	for i := range tables {
		parts[i] = outTablePart{RID: fmt.Sprintf("rId%d", i+1)}
	}
	return &outTableParts{Count: len(parts), TableParts: parts}
}

// ======================================================================
// Data validation XML generation
// ======================================================================

type outDataValidations struct {
	Count int                 `xml:"count,attr"`
	DVs   []outDataValidation `xml:"dataValidation"`
}

type outDataValidation struct {
	Type             string `xml:"type,attr"`
	Sqref            string `xml:"sqref,attr"`
	AllowBlank       int    `xml:"allowBlank,attr,omitempty"`
	ShowErrorMessage int    `xml:"showErrorMessage,attr,omitempty"`
	ErrorStyle       string `xml:"errorStyle,attr,omitempty"`
	ErrorTitle       string `xml:"errorTitle,attr,omitempty"`
	ErrorMessage     string `xml:"error,attr,omitempty"`
	Formula1         string `xml:"formula1,omitempty"`
	Formula2         string `xml:"formula2,omitempty"`
}

func buildDataValidations(dvs []*DataValidation) *outDataValidations {
	if len(dvs) == 0 {
		return nil
	}
	items := make([]outDataValidation, len(dvs))
	for i, dv := range dvs {
		ab := 0
		if dv.AllowBlank {
			ab = 1
		}
		se := 0
		if dv.ShowErrorMessage {
			se = 1
		}
		items[i] = outDataValidation{
			Type:             dv.Type,
			Sqref:            dv.Ref,
			AllowBlank:       ab,
			ShowErrorMessage: se,
			ErrorStyle:       dv.ErrorStyle,
			ErrorTitle:       dv.ErrorTitle,
			ErrorMessage:     dv.ErrorMessage,
			Formula1:         dv.Formula1,
			Formula2:         dv.Formula2,
		}
	}
	return &outDataValidations{Count: len(items), DVs: items}
}

// ======================================================================
// Drawing (picture) XML generation
// ======================================================================

type outDrawing struct {
	XMLName xml.Name    `xml:"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing wsDr"`
	A       string      `xml:"xmlns:a,attr"`
	Anchors []outAnchor `xml:"twoCellAnchor"`
}

type outAnchor struct {
	From       outPos      `xml:"from"`
	To         outPos      `xml:"to"`
	Pic        outPic      `xml:"pic"`
	ClientData outClientData `xml:"clientData"`
}

type outPos struct {
	Col    int   `xml:"col"`
	ColOff int64 `xml:"colOff"`
	Row    int   `xml:"row"`
	RowOff int64 `xml:"rowOff"`
}

type outPic struct {
	NvPicPr  outNvPicPr  `xml:"nvPicPr"`
	BlipFill outBlipFill `xml:"blipFill"`
	SpPr     outSpPr     `xml:"spPr"`
}

type outNvPicPr struct {
	CNvPr    outCNvPr    `xml:"cNvPr"`
	CNvPicPr outCNvPicPr `xml:"cNvPicPr"`
}

type outCNvPr struct {
	ID   int    `xml:"id,attr"`
	Name string `xml:"name,attr"`
}

type outCNvPicPr struct{}

type outBlipFill struct {
	Blip    outBlip   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main blip"`
	Stretch outStretch `xml:"http://schemas.openxmlformats.org/drawingml/2006/main stretch"`
}

type outBlip struct {
	Embed string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr"`
}

type outStretch struct {
	FillRect outFillRect `xml:"fillRect"`
}

type outFillRect struct{}

type outSpPr struct {
	Xfrm    outXfrm    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main xfrm"`
	PrstGeom outPrstGeom `xml:"http://schemas.openxmlformats.org/drawingml/2006/main prstGeom"`
}

type outXfrm struct {
	Off outPoint `xml:"off"`
	Ext outPoint `xml:"ext"`
}

type outPoint struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

type outPrstGeom struct {
	Prst string    `xml:"prst,attr"`
	AvLst outAvLst `xml:"avLst"`
}

type outAvLst struct{}

type outClientData struct{}

// generateDrawingXML produces the content of xl/drawings/drawingN.xml.
func generateDrawingXML(pictures []*Picture, idBase int) string {
	anchors := make([]outAnchor, len(pictures))
	for i, pic := range pictures {
		emuW := int64(pic.Width) * emuPerPixel
		emuH := int64(pic.Height) * emuPerPixel

		anchors[i] = outAnchor{
			From: outPos{
				Col:    pic.Col,
				ColOff: pic.ColOff,
				Row:    pic.Row,
				RowOff: pic.RowOff,
			},
			To: outPos{
				Col:    pic.Col,
				ColOff: pic.ColOff + emuW,
				Row:    pic.Row,
				RowOff: pic.RowOff + emuH,
			},
			Pic: outPic{
				NvPicPr: outNvPicPr{
					CNvPr:    outCNvPr{ID: idBase + i + 1, Name: pic.Name},
					CNvPicPr: outCNvPicPr{},
				},
				BlipFill: outBlipFill{
					Blip:    outBlip{Embed: fmt.Sprintf("rId%d", i+1)},
					Stretch: outStretch{FillRect: outFillRect{}},
				},
				SpPr: outSpPr{
					Xfrm: outXfrm{
						Off: outPoint{X: pic.ColOff, Y: pic.RowOff},
						Ext: outPoint{X: emuW, Y: emuH},
					},
					PrstGeom: outPrstGeom{Prst: "rect", AvLst: outAvLst{}},
				},
			},
			ClientData: outClientData{},
		}
	}

	d := outDrawing{
		A:       "http://schemas.openxmlformats.org/drawingml/2006/main",
		Anchors: anchors,
	}
	return marshalXML(d)
}

// generateDrawingRelsXML produces xl/drawings/_rels/drawingN.xml.rels.
func generateDrawingRelsXML(pictures []*Picture) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")
	for i, pic := range pictures {
		ext := pic.Format
		if ext == "jpeg" {
			ext = "jpg"
		}
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image%d.%s"/>`+"\n",
			i+1, i+1, ext)
	}
	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// generateUnifiedSheetRelsXML produces xl/worksheets/_rels/sheetN.xml.rels
// containing relationships for both tables and the drawing (when present).
// rIds are assigned sequentially: tables first, then drawing.
func generateUnifiedSheetRelsXML(tables []*Table, pictures []*Picture) string {
	if len(tables) == 0 && len(pictures) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")

	rid := 1
	for _, t := range tables {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/%s.xml"/>`+"\n",
			rid, strings.ToLower(t.Name))
		rid++
	}
	if len(pictures) > 0 {
		// The drawing file is named drawingN.xml; we use the first picture's
		// sheet context.  Each sheet has exactly one drawing file.
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>`+"\n", rid)
	}

	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// drawingRID returns the rId that the <drawing> element in the sheet XML
// should reference, or empty when there are no pictures.  This is always
// len(tables)+1 because tables occupy rIds 1..N.
func drawingRID(tables []*Table) string {
	if len(tables) > 0 {
		return fmt.Sprintf("rId%d", len(tables)+1)
	}
	return "rId1"
}

// generateContentTypesForDrawings returns <Override> elements for drawing
// and media parts.
func generateContentTypesForDrawings(worksheets []*Worksheet) string {
	var b strings.Builder
	hasDrawing := false
	hasPng := false
	hasJpeg := false
	for _, ws := range worksheets {
		if len(ws.Pictures) > 0 {
			hasDrawing = true
		}
		for _, pic := range ws.Pictures {
			if pic.Format == "png" {
				hasPng = true
			} else if pic.Format == "jpeg" {
				hasJpeg = true
			}
		}
	}
	if hasDrawing {
		b.WriteString(`  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>` + "\n")
	}
	if hasPng {
		b.WriteString(`  <Default Extension="png" ContentType="image/png"/>` + "\n")
	}
	if hasJpeg {
		b.WriteString(`  <Default Extension="jpg" ContentType="image/jpeg"/>` + "\n")
	}
	return b.String()
}
