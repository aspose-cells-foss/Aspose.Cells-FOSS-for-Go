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
	XMLName      xml.Name   `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	Fonts        outFonts   `xml:"fonts"`
	Fills        outFills   `xml:"fills"`
	Borders      outBorders `xml:"borders"`
	CellStyleXfs outCellXfs `xml:"cellStyleXfs"`
	CellXfs      outCellXfs `xml:"cellXfs"`
	Dxfs         *outDxfs   `xml:"dxfs,omitempty"`
}

type outDxfs struct {
	Count int      `xml:"count,attr"`
	Dxfs  []outDxf `xml:"dxf"`
}

type outDxf struct {
	Font      *outFont      `xml:"font,omitempty"`
	Fill      *outFill      `xml:"fill,omitempty"`
	Border    *outBorder    `xml:"border,omitempty"`
	Alignment *outAlign     `xml:"alignment,omitempty"`
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

	// ---- Build DXF table from conditional formatting rules ----
	dxfTable := buildDxfTable(wb.Worksheets)
	if len(dxfTable) > 0 {
		ss.Dxfs = &outDxfs{Count: len(dxfTable), Dxfs: dxfTable}
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
// DXF (Differential Formatting) table builder
// ======================================================================

// buildDxfTable collects all unique styles used in conditional formatting rules
// across all worksheets and returns them as DXF elements.
func buildDxfTable(worksheets []*Worksheet) []outDxf {
	seen := make(map[string]int)
	var table []outDxf

	for _, ws := range worksheets {
		for _, cf := range ws.ConditionalFormattings {
			for _, rule := range cf.Rules {
				if rule.Style == nil {
					continue
				}

				key := dxfKey(rule.Style)
				if idx, ok := seen[key]; ok {
					rule.StyleID = idx
					continue
				}

				// Create new DXF element
				dxf := outDxf{}

				// Font
				if rule.Style.Font != nil {
					font := &outFont{
						Sz:   outVal{Val: strconv.FormatFloat(rule.Style.Font.Size, 'f', -1, 64)},
						Name: outVal{Val: rule.Style.Font.Name},
					}
					if rule.Style.Font.Bold {
						font.B = &outEmpty{}
					}
					if rule.Style.Font.Italic {
						font.I = &outEmpty{}
					}
					if rule.Style.Font.Color != "" && rule.Style.Font.Color != "FF000000" {
						font.Color = &outColor{RGB: rule.Style.Font.Color}
					}
					dxf.Font = font
				}

				// Fill
				if rule.Style.Fill != nil && rule.Style.Fill.Type != "" && rule.Style.Fill.Type != "none" {
					fill := &outFill{PatternFill: &outPatternFill{PatternType: rule.Style.Fill.Type}}
					if rule.Style.Fill.Color != "" {
						fill.PatternFill.FgColor = &outColor{RGB: rule.Style.Fill.Color}
					}
					dxf.Fill = fill
				}

				// Border
				if rule.Style.Border != nil && (rule.Style.Border.Top || rule.Style.Border.Bottom || rule.Style.Border.Left || rule.Style.Border.Right) {
					border := &outBorder{}
					if rule.Style.Border.Left {
						border.Left = &outBorderSide{Style: "thin"}
					} else {
						border.Left = &outBorderSide{}
					}
					if rule.Style.Border.Right {
						border.Right = &outBorderSide{Style: "thin"}
					} else {
						border.Right = &outBorderSide{}
					}
					if rule.Style.Border.Top {
						border.Top = &outBorderSide{Style: "thin"}
					} else {
						border.Top = &outBorderSide{}
					}
					if rule.Style.Border.Bottom {
						border.Bottom = &outBorderSide{Style: "thin"}
					} else {
						border.Bottom = &outBorderSide{}
					}
					dxf.Border = border
				}

				// Alignment
				if rule.Style.Alignment != nil && (rule.Style.Alignment.Horizontal != "" || rule.Style.Alignment.Vertical != "" || rule.Style.Alignment.WrapText) {
					align := &outAlign{
						Horizontal: rule.Style.Alignment.Horizontal,
						Vertical:   rule.Style.Alignment.Vertical,
					}
					if rule.Style.Alignment.WrapText {
						align.WrapText = "1"
					}
					dxf.Alignment = align
				}

				seen[key] = len(table)
				rule.StyleID = len(table)
				table = append(table, dxf)
			}
		}
	}

	return table
}

// dxfKey creates a unique key for a Style to deduplicate DXF elements.
func dxfKey(s *Style) string {
	if s == nil {
		return ""
	}
	key := "DXF:"
	if s.Font != nil {
		key += "F:" + s.Font.Name + "|" +
			strconv.FormatFloat(s.Font.Size, 'f', -1, 64) + "|" +
			strconv.FormatBool(s.Font.Bold) + "|" +
			strconv.FormatBool(s.Font.Italic) + "|" +
			s.Font.Color + "|"
	}
	if s.Fill != nil {
		key += "L:" + s.Fill.Type + "|" + s.Fill.Color + "|"
	}
	if s.Border != nil {
		key += "B:" + strconv.FormatBool(s.Border.Top) + "|" +
			strconv.FormatBool(s.Border.Bottom) + "|" +
			strconv.FormatBool(s.Border.Left) + "|" +
			strconv.FormatBool(s.Border.Right) + "|" +
			s.Border.Color + "|"
	}
	if s.Alignment != nil {
		key += "A:" + s.Alignment.Horizontal + "|" +
			s.Alignment.Vertical + "|" +
			strconv.FormatBool(s.Alignment.WrapText)
	}
	return key
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
	From       outPos        `xml:"from"`
	To         outPos        `xml:"to"`
	Pic        outPic        `xml:"pic"`
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
	Blip    outBlip    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main blip"`
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
	Xfrm     outXfrm     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main xfrm"`
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
	Prst  string   `xml:"prst,attr"`
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
// globalPicIdx is the running picture counter so that media part names are
// unique across all sheets in the workbook.
func generateDrawingRelsXML(pictures []*Picture, globalPicIdx int) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")
	for i, pic := range pictures {
		ext := pic.Format
		if ext == "jpeg" {
			ext = "jpg"
		}
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image%d.%s"/>`+"\n",
			i+1, globalPicIdx+i+1, ext)
	}
	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// generateUnifiedSheetRelsXML produces xl/worksheets/_rels/sheetN.xml.rels
// containing relationships for tables, drawing, charts, and pivot tables (when present).
// rIds are assigned sequentially: tables first, then drawing, charts, pivot tables.
// sheetIndex is the 0-based sheet index used to name the drawing part.
func generateUnifiedSheetRelsXML(tables []*Table, pictures []*Picture, charts []*Chart, pivotTables []*PivotTable, sheetIndex int) string {
	if len(tables) == 0 && len(pictures) == 0 && len(charts) == 0 && len(pivotTables) == 0 {
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
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing%d.xml"/>`+"\n", rid, sheetIndex+1)
		rid++
	}
	for i := range charts {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart%d.xml"/>`+"\n", rid, i+1)
		rid++
	}
	for i := range pivotTables {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable%d.xml"/>`+"\n", rid, i+1)
		rid++
	}

	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// drawingRID returns the rId that the <drawing> element in the sheet XML
// should reference.  Callers only use it for sheets that have pictures, and
// tables occupy rIds 1..N, so the drawing always takes the next one.
func drawingRID(tables []*Table) string {
	return fmt.Sprintf("rId%d", len(tables)+1)
}

// generateContentTypesForDrawings returns <Override> elements for drawing
// and media parts.
func generateContentTypesForDrawings(worksheets []*Worksheet) string {
	var b strings.Builder
	hasPng := false
	hasJpeg := false
	for i, ws := range worksheets {
		if len(ws.Pictures) > 0 {
			fmt.Fprintf(&b, `  <Override PartName="/xl/drawings/drawing%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`+"\n", i+1)
		}
		for _, pic := range ws.Pictures {
			if pic.Format == "png" {
				hasPng = true
			} else if pic.Format == "jpeg" {
				hasJpeg = true
			}
		}
	}
	if hasPng {
		b.WriteString(`  <Default Extension="png" ContentType="image/png"/>` + "\n")
	}
	if hasJpeg {
		b.WriteString(`  <Default Extension="jpg" ContentType="image/jpeg"/>` + "\n")
	}
	return b.String()
}

// ======================================================================
// Conditional formatting XML generation
// ======================================================================

type outConditionalFormatting struct {
	Ref   string      `xml:"ref,attr"`
	Rules []outCFRule `xml:"cfRule"`
}

type outCFRule struct {
	Type       string `xml:"type,attr"`
	Operator   string `xml:"operator,attr,omitempty"`
	Priority   int    `xml:"priority,attr"`
	Formula    string `xml:"formula,omitempty"`
	Formula2   string `xml:"formula2,omitempty"`
	Text       string `xml:"text,attr,omitempty"`
	StopIfTrue int    `xml:"stopIfTrue,attr,omitempty"`
	DxfID      int    `xml:"dxfId,attr,omitempty"`
}

func buildConditionalFormattings(cfs []*ConditionalFormatting) []outConditionalFormatting {
	if len(cfs) == 0 {
		return nil
	}
	out := make([]outConditionalFormatting, len(cfs))
	for i, cf := range cfs {
		rules := make([]outCFRule, len(cf.Rules))
		for j, rule := range cf.Rules {
			stopIfTrue := 0
			if rule.StopIfTrue {
				stopIfTrue = 1
			}
			rules[j] = outCFRule{
				Type:       rule.Type,
				Operator:   rule.Operator,
				Priority:   rule.Priority,
				Formula:    rule.Formula,
				Formula2:   rule.Formula2,
				Text:       rule.Text,
				StopIfTrue: stopIfTrue,
				DxfID:      rule.StyleID,
			}
		}
		out[i] = outConditionalFormatting{
			Ref:   cf.Ref,
			Rules: rules,
		}
	}
	return out
}

// ======================================================================
// Chart XML generation
// ======================================================================

// generateChartXML produces the content of xl/charts/chartN.xml.
// DEPRECATED: Use generateChartXMLFull instead for complete chart support.
func generateChartXML(chart *Chart, sheetName string) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` + "\n")
	b.WriteString(`  <c:chart>` + "\n")

	if chart.Title != "" {
		b.WriteString(fmt.Sprintf(`    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>%s</a:t></a:r></a:p></c:rich></c:tx></c:title>`+"\n", xmlEscape(chart.Title)))
	}

	// Plot area with chart type.
	b.WriteString(`    <c:plotArea>` + "\n")
	b.WriteString(`      <c:layout/>` + "\n")

	switch chart.Type {
	case ChartTypeBar:
		b.WriteString(`      <c:barChart>` + "\n")
		b.WriteString(`        <c:barDir val="col"/>` + "\n")
		b.WriteString(`        <c:grouping val="clustered"/>` + "\n")
		for i, series := range chart.Series {
			b.WriteString(fmt.Sprintf(`        <c:ser><c:idx val="%d"/><c:order val="%d"/>`, i, i))
			if series.Name != "" {
				b.WriteString(fmt.Sprintf(`<c:tx><c:strRef><c:f>'%s'!%s</c:f></c:strRef></c:tx>`, xmlEscape(sheetName), xmlEscape(series.Name)))
			}
			if series.Categories != "" {
				b.WriteString(fmt.Sprintf(`<c:cat><c:strRef><c:f>'%s'!%s</c:f></c:strRef></c:cat>`, xmlEscape(sheetName), xmlEscape(series.Categories)))
			}
			if series.Values != "" {
				b.WriteString(fmt.Sprintf(`<c:val><c:numRef><c:f>'%s'!%s</c:f></c:numRef></c:val>`, xmlEscape(sheetName), xmlEscape(series.Values)))
			}
			b.WriteString(`</c:ser>` + "\n")
		}
		b.WriteString(`      </c:barChart>` + "\n")
	case ChartTypeLine:
		b.WriteString(`      <c:lineChart>` + "\n")
		b.WriteString(`        <c:grouping val="standard"/>` + "\n")
		for i, series := range chart.Series {
			b.WriteString(fmt.Sprintf(`        <c:ser><c:idx val="%d"/><c:order val="%d"/>`, i, i))
			if series.Name != "" {
				b.WriteString(fmt.Sprintf(`<c:tx><c:strRef><c:f>'%s'!%s</c:f></c:strRef></c:tx>`, xmlEscape(sheetName), xmlEscape(series.Name)))
			}
			if series.Categories != "" {
				b.WriteString(fmt.Sprintf(`<c:cat><c:strRef><c:f>'%s'!%s</c:f></c:strRef></c:cat>`, xmlEscape(sheetName), xmlEscape(series.Categories)))
			}
			if series.Values != "" {
				b.WriteString(fmt.Sprintf(`<c:val><c:numRef><c:f>'%s'!%s</c:f></c:numRef></c:val>`, xmlEscape(sheetName), xmlEscape(series.Values)))
			}
			b.WriteString(`</c:ser>` + "\n")
		}
		b.WriteString(`      </c:lineChart>` + "\n")
	case ChartTypePie:
		b.WriteString(`      <c:pieChart>` + "\n")
		if len(chart.Series) > 0 {
			series := chart.Series[0]
			b.WriteString(`        <c:ser><c:idx val="0"/><c:order val="0"/>`)
			if series.Categories != "" {
				b.WriteString(fmt.Sprintf(`<c:cat><c:strRef><c:f>'%s'!%s</c:f></c:strRef></c:cat>`, xmlEscape(sheetName), xmlEscape(series.Categories)))
			}
			if series.Values != "" {
				b.WriteString(fmt.Sprintf(`<c:val><c:numRef><c:f>'%s'!%s</c:f></c:numRef></c:val>`, xmlEscape(sheetName), xmlEscape(series.Values)))
			}
			b.WriteString(`</c:ser>` + "\n")
		}
		b.WriteString(`      </c:pieChart>` + "\n")
	}

	b.WriteString(`    </c:plotArea>` + "\n")
	b.WriteString(`  </c:chart>` + "\n")
	b.WriteString(`</c:chartSpace>` + "\n")
	return b.String()
}

// xmlEscape escapes special XML characters.
func xmlEscape(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	s = strings.ReplaceAll(s, "'", "&apos;")
	return s
}

// ======================================================================
// Chart integration helpers
// ======================================================================

// generateContentTypesForCharts returns <Override> elements for chart parts.
func generateContentTypesForCharts(worksheets []*Worksheet) string {
	var b strings.Builder
	for _, ws := range worksheets {
		for i := range ws.Charts {
			fmt.Fprintf(&b, `  <Override PartName="/xl/charts/chart%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`+"\n", i+1)
		}
	}
	return b.String()
}

// ======================================================================
// Pivot Table XML generation
// ======================================================================

type outPivotTableDefinition struct {
	XMLName         xml.Name                   `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main pivotTableDefinition"`
	Name            string                     `xml:"name,attr"`
	CacheID         int                        `xml:"cacheId,attr"`
	DataOnRows      int                        `xml:"dataOnRows,attr"`
	DataCaption     string                     `xml:"dataCaption,attr"`
	Location        outPivotLocation           `xml:"location"`
	PivotFields     outPivotFields             `xml:"pivotFields"`
	RowFields       *outPivotRowFields         `xml:"rowFields,omitempty"`
	ColFields       *outPivotColFields         `xml:"colFields,omitempty"`
	DataFields      *outPivotDataFields        `xml:"dataFields,omitempty"`
	PivotTableStyle outPivotTableStyleInfo     `xml:"pivotTableStyleInfo"`
}

type outPivotLocation struct {
	Ref         string `xml:"ref,attr"`
	FirstRowCol int    `xml:"firstRowCol,attr"`
}

type outPivotFields struct {
	Count  int              `xml:"count,attr"`
	Fields []outPivotField  `xml:"pivotField"`
}

type outPivotField struct {
	Name      string `xml:"name,attr,omitempty"`
	Axis      string `xml:"axis,attr,omitempty"`
	ShowAll   int    `xml:"showAll,attr"`
}

type outPivotRowFields struct {
	Count int              `xml:"count,attr"`
	Fields []outPivotFieldRef `xml:"field"`
}

type outPivotColFields struct {
	Count int              `xml:"count,attr"`
	Fields []outPivotFieldRef `xml:"field"`
}

type outPivotFieldRef struct {
	X int `xml:"x,attr"`
}

type outPivotDataFields struct {
	Count  int                  `xml:"count,attr"`
	Fields []outPivotDataField  `xml:"dataField"`
}

type outPivotDataField struct {
	Name     string `xml:"name,attr"`
	Field    int    `xml:"fld,attr"`
	Subtotal string `xml:"subtotal,attr,omitempty"`
}

type outPivotTableStyleInfo struct {
	Name              string `xml:"name,attr"`
	ShowRowHeaders    int    `xml:"showRowHeaders,attr"`
	ShowColHeaders    int    `xml:"showColHeaders,attr"`
	ShowRowStripes    int    `xml:"showRowStripes,attr"`
	ShowColStripes    int    `xml:"showColStripes,attr"`
}

// generatePivotTableXML produces the content of xl/pivotTables/pivotTableN.xml.
func generatePivotTableXML(pt *PivotTable, cacheID int) string {
	// Parse source range to get column count
	colCount := rangeColumnCount(pt.SourceRef)

	// Build pivot fields (one per source column)
	fields := make([]outPivotField, colCount)
	for i := 0; i < colCount; i++ {
		fields[i] = outPivotField{
			Name:    fmt.Sprintf("Field%d", i),
			ShowAll: 0,
		}
	}

	// Mark row/col/data fields with axis
	for _, rf := range pt.RowFields {
		for i := range fields {
			if fields[i].Name == rf {
				fields[i].Axis = "axisRow"
			}
		}
	}
	for _, cf := range pt.ColFields {
		for i := range fields {
			if fields[i].Name == cf {
				fields[i].Axis = "axisCol"
			}
		}
	}

	pivotDef := outPivotTableDefinition{
		Name:        pt.Name,
		CacheID:     cacheID,
		DataOnRows:  1,
		DataCaption: "Values",
		Location: outPivotLocation{
			Ref:         pt.Ref,
			FirstRowCol: 1,
		},
		PivotFields: outPivotFields{
			Count:  len(fields),
			Fields: fields,
		},
		PivotTableStyle: outPivotTableStyleInfo{
			Name:           "PivotTableStyleLight16",
			ShowRowHeaders: 1,
			ShowColHeaders: 1,
			ShowRowStripes: 1,
			ShowColStripes: 0,
		},
	}

	// Add row fields
	if len(pt.RowFields) > 0 {
		rowFieldRefs := make([]outPivotFieldRef, len(pt.RowFields))
		for i := range pt.RowFields {
			rowFieldRefs[i] = outPivotFieldRef{X: i}
		}
		pivotDef.RowFields = &outPivotRowFields{
			Count:  len(rowFieldRefs),
			Fields: rowFieldRefs,
		}
	}

	// Add col fields
	if len(pt.ColFields) > 0 {
		colFieldRefs := make([]outPivotFieldRef, len(pt.ColFields))
		for i := range pt.ColFields {
			colFieldRefs[i] = outPivotFieldRef{X: i}
		}
		pivotDef.ColFields = &outPivotColFields{
			Count:  len(colFieldRefs),
			Fields: colFieldRefs,
		}
	}

	// Add data fields
	if len(pt.DataFields) > 0 {
		dataFields := make([]outPivotDataField, len(pt.DataFields))
		for i, df := range pt.DataFields {
			dataFields[i] = outPivotDataField{
				Name:     df.DisplayName,
				Field:    i,
				Subtotal: df.Aggregation,
			}
		}
		pivotDef.DataFields = &outPivotDataFields{
			Count:  len(dataFields),
			Fields: dataFields,
		}
	}

	return marshalXML(pivotDef)
}

// generatePivotCacheDefinitionXML produces xl/pivotCache/pivotCacheDefinitionN.xml.
func generatePivotCacheDefinitionXML(pt *PivotTable, cacheID int) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` + "\n")
	b.WriteString(fmt.Sprintf(`  <cacheSource type="worksheet">` + "\n"))
	b.WriteString(fmt.Sprintf(`    <worksheetSource ref="%s" sheet="%s"/>` + "\n", xmlEscape(pt.SourceRef), xmlEscape(pt.Name)))
	b.WriteString(`  </cacheSource>` + "\n")
	b.WriteString(fmt.Sprintf(`  <cacheFields count="%d">` + "\n", rangeColumnCount(pt.SourceRef)))
	for i := 0; i < rangeColumnCount(pt.SourceRef); i++ {
		b.WriteString(fmt.Sprintf(`    <cacheField name="Field%d" numFmtId="0">` + "\n", i))
		b.WriteString(`      <sharedItems/>` + "\n")
		b.WriteString(`    </cacheField>` + "\n")
	}
	b.WriteString(`  </cacheFields>` + "\n")
	b.WriteString(`</pivotCacheDefinition>` + "\n")
	return b.String()
}

// generatePivotCacheRecordsXML produces xl/pivotCache/pivotCacheRecordsN.xml.
func generatePivotCacheRecordsXML(pt *PivotTable) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` + "\n")
	b.WriteString(`</pivotCacheRecords>` + "\n")
	return b.String()
}

// generateContentTypesForPivotTables returns <Override> elements for pivot table parts.
func generateContentTypesForPivotTables(worksheets []*Worksheet) string {
	var b strings.Builder
	ptIdx := 0
	for _, ws := range worksheets {
		for range ws.PivotTables {
			ptIdx++
			fmt.Fprintf(&b, `  <Override PartName="/xl/pivotCache/pivotCacheDefinition%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>`+"\n", ptIdx)
			fmt.Fprintf(&b, `  <Override PartName="/xl/pivotCache/pivotCacheRecords%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>`+"\n", ptIdx)
			fmt.Fprintf(&b, `  <Override PartName="/xl/pivotTables/pivotTable%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>`+"\n", ptIdx)
		}
	}
	return b.String()
}
