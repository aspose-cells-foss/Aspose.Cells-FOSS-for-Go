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
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	// ECMA-376 orders numFmts first, ahead of fonts.
	NumFmts      *outNumFmts `xml:"numFmts,omitempty"`
	Fonts        outFonts    `xml:"fonts"`
	Fills        outFills    `xml:"fills"`
	Borders      outBorders  `xml:"borders"`
	CellStyleXfs outCellXfs  `xml:"cellStyleXfs"`
	CellXfs      outCellXfs  `xml:"cellXfs"`
	// The <dxfs> block is injected as raw XML after marshaling (see
	// generateStylesXML) so cached entries from the loaded file can be
	// preserved verbatim ahead of freshly built ones.
}

type outNumFmts struct {
	Count   int         `xml:"count,attr"`
	NumFmts []outNumFmt `xml:"numFmt"`
}

type outNumFmt struct {
	NumFmtId   int    `xml:"numFmtId,attr"`
	FormatCode string `xml:"formatCode,attr"`
}

// outDxf is a differential format entry. Child order follows the CT_Dxf
// sequence: font, numFmt, fill, alignment, border.
type outDxf struct {
	XMLName   xml.Name   `xml:"dxf"`
	Font      *outFont   `xml:"font,omitempty"`
	NumFmt    *outNumFmt `xml:"numFmt,omitempty"`
	Fill      *outFill   `xml:"fill,omitempty"`
	Alignment *outAlign  `xml:"alignment,omitempty"`
	Border    *outBorder `xml:"border,omitempty"`
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
	NumFmtId          int       `xml:"numFmtId,attr"`
	FontId            int       `xml:"fontId,attr"`
	FillId            int       `xml:"fillId,attr"`
	BorderId          int       `xml:"borderId,attr"`
	XfId              int       `xml:"xfId,attr"`
	ApplyNumberFormat int       `xml:"applyNumberFormat,attr,omitempty"`
	ApplyFont         int       `xml:"applyFont,attr,omitempty"`
	ApplyFill         int       `xml:"applyFill,attr,omitempty"`
	ApplyBorder       int       `xml:"applyBorder,attr,omitempty"`
	ApplyAlignment    int       `xml:"applyAlignment,attr,omitempty"`
	Alignment         *outAlign `xml:"alignment,omitempty"`
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
func generateStylesXML(wb *Workbook) (string, error) {
	if len(wb.styles) == 0 {
		return "", nil
	}

	// ---- Font table ----
	fontTable, fontIndex := buildFontTable(wb.styles)

	// ---- Fill table ----
	fillTable, fillIndex := buildFillTable(wb.styles)

	// ---- Border table ----
	borderTable, borderIndex := buildBorderTable(wb.styles)

	// ---- cellXfs ----
	numFmts := newNumFmtAllocator()
	xfCount := len(wb.styles)
	cellXfs := make([]outXf, xfCount)
	for i, st := range wb.styles {
		fID := fontIndex[fontKey(st.Font)]
		fiID := fillIndex[fillKey(st.Fill)]
		bID := borderIndex[borderKey(st.Border)]

		xf := outXf{
			NumFmtId: numFmts.idFor(st.NumberFormat),
			FontId:   fID,
			FillId:   fiID,
			BorderId: bID,
			XfId:     0,
		}

		// Apply flags.
		if xf.NumFmtId != 0 {
			xf.ApplyNumberFormat = 1
		}
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

	// ---- DXF table ----
	// Cached entries from the loaded styles.xml keep their original
	// indices; fresh entries are appended after them, so loaded rules and
	// cached sheet XML keep pointing at the right <dxf>. Assign the rule
	// style ids here so sheet XML generated afterwards (buildConditional-
	// Formattings reads rule.StyleID) references the merged block.
	// DXFs share the workbook-wide numFmt table, so custom codes used by
	// rule styles are allocated from the same allocator as cellXfs.
	dxfTable, dxfIDs := buildDxfTable(wb.Worksheets, wb.dxfBaseCount, numFmts)
	for rule, id := range dxfIDs {
		rule.StyleID = id
	}
	if entries := numFmts.entries(); len(entries) > 0 {
		ss.NumFmts = &outNumFmts{Count: len(entries), NumFmts: entries}
	}

	body, err := marshalXML(ss)
	if err != nil {
		return "", err
	}
	if wb.dxfBaseCount > 0 || len(dxfTable) > 0 {
		var b strings.Builder
		fmt.Fprintf(&b, "<dxfs count=\"%d\">", wb.dxfBaseCount+len(dxfTable))
		b.WriteString(wb.dxfRaw)
		for _, d := range dxfTable {
			raw, err := xml.Marshal(d)
			if err != nil {
				return "", fmt.Errorf("dxf xml marshal failed: %w", err)
			}
			b.Write(raw)
		}
		b.WriteString("</dxfs>")
		// CT_Stylesheet places dxfs after cellXfs; no later elements are
		// emitted, so appending before the closing tag keeps the sequence.
		body = strings.Replace(body, "</styleSheet>", b.String()+"\n</styleSheet>", 1)
	}

	return body, nil
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

// hasDxfStyles reports whether any conditional-formatting rule carries a
// Style that generateStylesXML must emit as a <dxf> entry. The save planner
// uses it to ensure the styles part is written even when no cell styles are
// registered.
func hasDxfStyles(worksheets []*Worksheet) bool {
	for _, ws := range worksheets {
		for _, cf := range ws.ConditionalFormattings {
			for _, rule := range cf.Rules {
				if rule.Style != nil {
					return true
				}
			}
		}
	}
	return false
}

// buildDxfTable collects all unique styles used in conditional formatting
// rules across all worksheets and returns them as DXF elements, along with
// the dxfs-block index each rule's StyleID must take. base is the number of
// cached entries preserved from the loaded styles.xml, so fresh entries
// start at base. Custom number format codes are allocated through numFmts,
// the workbook-wide table shared with cellXfs. The caller applies the
// returned ids to the rules; this function does not mutate them.
func buildDxfTable(worksheets []*Worksheet, base int, numFmts *numFmtAllocator) ([]outDxf, map[*ConditionalFormattingRule]int) {
	seen := make(map[string]int)
	ids := make(map[*ConditionalFormattingRule]int)
	var table []outDxf

	for _, ws := range worksheets {
		for _, cf := range ws.ConditionalFormattings {
			for _, rule := range cf.Rules {
				if rule.Style == nil {
					continue
				}

				key := dxfKey(rule.Style)
				if idx, ok := seen[key]; ok {
					ids[rule] = base + idx
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

				// Number format
				if nfID := numFmts.idFor(rule.Style.NumberFormat); nfID != 0 {
					dxf.NumFmt = &outNumFmt{NumFmtId: nfID, FormatCode: numFmts.codeForID(nfID)}
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
				ids[rule] = base + len(table)
				table = append(table, dxf)
			}
		}
	}

	return table, ids
}

// dxfKey creates a unique key for a Style to deduplicate DXF elements.
func dxfKey(s *Style) string {
	if s == nil {
		return ""
	}
	key := "DXF:"
	key += "N:" + s.NumberFormat + "|"
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
func generateTableXML(t *Table, tableID int) (string, error) {
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
	Operator         string `xml:"operator,attr,omitempty"`
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
			Operator:         dv.Operator,
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

// outAnchor is a twoCellAnchor holding either a picture (pic) or a chart
// (graphicFrame), never both.
type outAnchor struct {
	From         outPos           `xml:"from"`
	To           outPos           `xml:"to"`
	Pic          *outPic          `xml:"pic,omitempty"`
	GraphicFrame *outGraphicFrame `xml:"graphicFrame,omitempty"`
	ClientData   outClientData    `xml:"clientData"`
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
	Off outPoint  `xml:"off"`
	Ext outExtent `xml:"ext"`
}

type outPoint struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// outExtent mirrors a:ext, which uses cx/cy attributes (unlike a:off).
type outExtent struct {
	CX int64 `xml:"cx,attr"`
	CY int64 `xml:"cy,attr"`
}

type outPrstGeom struct {
	Prst  string   `xml:"prst,attr"`
	AvLst outAvLst `xml:"avLst"`
}

type outAvLst struct{}

type outClientData struct{}

// outGraphicFrame embeds a chart part in a drawing (CT_GraphicalObjectFrame).
// The chart is referenced by relationship id; the relationship lives in the
// drawing's .rels file.
type outGraphicFrame struct {
	NvGraphicFramePr outNvGraphicFramePr `xml:"nvGraphicFramePr"`
	Xfrm             outGFXfrm           `xml:"xfrm"`
	Graphic          outGraphic          `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphic"`
}

type outNvGraphicFramePr struct {
	CNvPr             outCNvPr             `xml:"cNvPr"`
	CNvGraphicFramePr outCNvGraphicFramePr `xml:"cNvGraphicFramePr"`
}

type outCNvGraphicFramePr struct{}

type outGFXfrm struct {
	Off outPoint  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main off"`
	Ext outExtent `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ext"`
}

type outGraphic struct {
	GraphicData outGraphicData `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphicData"`
}

type outGraphicData struct {
	URI   string      `xml:"uri,attr"`
	Chart outChartRef `xml:"http://schemas.openxmlformats.org/drawingml/2006/chart chart"`
}

type outChartRef struct {
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// generateDrawingXML produces the content of xl/drawings/drawingN.xml.
// Pictures get a twoCellAnchor with a pic element; charts get a
// twoCellAnchor with a graphicFrame referencing the chart part. idBase
// offsets the cNvPr ids so they stay unique across the workbook.
//
// Relationship ids match generateDrawingRelsXML: pictures occupy
// rId1..len(pictures) and charts continue from rId(len(pictures)+1).
func generateDrawingXML(pictures []*Picture, charts []*Chart, idBase int) (string, error) {
	anchors := make([]outAnchor, 0, len(pictures)+len(charts))
	id := idBase

	for _, pic := range pictures {
		id++
		emuW := int64(pic.Width) * emuPerPixel
		emuH := int64(pic.Height) * emuPerPixel

		anchors = append(anchors, outAnchor{
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
			Pic: &outPic{
				NvPicPr: outNvPicPr{
					CNvPr:    outCNvPr{ID: id, Name: pic.Name},
					CNvPicPr: outCNvPicPr{},
				},
				BlipFill: outBlipFill{
					Blip:    outBlip{Embed: fmt.Sprintf("rId%d", len(anchors)+1)},
					Stretch: outStretch{FillRect: outFillRect{}},
				},
				SpPr: outSpPr{
					Xfrm: outXfrm{
						Off: outPoint{X: pic.ColOff, Y: pic.RowOff},
						Ext: outExtent{CX: emuW, CY: emuH},
					},
					PrstGeom: outPrstGeom{Prst: "rect", AvLst: outAvLst{}},
				},
			},
			ClientData: outClientData{},
		})
	}

	for j, ch := range charts {
		id++
		emuW := int64(ch.Width) * emuPerPixel
		emuH := int64(ch.Height) * emuPerPixel
		name := ch.Name
		if name == "" {
			name = fmt.Sprintf("Chart %d", j+1)
		}

		anchors = append(anchors, outAnchor{
			From: outPos{
				Col:    ch.Anchor.Col,
				ColOff: ch.Anchor.ColOff,
				Row:    ch.Anchor.Row,
				RowOff: ch.Anchor.RowOff,
			},
			To: outPos{
				Col:    ch.Anchor.Col,
				ColOff: ch.Anchor.ColOff + emuW,
				Row:    ch.Anchor.Row,
				RowOff: ch.Anchor.RowOff + emuH,
			},
			GraphicFrame: &outGraphicFrame{
				NvGraphicFramePr: outNvGraphicFramePr{
					CNvPr:             outCNvPr{ID: id, Name: name},
					CNvGraphicFramePr: outCNvGraphicFramePr{},
				},
				Xfrm: outGFXfrm{
					Off: outPoint{X: ch.Anchor.ColOff, Y: ch.Anchor.RowOff},
					Ext: outExtent{CX: emuW, CY: emuH},
				},
				Graphic: outGraphic{
					GraphicData: outGraphicData{
						URI:   "http://schemas.openxmlformats.org/drawingml/2006/chart",
						Chart: outChartRef{RID: fmt.Sprintf("rId%d", len(pictures)+j+1)},
					},
				},
			},
			ClientData: outClientData{},
		})
	}

	d := outDrawing{
		A:       "http://schemas.openxmlformats.org/drawingml/2006/main",
		Anchors: anchors,
	}
	return marshalXML(d)
}

// generateDrawingRelsXML produces xl/drawings/_rels/drawingN.xml.rels.
// mediaTargets and chartTargets are part-relative targets ("../media/…",
// "../charts/…") resolved by the saver; images take rId1..len(mediaTargets)
// and charts continue from there, matching the embed ids written by
// generateDrawingXML.
func generateDrawingRelsXML(mediaTargets, chartTargets []string) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")
	rid := 1
	for _, t := range mediaTargets {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="%s"/>`+"\n", rid, t)
		rid++
	}
	for _, t := range chartTargets {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="%s"/>`+"\n", rid, t)
		rid++
	}
	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// generateUnifiedSheetRelsXML produces xl/worksheets/_rels/sheetN.xml.rels.
// The saver passes the resolved package part paths ("xl/tables/…",
// "xl/drawings/…", "xl/pivotTables/…") it actually wrote, so relationship
// targets can never disagree with the parts in the package. Charts are not
// listed here: they are referenced from the drawing part's own rels.
//
// rIds are assigned sequentially — tables first, then the drawing, then
// pivot tables, then external hyperlinks — matching buildTableParts,
// drawingRID, and buildHyperlinks in the sheet XML.
func generateUnifiedSheetRelsXML(tablePaths []string, drawingPath string, pivotPaths []string, hyperlinks []*Hyperlink) string {
	hasExternalLink := false
	for _, h := range hyperlinks {
		if h.Target != "" {
			hasExternalLink = true
			break
		}
	}
	if len(tablePaths) == 0 && drawingPath == "" && len(pivotPaths) == 0 && !hasExternalLink {
		return ""
	}
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")

	rid := 1
	for _, p := range tablePaths {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="%s"/>`+"\n",
			rid, sheetRelTarget(p))
		rid++
	}
	if drawingPath != "" {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="%s"/>`+"\n",
			rid, sheetRelTarget(drawingPath))
		rid++
	}
	for _, p := range pivotPaths {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="%s"/>`+"\n",
			rid, sheetRelTarget(p))
		rid++
	}
	for _, h := range hyperlinks {
		if h.Target == "" {
			continue // internal links carry no relationship
		}
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>`+"\n",
			rid, xmlEscapeAttr(h.Target))
		rid++
	}

	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// xmlEscapeAttr escapes user-provided text (e.g. a hyperlink URL containing
// & or ") for inclusion in a double-quoted XML attribute.
func xmlEscapeAttr(s string) string {
	var b strings.Builder
	if err := xml.EscapeText(&b, []byte(s)); err != nil {
		return s // strings.Builder never fails; keeps the signature simple
	}
	return b.String()
}

// sheetRelTarget rewrites a package part path like "xl/tables/table1.xml"
// into the sheet-relative form "../tables/table1.xml" used in worksheet
// relationship files.
func sheetRelTarget(partPath string) string {
	return "../" + strings.TrimPrefix(partPath, "xl/")
}

// drawingRID returns the rId that the <drawing> element in the sheet XML
// should reference.  Callers only use it for sheets that have pictures, and
// tables occupy rIds 1..N, so the drawing always takes the next one.
func drawingRID(tables []*Table) string {
	return fmt.Sprintf("rId%d", len(tables)+1)
}

// ======================================================================
// Conditional formatting XML generation
// ======================================================================

type outConditionalFormatting struct {
	// CT_ConditionalFormatting addresses its range through sqref.
	Ref   string      `xml:"sqref,attr"`
	Rules []outCFRule `xml:"cfRule"`
}

type outCFRule struct {
	Type       string `xml:"type,attr"`
	Operator   string `xml:"operator,attr,omitempty"`
	Priority   int    `xml:"priority,attr"`
	Rank       int    `xml:"rank,attr,omitempty"`
	Bottom     int    `xml:"bottom,attr,omitempty"`
	Percent    int    `xml:"percent,attr,omitempty"`
	Text       string `xml:"text,attr,omitempty"`
	StopIfTrue int    `xml:"stopIfTrue,attr,omitempty"`
	// DxfID is a pointer so that a rule without any DXF backing omits the
	// attribute entirely instead of emitting a dangling dxfId="0".
	DxfID *int `xml:"dxfId,attr,omitempty"`
	// CT_CfRule children: formula*, then at most one of colorScale,
	// dataBar, iconSet.
	Formula    string         `xml:"formula,omitempty"`
	Formula2   string         `xml:"formula2,omitempty"`
	ColorScale *outColorScale `xml:"colorScale,omitempty"`
	DataBar    *outDataBar    `xml:"dataBar,omitempty"`
	IconSet    *outIconSet    `xml:"iconSet,omitempty"`
}

type outColorScale struct {
	Cfvo   []outCfvo    `xml:"cfvo"`
	Colors []outCFColor `xml:"color"`
}

type outDataBar struct {
	MinLength int `xml:"minLength,attr,omitempty"`
	MaxLength int `xml:"maxLength,attr,omitempty"`
	// ShowValue is always written: the schema default is 1, which a
	// zero-valued Go int would contradict.
	ShowValue int        `xml:"showValue,attr"`
	Cfvo      []outCfvo  `xml:"cfvo"`
	Color     outCFColor `xml:"color"`
}

type outIconSet struct {
	IconSet   string    `xml:"iconSet,attr,omitempty"`
	ShowValue int       `xml:"showValue,attr"`
	Percent   int       `xml:"percent,attr"`
	Reverse   int       `xml:"reverse,attr,omitempty"`
	Cfvo      []outCfvo `xml:"cfvo"`
}

type outCfvo struct {
	Type string `xml:"type,attr"`
	Val  string `xml:"val,attr,omitempty"`
}

type outCFColor struct {
	RGB string `xml:"rgb,attr"`
}

// cfDxfID returns the dxfId attribute value for a rule, or nil when no
// <dxf> entry backs it. Rules with a Style were assigned an id during
// styles.xml generation (see buildDxfTable); rules loaded from a file
// re-emit their original id, which stays valid because the merged dxfs
// block preserves the cached entries at their original indices.
func cfDxfID(rule *ConditionalFormattingRule) *int {
	if rule.Style != nil || rule.hasDxfID {
		id := rule.StyleID
		return &id
	}
	return nil
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
			ocr := outCFRule{
				Type:       rule.Type,
				Operator:   rule.Operator,
				Priority:   rule.Priority,
				Formula:    rule.Formula,
				Formula2:   rule.Formula2,
				Text:       rule.Text,
				StopIfTrue: stopIfTrue,
				DxfID:      cfDxfID(rule),
			}
			switch rule.Type {
			case CondTypeTop10:
				rank := rule.Rank
				if rank == 0 {
					rank = 10
				}
				ocr.Rank = rank
				if rule.Bottom {
					ocr.Bottom = 1
				}
				if rule.Percent {
					ocr.Percent = 1
				}
			case CondTypeColorScale:
				if rule.ColorScale != nil {
					ocr.ColorScale = outColorScaleFrom(rule.ColorScale)
				}
			case CondTypeDataBar:
				if rule.DataBar != nil {
					ocr.DataBar = outDataBarFrom(rule.DataBar)
				}
			case CondTypeIconSet:
				if rule.IconSet != nil {
					ocr.IconSet = outIconSetFrom(rule.IconSet)
				}
			}
			rules[j] = ocr
		}
		out[i] = outConditionalFormatting{
			Ref:   cf.Ref,
			Rules: rules,
		}
	}
	return out
}

func outColorScaleFrom(cs *ColorScale) *outColorScale {
	out := &outColorScale{}
	for _, t := range cs.Thresholds {
		out.Cfvo = append(out.Cfvo, outCfvo{Type: t.Type, Val: t.Val})
	}
	for _, c := range cs.Colors {
		out.Colors = append(out.Colors, outCFColor{RGB: c})
	}
	return out
}

func outDataBarFrom(db *DataBar) *outDataBar {
	out := &outDataBar{
		MinLength: db.MinLength,
		MaxLength: db.MaxLength,
		ShowValue: 1,
		Cfvo: []outCfvo{
			{Type: db.Min.Type, Val: db.Min.Val},
			{Type: db.Max.Type, Val: db.Max.Val},
		},
		Color: outCFColor{RGB: db.Color},
	}
	if db.HideValue {
		out.ShowValue = 0
	}
	return out
}

func outIconSetFrom(is *IconSet) *outIconSet {
	out := &outIconSet{IconSet: is.IconSet, ShowValue: 1, Percent: 1}
	if is.HideValue {
		out.ShowValue = 0
	}
	if is.Reverse {
		out.Reverse = 1
	}
	for _, t := range is.Thresholds {
		out.Cfvo = append(out.Cfvo, outCfvo{Type: t.Type, Val: t.Val})
		// The percent attribute describes the cfvo values; a cutoff given
		// as a number or formula turns it off.
		if t.Type != CondValuePercent {
			out.Percent = 0
		}
	}
	return out
}

// ======================================================================
// Chart XML generation
// ======================================================================

// Fixed axis ids for cartesian charts. Each chart-type group references
// them via the mandatory <c:axId> pair, and the plot area's axis elements
// carry them back.
const (
	chartCatAxisID = "111111111"
	chartValAxisID = "222222222"
)

// generateChartXML produces the content of xl/charts/chartN.xml.
func generateChartXML(chart *Chart, sheetName string) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` + "\n")
	b.WriteString(`  <c:chart>` + "\n")

	if chart.Title != "" {
		b.WriteString("    " + chartTitleXML(chart.Title) + "\n")
	}

	// Plot area with chart type, then the axes the group references.
	b.WriteString(`    <c:plotArea>` + "\n")
	b.WriteString(`      <c:layout/>` + "\n")

	switch chart.Type {
	case ChartTypeBar:
		b.WriteString(`      <c:barChart>` + "\n")
		b.WriteString(`        <c:barDir val="col"/>` + "\n")
		b.WriteString(`        <c:grouping val="clustered"/>` + "\n")
		writeChartSeries(&b, chart, sheetName, false, len(chart.Series))
		writeChartAxisIDs(&b)
		b.WriteString(`      </c:barChart>` + "\n")
		writeChartAxis(&b, "catAx", chartCatAxisID, chartValAxisID, "b", chart.XAxis)
		writeChartAxis(&b, "valAx", chartValAxisID, chartCatAxisID, "l", chart.YAxis)
	case ChartTypeLine:
		b.WriteString(`      <c:lineChart>` + "\n")
		b.WriteString(`        <c:grouping val="standard"/>` + "\n")
		writeChartSeries(&b, chart, sheetName, false, len(chart.Series))
		writeChartAxisIDs(&b)
		b.WriteString(`      </c:lineChart>` + "\n")
		writeChartAxis(&b, "catAx", chartCatAxisID, chartValAxisID, "b", chart.XAxis)
		writeChartAxis(&b, "valAx", chartValAxisID, chartCatAxisID, "l", chart.YAxis)
	case ChartTypeArea:
		b.WriteString(`      <c:areaChart>` + "\n")
		b.WriteString(`        <c:grouping val="standard"/>` + "\n")
		writeChartSeries(&b, chart, sheetName, false, len(chart.Series))
		writeChartAxisIDs(&b)
		b.WriteString(`      </c:areaChart>` + "\n")
		writeChartAxis(&b, "catAx", chartCatAxisID, chartValAxisID, "b", chart.XAxis)
		writeChartAxis(&b, "valAx", chartValAxisID, chartCatAxisID, "l", chart.YAxis)
	case ChartTypeScatter:
		b.WriteString(`      <c:scatterChart>` + "\n")
		b.WriteString(`        <c:scatterStyle val="lineMarker"/>` + "\n")
		writeChartSeries(&b, chart, sheetName, true, len(chart.Series))
		writeChartAxisIDs(&b)
		b.WriteString(`      </c:scatterChart>` + "\n")
		// Scatter plots both axes as value axes (CT_ScatterChart).
		writeChartAxis(&b, "valAx", chartCatAxisID, chartValAxisID, "b", chart.XAxis)
		writeChartAxis(&b, "valAx", chartValAxisID, chartCatAxisID, "l", chart.YAxis)
	case ChartTypePie:
		b.WriteString(`      <c:pieChart>` + "\n")
		writeChartSeries(&b, chart, sheetName, false, 1)
		b.WriteString(`      </c:pieChart>` + "\n")
	}

	b.WriteString(`    </c:plotArea>` + "\n")

	if chart.Legend != nil {
		b.WriteString(chartLegendXML(chart.Legend))
	}

	b.WriteString(`  </c:chart>` + "\n")
	b.WriteString(`</c:chartSpace>` + "\n")
	return b.String()
}

// chartTitleXML renders the rich-text title block shared by chart and axis
// titles (CT_Title).
func chartTitleXML(title string) string {
	return fmt.Sprintf(`<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>%s</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`, xmlEscape(title))
}

// writeChartAxisIDs appends the <c:axId> pair required at the end of every
// cartesian chart group (CT_BarChart, CT_LineChart, CT_AreaChart,
// CT_ScatterChart).
func writeChartAxisIDs(b *strings.Builder) {
	fmt.Fprintf(b, "        <c:axId val=\"%s\"/>\n", chartCatAxisID)
	fmt.Fprintf(b, "        <c:axId val=\"%s\"/>\n", chartValAxisID)
}

// writeChartAxis emits one axis element inside the plot area, after the
// chart-type group. tag is "catAx" or "valAx" and pos the CT_AxPos value
// ("b" bottom, "l" left). A nil axis writes Excel defaults: visible, no
// title.
func writeChartAxis(b *strings.Builder, tag, axID, crossID, pos string, ax *ChartAxis) {
	del := "0"
	title := ""
	if ax != nil {
		if ax.Hidden {
			del = "1"
		}
		if ax.Title != "" {
			title = chartTitleXML(ax.Title)
		}
	}
	// CT_CatAx/CT_ValAx order: axId, scaling, delete, axPos, title?, crossAx.
	fmt.Fprintf(b, "      <c:%s><c:axId val=\"%s\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"%s\"/><c:axPos val=\"%s\"/>%s<c:crossAx val=\"%s\"/></c:%s>\n",
		tag, axID, del, pos, title, crossID, tag)
}

// writeChartSeries emits up to limit series elements. Scatter series carry
// numeric <c:xVal>/<c:yVal> references instead of <c:cat>/<c:val>.
func writeChartSeries(b *strings.Builder, chart *Chart, sheetName string, scatter bool, limit int) {
	sheet := xmlEscape(sheetName)
	for i, series := range chart.Series {
		if i >= limit {
			break
		}
		b.WriteString(fmt.Sprintf(`        <c:ser><c:idx val="%d"/><c:order val="%d"/>`, i, i))
		if series.Name != "" {
			// A literal name belongs in c:v; c:strRef would make Excel
			// resolve it as a cell reference or defined name.
			b.WriteString(fmt.Sprintf(`<c:tx><c:v>%s</c:v></c:tx>`, xmlEscape(series.Name)))
		}
		if scatter {
			if series.Categories != "" {
				b.WriteString(fmt.Sprintf(`<c:xVal><c:numRef><c:f>'%s'!%s</c:f></c:numRef></c:xVal>`, sheet, xmlEscape(absoluteRef(series.Categories))))
			}
			if series.Values != "" {
				b.WriteString(fmt.Sprintf(`<c:yVal><c:numRef><c:f>'%s'!%s</c:f></c:numRef></c:yVal>`, sheet, xmlEscape(absoluteRef(series.Values))))
			}
		} else {
			if series.Categories != "" {
				b.WriteString(fmt.Sprintf(`<c:cat><c:strRef><c:f>'%s'!%s</c:f></c:strRef></c:cat>`, sheet, xmlEscape(absoluteRef(series.Categories))))
			}
			if series.Values != "" {
				b.WriteString(fmt.Sprintf(`<c:val><c:numRef><c:f>'%s'!%s</c:f></c:numRef></c:val>`, sheet, xmlEscape(absoluteRef(series.Values))))
			}
		}
		b.WriteString(`</c:ser>` + "\n")
	}
}

// chartLegendXML renders the legend element that follows plotArea in
// CT_Chart.
func chartLegendXML(legend *ChartLegend) string {
	pos := legend.Position
	switch pos {
	case LegendRight, LegendLeft, LegendTop, LegendBottom:
	default:
		pos = LegendRight
	}
	overlay := "0"
	if legend.Overlay {
		overlay = "1"
	}
	return fmt.Sprintf(`    <c:legend><c:legendPos val="%s"/><c:overlay val="%s"/></c:legend>`+"\n", pos, overlay)
}

// absoluteRef rewrites "A2:B5" as "$A$2:$B$5" so chart references match
// Excel's own output. Anything that is not a plain A1-style reference (or is
// already anchored) is returned unchanged.
func absoluteRef(ref string) string {
	parts := strings.Split(ref, ":")
	for i, p := range parts {
		if strings.HasPrefix(p, "$") {
			return ref
		}
		j := 0
		for j < len(p) && isASCIILetter(p[j]) {
			j++
		}
		if j == 0 || j == len(p) {
			return ref
		}
		parts[i] = "$" + p[:j] + "$" + p[j:]
	}
	return strings.Join(parts, ":")
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
// Pivot Table XML generation
// ======================================================================

type outPivotTableDefinition struct {
	XMLName         xml.Name               `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main pivotTableDefinition"`
	Name            string                 `xml:"name,attr"`
	CacheID         int                    `xml:"cacheId,attr"`
	DataOnRows      int                    `xml:"dataOnRows,attr"`
	DataCaption     string                 `xml:"dataCaption,attr"`
	Location        outPivotLocation       `xml:"location"`
	PivotFields     outPivotFields         `xml:"pivotFields"`
	RowFields       *outPivotRowFields     `xml:"rowFields,omitempty"`
	ColFields       *outPivotColFields     `xml:"colFields,omitempty"`
	DataFields      *outPivotDataFields    `xml:"dataFields,omitempty"`
	PivotTableStyle outPivotTableStyleInfo `xml:"pivotTableStyleInfo"`
}

type outPivotLocation struct {
	Ref         string `xml:"ref,attr"`
	FirstRowCol int    `xml:"firstRowCol,attr"`
}

type outPivotFields struct {
	Count  int             `xml:"count,attr"`
	Fields []outPivotField `xml:"pivotField"`
}

// outPivotField mirrors CT_pivotField, which carries no name attribute:
// field names live in the cache definition's <cacheField> elements and are
// referenced from the pivot table by index.
type outPivotField struct {
	Axis    string `xml:"axis,attr,omitempty"`
	ShowAll int    `xml:"showAll,attr"`
}

type outPivotRowFields struct {
	Count  int                `xml:"count,attr"`
	Fields []outPivotFieldRef `xml:"field"`
}

type outPivotColFields struct {
	Count  int                `xml:"count,attr"`
	Fields []outPivotFieldRef `xml:"field"`
}

type outPivotFieldRef struct {
	X int `xml:"x,attr"`
}

type outPivotDataFields struct {
	Count  int                 `xml:"count,attr"`
	Fields []outPivotDataField `xml:"dataField"`
}

type outPivotDataField struct {
	Name     string `xml:"name,attr"`
	Field    int    `xml:"fld,attr"`
	Subtotal string `xml:"subtotal,attr,omitempty"`
}

type outPivotTableStyleInfo struct {
	Name           string `xml:"name,attr"`
	ShowRowHeaders int    `xml:"showRowHeaders,attr"`
	ShowColHeaders int    `xml:"showColHeaders,attr"`
	ShowRowStripes int    `xml:"showRowStripes,attr"`
	ShowColStripes int    `xml:"showColStripes,attr"`
}

// generatePivotTableXML produces the content of xl/pivotTables/pivotTableN.xml.
func generatePivotTableXML(pt *PivotTable, ws *Worksheet, cacheID int) (string, error) {
	// Field names come from the source range's header row; row/col/data
	// fields reference them by source column index.
	headers := pivotSourceHeaders(ws, pt.SourceRef)
	colCount := len(headers)
	fieldIdx := make(map[string]int, colCount)
	for i, h := range headers {
		if _, ok := fieldIdx[h]; !ok {
			fieldIdx[h] = i
		}
	}
	// resolve maps a field name to its source column index, falling back to
	// the position in the user's list when the name is not a source header.
	resolve := func(name string, pos int) int {
		if idx, ok := fieldIdx[name]; ok {
			return idx
		}
		if pos < colCount {
			return pos
		}
		return 0
	}

	fields := make([]outPivotField, colCount)
	for i := range fields {
		fields[i] = outPivotField{ShowAll: 0}
	}

	// Row and column field references, with the matching axis marker on the
	// pivotFields entry they point at.
	buildRefs := func(names []string, axis string) []outPivotFieldRef {
		refs := make([]outPivotFieldRef, 0, len(names))
		for i, name := range names {
			idx := resolve(name, i)
			fields[idx].Axis = axis
			refs = append(refs, outPivotFieldRef{X: idx})
		}
		return refs
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
		rowFieldRefs := buildRefs(pt.RowFields, "axisRow")
		pivotDef.RowFields = &outPivotRowFields{
			Count:  len(rowFieldRefs),
			Fields: rowFieldRefs,
		}
	}

	// Add col fields
	if len(pt.ColFields) > 0 {
		colFieldRefs := buildRefs(pt.ColFields, "axisCol")
		pivotDef.ColFields = &outPivotColFields{
			Count:  len(colFieldRefs),
			Fields: colFieldRefs,
		}
	}

	// Add data fields
	if len(pt.DataFields) > 0 {
		dataFields := make([]outPivotDataField, len(pt.DataFields))
		for i, df := range pt.DataFields {
			// The name attribute is the caption shown in the pivot; the fld
			// attribute points at the aggregated source column.
			name := df.DisplayName
			if name == "" {
				name = df.Name
			}
			dataFields[i] = outPivotDataField{
				Name:     name,
				Field:    resolve(df.Name, i),
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

// pivotSourceHeaders returns the field names of a pivot source range: the
// cell texts of its first row, falling back to "Field<N>" for columns
// without a usable header.
func pivotSourceHeaders(ws *Worksheet, sourceRef string) []string {
	n := rangeColumnCount(sourceRef)
	if n < 1 {
		n = 1
	}
	headers := make([]string, n)
	for i := range headers {
		headers[i] = fmt.Sprintf("Field%d", i)
	}
	parts := splitRange(sourceRef)
	if ws == nil || len(parts) == 0 {
		return headers
	}
	startCol, startRow := splitRef(parts[0])
	if startRow < 1 {
		return headers
	}
	first := colToNum(startCol)
	for i := range headers {
		cell, err := ws.Cells().Get(fmt.Sprintf("%s%d", numToCol(first+i), startRow))
		if err != nil {
			continue
		}
		if s := strings.TrimSpace(cell.AsString()); s != "" {
			headers[i] = s
		}
	}
	return headers
}

// generatePivotCacheDefinitionXML produces xl/pivotCache/pivotCacheDefinitionN.xml.
// ws is the worksheet the source range lives on (its name goes into
// <worksheetSource> and its header row provides the cache field names); the
// r:id="rId1" attribute points at the cache records part through the
// definition's rels.
func generatePivotCacheDefinitionXML(pt *PivotTable, ws *Worksheet) string {
	headers := pivotSourceHeaders(ws, pt.SourceRef)
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` + "\n")
	b.WriteString(`  <cacheSource type="worksheet">` + "\n")
	b.WriteString(fmt.Sprintf(`    <worksheetSource ref="%s" sheet="%s"/>`+"\n", xmlEscape(pt.SourceRef), xmlEscape(ws.Name)))
	b.WriteString(`  </cacheSource>` + "\n")
	b.WriteString(fmt.Sprintf(`  <cacheFields count="%d">`+"\n", len(headers)))
	for _, h := range headers {
		b.WriteString(fmt.Sprintf(`    <cacheField name="%s" numFmtId="0">`+"\n", xmlEscape(h)))
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

// pivotTableRelsXML produces xl/pivotTables/_rels/pivotTableN.xml.rels,
// linking the pivot table to its cache definition. cacheDefTarget is
// relative to the pivotTables folder ("../pivotCache/…").
func pivotTableRelsXML(cacheDefTarget string) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")
	fmt.Fprintf(&b, `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="%s"/>`+"\n", cacheDefTarget)
	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// pivotCacheRelsXML produces xl/pivotCache/_rels/pivotCacheDefinitionN.xml.rels,
// backing the r:id="rId1" attribute on the cache definition. recordsTarget
// is relative to the pivotCache folder ("pivotCacheRecordsN.xml").
func pivotCacheRelsXML(recordsTarget string) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")
	fmt.Fprintf(&b, `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="%s"/>`+"\n", recordsTarget)
	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}
