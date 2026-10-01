package cells_foss

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"path"
	"strconv"
	"strings"
)

// ======================================================================
// Internal XML types for decoding styles.xml
// ======================================================================

type xmlStyleSheet struct {
	XMLName      xml.Name   `xml:"styleSheet"`
	Fonts        xmlFonts   `xml:"fonts"`
	Fills        xmlFills   `xml:"fills"`
	Borders      xmlBorders `xml:"borders"`
	CellStyleXfs xmlXfs     `xml:"cellStyleXfs"`
	CellXfs      xmlXfs     `xml:"cellXfs"`
}

type xmlFonts struct {
	Count int       `xml:"count,attr"`
	Fonts []xmlFont `xml:"font"`
}

type xmlFont struct {
	Name   xmlVal    `xml:"name"`
	Size   xmlVal    `xml:"sz"`
	Color  *xmlColor `xml:"color"`
	Bold   *xmlEmpty `xml:"b"`
	Italic *xmlEmpty `xml:"i"`
}

type xmlFills struct {
	Count int       `xml:"count,attr"`
	Fills []xmlFill `xml:"fill"`
}

type xmlFill struct {
	PatternFill *xmlPatternFill `xml:"patternFill"`
}

type xmlPatternFill struct {
	PatternType string    `xml:"patternType,attr"`
	FgColor     *xmlColor `xml:"fgColor"`
}

type xmlBorders struct {
	Count   int         `xml:"count,attr"`
	Borders []xmlBorder `xml:"border"`
}

type xmlBorder struct {
	Left   *xmlBorderSide `xml:"left"`
	Right  *xmlBorderSide `xml:"right"`
	Top    *xmlBorderSide `xml:"top"`
	Bottom *xmlBorderSide `xml:"bottom"`
}

type xmlBorderSide struct {
	Style string `xml:"style,attr,omitempty"`
}

type xmlXfs struct {
	Count int     `xml:"count,attr"`
	Xfs   []xmlXf `xml:"xf"`
}

type xmlXf struct {
	NumFmtId       int       `xml:"numFmtId,attr"`
	FontId         int       `xml:"fontId,attr"`
	FillId         int       `xml:"fillId,attr"`
	BorderId       int       `xml:"borderId,attr"`
	XfId           int       `xml:"xfId,attr"`
	ApplyFont      int       `xml:"applyFont,attr,omitempty"`
	ApplyFill      int       `xml:"applyFill,attr,omitempty"`
	ApplyBorder    int       `xml:"applyBorder,attr,omitempty"`
	ApplyAlignment int       `xml:"applyAlignment,attr,omitempty"`
	Alignment      *xmlAlign `xml:"alignment"`
}

type xmlAlign struct {
	Horizontal string `xml:"horizontal,attr,omitempty"`
	Vertical   string `xml:"vertical,attr,omitempty"`
	WrapText   int    `xml:"wrapText,attr,omitempty"`
}

type xmlColor struct {
	RGB     string `xml:"rgb,attr,omitempty"`
	Theme   int    `xml:"theme,attr,omitempty"`
	Indexed int    `xml:"indexed,attr,omitempty"`
}

type xmlVal struct {
	Val string `xml:"val,attr"`
}

type xmlEmpty struct{}

// parseFonts converts the parsed <font> elements into Font objects.
func parseFonts(fonts []xmlFont) []*Font {
	out := make([]*Font, len(fonts))
	for i, f := range fonts {
		ft := &Font{
			Name:   f.Name.Val,
			Bold:   f.Bold != nil,
			Italic: f.Italic != nil,
		}
		if f.Size.Val != "" {
			if sz, err := strconv.ParseFloat(f.Size.Val, 64); err == nil {
				ft.Size = sz
			}
		}
		if f.Color != nil {
			ft.Color = resolveColor(*f.Color)
		}
		out[i] = ft
	}
	return out
}

// parseFills converts the parsed <fill> elements into Fill objects.
func parseFills(fills []xmlFill) []*Fill {
	out := make([]*Fill, len(fills))
	for i, f := range fills {
		fl := &Fill{Type: FillTypeNone}
		if f.PatternFill != nil {
			fl.Type = f.PatternFill.PatternType
			if f.PatternFill.FgColor != nil {
				fl.Color = resolveColor(*f.PatternFill.FgColor)
			}
		}
		out[i] = fl
	}
	return out
}

// parseBorders converts the parsed <border> elements into Border objects.
func parseBorders(borders []xmlBorder) []*Border {
	out := make([]*Border, len(borders))
	for i, b := range borders {
		bd := &Border{
			Left:   hasBorderSide(b.Left),
			Right:  hasBorderSide(b.Right),
			Top:    hasBorderSide(b.Top),
			Bottom: hasBorderSide(b.Bottom),
		}
		out[i] = bd
	}
	return out
}

// hasBorderSide returns true when the side has a non-empty style attribute.
func hasBorderSide(s *xmlBorderSide) bool {
	return s != nil && s.Style != ""
}

// resolveColor picks the best colour representation from an xmlColor.
func resolveColor(c xmlColor) string {
	if c.RGB != "" {
		return c.RGB
	}
	if c.Theme != 0 {
		return fmt.Sprintf("theme:%d", c.Theme)
	}
	if c.Indexed != 0 {
		return fmt.Sprintf("indexed:%d", c.Indexed)
	}
	return ""
}

// ======================================================================
// loadStyles
// ======================================================================

// loadStyles reads xl/styles.xml from the archive and populates the
// Workbook's style registry.  It returns the raw bytes for caching so
// unmodified workbooks can reuse the original XML.
//
// When xl/styles.xml is absent the workbook is seeded with the single
// default style at index 0.
func loadStyles(zr *zip.Reader, wb *Workbook) ([]byte, error) {
	raw, err := readZipFile(zr, "xl/styles.xml")
	if err != nil {
		// No styles.xml — seed the default style only.
		wb.styles = []*Style{DefaultStyle()}
		return nil, nil
	}

	var ss xmlStyleSheet
	if err := xml.Unmarshal(raw, &ss); err != nil {
		return nil, fmt.Errorf("cannot parse styles.xml: %w", err)
	}

	fonts := parseFonts(ss.Fonts.Fonts)
	fills := parseFills(ss.Fills.Fills)
	borders := parseBorders(ss.Borders.Borders)

	styles := make([]*Style, 0, len(ss.CellXfs.Xfs))
	for _, xf := range ss.CellXfs.Xfs {
		st := DefaultStyle()

		if xf.FontId >= 0 && xf.FontId < len(fonts) {
			st.Font = fonts[xf.FontId]
		}
		if xf.FillId >= 0 && xf.FillId < len(fills) {
			st.Fill = fills[xf.FillId]
		}
		if xf.BorderId >= 0 && xf.BorderId < len(borders) {
			st.Border = borders[xf.BorderId]
		}
		if xf.Alignment != nil {
			st.Alignment = &Alignment{
				Horizontal: xf.Alignment.Horizontal,
				Vertical:   xf.Alignment.Vertical,
				WrapText:   xf.Alignment.WrapText == 1,
			}
		}

		styles = append(styles, st)
	}

	if len(styles) == 0 {
		styles = append(styles, DefaultStyle())
	}
	wb.styles = styles

	return raw, nil
}

// ======================================================================
// Table loading
// ======================================================================

type xmlTable struct {
	XMLName        xml.Name `xml:"table"`
	ID             string   `xml:"id,attr"`
	Name           string   `xml:"name,attr"`
	DisplayName    string   `xml:"displayName,attr"`
	Ref            string   `xml:"ref,attr"`
	HeaderRowCount int      `xml:"headerRowCount,attr"`
}

// loadTables reads the table definitions for the given sheet index.  It looks
// for the sheet's rels file, discovers table relationships, then parses the
// corresponding table XML files.  A missing rels file means the sheet has no
// tables; individual unresolvable tables are skipped.
func loadTables(zr *zip.Reader, sheetIndex int) ([]*Table, error) {
	relsPath := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetIndex+1)
	relsRaw, err := readZipFile(zr, relsPath)
	if err != nil {
		return nil, nil
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(relsRaw, &rels); err != nil {
		return nil, fmt.Errorf("cannot parse %s: %w", relsPath, err)
	}

	var tables []*Table
	for _, rel := range rels.Rels {
		if !strings.Contains(rel.Target, "tables/") {
			continue
		}

		// Targets are relative to the rels file's directory, e.g.
		// "../tables/table1.xml" resolves to "xl/tables/table1.xml".
		tablePath := path.Clean(path.Join("xl/worksheets", rel.Target))
		raw, err := readZipFile(zr, tablePath)
		if err != nil {
			continue
		}

		var xt xmlTable
		if err := xml.Unmarshal(raw, &xt); err != nil {
			continue
		}

		tables = append(tables, &Table{
			Name:         xt.Name,
			Range:        xt.Ref,
			HasHeaderRow: xt.HeaderRowCount > 0,
			StyleName:    "TableStyleMedium9", // default; style parsing deferred
		})
	}
	return tables, nil
}

// ======================================================================
// Internal XML types for decoding drawing XML
// ======================================================================

type xmlDrawingWS struct {
	XMLName  xml.Name    `xml:"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing wsDr"`
	Anchors  []xmlAnchor `xml:"twoCellAnchor"`
	Anchors2 []xmlAnchor `xml:"oneCellAnchor"`
}

type xmlAnchor struct {
	From xmlPos `xml:"from"`
	To   xmlPos `xml:"to"`
	Pic  xmlPic `xml:"pic"`
}

type xmlPos struct {
	Col    int   `xml:"col"`
	ColOff int64 `xml:"colOff"`
	Row    int   `xml:"row"`
	RowOff int64 `xml:"rowOff"`
}

type xmlPic struct {
	NvPicPr  xmlNvPicPr  `xml:"nvPicPr"`
	BlipFill xmlBlipFill `xml:"blipFill"`
	SpPr     xmlSpPr     `xml:"spPr"`
}

type xmlNvPicPr struct {
	CNvPr xmlCNvPr `xml:"cNvPr"`
}

type xmlCNvPr struct {
	ID   int    `xml:"id,attr"`
	Name string `xml:"name,attr"`
}

type xmlBlipFill struct {
	Blip xmlBlip `xml:"http://schemas.openxmlformats.org/drawingml/2006/main blip"`
}

type xmlBlip struct {
	Embed string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr"`
}

type xmlSpPr struct {
	Xfrm xmlXfrm `xml:"http://schemas.openxmlformats.org/drawingml/2006/main xfrm"`
}

type xmlXfrm struct {
	Off xmlPoint `xml:"off"`
	Ext xmlPoint `xml:"ext"`
}

type xmlPoint struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// loadDrawings reads the drawing XML and relationships for the given sheet
// index, then loads the referenced images from xl/media/.
func loadDrawings(zr *zip.Reader, sheetIndex int) ([]*Picture, error) {
	// 1. Find the drawing relationship in sheet rels.
	relsPath := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetIndex+1)
	relsRaw, err := readZipFile(zr, relsPath)
	if err != nil {
		return nil, nil // No rels means no drawings
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(relsRaw, &rels); err != nil {
		return nil, fmt.Errorf("cannot parse %s: %w", relsPath, err)
	}

	// Find the drawing relationship.
	var drawingRID, drawingTarget string
	for _, rel := range rels.Rels {
		if strings.Contains(rel.Target, "drawings/") {
			drawingRID = rel.ID
			drawingTarget = rel.Target
			break
		}
	}
	if drawingRID == "" {
		return nil, nil // No drawing relationship
	}

	// 2. Parse the drawing XML.
	drawingPath := path.Clean(path.Join("xl/worksheets", drawingTarget))
	drawingRaw, err := readZipFile(zr, drawingPath)
	if err != nil {
		return nil, fmt.Errorf("cannot read %s: %w", drawingPath, err)
	}

	var drawing xmlDrawingWS
	if err := xml.Unmarshal(drawingRaw, &drawing); err != nil {
		return nil, fmt.Errorf("cannot parse %s: %w", drawingPath, err)
	}

	// 3. Parse drawing relationships to map rId to media paths.
	drawingRelsPath := path.Clean(path.Join(path.Dir(drawingPath), "_rels", path.Base(drawingPath)+".rels"))
	drawingRelsRaw, err := readZipFile(zr, drawingRelsPath)
	if err != nil {
		return nil, fmt.Errorf("cannot read %s: %w", drawingRelsPath, err)
	}

	var drawingRels xmlRelationships
	if err := xml.Unmarshal(drawingRelsRaw, &drawingRels); err != nil {
		return nil, fmt.Errorf("cannot parse %s: %w", drawingRelsPath, err)
	}

	mediaPathByRID := make(map[string]string)
	for _, rel := range drawingRels.Rels {
		if strings.Contains(rel.Target, "media/") {
			mediaPath := path.Clean(path.Join(path.Dir(drawingPath), rel.Target))
			mediaPathByRID[rel.ID] = mediaPath
		}
	}

	// 4. Load each picture.
	var pictures []*Picture
	anchors := append(drawing.Anchors, drawing.Anchors2...)
	for _, anchor := range anchors {
		// Get the image relationship ID.
		embedRID := anchor.Pic.BlipFill.Blip.Embed
		if embedRID == "" {
			continue
		}

		// Get the media path.
		mediaPath, ok := mediaPathByRID[embedRID]
		if !ok {
			continue
		}

		// Load the image data.
		imageData, err := readZipFile(zr, mediaPath)
		if err != nil {
			continue
		}

		// Determine format from extension.
		format := "png"
		if strings.HasSuffix(strings.ToLower(mediaPath), ".jpg") || strings.HasSuffix(strings.ToLower(mediaPath), ".jpeg") {
			format = "jpeg"
		}

		// Calculate width/height from EMUs.
		width := int(anchor.Pic.SpPr.Xfrm.Ext.X / emuPerPixel)
		height := int(anchor.Pic.SpPr.Xfrm.Ext.Y / emuPerPixel)

		pic := &Picture{
			Data:   imageData,
			Format: format,
			Row:    anchor.From.Row,
			Col:    anchor.From.Col,
			RowOff: anchor.From.RowOff,
			ColOff: anchor.From.ColOff,
			Width:  width,
			Height: height,
			Name:   anchor.Pic.NvPicPr.CNvPr.Name,
		}
		pictures = append(pictures, pic)
	}

	return pictures, nil
}

// ======================================================================
// Internal XML types for decoding conditional formatting
// ======================================================================

type xmlConditionalFormatting struct {
	Ref   string      `xml:"ref,attr"`
	Rules []xmlCFRule `xml:"cfRule"`
}

type xmlCFRule struct {
	Type       string `xml:"type,attr"`
	Operator   string `xml:"operator,attr,omitempty"`
	Priority   int    `xml:"priority,attr"`
	Formula    string `xml:"formula,omitempty"`
	Formula2   string `xml:"formula2,omitempty"`
	Text       string `xml:"text,attr,omitempty"`
	StopIfTrue int    `xml:"stopIfTrue,attr,omitempty"`
	DxfID      int    `xml:"dxfId,attr,omitempty"`
}

// loadConditionalFormattings reads the conditional formatting rules for the
// given sheet from the worksheet XML.
func loadConditionalFormattings(ws *Worksheet, rawXML []byte) error {
	// Parse the worksheet XML to extract conditional formatting.
	var sheet struct {
		ConditionalFormattings []xmlConditionalFormatting `xml:"conditionalFormatting"`
	}
	if err := xml.Unmarshal(rawXML, &sheet); err != nil {
		return nil // Not an error; sheet may not have conditional formatting
	}

	for _, xcf := range sheet.ConditionalFormattings {
		cf := &ConditionalFormatting{
			Ref: xcf.Ref,
		}
		for _, xrule := range xcf.Rules {
			rule := &ConditionalFormattingRule{
				Type:       xrule.Type,
				Operator:   xrule.Operator,
				Priority:   xrule.Priority,
				Formula:    xrule.Formula,
				Formula2:   xrule.Formula2,
				Text:       xrule.Text,
				StopIfTrue: xrule.StopIfTrue != 0,
				StyleID:    xrule.DxfID,
			}
			cf.Rules = append(cf.Rules, rule)
		}
		ws.ConditionalFormattings = append(ws.ConditionalFormattings, cf)
	}
	return nil
}
