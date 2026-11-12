package cells_foss

import (
	"archive/zip"
	"bytes"
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
	NumFmts      xmlNumFmts `xml:"numFmts"`
	Fonts        xmlFonts   `xml:"fonts"`
	Fills        xmlFills   `xml:"fills"`
	Borders      xmlBorders `xml:"borders"`
	CellStyleXfs xmlXfs     `xml:"cellStyleXfs"`
	CellXfs      xmlXfs     `xml:"cellXfs"`
}

type xmlNumFmts struct {
	Count   int         `xml:"count,attr"`
	NumFmts []xmlNumFmt `xml:"numFmt"`
}

type xmlNumFmt struct {
	NumFmtId   int    `xml:"numFmtId,attr"`
	FormatCode string `xml:"formatCode,attr"`
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

	// numFmtId -> format code. The file's own <numFmts> entries override
	// the standard built-in ids; built-ins cover ids the file does not
	// declare explicitly.
	numFmtCodes := make(map[int]string, len(builtinNumFmtCodes)+len(ss.NumFmts.NumFmts))
	for id, code := range builtinNumFmtCodes {
		numFmtCodes[id] = code
	}
	for _, nf := range ss.NumFmts.NumFmts {
		numFmtCodes[nf.NumFmtId] = nf.FormatCode
	}

	styles := make([]*Style, 0, len(ss.CellXfs.Xfs))
	for _, xf := range ss.CellXfs.Xfs {
		st := DefaultStyle()

		if code := numFmtCodes[xf.NumFmtId]; code != "" && !strings.EqualFold(code, NumFmtGeneral) {
			st.NumberFormat = code
		}
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

	// Cache the original <dxfs> block verbatim. When styles.xml has to be
	// regenerated these entries are kept at their original indices so
	// dxfIds in cached sheet XML and in loaded conditional-formatting
	// rules stay valid.
	cacheDxfBlock(wb, raw)

	return raw, nil
}

// cacheDxfBlock extracts the inner XML of the <dxfs> block from a styles
// part and caches it on the workbook as the baseline for future style-sheet
// regeneration. A styles part without dxfs clears the baseline.
func cacheDxfBlock(wb *Workbook, stylesXML []byte) {
	wb.dxfRaw = ""
	wb.dxfBaseCount = 0

	start := bytes.Index(stylesXML, []byte("<dxfs"))
	if start < 0 {
		return
	}
	gt := bytes.IndexByte(stylesXML[start:], '>')
	if gt < 0 {
		return
	}
	inner := start + gt + 1
	end := bytes.Index(stylesXML[inner:], []byte("</dxfs>"))
	if end < 0 {
		return
	}
	block := string(stylesXML[inner : inner+end])
	wb.dxfRaw = block
	wb.dxfBaseCount = strings.Count(block, "<dxf")
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
	From         xmlPos           `xml:"from"`
	To           xmlPos           `xml:"to"`
	Pic          xmlPic           `xml:"pic"`
	GraphicFrame *xmlGraphicFrame `xml:"graphicFrame"`
}

// xmlGraphicFrame parses the drawing-level chart reference
// (CT_GraphicalObjectFrame) including its transform for size restoration.
type xmlGraphicFrame struct {
	Xfrm    xmlGFXfrm    `xml:"xfrm"`
	Graphic xmlGraphicEl `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphic"`
}

type xmlGFXfrm struct {
	Ext xmlExtent `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ext"`
}

type xmlGraphicEl struct {
	GraphicData xmlGraphicDataEl `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphicData"`
}

type xmlGraphicDataEl struct {
	Chart xmlChartRefEl `xml:"http://schemas.openxmlformats.org/drawingml/2006/chart chart"`
}

type xmlChartRefEl struct {
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
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
	Off xmlPoint  `xml:"off"`
	Ext xmlExtent `xml:"ext"`
}

type xmlPoint struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// xmlExtent parses a:ext. Spec-compliant files use cx/cy attributes; files
// produced by older versions of this library used x/y. Accept both.
type xmlExtent struct {
	CX int64 `xml:"cx,attr"`
	CY int64 `xml:"cy,attr"`
	X  int64 `xml:"x,attr"`
	Y  int64 `xml:"y,attr"`
}

func (e xmlExtent) width() int64 {
	if e.CX != 0 {
		return e.CX
	}
	return e.X
}

func (e xmlExtent) height() int64 {
	if e.CY != 0 {
		return e.CY
	}
	return e.Y
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
		width := int(anchor.Pic.SpPr.Xfrm.Ext.width() / emuPerPixel)
		height := int(anchor.Pic.SpPr.Xfrm.Ext.height() / emuPerPixel)

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
	// Excel writes sqref; files from older library versions wrote ref.
	Sqref string      `xml:"sqref,attr"`
	Ref   string      `xml:"ref,attr"`
	Rules []xmlCFRule `xml:"cfRule"`
}

func (x xmlConditionalFormatting) rangeRef() string {
	if x.Sqref != "" {
		return x.Sqref
	}
	return x.Ref
}

type xmlCFRule struct {
	Type       string `xml:"type,attr"`
	Operator   string `xml:"operator,attr,omitempty"`
	Priority   int    `xml:"priority,attr"`
	Rank       int    `xml:"rank,attr"`
	Bottom     int    `xml:"bottom,attr"`
	Percent    int    `xml:"percent,attr"`
	Text       string `xml:"text,attr,omitempty"`
	StopIfTrue int    `xml:"stopIfTrue,attr,omitempty"`
	// DxfID is a pointer so a literal dxfId="0" is distinguishable from a
	// missing attribute.
	DxfID      *int           `xml:"dxfId,attr"`
	Formula    string         `xml:"formula,omitempty"`
	Formula2   string         `xml:"formula2,omitempty"`
	ColorScale *xmlColorScale `xml:"colorScale"`
	DataBar    *xmlDataBar    `xml:"dataBar"`
	IconSet    *xmlIconSet    `xml:"iconSet"`
}

type xmlColorScale struct {
	Cfvo   []xmlCfvo    `xml:"cfvo"`
	Colors []xmlCFColor `xml:"color"`
}

type xmlDataBar struct {
	MinLength int        `xml:"minLength,attr"`
	MaxLength int        `xml:"maxLength,attr"`
	ShowValue *int       `xml:"showValue,attr"`
	Cfvo      []xmlCfvo  `xml:"cfvo"`
	Color     xmlCFColor `xml:"color"`
}

type xmlIconSet struct {
	IconSet   string    `xml:"iconSet,attr"`
	ShowValue *int      `xml:"showValue,attr"`
	Reverse   int       `xml:"reverse,attr"`
	Cfvo      []xmlCfvo `xml:"cfvo"`
}

type xmlCfvo struct {
	Type string `xml:"type,attr"`
	Val  string `xml:"val,attr"`
}

type xmlCFColor struct {
	RGB string `xml:"rgb,attr"`
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
			Ref: xcf.rangeRef(),
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
				Rank:       xrule.Rank,
				Bottom:     xrule.Bottom != 0,
				Percent:    xrule.Percent != 0,
			}
			if xrule.DxfID != nil {
				rule.StyleID = *xrule.DxfID
				rule.hasDxfID = true
			}
			if xrule.ColorScale != nil {
				cs := &ColorScale{}
				for _, c := range xrule.ColorScale.Cfvo {
					cs.Thresholds = append(cs.Thresholds, CondThreshold{Type: c.Type, Val: c.Val})
				}
				for _, c := range xrule.ColorScale.Colors {
					cs.Colors = append(cs.Colors, c.RGB)
				}
				rule.ColorScale = cs
			}
			if xrule.DataBar != nil && len(xrule.DataBar.Cfvo) >= 2 {
				db := &DataBar{
					Min:       CondThreshold{Type: xrule.DataBar.Cfvo[0].Type, Val: xrule.DataBar.Cfvo[0].Val},
					Max:       CondThreshold{Type: xrule.DataBar.Cfvo[1].Type, Val: xrule.DataBar.Cfvo[1].Val},
					Color:     xrule.DataBar.Color.RGB,
					MinLength: xrule.DataBar.MinLength,
					MaxLength: xrule.DataBar.MaxLength,
					HideValue: xrule.DataBar.ShowValue != nil && *xrule.DataBar.ShowValue == 0,
				}
				rule.DataBar = db
			}
			if xrule.IconSet != nil {
				is := &IconSet{
					IconSet:   xrule.IconSet.IconSet,
					HideValue: xrule.IconSet.ShowValue != nil && *xrule.IconSet.ShowValue == 0,
					Reverse:   xrule.IconSet.Reverse != 0,
				}
				if is.IconSet == "" {
					is.IconSet = "3TrafficLights1" // schema default
				}
				for _, c := range xrule.IconSet.Cfvo {
					is.Thresholds = append(is.Thresholds, CondThreshold{Type: c.Type, Val: c.Val})
				}
				rule.IconSet = is
			}
			cf.Rules = append(cf.Rules, rule)
		}
		ws.ConditionalFormattings = append(ws.ConditionalFormattings, cf)
	}
	return nil
}

// ======================================================================
// Internal XML types for decoding chart XML
// ======================================================================

type xmlChartSpace struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/chart chartSpace"`
	Chart   xmlChart `xml:"chart"`
}

type xmlChart struct {
	Title    *xmlChartTitle `xml:"title,omitempty"`
	PlotArea xmlPlotArea    `xml:"plotArea"`
	Legend   *xmlLegend     `xml:"legend,omitempty"`
}

// xmlLegend decodes CT_Legend; only position and overlay are recovered.
type xmlLegend struct {
	Pos     *xmlChartStrAttr  `xml:"legendPos,omitempty"`
	Overlay *xmlChartBoolAttr `xml:"overlay,omitempty"`
}

// xmlChartAxis decodes CT_CatAx/CT_ValAx; only the fields the model carries
// are recovered.
type xmlChartAxis struct {
	Delete *xmlChartBoolAttr `xml:"delete,omitempty"`
	AxPos  *xmlChartStrAttr  `xml:"axPos,omitempty"`
	Title  *xmlChartTitle    `xml:"title,omitempty"`
}

type xmlChartStrAttr struct {
	Val string `xml:"val,attr"`
}

type xmlChartBoolAttr struct {
	Val bool `xml:"val,attr"`
}

type xmlChartTitle struct {
	Tx xmlChartTx `xml:"tx"`
}

type xmlChartTx struct {
	Rich xmlChartRich `xml:"rich"`
}

type xmlChartRich struct {
	P xmlChartP `xml:"p"`
}

type xmlChartP struct {
	R []xmlChartR `xml:"r"`
}

type xmlChartR struct {
	T string `xml:"t"`
}

type xmlPlotArea struct {
	BarChart     *xmlBarChart     `xml:"barChart,omitempty"`
	LineChart    *xmlLineChart    `xml:"lineChart,omitempty"`
	PieChart     *xmlPieChart     `xml:"pieChart,omitempty"`
	ScatterChart *xmlScatterChart `xml:"scatterChart,omitempty"`
	AreaChart    *xmlAreaChart    `xml:"areaChart,omitempty"`
	CatAx        []xmlChartAxis   `xml:"catAx"`
	ValAx        []xmlChartAxis   `xml:"valAx"`
}

type xmlBarChart struct {
	Series []xmlChartSeries `xml:"ser"`
}

type xmlLineChart struct {
	Series []xmlChartSeries `xml:"ser"`
}

type xmlPieChart struct {
	Series []xmlChartSeries `xml:"ser"`
}

type xmlAreaChart struct {
	Series []xmlChartSeries `xml:"ser"`
}

// xmlScatterChart decodes CT_ScatterChart; scatter series reference their
// data through xVal/yVal instead of cat/val.
type xmlScatterChart struct {
	Series []xmlScatterSeries `xml:"ser"`
}

type xmlScatterSeries struct {
	Tx   xmlChartSeriesTx  `xml:"tx,omitempty"`
	XVal xmlChartSeriesRef `xml:"xVal,omitempty"`
	YVal xmlChartSeriesRef `xml:"yVal,omitempty"`
}

type xmlChartSeries struct {
	Tx  xmlChartSeriesTx  `xml:"tx,omitempty"`
	Cat xmlChartSeriesRef `xml:"cat,omitempty"`
	Val xmlChartSeriesRef `xml:"val,omitempty"`
}

// xmlChartSeriesTx decodes CT_SerTx: a literal name in c:v, or a cell
// reference in c:strRef.
type xmlChartSeriesTx struct {
	StrRef xmlChartStrRef `xml:"strRef"`
	V      string         `xml:"v"`
}

type xmlChartSeriesRef struct {
	StrRef xmlChartStrRef `xml:"strRef,omitempty"`
	NumRef xmlChartNumRef `xml:"numRef,omitempty"`
}

type xmlChartStrRef struct {
	F string `xml:"f"`
}

type xmlChartNumRef struct {
	F string `xml:"f"`
}

// chartRef is a discovered chart part plus the placement recovered from its
// drawing graphicFrame, when available.
type chartRef struct {
	path   string
	anchor *ChartAnchor
	width  int
	height int
}

// loadCharts reads the chart definitions for the given sheet index.
//
// Spec-compliant packages reference charts from the sheet's drawing part
// through a graphicFrame; files written by older versions of this library
// referenced them directly from the sheet rels. Both layouts are accepted.
// Charts reached through a drawing also recover their anchor and size from
// the graphicFrame.
func loadCharts(zr *zip.Reader, sheetIndex int) ([]*Chart, error) {
	// 1. Find the chart relationships reachable from the sheet.
	relsPath := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetIndex+1)
	relsRaw, err := readZipFile(zr, relsPath)
	if err != nil {
		return nil, nil // No rels means no charts
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(relsRaw, &rels); err != nil {
		return nil, fmt.Errorf("cannot parse %s: %w", relsPath, err)
	}

	var refs []chartRef
	seen := make(map[string]bool)

	// Direct sheet-rels references (legacy layout written by this library).
	for _, rel := range rels.Rels {
		if rel.TargetMode == "External" || !strings.Contains(rel.Target, "charts/") {
			continue
		}
		cp := path.Clean(path.Join("xl/worksheets", rel.Target))
		if !seen[cp] {
			seen[cp] = true
			refs = append(refs, chartRef{path: cp})
		}
	}

	// Drawing-mediated references (spec layout).
	for _, rel := range rels.Rels {
		if rel.TargetMode == "External" || !strings.Contains(rel.Target, "drawings/") {
			continue
		}
		drawingPath := path.Clean(path.Join("xl/worksheets", rel.Target))
		refs = append(refs, loadDrawingChartRefs(zr, drawingPath, seen)...)
	}

	// 2. Parse each chart part.
	var charts []*Chart
	for _, ref := range refs {
		raw, err := readZipFile(zr, ref.path)
		if err != nil {
			continue
		}

		var chartSpace xmlChartSpace
		if err := xml.Unmarshal(raw, &chartSpace); err != nil {
			continue
		}

		chart := &Chart{
			Width:  400, // Default width
			Height: 300, // Default height
		}
		if ref.anchor != nil {
			chart.Anchor = *ref.anchor
			if ref.width > 0 {
				chart.Width = ref.width
			}
			if ref.height > 0 {
				chart.Height = ref.height
			}
		}

		// Extract title
		chart.Title = chartTitleText(chartSpace.Chart.Title)

		// Determine chart type and extract series
		pa := &chartSpace.Chart.PlotArea
		switch {
		case pa.BarChart != nil:
			chart.Type = ChartTypeBar
			for _, ser := range pa.BarChart.Series {
				chart.Series = append(chart.Series, extractChartSeries(ser))
			}
		case pa.LineChart != nil:
			chart.Type = ChartTypeLine
			for _, ser := range pa.LineChart.Series {
				chart.Series = append(chart.Series, extractChartSeries(ser))
			}
		case pa.AreaChart != nil:
			chart.Type = ChartTypeArea
			for _, ser := range pa.AreaChart.Series {
				chart.Series = append(chart.Series, extractChartSeries(ser))
			}
		case pa.ScatterChart != nil:
			chart.Type = ChartTypeScatter
			for _, ser := range pa.ScatterChart.Series {
				chart.Series = append(chart.Series, extractScatterSeries(ser))
			}
		case pa.PieChart != nil:
			chart.Type = ChartTypePie
			if len(pa.PieChart.Series) > 0 {
				chart.Series = append(chart.Series, extractChartSeries(pa.PieChart.Series[0]))
			}
		}

		chart.XAxis, chart.YAxis = extractChartAxes(pa, chart.Type)
		if l := chartSpace.Chart.Legend; l != nil {
			legend := &ChartLegend{Position: LegendRight}
			if l.Pos != nil {
				switch LegendPosition(l.Pos.Val) {
				case LegendRight, LegendLeft, LegendTop, LegendBottom:
					legend.Position = LegendPosition(l.Pos.Val)
				}
			}
			if l.Overlay != nil {
				legend.Overlay = l.Overlay.Val
			}
			chart.Legend = legend
		}

		charts = append(charts, chart)
	}

	return charts, nil
}

// loadDrawingChartRefs collects the charts referenced by one drawing part,
// recovering each chart's placement from its graphicFrame anchor. Paths
// already in seen are skipped and new ones are added to it.
func loadDrawingChartRefs(zr *zip.Reader, drawingPath string, seen map[string]bool) []chartRef {
	drawingRelsRaw, err := readZipFile(zr, relsPathFor(drawingPath))
	if err != nil {
		return nil
	}
	var drels xmlRelationships
	if err := xml.Unmarshal(drawingRelsRaw, &drels); err != nil {
		return nil
	}
	chartPathByRID := make(map[string]string)
	for _, drel := range drels.Rels {
		if drel.TargetMode == "External" || !strings.Contains(drel.Target, "charts/") {
			continue
		}
		chartPathByRID[drel.ID] = resolvePartPath(drawingPath, drel.Target)
	}
	if len(chartPathByRID) == 0 {
		return nil
	}

	drawingRaw, err := readZipFile(zr, drawingPath)
	if err != nil {
		return nil
	}
	var drawing xmlDrawingWS
	if err := xml.Unmarshal(drawingRaw, &drawing); err != nil {
		return nil
	}

	var refs []chartRef
	for _, anchor := range append(drawing.Anchors, drawing.Anchors2...) {
		gf := anchor.GraphicFrame
		if gf == nil {
			continue
		}
		cp := chartPathByRID[gf.Graphic.GraphicData.Chart.RID]
		if cp == "" || seen[cp] {
			continue
		}
		seen[cp] = true
		refs = append(refs, chartRef{
			path: cp,
			anchor: &ChartAnchor{
				Row:    anchor.From.Row,
				Col:    anchor.From.Col,
				RowOff: anchor.From.RowOff,
				ColOff: anchor.From.ColOff,
			},
			width:  int(gf.Xfrm.Ext.width() / emuPerPixel),
			height: int(gf.Xfrm.Ext.height() / emuPerPixel),
		})
	}
	return refs
}

// extractChartSeries extracts a ChartSeries from XML.
func extractChartSeries(ser xmlChartSeries) *ChartSeries {
	series := &ChartSeries{Name: seriesNameFromTx(ser.Tx)}

	// Extract categories
	if ser.Cat.StrRef.F != "" {
		series.Categories = extractRangeFromFormula(ser.Cat.StrRef.F)
	} else if ser.Cat.NumRef.F != "" {
		series.Categories = extractRangeFromFormula(ser.Cat.NumRef.F)
	}

	// Extract values
	series.Values = refFromSeriesRef(ser.Val)

	return series
}

// extractScatterSeries maps a scatter series' xVal/yVal numeric references
// back onto Categories/Values.
func extractScatterSeries(ser xmlScatterSeries) *ChartSeries {
	return &ChartSeries{
		Name:       seriesNameFromTx(ser.Tx),
		Categories: refFromSeriesRef(ser.XVal),
		Values:     refFromSeriesRef(ser.YVal),
	}
}

// seriesNameFromTx recovers a series name: literal text from c:v, or the
// reference form older files wrote into c:strRef.
func seriesNameFromTx(tx xmlChartSeriesTx) string {
	if tx.V != "" {
		return tx.V
	}
	if tx.StrRef.F != "" {
		return extractRangeFromFormula(tx.StrRef.F)
	}
	return ""
}

// refFromSeriesRef extracts the range from a numRef-first series reference.
func refFromSeriesRef(r xmlChartSeriesRef) string {
	if r.NumRef.F != "" {
		return extractRangeFromFormula(r.NumRef.F)
	}
	if r.StrRef.F != "" {
		return extractRangeFromFormula(r.StrRef.F)
	}
	return ""
}

// extractChartAxes recovers the axis configuration for a loaded chart.
// Scatter charts encode both axes as valAx, distinguished by axPos; other
// cartesian charts use catAx for X and the first valAx for Y. Pie charts
// have no axes.
func extractChartAxes(pa *xmlPlotArea, ct ChartType) (x, y *ChartAxis) {
	if ct == ChartTypeScatter {
		for i := range pa.ValAx {
			pos := ""
			if pa.ValAx[i].AxPos != nil {
				pos = pa.ValAx[i].AxPos.Val
			}
			switch pos {
			case "b":
				if x == nil {
					x = axisFromXML(&pa.ValAx[i])
				}
			case "l":
				if y == nil {
					y = axisFromXML(&pa.ValAx[i])
				}
			default:
				if x == nil {
					x = axisFromXML(&pa.ValAx[i])
				} else if y == nil {
					y = axisFromXML(&pa.ValAx[i])
				}
			}
		}
		return x, y
	}
	if len(pa.CatAx) > 0 {
		x = axisFromXML(&pa.CatAx[0])
	}
	if len(pa.ValAx) > 0 {
		y = axisFromXML(&pa.ValAx[0])
	}
	return x, y
}

func axisFromXML(ax *xmlChartAxis) *ChartAxis {
	out := &ChartAxis{Title: chartTitleText(ax.Title)}
	if ax.Delete != nil {
		out.Hidden = ax.Delete.Val
	}
	return out
}

// chartTitleText concatenates the runs of a rich-text chart or axis title.
func chartTitleText(t *xmlChartTitle) string {
	if t == nil {
		return ""
	}
	var sb strings.Builder
	for _, r := range t.Tx.Rich.P.R {
		sb.WriteString(r.T)
	}
	return sb.String()
}

// extractRangeFromFormula extracts a cell range from a formula like "'Sheet1'!$A$1:$A$10"
func extractRangeFromFormula(formula string) string {
	// Remove sheet name and quotes
	if idx := strings.Index(formula, "!"); idx >= 0 {
		formula = formula[idx+1:]
	}

	// Remove $ signs
	formula = strings.ReplaceAll(formula, "$", "")

	return formula
}

// ======================================================================
// Internal XML types for decoding pivot table XML
// ======================================================================

type xmlPivotTableDefinition struct {
	XMLName     xml.Name            `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main pivotTableDefinition"`
	Name        string              `xml:"name,attr"`
	CacheID     int                 `xml:"cacheId,attr"`
	Location    xmlPivotLocation    `xml:"location"`
	PivotFields xmlPivotFields      `xml:"pivotFields"`
	RowFields   *xmlPivotFieldRefs  `xml:"rowFields,omitempty"`
	ColFields   *xmlPivotFieldRefs  `xml:"colFields,omitempty"`
	DataFields  *xmlPivotDataFields `xml:"dataFields,omitempty"`
}

type xmlPivotLocation struct {
	Ref string `xml:"ref,attr"`
}

type xmlPivotFields struct {
	Count  int             `xml:"count,attr"`
	Fields []xmlPivotField `xml:"pivotField"`
}

type xmlPivotField struct {
	Name string `xml:"name,attr,omitempty"`
	Axis string `xml:"axis,attr,omitempty"`
}

type xmlPivotFieldRefs struct {
	Count  int                `xml:"count,attr"`
	Fields []xmlPivotFieldRef `xml:"field"`
}

type xmlPivotFieldRef struct {
	X int `xml:"x,attr"`
}

type xmlPivotDataFields struct {
	Count  int                 `xml:"count,attr"`
	Fields []xmlPivotDataField `xml:"dataField"`
}

type xmlPivotDataField struct {
	Name     string `xml:"name,attr"`
	Field    int    `xml:"fld,attr"`
	Subtotal string `xml:"subtotal,attr,omitempty"`
}

type xmlPivotCacheDefinition struct {
	XMLName     xml.Name       `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main pivotCacheDefinition"`
	CacheSource xmlCacheSource `xml:"cacheSource"`
	CacheFields xmlCacheFields `xml:"cacheFields"`
}

type xmlCacheSource struct {
	Type            string          `xml:"type,attr"`
	WorksheetSource xmlWorksheetSrc `xml:"worksheetSource"`
}

type xmlWorksheetSrc struct {
	Ref   string `xml:"ref,attr"`
	Sheet string `xml:"sheet,attr"`
}

type xmlCacheFields struct {
	Count  int             `xml:"count,attr"`
	Fields []xmlCacheField `xml:"cacheField"`
}

type xmlCacheField struct {
	Name string `xml:"name,attr"`
}

// loadPivotTables reads the pivot table definitions for the given sheet index.
func loadPivotTables(zr *zip.Reader, sheetIndex int) ([]*PivotTable, error) {
	// 1. Find the pivot table relationships in sheet rels.
	sheetPath := fmt.Sprintf("xl/worksheets/sheet%d.xml", sheetIndex+1)
	relsPath := relsPathFor(sheetPath)
	relsRaw, err := readZipFile(zr, relsPath)
	if err != nil {
		return nil, nil // No rels means no pivot tables
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(relsRaw, &rels); err != nil {
		return nil, fmt.Errorf("cannot parse %s: %w", relsPath, err)
	}

	var pivotTables []*PivotTable
	for _, rel := range rels.Rels {
		if rel.TargetMode == "External" {
			continue
		}
		if !strings.HasSuffix(rel.Type, "/pivotTable") && !strings.Contains(rel.Target, "pivotTables/") {
			continue
		}

		// Resolve pivot table path
		ptPath := resolvePartPath(sheetPath, rel.Target)
		if ptPath == "" {
			continue
		}
		raw, err := readZipFile(zr, ptPath)
		if err != nil {
			continue
		}

		// Parse pivot table XML
		var ptDef xmlPivotTableDefinition
		if err := xml.Unmarshal(raw, &ptDef); err != nil {
			continue
		}

		pt := &PivotTable{
			Name: ptDef.Name,
			Ref:  ptDef.Location.Ref,
		}

		// Field names live in the cache definition; pivotField/dataField
		// elements reference them by index.
		var cacheNames []string
		cachePath := pivotCacheDefPath(zr, ptDef.CacheID)
		if cacheRaw, err := readZipFile(zr, cachePath); err == nil {
			var cacheDef xmlPivotCacheDefinition
			if err := xml.Unmarshal(cacheRaw, &cacheDef); err == nil {
				pt.SourceRef = cacheDef.CacheSource.WorksheetSource.Ref
				for _, cf := range cacheDef.CacheFields.Fields {
					cacheNames = append(cacheNames, cf.Name)
				}
			}
		}
		fieldName := func(x int) string {
			if x >= 0 && x < len(cacheNames) && cacheNames[x] != "" {
				return cacheNames[x]
			}
			// Files written by older versions of this library stored the
			// name on the pivotField element instead.
			if x >= 0 && x < len(ptDef.PivotFields.Fields) {
				return ptDef.PivotFields.Fields[x].Name
			}
			return ""
		}

		// Extract row fields
		if ptDef.RowFields != nil {
			for _, fieldRef := range ptDef.RowFields.Fields {
				if name := fieldName(fieldRef.X); name != "" {
					pt.RowFields = append(pt.RowFields, name)
				}
			}
		}

		// Extract col fields
		if ptDef.ColFields != nil {
			for _, fieldRef := range ptDef.ColFields.Fields {
				if name := fieldName(fieldRef.X); name != "" {
					pt.ColFields = append(pt.ColFields, name)
				}
			}
		}

		// Extract data fields. The name attribute is the display caption;
		// the aggregated source column is the fld index.
		if ptDef.DataFields != nil {
			for _, df := range ptDef.DataFields.Fields {
				name := fieldName(df.Field)
				if name == "" {
					name = df.Name
				}
				subtotal := df.Subtotal
				if subtotal == "" {
					subtotal = PivotAggregationSum // schema default
				}
				pt.DataFields = append(pt.DataFields, &PivotDataField{
					Name:        name,
					DisplayName: df.Name,
					Aggregation: subtotal,
				})
			}
		}

		pivotTables = append(pivotTables, pt)
	}

	return pivotTables, nil
}

// pivotCacheDefPath locates the pivot cache definition part for a cacheId,
// preferring the workbook-level <pivotCaches> mapping and falling back to
// the conventional part name.
func pivotCacheDefPath(zr *zip.Reader, cacheID int) string {
	if wbRaw, err := readZipFile(zr, "xl/workbook.xml"); err == nil {
		var wbx struct {
			PivotCaches struct {
				Caches []struct {
					CacheID int    `xml:"cacheId,attr"`
					RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
				} `xml:"pivotCache"`
			} `xml:"pivotCaches"`
		}
		if xml.Unmarshal(wbRaw, &wbx) == nil {
			for _, c := range wbx.PivotCaches.Caches {
				if c.CacheID != cacheID || c.RID == "" {
					continue
				}
				relsRaw, err := readZipFile(zr, "xl/_rels/workbook.xml.rels")
				if err != nil {
					break
				}
				var rels xmlRelationships
				if xml.Unmarshal(relsRaw, &rels) != nil {
					break
				}
				for _, rel := range rels.Rels {
					if rel.ID == c.RID {
						if p := resolvePartPath("xl/workbook.xml", rel.Target); p != "" {
							return p
						}
					}
				}
				break
			}
		}
	}
	return fmt.Sprintf("xl/pivotCache/pivotCacheDefinition%d.xml", cacheID)
}
