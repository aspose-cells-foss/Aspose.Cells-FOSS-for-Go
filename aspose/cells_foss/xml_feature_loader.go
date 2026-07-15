package cells_foss

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// ======================================================================
// Internal XML types for decoding styles.xml
// ======================================================================

type xmlStyleSheet struct {
	XMLName      xml.Name    `xml:"styleSheet"`
	Fonts        xmlFonts    `xml:"fonts"`
	Fills        xmlFills    `xml:"fills"`
	Borders      xmlBorders  `xml:"borders"`
	CellStyleXfs xmlXfs      `xml:"cellStyleXfs"`
	CellXfs      xmlXfs      `xml:"cellXfs"`
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
	NumFmtId       int          `xml:"numFmtId,attr"`
	FontId         int          `xml:"fontId,attr"`
	FillId         int          `xml:"fillId,attr"`
	BorderId       int          `xml:"borderId,attr"`
	XfId           int          `xml:"xfId,attr"`
	ApplyFont      int          `xml:"applyFont,attr,omitempty"`
	ApplyFill      int          `xml:"applyFill,attr,omitempty"`
	ApplyBorder    int          `xml:"applyBorder,attr,omitempty"`
	ApplyAlignment int          `xml:"applyAlignment,attr,omitempty"`
	Alignment      *xmlAlign    `xml:"alignment"`
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

// ======================================================================
// loadStyles
// ======================================================================

// loadStyles reads xl/styles.xml from the archive and populates the
// Workbook's style registry.  It returns the raw bytes for caching so
// unmodified workbooks can reuse the original XML.
//
// When xl/styles.xml is absent the workbook is seeded with the single
// default style at index 0.
func loadStyles(r *zip.ReadCloser, wb *Workbook) ([]byte, error) {
	raw, err := readZipFile(r, "xl/styles.xml")
	if err != nil {
		// No styles.xml — seed the default style only.
		wb.styles = []*Style{DefaultStyle()}
		return nil, nil
	}

	var ss xmlStyleSheet
	if err := xml.Unmarshal(raw, &ss); err != nil {
		return nil, fmt.Errorf("cannot parse styles.xml: %w", err)
	}

	// Build lookups: index → parsed object.
	fonts := parseFonts(ss.Fonts.Fonts)
	fills := parseFills(ss.Fills.Fills)
	borders := parseBorders(ss.Borders.Borders)

	// Build Style for each cellXf entry.
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

// parseFonts converts the parsed <font> elements into Font objects.
func parseFonts(fonts []xmlFont) []*Font {
	out := make([]*Font, len(fonts))
	for i, f := range fonts {
		ft := &Font{
			Name:  f.Name.Val,
			Bold:  f.Bold != nil,
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

// readZipBytes reads a named entry from the archive.  (Thin wrapper over
// the same helper in xmlloader.go; duplicated to keep loading concerns
// together.)
func readZipBytes(r *zip.ReadCloser, name string) ([]byte, error) {
	for _, f := range r.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("file %q not found in archive", name)
}

// ======================================================================
// Table loading
// ======================================================================

type xmlTable struct {
	XMLName        xml.Name        `xml:"table"`
	ID             string          `xml:"id,attr"`
	Name           string          `xml:"name,attr"`
	DisplayName    string          `xml:"displayName,attr"`
	Ref            string          `xml:"ref,attr"`
	HeaderRowCount int             `xml:"headerRowCount,attr"`
}

// loadTables reads table definitions for the given sheet index.  It first
// looks for the sheet's rels file, discovers table relationships, then
// parses the corresponding table XML files.
func loadTables(r *zip.ReadCloser, sheetIndex int) ([]*Table, error) {
	relsPath := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetIndex+1)
	relsRaw, err := readZipBytes(r, relsPath)
	if err != nil {
		// No rels file — no tables.
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

		// Resolve the target path relative to xl/worksheets/.
		tablePath := "xl/worksheets/" + rel.Target
		// Normalise "../tables/table1.xml" → "xl/tables/table1.xml".
		tablePath = strings.Replace(tablePath, "/../", "/", 1)
		// Clean any remaining .. segments.
		for strings.Contains(tablePath, "..") {
			tablePath = strings.Replace(tablePath, "/../", "/", 1)
		}
		// Simpler normalisation: "../tables/" → just look for "xl/tables/".
		parts := strings.Split(rel.Target, "/")
		fileName := parts[len(parts)-1]
		simplePath := "xl/tables/" + fileName

		raw, err := readZipBytes(r, simplePath)
		if err != nil {
			raw, err = readZipBytes(r, tablePath)
			if err != nil {
				continue // skip unresolvable table
			}
		}

		var xt xmlTable
		if err := xml.Unmarshal(raw, &xt); err != nil {
			continue
		}

		t := &Table{
			Name:         xt.Name,
			Range:        xt.Ref,
			HasHeaderRow: xt.HeaderRowCount > 0,
			StyleName:    "TableStyleMedium9", // default; style parsing deferred
		}
		tables = append(tables, t)
	}
	return tables, nil
}
