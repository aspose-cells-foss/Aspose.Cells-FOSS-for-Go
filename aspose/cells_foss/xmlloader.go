package cells_foss

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
)

// ---------------------------------------------------------------------------
// Internal XML types for decoding workbook.xml
// ---------------------------------------------------------------------------

type xmlWorkbook struct {
	XMLName xml.Name   `xml:"workbook"`
	Sheets  xmlSheets  `xml:"sheets"`
}

type xmlSheets struct {
	Sheet []xmlSheet `xml:"sheet"`
}

// xmlSheet mirrors a <sheet> element in workbook.xml.
// The RID field uses the relationships namespace; the decoder matches on the
// local name "id" within that namespace regardless of the prefix used in-file.
type xmlSheet struct {
	Name    string `xml:"name,attr"`
	SheetID string `xml:"sheetId,attr"`
	RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// ---------------------------------------------------------------------------
// Internal XML types for decoding workbook.xml.rels
// ---------------------------------------------------------------------------

type xmlRelationships struct {
	XMLName xml.Name           `xml:"Relationships"`
	Rels    []xmlRelationship  `xml:"Relationship"`
}

type xmlRelationship struct {
	ID     string `xml:"Id,attr"`
	Target string `xml:"Target,attr"`
}

// ---------------------------------------------------------------------------
// Internal XML types for decoding sheetN.xml
// ---------------------------------------------------------------------------

type xmlWorksheet struct {
	XMLName         xml.Name          `xml:"worksheet"`
	SheetData       xmlSheetData      `xml:"sheetData"`
	DataValidations *xmlDataValidations `xml:"dataValidations"`
}

type xmlDataValidations struct {
	Count int               `xml:"count,attr"`
	DVs   []xmlDataValidation `xml:"dataValidation"`
}

type xmlDataValidation struct {
	Type             string `xml:"type,attr"`
	Sqref            string `xml:"sqref,attr"`
	AllowBlank       int    `xml:"allowBlank,attr,omitempty"`
	ShowErrorMessage int    `xml:"showErrorMessage,attr,omitempty"`
	ErrorStyle       string `xml:"errorStyle,attr,omitempty"`
	ErrorTitle       string `xml:"errorTitle,attr,omitempty"`
	ErrorMessage     string `xml:"error,attr,omitempty"`
	Formula1         string `xml:"formula1"`
	Formula2         string `xml:"formula2"`
}

type xmlSheetData struct {
	Rows []xmlRow `xml:"row"`
}

type xmlRow struct {
	R     int      `xml:"r,attr"`
	Cells []xmlCell `xml:"c"`
}

type xmlCell struct {
	Ref string `xml:"r,attr"`
	T   string `xml:"t,attr,omitempty"`
	S   int    `xml:"s,attr,omitempty"`
	V   string `xml:"v"`
	F   string `xml:"f"`
}

// ---------------------------------------------------------------------------
// LoadWorkbook
// ---------------------------------------------------------------------------

// LoadWorkbook opens the .xlsx file at path, parses its XML parts, and
// returns a populated *Workbook.
//
// Shared strings are resolved automatically when xl/sharedStrings.xml is
// present; cells with t="s" receive the corresponding string value rather
// than the numeric index.
//
// When the file is encrypted (detected by the "ECRX" magic header),
// LoadWorkbook returns an error directing the caller to use LoadWithPassword.
func LoadWorkbook(path string) (*Workbook, error) {
	raw, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	// Detect encrypted files.
	if isEncryptedFile(raw) {
		return nil, fmt.Errorf("loading workbook: file is encrypted — use LoadWithPassword(path, password)")
	}

	return loadWorkbookFromBytes(raw, path)
}

// LoadWithPassword opens an encrypted .xlsx file using the given password.
// On success the returned Workbook has its password field set so that
// subsequent saves preserve encryption.
func LoadWithPassword(path string, password string) (*Workbook, error) {
	raw, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	if !isEncryptedFile(raw) {
		// Not encrypted — delegate to standard load.
		return loadWorkbookFromBytes(raw, path)
	}

	infoXML, encPkg, err := readEncryptedFile(raw)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	zipBytes, err := decryptPackage(infoXML, encPkg, password)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	wb, err := loadWorkbookFromBytes(zipBytes, path)
	if err != nil {
		return nil, err
	}
	wb.password = password
	return wb, nil
}

// loadWorkbookFromBytes is the shared loading logic used by both
// LoadWorkbook and LoadWithPassword.
func loadWorkbookFromBytes(data []byte, path string) (*Workbook, error) {
	br := bytes.NewReader(data)
	r, err := zip.NewReader(br, int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	// We need a zip.ReadCloser-like interface, but zip.NewReader returns
	// *zip.Reader (not ReadCloser).  Adapt the helpers.
	return loadWorkbookFromReader(r, path)
}

// zipReadCloserAdapter wraps a zip.Reader to satisfy the zip.ReadCloser
// interface used by the helper functions.
type zipReadCloserAdapter struct {
	*zip.Reader
	files map[string]*zip.File
}

func (a *zipReadCloserAdapter) Close() error { return nil }

func loadWorkbookFromReader(zr *zip.Reader, path string) (*Workbook, error) {
	// Build a compatible adapter.
	adapter := &zipReadCloserAdapter{
		Reader: zr,
	}
	// Populate adapter.File from zr.File (both are []*zip.File).
	adapter.File = zr.File

	// 1. Resolve sheet file paths via relationships.
	sheetPathByRID, err := parseWorkbookRelsFromReader(adapter)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	// 2. Parse workbook.xml.
	sheetDefs, err := parseWorkbookXMLFromReader(adapter)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	// 3. Load shared strings.
	sharedStrings, _ := loadSharedStringsFromReader(adapter)

	// 4. Build worksheets.
	worksheets := make([]*Worksheet, 0, len(sheetDefs))
	var firstSheetRaw []byte

	for i, def := range sheetDefs {
		relPath, ok := sheetPathByRID[def.RID]
		if !ok {
			return nil, fmt.Errorf("loading workbook: no relationship target for sheet %q (r:id=%q)", def.Name, def.RID)
		}

		fullPath := "xl/" + relPath
		ws, rawXML, err := parseWorksheetXMLFromReader(adapter, fullPath, sharedStrings)
		if err != nil {
			return nil, fmt.Errorf("loading workbook: parsing sheet %q: %w", def.Name, err)
		}
		ws.Name = def.Name
		ws.Index = i
		worksheets = append(worksheets, ws)

		if i == 0 {
			firstSheetRaw = rawXML
		}
	}

	wb := &Workbook{
		Worksheets: worksheets,
		SourceXML:  firstSheetRaw,
		FilePath:   path,
	}

	// 5. Load styles.
	stylesRaw, _ := loadStylesFromReader(adapter, wb)
	wb.StylesXML = stylesRaw

	// 6. Load tables.
	for i := range wb.Worksheets {
		tables, _ := loadTablesFromReader(adapter, i)
		wb.Worksheets[i].Tables = tables
	}

	// 7. Wire parent references.
	for _, ws := range wb.Worksheets {
		ws.cells.setParent(wb)
	}

	return wb, nil
}

// Reader-based versions of the helper functions (work with *zip.ReadCloser adapter).
func readZipFileFromReader(r *zipReadCloserAdapter, name string) ([]byte, error) {
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

func parseWorkbookRelsFromReader(r *zipReadCloserAdapter) (map[string]string, error) {
	raw, err := readZipFileFromReader(r, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, fmt.Errorf("cannot read workbook relationships: %w", err)
	}
	var rels xmlRelationships
	if err := xml.Unmarshal(raw, &rels); err != nil {
		return nil, fmt.Errorf("cannot parse workbook relationships: %w", err)
	}
	out := make(map[string]string, len(rels.Rels))
	for _, rel := range rels.Rels {
		out[rel.ID] = rel.Target
	}
	return out, nil
}

func parseWorkbookXMLFromReader(r *zipReadCloserAdapter) ([]sheetDef, error) {
	raw, err := readZipFileFromReader(r, "xl/workbook.xml")
	if err != nil {
		return nil, fmt.Errorf("cannot read workbook.xml: %w", err)
	}
	var wb xmlWorkbook
	if err := xml.Unmarshal(raw, &wb); err != nil {
		return nil, fmt.Errorf("cannot parse workbook.xml: %w", err)
	}
	defs := make([]sheetDef, 0, len(wb.Sheets.Sheet))
	for _, s := range wb.Sheets.Sheet {
		defs = append(defs, sheetDef{Name: s.Name, SheetID: s.SheetID, RID: s.RID})
	}
	return defs, nil
}

func parseWorksheetXMLFromReader(r *zipReadCloserAdapter, fullPath string, ss map[int]string) (*Worksheet, []byte, error) {
	raw, err := readZipFileFromReader(r, fullPath)
	if err != nil {
		return nil, nil, fmt.Errorf("cannot read %s: %w", fullPath, err)
	}
	var sheet xmlWorksheet
	if err := xml.Unmarshal(raw, &sheet); err != nil {
		return nil, nil, fmt.Errorf("cannot parse %s: %w", fullPath, err)
	}

	cl := &Cells{cells: make(map[string]*Cell)}
	ws := &Worksheet{cells: cl}
	for _, row := range sheet.SheetData.Rows {
		for _, rc := range row.Cells {
			cell := &Cell{
				Ref:     rc.Ref,
				StyleID: rc.S,
				Formula: rc.F,
				cells:   cl,
			}
			cell.Value = resolveCellValue(rc, ss)
			ws.cells.cells[rc.Ref] = cell
		}
	}
	if sheet.DataValidations != nil {
		for _, xdv := range sheet.DataValidations.DVs {
			dv := &DataValidation{
				Type:             xdv.Type,
				Ref:              xdv.Sqref,
				Formula1:         strings.TrimSpace(xdv.Formula1),
				Formula2:         strings.TrimSpace(xdv.Formula2),
				AllowBlank:       xdv.AllowBlank != 0,
				ShowErrorMessage: xdv.ShowErrorMessage != 0,
				ErrorStyle:       xdv.ErrorStyle,
				ErrorTitle:       xdv.ErrorTitle,
				ErrorMessage:     xdv.ErrorMessage,
			}
			ws.DataValidations = append(ws.DataValidations, dv)
		}
	}
	return ws, raw, nil
}

func loadSharedStringsFromReader(r *zipReadCloserAdapter) (map[int]string, error) {
	raw, err := readZipFileFromReader(r, "xl/sharedStrings.xml")
	if err != nil {
		return map[int]string{}, nil
	}
	var sst xmlSST
	if err := xml.Unmarshal(raw, &sst); err != nil {
		return nil, fmt.Errorf("shared strings: %w", err)
	}
	out := make(map[int]string, len(sst.Items))
	for i, si := range sst.Items {
		out[i] = resolveSIText(si)
	}
	return out, nil
}

func loadStylesFromReader(r *zipReadCloserAdapter, wb *Workbook) ([]byte, error) {
	raw, err := readZipFileFromReader(r, "xl/styles.xml")
	if err != nil {
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

func loadTablesFromReader(r *zipReadCloserAdapter, sheetIndex int) ([]*Table, error) {
	relsPath := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetIndex+1)
	relsRaw, err := readZipFileFromReader(r, relsPath)
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
		parts := strings.Split(rel.Target, "/")
		fileName := parts[len(parts)-1]
		simplePath := "xl/tables/" + fileName
		raw, err := readZipFileFromReader(r, simplePath)
		if err != nil {
			continue
		}
		var xt xmlTable
		if err := xml.Unmarshal(raw, &xt); err != nil {
			continue
		}
		t := &Table{
			Name:         xt.Name,
			Range:        xt.Ref,
			HasHeaderRow: xt.HeaderRowCount > 0,
			StyleName:    "TableStyleMedium9",
		}
		tables = append(tables, t)
	}
	return tables, nil
}

// sheetDef holds the parsed metadata for a single sheet from workbook.xml.
type sheetDef struct {
	Name    string
	SheetID string
	RID     string
}

// ---------------------------------------------------------------------------
// Helper: resolve the effective cell value (shared string or literal)
// ---------------------------------------------------------------------------

func resolveCellValue(c xmlCell, ss map[int]string) interface{} {
	if c.V == "" {
		return nil
	}

	// Shared string index.
	if c.T == "s" {
		idx, err := strconv.Atoi(strings.TrimSpace(c.V))
		if err != nil {
			return c.V // fallback: return the raw index
		}
		if s, ok := ss[idx]; ok {
			return s
		}
		return c.V
	}

	// Boolean.
	if c.T == "b" {
		if c.V == "1" {
			return true
		}
		return false
	}

	// Numeric or untyped — return as string; callers may parse with strconv.
	return c.V
}

// ---------------------------------------------------------------------------
// Helper: read a single file from the ZIP archive into a []byte
// ---------------------------------------------------------------------------

func readZipFile(r *zip.ReadCloser, name string) ([]byte, error) {
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
