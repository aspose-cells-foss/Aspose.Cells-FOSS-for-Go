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
	XMLName xml.Name  `xml:"workbook"`
	Sheets  xmlSheets `xml:"sheets"`
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
	XMLName xml.Name          `xml:"Relationships"`
	Rels    []xmlRelationship `xml:"Relationship"`
}

type xmlRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr"`
}

// ---------------------------------------------------------------------------
// Internal XML types for decoding [Content_Types].xml
// ---------------------------------------------------------------------------

type xmlContentTypes struct {
	Defaults  []xmlCTDefault  `xml:"Default"`
	Overrides []xmlCTOverride `xml:"Override"`
}

type xmlCTDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type xmlCTOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// contentTypesCache holds the content-type declarations of a loaded package
// so that parts written back verbatim keep their original declarations.
type contentTypesCache struct {
	defaults  map[string]string // lower-case extension → content type
	overrides map[string]string // part path (no leading slash) → content type
}

// loadContentTypesCache parses [Content_Types].xml. It returns nil when the
// part is missing or unreadable; callers treat a nil cache as "no original
// declarations available".
func loadContentTypesCache(zr *zip.Reader) *contentTypesCache {
	raw, err := readZipFile(zr, "[Content_Types].xml")
	if err != nil {
		return nil
	}
	var ct xmlContentTypes
	if err := xml.Unmarshal(raw, &ct); err != nil {
		return nil
	}
	out := &contentTypesCache{
		defaults:  make(map[string]string, len(ct.Defaults)),
		overrides: make(map[string]string, len(ct.Overrides)),
	}
	for _, d := range ct.Defaults {
		out.defaults[strings.ToLower(d.Extension)] = d.ContentType
	}
	for _, o := range ct.Overrides {
		out.overrides[strings.TrimPrefix(o.PartName, "/")] = o.ContentType
	}
	return out
}

// ---------------------------------------------------------------------------
// Internal XML types for decoding sheetN.xml
// ---------------------------------------------------------------------------

type xmlWorksheet struct {
	XMLName         xml.Name            `xml:"worksheet"`
	SheetViews      *xmlSheetViews      `xml:"sheetViews"`
	Cols            *xmlCols            `xml:"cols"`
	SheetData       xmlSheetData        `xml:"sheetData"`
	AutoFilter      *xmlAutoFilter      `xml:"autoFilter"`
	MergeCells      *xmlMergeCells      `xml:"mergeCells"`
	DataValidations *xmlDataValidations `xml:"dataValidations"`
	Hyperlinks      *xmlHyperlinks      `xml:"hyperlinks"`
}

type xmlSheetViews struct {
	Views []xmlSheetView `xml:"sheetView"`
}

type xmlSheetView struct {
	Pane *xmlPane `xml:"pane"`
}

type xmlPane struct {
	XSplit int    `xml:"xSplit,attr"`
	YSplit int    `xml:"ySplit,attr"`
	State  string `xml:"state,attr"`
}

type xmlCols struct {
	Cols []xmlCol `xml:"col"`
}

type xmlCol struct {
	Min         int     `xml:"min,attr"`
	Max         int     `xml:"max,attr"`
	Width       float64 `xml:"width,attr"`
	CustomWidth int     `xml:"customWidth,attr"`
}

type xmlAutoFilter struct {
	Ref string `xml:"ref,attr"`
}

type xmlMergeCells struct {
	Merges []xmlMergeCell `xml:"mergeCell"`
}

type xmlMergeCell struct {
	Ref string `xml:"ref,attr"`
}

type xmlHyperlinks struct {
	Links []xmlHyperlink `xml:"hyperlink"`
}

type xmlHyperlink struct {
	Ref      string `xml:"ref,attr"`
	RID      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	Location string `xml:"location,attr"`
	Tooltip  string `xml:"tooltip,attr"`
	Display  string `xml:"display,attr"`
}

type xmlDataValidations struct {
	Count int                 `xml:"count,attr"`
	DVs   []xmlDataValidation `xml:"dataValidation"`
}

type xmlDataValidation struct {
	Type             string `xml:"type,attr"`
	Operator         string `xml:"operator,attr"`
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
	R            int       `xml:"r,attr"`
	Ht           float64   `xml:"ht,attr"`
	CustomHeight int       `xml:"customHeight,attr"`
	Cells        []xmlCell `xml:"c"`
}

type xmlCell struct {
	Ref string `xml:"r,attr"`
	T   string `xml:"t,attr,omitempty"`
	S   int    `xml:"s,attr,omitempty"`
	V   string `xml:"v"`
	F   string `xml:"f"`
	// Is carries the inline string payload for cells with t="inlineStr".
	// The <is> element has the same CT_Rst structure as <si>.
	Is *xmlSI `xml:"is"`
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
// When the file is encrypted (detected by the OLE/CFB magic header or the
// legacy "ECRX" magic header), LoadWorkbook returns an error directing the
// caller to use LoadWithPassword.
func LoadWorkbook(path string) (*Workbook, error) {
	raw, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	// Detect encrypted files.
	if isEncryptedFile(raw) || isCFBFile(raw) {
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
	return loadWorkbookMaybeEncrypted(raw, password, path)
}

// LoadFrom reads an .xlsx package from r and returns the populated Workbook.
// It is the stream counterpart of LoadWorkbook: encrypted packages are
// rejected — use LoadFromWithPassword for those. The returned Workbook has
// an empty FilePath since a stream has no file name.
func LoadFrom(r io.Reader) (*Workbook, error) {
	data, err := io.ReadAll(r)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}
	if isEncryptedFile(data) || isCFBFile(data) {
		return nil, fmt.Errorf("loading workbook: stream is encrypted — use LoadFromWithPassword(r, password)")
	}
	return loadWorkbookFromBytes(data, "")
}

// LoadFromWithPassword reads a possibly encrypted .xlsx package from r using
// the given password. It is the stream counterpart of LoadWithPassword:
// unencrypted packages load directly, encrypted ones are decrypted first.
// On success the returned Workbook has its password field set so that
// subsequent saves preserve encryption.
func LoadFromWithPassword(r io.Reader, password string) (*Workbook, error) {
	data, err := io.ReadAll(r)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}
	return loadWorkbookMaybeEncrypted(data, password, "")
}

// loadWorkbookMaybeEncrypted loads package bytes that may or may not be
// encrypted. Unencrypted bytes load directly; encrypted bytes are decrypted
// with password first and the password is kept on the Workbook so later
// saves re-encrypt. path is recorded in Workbook.FilePath (empty for
// streams).
func loadWorkbookMaybeEncrypted(raw []byte, password, path string) (*Workbook, error) {
	if !isEncryptedFile(raw) && !isCFBFile(raw) {
		// Not encrypted — standard load.
		return loadWorkbookFromBytes(raw, path)
	}

	var infoXML, encPkg []byte
	var origSize uint64
	var err error
	if isCFBFile(raw) {
		infoXML, encPkg, origSize, err = readCFBEncrypted(raw)
	} else {
		infoXML, encPkg, err = readEncryptedFile(raw)
	}
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	zipBytes, err := decryptPackage(infoXML, encPkg, password, origSize)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	// loadWorkbookFromBytes caches zipBytes as the verbatim round-trip
	// source; re-saving with the password re-encrypts those bytes.
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
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}
	wb, err := loadWorkbookFromReader(zr, path)
	if err != nil {
		return nil, err
	}
	wb.sourceZip = data
	return wb, nil
}

func loadWorkbookFromReader(zr *zip.Reader, path string) (*Workbook, error) {
	// 1. Resolve sheet file paths via relationships.
	sheetPathByRID, err := parseWorkbookRels(zr)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	// Content-type declarations of the original package, reused when parts
	// are written back verbatim.
	origCT := loadContentTypesCache(zr)

	// 2. Parse workbook.xml.
	sheetDefs, err := parseWorkbookXML(zr)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	// 3. Load shared strings. A corrupt sharedStrings.xml makes every string
	// cell unreadable, so the load fails instead of silently degrading.
	sstRaw, sharedStrings, err := loadSharedStringsRaw(zr)
	if err != nil {
		return nil, fmt.Errorf("loading workbook: %w", err)
	}

	// 4. Build worksheets.
	worksheets := make([]*Worksheet, 0, len(sheetDefs))
	var firstSheetRaw []byte

	for i, def := range sheetDefs {
		relPath, ok := sheetPathByRID[def.RID]
		if !ok {
			return nil, fmt.Errorf("loading workbook: no relationship target for sheet %q (r:id=%q)", def.Name, def.RID)
		}

		// Relationship targets may be relative ("worksheets/sheet1.xml") or
		// absolute ("/xl/worksheets/sheet1.xml"); both resolve against the
		// workbook part.
		fullPath := resolvePartPath("xl/workbook.xml", relPath)
		ws, rawXML, err := parseWorksheetXML(zr, fullPath, sharedStrings)
		if err != nil {
			return nil, fmt.Errorf("loading workbook: parsing sheet %q: %w", def.Name, err)
		}
		ws.Name = def.Name
		ws.Index = i
		ws.sourceXML = rawXML
		ws.cachedParts = cacheSheetPartClosure(zr, fullPath)
		worksheets = append(worksheets, ws)

		if i == 0 {
			firstSheetRaw = rawXML
		}
	}

	wb := &Workbook{
		Worksheets: worksheets,
		SourceXML:  firstSheetRaw,
		FilePath:   path,
		sstRaw:     sstRaw,
		sstBase:    sharedStrings,
		origCT:     origCT,
	}

	// 5. Load the VBA project part (macro-enabled workbooks). The bytes are
	// stored as-is; this library cannot parse or edit VBA code.
	if vbaRaw, err := readZipFile(zr, "xl/vbaProject.bin"); err == nil && len(vbaRaw) > 0 {
		wb.vbaProject = &VBAProject{Data: vbaRaw}
	}

	// 6. Load styles.
	stylesRaw, err := loadStyles(zr, wb)
	if err != nil {
		return nil, err
	}
	wb.StylesXML = stylesRaw

	// 7-10. Load per-sheet feature parts. A missing part is not an error
	// (the loaders return nil, nil), but a corrupt part fails the load
	// instead of silently dropping content.
	for i := range wb.Worksheets {
		tables, err := loadTables(zr, i)
		if err != nil {
			return nil, fmt.Errorf("sheet %d tables: %w", i+1, err)
		}
		wb.Worksheets[i].Tables = tables

		pictures, err := loadDrawings(zr, i)
		if err != nil {
			return nil, fmt.Errorf("sheet %d drawings: %w", i+1, err)
		}
		wb.Worksheets[i].Pictures = pictures

		charts, err := loadCharts(zr, i)
		if err != nil {
			return nil, fmt.Errorf("sheet %d charts: %w", i+1, err)
		}
		wb.Worksheets[i].Charts = charts

		pivotTables, err := loadPivotTables(zr, i)
		if err != nil {
			return nil, fmt.Errorf("sheet %d pivot tables: %w", i+1, err)
		}
		wb.Worksheets[i].PivotTables = pivotTables
	}

	// 11. Wire parent references.
	for _, ws := range wb.Worksheets {
		ws.cells.setParent(wb)
		ws.cells.setWorksheet(ws)
	}

	return wb, nil
}

func parseWorkbookRels(zr *zip.Reader) (map[string]string, error) {
	raw, err := readZipFile(zr, "xl/_rels/workbook.xml.rels")
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

func parseWorkbookXML(zr *zip.Reader) ([]sheetDef, error) {
	raw, err := readZipFile(zr, "xl/workbook.xml")
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

func parseWorksheetXML(zr *zip.Reader, fullPath string, ss []string) (*Worksheet, []byte, error) {
	raw, err := readZipFile(zr, fullPath)
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
		if row.CustomHeight == 1 && row.Ht > 0 {
			if ws.RowHeights == nil {
				ws.RowHeights = make(map[int]float64)
			}
			ws.RowHeights[row.R] = row.Ht
		}
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
	loadSheetLayout(ws, &sheet, zr, fullPath)
	if sheet.DataValidations != nil {
		for _, xdv := range sheet.DataValidations.DVs {
			dv := &DataValidation{
				Type:             xdv.Type,
				Operator:         xdv.Operator,
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
	// Load conditional formatting.
	loadConditionalFormattings(ws, raw)
	return ws, raw, nil
}

// loadSheetLayout parses the layout elements of a worksheet part — frozen
// panes, column widths, autoFilter, merged cells, and hyperlinks — into the
// model so regenerated sheet XML keeps them. External hyperlink targets are
// resolved through the sheet's .rels part.
func loadSheetLayout(ws *Worksheet, sheet *xmlWorksheet, zr *zip.Reader, sheetPath string) {
	if sheet.SheetViews != nil && len(sheet.SheetViews.Views) > 0 {
		if p := sheet.SheetViews.Views[0].Pane; p != nil && p.State == "frozen" {
			ws.FrozenRow = p.YSplit
			ws.FrozenCol = p.XSplit
		}
	}

	if sheet.Cols != nil {
		for _, c := range sheet.Cols.Cols {
			if c.Width <= 0 || c.Min < 1 || c.Max < c.Min {
				continue
			}
			last := c.Max
			if last > maxColCount {
				last = maxColCount
			}
			if ws.ColumnWidths == nil {
				ws.ColumnWidths = make(map[string]float64)
			}
			for n := c.Min; n <= last; n++ {
				ws.ColumnWidths[numToCol(n-1)] = c.Width // <col min/max> is 1-based; numToCol 0-based
			}
		}
	}

	if sheet.AutoFilter != nil {
		ws.AutoFilterRange = sheet.AutoFilter.Ref
	}

	if sheet.MergeCells != nil {
		for _, m := range sheet.MergeCells.Merges {
			if m.Ref != "" {
				ws.MergedCells = append(ws.MergedCells, m.Ref)
			}
		}
	}

	if sheet.Hyperlinks != nil && len(sheet.Hyperlinks.Links) > 0 {
		extTargets := loadExternalRelTargets(zr, relsPathFor(sheetPath))
		for _, xh := range sheet.Hyperlinks.Links {
			h := &Hyperlink{
				Ref:      xh.Ref,
				Location: xh.Location,
				Tooltip:  xh.Tooltip,
				Display:  xh.Display,
			}
			if xh.RID != "" {
				h.Target = extTargets[xh.RID]
			}
			if h.Target == "" && h.Location == "" {
				continue // unresolvable link — dropping beats dangling r:id
			}
			ws.Hyperlinks = append(ws.Hyperlinks, h)
		}
	}
}

// loadExternalRelTargets reads a .rels part and returns the targets of its
// external relationships keyed by id. A missing or corrupt part yields an
// empty map.
func loadExternalRelTargets(zr *zip.Reader, relsPath string) map[string]string {
	out := make(map[string]string)
	raw, err := readZipFile(zr, relsPath)
	if err != nil {
		return out
	}
	var rels xmlRelationships
	if err := xml.Unmarshal(raw, &rels); err != nil {
		return out
	}
	for _, rel := range rels.Rels {
		if rel.TargetMode == "External" {
			out[rel.ID] = rel.Target
		}
	}
	return out
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

func resolveCellValue(c xmlCell, ss []string) interface{} {
	// Inline strings carry no <v>; their text lives in <is>.
	if c.T == "inlineStr" {
		if c.Is != nil {
			return resolveSIText(*c.Is)
		}
		return nil
	}

	if c.V == "" {
		return nil
	}

	// Shared string index.
	if c.T == "s" {
		idx, err := strconv.Atoi(strings.TrimSpace(c.V))
		if err != nil {
			return c.V // fallback: return the raw index
		}
		if idx >= 0 && idx < len(ss) {
			return ss[idx]
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

	// String formula result — keep as string.
	if c.T == "str" {
		return c.V
	}

	// Error value (#DIV/0!, #N/A, ...).
	if c.T == "e" {
		return ErrorValue(c.V)
	}

	// Numeric (no t attribute, or t="n") — parse as float64 to preserve
	// round-trip fidelity: saving a float64 writes <v> without t="s",
	// whereas a string would be written as a shared string.
	if f, err := strconv.ParseFloat(c.V, 64); err == nil {
		return f
	}

	// Fallback: non-numeric untyped value — return as string.
	return c.V
}

// cacheSheetPartClosure caches a sheet's relationship closure: its .rels
// file plus every part reachable through relationship files (recursively),
// e.g. drawings and their media, charts, tables, pivot tables and caches,
// comments, and printer settings. Unmodified sheets re-emit these parts
// verbatim on save so rIds in the cached sheet XML stay valid.
//
// External relationships are skipped, and so are relationship targets that
// do not exist in the package (a malformed input file must not fail the
// load — the missing part simply cannot be preserved).
func cacheSheetPartClosure(zr *zip.Reader, sheetPath string) map[string][]byte {
	parts := make(map[string][]byte)
	visited := make(map[string]bool)

	var walk func(partPath string)
	walk = func(partPath string) {
		if visited[partPath] {
			return
		}
		visited[partPath] = true

		relsPath := relsPathFor(partPath)
		raw, err := readZipFile(zr, relsPath)
		if err != nil {
			return // no relationships for this part
		}
		parts[relsPath] = raw

		var rels xmlRelationships
		if err := xml.Unmarshal(raw, &rels); err != nil {
			return
		}
		for _, rel := range rels.Rels {
			if rel.TargetMode == "External" {
				continue
			}
			tp := resolvePartPath(partPath, rel.Target)
			if tp == "" {
				continue
			}
			if _, ok := parts[tp]; !ok {
				data, err := readZipFile(zr, tp)
				if err != nil {
					continue
				}
				parts[tp] = data
			}
			walk(tp)
		}
	}

	walk(sheetPath)
	return parts
}

// relsPathFor returns the relationship part path for a package part, e.g.
// "xl/worksheets/sheet1.xml" → "xl/worksheets/_rels/sheet1.xml.rels".
func relsPathFor(partPath string) string {
	if i := strings.LastIndex(partPath, "/"); i >= 0 {
		return partPath[:i] + "/_rels/" + partPath[i+1:] + ".rels"
	}
	return "_rels/" + partPath + ".rels"
}

// resolvePartPath resolves a relationship target against the part that
// declares it. Absolute targets ("/xl/…") are package-rooted; relative
// targets ("../tables/table1.xml") are resolved against the declaring
// part's folder. External or otherwise unusable targets resolve to "".
func resolvePartPath(basePart, target string) string {
	target = strings.TrimSpace(target)
	if target == "" || strings.ContainsRune(target, '\\') {
		return ""
	}
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}
	if i := strings.IndexAny(target, "?#"); i >= 0 {
		target = target[:i]
	}
	if strings.Contains(target, "://") {
		return "" // external URL without TargetMode="External"
	}

	var segs []string
	if i := strings.LastIndex(basePart, "/"); i >= 0 {
		segs = append(segs, strings.Split(basePart[:i], "/")...)
	}
	segs = append(segs, strings.Split(target, "/")...)

	out := make([]string, 0, len(segs))
	for _, s := range segs {
		switch s {
		case "", ".":
		case "..":
			if len(out) > 0 {
				out = out[:len(out)-1]
			}
		default:
			out = append(out, s)
		}
	}
	return strings.Join(out, "/")
}

// readZipFile reads a named entry from the zip archive into memory.
func readZipFile(zr *zip.Reader, name string) ([]byte, error) {
	for _, f := range zr.File {
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
