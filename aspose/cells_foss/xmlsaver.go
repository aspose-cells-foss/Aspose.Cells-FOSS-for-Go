package cells_foss

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"sort"
	"strconv"
	"strings"
)

// ---------------------------------------------------------------------------
// Output XML types (used exclusively for marshaling, never for parsing)
// ---------------------------------------------------------------------------

type outWorkbookXML struct {
	XMLName     xml.Name        `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	R           string          `xml:"xmlns:r,attr"`
	Sheets      outSheets       `xml:"sheets"`
	PivotCaches *outPivotCaches `xml:"pivotCaches,omitempty"`
}

type outSheets struct {
	Sheet []outSheet `xml:"sheet"`
}

type outSheet struct {
	Name    string `xml:"name,attr"`
	SheetID string `xml:"sheetId,attr"`
	RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// outPivotCaches lists one <pivotCache> per pivot cache definition in the
// package, both freshly generated and preserved from the loaded file.
type outPivotCaches struct {
	Caches []outPivotCacheRef `xml:"pivotCache"`
}

type outPivotCacheRef struct {
	CacheID int    `xml:"cacheId,attr"`
	RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// outWorksheet field order must follow the ECMA-376 CT_Worksheet element
// sequence: sheetViews → cols → sheetData → autoFilter → mergeCells →
// conditionalFormatting → dataValidations → hyperlinks → drawing →
// tableParts. encoding/xml marshals fields in declaration order.
type outWorksheet struct {
	XMLName               xml.Name                   `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	SheetViews            *outSheetViews             `xml:"sheetViews,omitempty"`
	Cols                  *outCols                   `xml:"cols,omitempty"`
	SheetData             outSheetData               `xml:"sheetData"`
	AutoFilter            *outAutoFilter             `xml:"autoFilter,omitempty"`
	MergeCells            *outMergeCells             `xml:"mergeCells,omitempty"`
	ConditionalFormatting []outConditionalFormatting `xml:"conditionalFormatting,omitempty"`
	DataValidations       *outDataValidations        `xml:"dataValidations,omitempty"`
	Hyperlinks            *outHyperlinks             `xml:"hyperlinks,omitempty"`
	Drawing               *outDrawingRef             `xml:"drawing,omitempty"`
	TableParts            *outTableParts             `xml:"tableParts,omitempty"`
}

type outDrawingRef struct {
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

type outSheetViews struct {
	Views []outSheetView `xml:"sheetView"`
}

type outSheetView struct {
	WorkbookViewID int      `xml:"workbookViewId,attr"`
	Pane           *outPane `xml:"pane,omitempty"`
}

type outPane struct {
	XSplit      int    `xml:"xSplit,attr,omitempty"`
	YSplit      int    `xml:"ySplit,attr,omitempty"`
	TopLeftCell string `xml:"topLeftCell,attr,omitempty"`
	ActivePane  string `xml:"activePane,attr,omitempty"`
	State       string `xml:"state,attr"`
}

type outCols struct {
	Cols []outCol `xml:"col"`
}

type outCol struct {
	Min         int     `xml:"min,attr"`
	Max         int     `xml:"max,attr"`
	Width       float64 `xml:"width,attr"`
	CustomWidth int     `xml:"customWidth,attr,omitempty"`
}

type outMergeCells struct {
	Count  int            `xml:"count,attr"`
	Merges []outMergeCell `xml:"mergeCell"`
}

type outMergeCell struct {
	Ref string `xml:"ref,attr"`
}

type outHyperlinks struct {
	Links []outHyperlink `xml:"hyperlink"`
}

type outHyperlink struct {
	Ref      string `xml:"ref,attr"`
	RID      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	Location string `xml:"location,attr,omitempty"`
	Tooltip  string `xml:"tooltip,attr,omitempty"`
	Display  string `xml:"display,attr,omitempty"`
}

type outSheetData struct {
	Rows []outRow `xml:"row"`
}

type outRow struct {
	R            int       `xml:"r,attr"`
	Ht           float64   `xml:"ht,attr,omitempty"`
	CustomHeight int       `xml:"customHeight,attr,omitempty"`
	Cells        []outCell `xml:"c"`
}

// outCell field order must follow the ECMA-376 CT_Cell sequence: the <f>
// element precedes <v>. encoding/xml marshals fields in declaration order.
type outCell struct {
	Ref string `xml:"r,attr"`
	T   string `xml:"t,attr,omitempty"`
	S   int    `xml:"s,attr,omitempty"`
	F   string `xml:"f,omitempty"`
	V   string `xml:"v,omitempty"`
}

type outSST struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`
	Items       []outSI  `xml:"si"`
}

type outSI struct {
	T string `xml:"t"`
}

// ---------------------------------------------------------------------------
// Content types
// ---------------------------------------------------------------------------

// OPC content types of the parts this saver generates.
const (
	ctWorkbookMain      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	ctWorksheet         = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	ctStyles            = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
	ctSharedStrings     = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
	ctTable             = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
	ctDrawing           = "application/vnd.openxmlformats-officedocument.drawing+xml"
	ctChart             = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
	ctPivotTable        = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
	ctPivotCacheDefn    = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
	ctPivotCacheRecords = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
	ctWorkbookMacro     = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
	ctVBAProject        = "application/vnd.ms-office.vbaProject"
)

// relTypeVBAProject relates workbook.xml to xl/vbaProject.bin.
const relTypeVBAProject = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject"

// builtinImageCT covers common image extensions for parts whose original
// package declaration is not available.
var builtinImageCT = map[string]string{
	"png":  "image/png",
	"jpg":  "image/jpeg",
	"jpeg": "image/jpeg",
	"gif":  "image/gif",
	"bmp":  "image/bmp",
	"tif":  "image/tiff",
	"tiff": "image/tiff",
	"emf":  "image/x-emf",
	"wmf":  "image/x-wmf",
}

// ---------------------------------------------------------------------------
// Write plan
// ---------------------------------------------------------------------------

// zipPart is one package part queued for writing.
type zipPart struct {
	name string
	data []byte
}

// ctOverride records a part that needs an <Override> entry in
// [Content_Types].xml.
type ctOverride struct {
	name        string
	contentType string
}

// packagePlan accumulates every part of the output package together with the
// content-type declarations required for them. [Content_Types].xml is
// derived from the plan when it is written, so the package can never declare
// a part that does not exist, nor omit a declaration for one that does —
// either mistake makes Excel report the file as corrupt.
//
// Part names are deduplicated: the first writer of a name wins, which lets
// parts shared between sheets (e.g. one pivot cache behind two pivot tables)
// be written exactly once.
type packagePlan struct {
	wb        *Workbook
	parts     []zipPart
	index     map[string]int
	overrides []ctOverride
	defaults  map[string]string // lower-case extension → content type

	// rootExtras and wbRelExtras are relationships toward parts preserved
	// verbatim from the loaded package (docProps, theme, VBA). They are
	// appended to the regenerated _rels/.rels and xl/_rels/workbook.xml.rels
	// with fresh sequential ids.
	rootExtras  []preservedRel
	wbRelExtras []preservedRel
}

// preservedRel is a harvested relationship (type + target) pointing at a
// part that is re-emitted verbatim.
type preservedRel struct {
	relType string
	target  string
}

func newPackagePlan(wb *Workbook) *packagePlan {
	return &packagePlan{
		wb:    wb,
		index: make(map[string]int),
		defaults: map[string]string{
			"rels": "application/vnd.openxmlformats-package.relationships+xml",
			"xml":  "application/xml",
		},
	}
}

// has reports whether a part name is already claimed in the plan.
func (p *packagePlan) has(name string) bool {
	_, ok := p.index[name]
	return ok
}

// unique returns candidate with a numeric suffix appended when the name is
// already claimed, e.g. "xl/tables/table1_2.xml".
func (p *packagePlan) unique(candidate string) string {
	if !p.has(candidate) {
		return candidate
	}
	base, ext := candidate, ""
	if i := strings.LastIndex(candidate, "."); i > strings.LastIndex(candidate, "/") {
		base, ext = candidate[:i], candidate[i:]
	}
	for n := 2; ; n++ {
		c := fmt.Sprintf("%s_%d%s", base, n, ext)
		if !p.has(c) {
			return c
		}
	}
}

// add queues a part. contentType, when non-empty, is declared with an
// <Override>; otherwise the part is covered by a <Default> extension mapping
// taken from the original package declarations or the built-in image table.
// Duplicate names are ignored.
func (p *packagePlan) add(name string, data []byte, contentType string) {
	if p.has(name) {
		return
	}
	p.index[name] = len(p.parts)
	p.parts = append(p.parts, zipPart{name: name, data: data})
	if contentType != "" {
		p.overrides = append(p.overrides, ctOverride{name: name, contentType: contentType})
		return
	}
	ext := strings.ToLower(partExtension(name))
	if _, ok := p.defaults[ext]; ok {
		return
	}
	if p.wb != nil && p.wb.origCT != nil {
		if ct, ok := p.wb.origCT.defaults[ext]; ok {
			p.defaults[ext] = ct
			return
		}
	}
	if ct, ok := builtinImageCT[ext]; ok {
		p.defaults[ext] = ct
	}
}

// addCached queues a part that was loaded verbatim from the original
// package, reusing its original <Override> declaration when it had one.
func (p *packagePlan) addCached(name string, data []byte) {
	ct := ""
	if p.wb != nil && p.wb.origCT != nil {
		ct = p.wb.origCT.overrides[name]
	}
	p.add(name, data, ct)
}

// contentTypes snapshots the declarations accumulated in the plan so they
// can serve as the "original" declarations for a subsequent save.
func (p *packagePlan) contentTypes() *contentTypesCache {
	ct := &contentTypesCache{
		defaults:  make(map[string]string, len(p.defaults)),
		overrides: make(map[string]string, len(p.overrides)),
	}
	for ext, c := range p.defaults {
		ct.defaults[ext] = c
	}
	for _, o := range p.overrides {
		ct.overrides[o.name] = o.contentType
	}
	return ct
}

// writeTo emits the whole package: [Content_Types].xml derived from the
// plan, the fixed package relationships, then every part in queue order.
func (p *packagePlan) writeTo(zw *zip.Writer) error {
	var ct strings.Builder
	ct.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	ct.WriteString(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` + "\n")
	exts := make([]string, 0, len(p.defaults))
	for ext := range p.defaults {
		exts = append(exts, ext)
	}
	sort.Strings(exts)
	for _, ext := range exts {
		fmt.Fprintf(&ct, `  <Default Extension="%s" ContentType="%s"/>`+"\n", ext, p.defaults[ext])
	}
	for _, o := range p.overrides {
		fmt.Fprintf(&ct, `  <Override PartName="/%s" ContentType="%s"/>`+"\n", o.name, o.contentType)
	}
	ct.WriteString(`</Types>` + "\n")
	if err := writeZipString(zw, "[Content_Types].xml", ct.String()); err != nil {
		return err
	}

	var rootRels strings.Builder
	rootRels.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
`)
	fmt.Fprintf(&rootRels, `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`+"\n")
	for i, rel := range p.rootExtras {
		fmt.Fprintf(&rootRels, `  <Relationship Id="rId%d" Type="%s" Target="%s"/>`+"\n", i+2, rel.relType, rel.target)
	}
	rootRels.WriteString(`</Relationships>`)
	if err := writeZipString(zw, "_rels/.rels", rootRels.String()); err != nil {
		return err
	}

	for _, part := range p.parts {
		fw, err := zw.Create(part.name)
		if err != nil {
			return err
		}
		if _, err := fw.Write(part.data); err != nil {
			return err
		}
	}
	return nil
}

// partExtension returns the lower-case-caller-normalized extension of a
// part path, or "" when it has none.
func partExtension(name string) string {
	if i := strings.LastIndex(name, "."); i > strings.LastIndex(name, "/") {
		return name[i+1:]
	}
	return ""
}

// sortedPartNames returns the map keys in sorted order for deterministic
// package output.
func sortedPartNames(parts map[string][]byte) []string {
	names := make([]string, 0, len(parts))
	for n := range parts {
		names = append(names, n)
	}
	sort.Strings(names)
	return names
}

// seedCounter returns the largest numeric suffix used by cached parts under
// the given path prefix (e.g. "xl/charts/chart"), so regenerated parts can
// continue the numbering without colliding with preserved ones.
func seedCounter(parts map[string][]byte, prefix string, cur int) int {
	for name := range parts {
		if !strings.HasPrefix(name, prefix) {
			continue
		}
		rest := name[len(prefix):]
		n, digits := 0, 0
		for _, c := range rest {
			if c < '0' || c > '9' {
				break
			}
			n = n*10 + int(c-'0')
			digits++
		}
		if digits > 0 && n > cur {
			cur = n
		}
	}
	return cur
}

// ---------------------------------------------------------------------------
// Workbook-level part preservation
// ---------------------------------------------------------------------------

// preserveWorkbookParts re-emits workbook-scoped parts that this library does
// not model but that cell edits never invalidate: docProps (core/app/custom/
// thumbnail), xl/theme, and customXml. They are copied verbatim from the
// loaded package, and the relationships that pointed at them in the original
// _rels/.rels and xl/_rels/workbook.xml.rels are harvested so the regenerated
// relationship parts still reach them.
//
// The VBA project is handled separately (see SaveWorkbook): its bytes live
// on wb.vbaProject so SetVBAProject/ClearVBAProject are authoritative.
func preserveWorkbookParts(wb *Workbook, plan *packagePlan) {
	if len(wb.sourceZip) == 0 {
		return
	}
	zr, err := zip.NewReader(bytes.NewReader(wb.sourceZip), int64(len(wb.sourceZip)))
	if err != nil {
		return
	}

	preserved := make(map[string]bool)
	for _, f := range zr.File {
		if !isPreservedWorkbookPart(f.Name) {
			continue
		}
		data, err := readZipFile(zr, f.Name)
		if err != nil {
			continue
		}
		plan.addCached(f.Name, data)
		preserved[f.Name] = true
	}
	if len(preserved) == 0 {
		return
	}

	if raw, err := readZipFile(zr, "_rels/.rels"); err == nil {
		var rels xmlRelationships
		if xml.Unmarshal(raw, &rels) == nil {
			for _, rel := range rels.Rels {
				if rel.TargetMode == "External" {
					continue
				}
				if preserved[resolvePartPath("", rel.Target)] {
					plan.rootExtras = append(plan.rootExtras, preservedRel{relType: rel.Type, target: rel.Target})
				}
			}
		}
	}

	if raw, err := readZipFile(zr, "xl/_rels/workbook.xml.rels"); err == nil {
		var rels xmlRelationships
		if xml.Unmarshal(raw, &rels) == nil {
			for _, rel := range rels.Rels {
				if rel.TargetMode == "External" {
					continue
				}
				tp := resolvePartPath("xl/workbook.xml", rel.Target)
				if strings.HasPrefix(tp, "xl/theme/") && preserved[tp] {
					plan.wbRelExtras = append(plan.wbRelExtras, preservedRel{relType: rel.Type, target: rel.Target})
				}
			}
		}
	}
}

// isPreservedWorkbookPart reports whether a package part is a workbook-scoped
// part preserved verbatim across modified saves.
func isPreservedWorkbookPart(name string) bool {
	return strings.HasPrefix(name, "docProps/") ||
		strings.HasPrefix(name, "xl/theme/") ||
		strings.HasPrefix(name, "customXml/")
}

// ---------------------------------------------------------------------------
// SaveWorkbook
// ---------------------------------------------------------------------------

// SaveWorkbook writes the workbook to an .xlsx file at path.
//
// When nothing has been modified since the workbook was loaded (or last
// saved), the original package bytes are written back verbatim, preserving
// every part byte-for-byte — including parts this library does not model.
//
// Otherwise the package is rebuilt: unmodified sheets re-emit their cached
// XML plus the original companion parts (rels, drawings, media, charts,
// tables, pivot tables and caches, comments, …), while modified sheets are
// regenerated from the in-memory model with globally unique part names.
// [Content_Types].xml and the workbook relationships are derived from the
// parts actually written.
func SaveWorkbook(wb *Workbook, path string) error {
	return saveWorkbookTo(wb, path, nil)
}

// saveWorkbookTo builds the complete package for wb and delivers it: to w
// when non-nil (a stream save — wb.FilePath is left untouched), otherwise to
// the file at path. Package construction, encryption and cache refresh are
// identical for both destinations.
func saveWorkbookTo(wb *Workbook, path string, w io.Writer) error {
	if wb == nil {
		return fmt.Errorf("saving workbook: workbook is nil")
	}

	// ---- Fast path: byte-identical round-trip of the loaded package ----
	dirty := wb.Modified
	if !dirty {
		for _, ws := range wb.Worksheets {
			if ws.Modified {
				dirty = true
				break
			}
		}
	}
	if !dirty && wb.sourceZip != nil {
		outData, err := packageBytes(wb, wb.sourceZip)
		if err != nil {
			return fmt.Errorf("saving workbook: %w", err)
		}
		return emitWorkbook(wb, path, w, outData)
	}

	// Decide up front which optional parts will be written so that the
	// workbook relationships only declare parts that actually exist.
	// Conditional-formatting rules with a Style also require the part:
	// generateStylesXML builds their <dxf> entries and assigns the ids the
	// sheet XML references.
	hasStyles := len(wb.styles) > 1 || wb.StylesXML != nil || hasDxfStyles(wb.Worksheets)

	// Shared strings: when the workbook is unmodified and the original
	// sharedStrings.xml is cached, its bytes are reused verbatim. Otherwise
	// the table is rebuilt with the original entries kept at their original
	// indices, so cached worksheet XML (which references strings by index)
	// remains valid even for sheets that are not regenerated.
	ssTable, ssIndex := buildSharedStrings(wb)
	reuseSST := !wb.Modified && wb.sstRaw != nil && len(ssTable) == len(wb.sstBase)
	hasSST := reuseSST || len(ssTable) > 0

	plan := newPackagePlan(wb)

	// ---- Workbook-scoped parts kept verbatim (docProps, theme, customXml) ----
	preserveWorkbookParts(wb, plan)

	// VBA project: bytes are authoritative on wb.vbaProject so that
	// SetVBAProject/ClearVBAProject control the part even when the workbook
	// was not loaded from a package.
	hasVBA := wb.HasVBAProject()
	if hasVBA {
		plan.add("xl/vbaProject.bin", wb.vbaProject.Data, ctVBAProject)
		plan.wbRelExtras = append(plan.wbRelExtras, preservedRel{relType: relTypeVBAProject, target: "vbaProject.bin"})
	}

	// ---- styles.xml (before sheets: generating it assigns CF dxf ids) ----
	var stylesRaw []byte
	stylesRegenerated := false
	// stylesWritten tracks whether the part actually made it into the
	// package; the workbook relationships must only declare it when it did.
	stylesWritten := false
	if hasStyles {
		var ssXML string
		// Reuse the cached part only when nothing changed anywhere; a
		// sheet-level change (e.g. AddConditionalFormatting, which does
		// not set wb.Modified) requires regeneration so new <dxf>
		// entries are built and their ids assigned before sheet XML is
		// written. Cached dxf entries survive regeneration via dxfRaw.
		if !dirty && wb.StylesXML != nil {
			ssXML = string(wb.StylesXML)
		} else {
			generated, err := generateStylesXML(wb)
			if err != nil {
				return fmt.Errorf("saving workbook: styles.xml: %w", err)
			}
			ssXML = generated
			stylesRegenerated = true
		}
		if ssXML != "" {
			plan.add("xl/styles.xml", []byte(ssXML), ctStyles)
			stylesRaw = []byte(ssXML)
			stylesWritten = true
		}
	}

	// ---- Preserve unmodified sheets and reserve their part names ----
	// Done before regenerating anything so that freshly assigned part
	// numbers can never collide with parts kept verbatim.
	var globalPic, globalDraw, globalChart, globalPivot int
	for i, ws := range wb.Worksheets {
		if ws.Modified || ws.sourceXML == nil {
			continue
		}
		sheetPath := fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)
		plan.add(sheetPath, ws.sourceXML, ctWorksheet)
		for _, name := range sortedPartNames(ws.cachedParts) {
			data := ws.cachedParts[name]
			if strings.HasPrefix(name, "xl/worksheets/_rels/") {
				// The sheet is written at its canonical path; its rels
				// part must sit next to that path even when the original
				// package used a different sheet file name.
				name = relsPathFor(sheetPath)
			}
			plan.addCached(name, data)
		}

		globalPic = seedCounter(ws.cachedParts, "xl/media/image", globalPic)
		globalDraw = seedCounter(ws.cachedParts, "xl/drawings/drawing", globalDraw)
		globalChart = seedCounter(ws.cachedParts, "xl/charts/chart", globalChart)
		globalPivot = seedCounter(ws.cachedParts, "xl/pivotTables/pivotTable", globalPivot)
		globalPivot = seedCounter(ws.cachedParts, "xl/pivotCache/pivotCacheDefinition", globalPivot)
	}

	// Pivot caches referenced by preserved pivot tables must keep their
	// <pivotCaches> entries in the regenerated workbook.xml.
	pivotCaches := collectPreservedPivotCaches(wb)

	// ---- Regenerate modified sheets and their feature parts ----
	globalTable := 0
	drawObjID := 0
	for i, ws := range wb.Worksheets {
		if !ws.Modified && ws.sourceXML != nil {
			continue // preserved above
		}
		sheetPath := fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)

		// Companion parts written for this sheet are cached on it so a
		// later unmodified save re-emits exactly these bytes.
		companions := make(map[string][]byte)
		addCompanion := func(name string, data []byte, contentType string) {
			plan.add(name, data, contentType)
			companions[name] = data
		}

		// Tables. Names are lower-cased into part paths and deduplicated
		// across the workbook; the id attribute is globally unique.
		tablePaths := make([]string, 0, len(ws.Tables))
		for _, t := range ws.Tables {
			globalTable++
			base := strings.ToLower(strings.TrimSpace(t.Name))
			if base == "" {
				base = fmt.Sprintf("table%d", globalTable)
			}
			tp := plan.unique(fmt.Sprintf("xl/tables/%s.xml", base))
			tableXML, err := generateTableXML(t, globalTable)
			if err != nil {
				return fmt.Errorf("saving workbook: table %q: %w", t.Name, err)
			}
			addCompanion(tp, []byte(tableXML), ctTable)
			tablePaths = append(tablePaths, tp)
		}

		// Charts. They are referenced from the drawing part (graphicFrame),
		// not from the sheet rels.
		chartTargets := make([]string, 0, len(ws.Charts))
		for _, ch := range ws.Charts {
			globalChart++
			cp := fmt.Sprintf("xl/charts/chart%d.xml", globalChart)
			addCompanion(cp, []byte(generateChartXML(ch, ws.Name)), ctChart)
			chartTargets = append(chartTargets, fmt.Sprintf("../charts/chart%d.xml", globalChart))
		}

		// Drawing: pictures and charts share one drawing part per sheet.
		var drawingPath string
		if len(ws.Pictures) > 0 || len(ws.Charts) > 0 {
			globalDraw++
			drawingPath = fmt.Sprintf("xl/drawings/drawing%d.xml", globalDraw)

			mediaTargets := make([]string, 0, len(ws.Pictures))
			for _, pic := range ws.Pictures {
				ext := pic.Format
				if ext == "jpeg" {
					ext = "jpg"
				}
				globalPic++
				mediaPath := fmt.Sprintf("xl/media/image%d.%s", globalPic, ext)
				addCompanion(mediaPath, pic.Data, "")
				mediaTargets = append(mediaTargets, fmt.Sprintf("../media/image%d.%s", globalPic, ext))
			}

			drawingXML, err := generateDrawingXML(ws.Pictures, ws.Charts, drawObjID)
			if err != nil {
				return fmt.Errorf("saving workbook: sheet %q drawing: %w", ws.Name, err)
			}
			addCompanion(drawingPath, []byte(drawingXML), ctDrawing)
			drawObjID += len(ws.Pictures) + len(ws.Charts)
			addCompanion(relsPathFor(drawingPath),
				[]byte(generateDrawingRelsXML(mediaTargets, chartTargets)), "")
		}

		// Pivot tables with their cache definition and records. One global
		// counter keeps part names, cache ids, and relationship wiring in
		// sync across sheets.
		pivotPaths := make([]string, 0, len(ws.PivotTables))
		for _, pt := range ws.PivotTables {
			globalPivot++
			n := globalPivot

			ptPath := fmt.Sprintf("xl/pivotTables/pivotTable%d.xml", n)
			pivotXML, err := generatePivotTableXML(pt, ws, n)
			if err != nil {
				return fmt.Errorf("saving workbook: pivot table on sheet %q: %w", ws.Name, err)
			}
			addCompanion(ptPath, []byte(pivotXML), ctPivotTable)
			addCompanion(relsPathFor(ptPath),
				[]byte(pivotTableRelsXML(fmt.Sprintf("../pivotCache/pivotCacheDefinition%d.xml", n))), "")

			defPath := fmt.Sprintf("xl/pivotCache/pivotCacheDefinition%d.xml", n)
			addCompanion(defPath, []byte(generatePivotCacheDefinitionXML(pt, ws)), ctPivotCacheDefn)
			addCompanion(relsPathFor(defPath),
				[]byte(pivotCacheRelsXML(fmt.Sprintf("pivotCacheRecords%d.xml", n))), "")

			recPath := fmt.Sprintf("xl/pivotCache/pivotCacheRecords%d.xml", n)
			addCompanion(recPath, []byte(generatePivotCacheRecordsXML(pt)), ctPivotCacheRecords)

			pivotPaths = append(pivotPaths, ptPath)
			pivotCaches = append(pivotCaches, workbookPivotCache{cacheID: n, defPath: defPath})
		}

		// Sheet rels.
		if relsXML := generateUnifiedSheetRelsXML(tablePaths, drawingPath, pivotPaths, ws.Hyperlinks); relsXML != "" {
			addCompanion(relsPathFor(sheetPath), []byte(relsXML), "")
		}

		// Sheet XML.
		sheetContent, err := generateSheetXML(ws, ssIndex)
		if err != nil {
			return fmt.Errorf("saving workbook: sheet %q: %w", ws.Name, err)
		}
		plan.add(sheetPath, []byte(sheetContent), ctWorksheet)

		// Refresh the per-sheet caches so a subsequent unmodified save
		// reuses exactly what was just written.
		ws.sourceXML = []byte(sheetContent)
		ws.cachedParts = companions
	}

	// ---- workbook.xml + rels ----
	// A macro-enabled workbook must declare the macro-enabled main content
	// type; Excel refuses to load VBA from a plain .xlsx-typed package.
	wbCT := ctWorkbookMain
	if hasVBA {
		wbCT = ctWorkbookMacro
	}
	wbXML, err := generateWorkbookXML(wb.Worksheets, hasSST, stylesWritten, pivotCaches)
	if err != nil {
		return fmt.Errorf("saving workbook: workbook.xml: %w", err)
	}
	plan.add("xl/workbook.xml", []byte(wbXML), wbCT)
	plan.add("xl/_rels/workbook.xml.rels",
		[]byte(generateWorkbookRelsXML(wb.Worksheets, hasSST, stylesWritten, pivotCaches, plan.wbRelExtras)), "")

	// ---- sharedStrings.xml ----
	var sstRaw []byte
	switch {
	case reuseSST:
		sstRaw = wb.sstRaw
		plan.add("xl/sharedStrings.xml", sstRaw, ctSharedStrings)
	case len(ssTable) > 0:
		ssXML, err := generateSSTXML(ssTable)
		if err != nil {
			return fmt.Errorf("saving workbook: sharedStrings.xml: %w", err)
		}
		plan.add("xl/sharedStrings.xml", []byte(ssXML), ctSharedStrings)
		sstRaw = []byte(ssXML)
	}

	// ---- Write the package ----
	var zipBuf bytes.Buffer
	zw := zip.NewWriter(&zipBuf)
	if err := plan.writeTo(zw); err != nil {
		zw.Close()
		return fmt.Errorf("saving workbook: %w", err)
	}
	if err := zw.Close(); err != nil {
		return fmt.Errorf("saving workbook: %w", err)
	}
	zipBytes := zipBuf.Bytes()

	outData, err := packageBytes(wb, zipBytes)
	if err != nil {
		return fmt.Errorf("saving workbook: %w", err)
	}
	if err := emitWorkbook(wb, path, w, outData); err != nil {
		return err
	}

	// After a successful save the workbook is no longer "dirty". The cached
	// snapshots are refreshed to what was just written so a subsequent
	// unmodified save takes the verbatim fast path.
	wb.Modified = false
	wb.SourceXML = nil
	if len(wb.Worksheets) > 0 {
		wb.SourceXML = wb.Worksheets[0].sourceXML
	}
	wb.StylesXML = stylesRaw
	wb.sstRaw = sstRaw
	if sstRaw != nil {
		wb.sstBase = ssTable
	} else {
		wb.sstBase = nil
	}
	wb.sourceZip = zipBytes
	wb.origCT = plan.contentTypes()
	if stylesRegenerated && stylesRaw != nil {
		// The merged dxfs block just written becomes the baseline for the
		// next regeneration, keeping cached dxfIds in saved sheet XML valid.
		cacheDxfBlock(wb, stylesRaw)
	}
	for _, ws := range wb.Worksheets {
		ws.Modified = false
	}

	return nil
}

// emitWorkbook delivers finished package bytes to their destination: the
// writer w when one is given, otherwise the file at path. wb.FilePath is
// updated for file saves only — a stream has no file location.
func emitWorkbook(wb *Workbook, path string, w io.Writer, outData []byte) error {
	if w != nil {
		if _, err := w.Write(outData); err != nil {
			return fmt.Errorf("saving workbook: %w", err)
		}
		return nil
	}
	if err := os.WriteFile(path, outData, 0644); err != nil {
		return fmt.Errorf("saving workbook: %w", err)
	}
	wb.FilePath = path
	return nil
}

// packageBytes applies password encryption to finished package bytes when a
// password is set, and returns them unchanged otherwise.
func packageBytes(wb *Workbook, zipBytes []byte) ([]byte, error) {
	if wb.password == "" {
		return zipBytes, nil
	}
	infoXML, encPkg, origSize, err := encryptPackage(zipBytes, wb.password)
	if err != nil {
		return nil, fmt.Errorf("encryption: %w", err)
	}
	// Wrap in an OLE/CFB container so Excel can open it.
	out, err := writeEncryptedCFB(infoXML, encPkg, origSize)
	if err != nil {
		return nil, fmt.Errorf("CFB wrap: %w", err)
	}
	return out, nil
}

// ---------------------------------------------------------------------------
// Preserved pivot caches
// ---------------------------------------------------------------------------

// workbookPivotCache is one <pivotCache> entry for workbook.xml: a cache id
// and the package path of its definition part.
type workbookPivotCache struct {
	cacheID int
	defPath string
}

// collectPreservedPivotCaches finds the pivot cache definitions referenced
// by pivot tables preserved verbatim from unmodified sheets. The regenerated
// workbook.xml needs a <pivotCaches> entry — and workbook.xml.rels a
// relationship — for each cache in the package, including preserved ones.
func collectPreservedPivotCaches(wb *Workbook) []workbookPivotCache {
	seen := make(map[string]bool)
	var out []workbookPivotCache
	for _, ws := range wb.Worksheets {
		if ws.Modified || ws.sourceXML == nil {
			continue
		}
		for _, name := range sortedPartNames(ws.cachedParts) {
			if !strings.HasPrefix(name, "xl/pivotTables/") || !strings.HasSuffix(name, ".xml") {
				continue
			}
			id := extractCacheID(ws.cachedParts[name])
			if id < 0 {
				continue
			}
			def := pivotCacheDefTarget(name, ws.cachedParts[relsPathFor(name)])
			if def == "" || seen[def] {
				continue
			}
			seen[def] = true
			out = append(out, workbookPivotCache{cacheID: id, defPath: def})
		}
	}
	return out
}

// extractCacheID reads the cacheId attribute from a pivotTableDefinition
// part, returning -1 when absent or unparsable.
func extractCacheID(data []byte) int {
	const attr = `cacheId="`
	i := bytes.Index(data, []byte(attr))
	if i < 0 {
		return -1
	}
	rest := data[i+len(attr):]
	j := bytes.IndexByte(rest, '"')
	if j < 0 {
		return -1
	}
	n, err := strconv.Atoi(string(rest[:j]))
	if err != nil {
		return -1
	}
	return n
}

// pivotCacheDefTarget resolves the pivotCacheDefinition relationship of a
// pivot table part from its .rels bytes.
func pivotCacheDefTarget(pivotTablePath string, relsData []byte) string {
	if len(relsData) == 0 {
		return ""
	}
	var rels xmlRelationships
	if err := xml.Unmarshal(relsData, &rels); err != nil {
		return ""
	}
	for _, rel := range rels.Rels {
		if rel.TargetMode == "External" {
			continue
		}
		if strings.HasSuffix(rel.Type, "/pivotCacheDefinition") {
			return resolvePartPath(pivotTablePath, rel.Target)
		}
	}
	return ""
}

// ---------------------------------------------------------------------------
// workbook.xml generation
// ---------------------------------------------------------------------------

// workbookRelsNextID returns the first rId available after the sheet,
// sharedStrings, and styles relationships in workbook.xml.rels.
func workbookRelsNextID(worksheets []*Worksheet, hasSST, hasStyles bool) int {
	n := len(worksheets) + 1
	if hasSST {
		n++
	}
	if hasStyles {
		n++
	}
	return n
}

func generateWorkbookXML(worksheets []*Worksheet, hasSST, hasStyles bool, pivotCaches []workbookPivotCache) (string, error) {
	sheets := make([]outSheet, 0, len(worksheets))
	for i, ws := range worksheets {
		sheets = append(sheets, outSheet{
			Name:    ws.Name,
			SheetID: strconv.Itoa(i + 1),
			RID:     fmt.Sprintf("rId%d", i+1),
		})
	}

	wb := outWorkbookXML{
		R:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		Sheets: outSheets{Sheet: sheets},
	}
	if len(pivotCaches) > 0 {
		pc := &outPivotCaches{Caches: make([]outPivotCacheRef, len(pivotCaches))}
		rid := workbookRelsNextID(worksheets, hasSST, hasStyles)
		for k, c := range pivotCaches {
			pc.Caches[k] = outPivotCacheRef{
				CacheID: c.cacheID,
				RID:     fmt.Sprintf("rId%d", rid+k),
			}
		}
		wb.PivotCaches = pc
	}

	return marshalXML(wb)
}

// ---------------------------------------------------------------------------
// workbook.xml.rels generation
// ---------------------------------------------------------------------------

// generateWorkbookRelsXML builds xl/_rels/workbook.xml.rels. extras holds
// relationships toward parts preserved verbatim from the loaded package
// (theme, VBA project); they are appended with fresh sequential ids since
// nothing inside workbook.xml references them by r:id.
func generateWorkbookRelsXML(worksheets []*Worksheet, hasSST, hasStyles bool, pivotCaches []workbookPivotCache, extras []preservedRel) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")
	nextID := 1
	for i := range worksheets {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>`+"\n", nextID, i+1)
		nextID++
	}
	// Only reference parts that are actually written (see packagePlan).
	if hasSST {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`+"\n", nextID)
		nextID++
	}
	if hasStyles {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`+"\n", nextID)
		nextID++
	}
	for _, c := range pivotCaches {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="%s"/>`+"\n",
			nextID, strings.TrimPrefix(c.defPath, "xl/"))
		nextID++
	}
	for _, rel := range extras {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="%s" Target="%s"/>`+"\n", nextID, rel.relType, rel.target)
		nextID++
	}
	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// ---------------------------------------------------------------------------
// sheet XML generation
// ---------------------------------------------------------------------------

// rowCell pairs a cell with its numeric column index, so that ordering by
// column does not re-parse the A1 reference on every comparison.
type rowCell struct {
	col  int
	cell *Cell
}

func generateSheetXML(ws *Worksheet, ssIndex map[string]int) (string, error) {
	// Group cells by row, parsing each A1 reference exactly once.
	rows := make(map[int][]rowCell)
	for _, c := range ws.cells.All() {
		colName, row := splitRef(c.Ref)
		rows[row] = append(rows[row], rowCell{col: colToNum(colName), cell: c})
	}

	// Rows with an explicit height but no cells still need a <row> element
	// so the height survives.
	rowSet := make(map[int]bool, len(rows)+len(ws.RowHeights))
	for r := range rows {
		rowSet[r] = true
	}
	for r, h := range ws.RowHeights {
		if h > 0 {
			rowSet[r] = true
		}
	}
	rowNums := make([]int, 0, len(rowSet))
	for r := range rowSet {
		rowNums = append(rowNums, r)
	}
	sort.Ints(rowNums)

	// Build output rows with cells in ascending column order.
	outRows := make([]outRow, 0, len(rowNums))
	for _, r := range rowNums {
		cells := rows[r]
		sort.Slice(cells, func(i, j int) bool { return cells[i].col < cells[j].col })

		outCells := make([]outCell, 0, len(cells))
		for _, rc := range cells {
			outCells = append(outCells, cellToOutCell(rc.cell, ssIndex))
		}

		or := outRow{R: r, Cells: outCells}
		if h, ok := ws.RowHeights[r]; ok && h > 0 {
			or.Ht = h
			or.CustomHeight = 1
		}
		outRows = append(outRows, or)
	}

	wsOut := outWorksheet{
		SheetViews:            buildSheetViews(ws),
		Cols:                  buildCols(ws),
		SheetData:             outSheetData{Rows: outRows},
		ConditionalFormatting: buildConditionalFormattings(ws.ConditionalFormattings),
		DataValidations:       buildDataValidations(ws.DataValidations),
		Hyperlinks:            buildHyperlinks(ws),
		TableParts:            buildTableParts(ws.Tables),
	}
	if ws.AutoFilterRange != "" {
		wsOut.AutoFilter = &outAutoFilter{Ref: ws.AutoFilterRange}
	}
	if len(ws.MergedCells) > 0 {
		merges := make([]outMergeCell, 0, len(ws.MergedCells))
		for _, ref := range ws.MergedCells {
			merges = append(merges, outMergeCell{Ref: ref})
		}
		wsOut.MergeCells = &outMergeCells{Count: len(merges), Merges: merges}
	}
	if len(ws.Pictures) > 0 || len(ws.Charts) > 0 {
		wsOut.Drawing = &outDrawingRef{RID: drawingRID(ws.Tables)}
	}
	return marshalXML(wsOut)
}

// buildSheetViews emits the <sheetViews> block with a frozen <pane> when
// the sheet has frozen rows or columns, and nil otherwise.
func buildSheetViews(ws *Worksheet) *outSheetViews {
	if ws.FrozenRow <= 0 && ws.FrozenCol <= 0 {
		return nil
	}
	pane := &outPane{
		State: "frozen",
		// numToCol is 0-based: the first unfrozen column index equals the
		// frozen column count; rows are 1-based.
		TopLeftCell: numToCol(ws.FrozenCol) + strconv.Itoa(ws.FrozenRow+1),
	}
	switch {
	case ws.FrozenCol > 0 && ws.FrozenRow > 0:
		pane.XSplit = ws.FrozenCol
		pane.YSplit = ws.FrozenRow
		pane.ActivePane = "bottomRight"
	case ws.FrozenCol > 0:
		pane.XSplit = ws.FrozenCol
		pane.ActivePane = "topRight"
	default:
		pane.YSplit = ws.FrozenRow
		pane.ActivePane = "bottomLeft"
	}
	return &outSheetViews{Views: []outSheetView{{WorkbookViewID: 0, Pane: pane}}}
}

// buildCols emits one <col> entry per sized column, in ascending column
// order, and nil when no widths are set.
func buildCols(ws *Worksheet) *outCols {
	if len(ws.ColumnWidths) == 0 {
		return nil
	}
	nums := make([]int, 0, len(ws.ColumnWidths))
	byNum := make(map[int]float64, len(ws.ColumnWidths))
	for letter, w := range ws.ColumnWidths {
		if w <= 0 {
			continue
		}
		n := colToNum(strings.ToUpper(strings.TrimSpace(letter))) + 1 // 1-based <col min/max>
		if n < 1 || n > maxColCount {
			continue
		}
		if _, seen := byNum[n]; !seen {
			nums = append(nums, n)
		}
		byNum[n] = w
	}
	if len(nums) == 0 {
		return nil
	}
	sort.Ints(nums)
	cols := make([]outCol, 0, len(nums))
	for _, n := range nums {
		cols = append(cols, outCol{Min: n, Max: n, Width: byNum[n], CustomWidth: 1})
	}
	return &outCols{Cols: cols}
}

// buildHyperlinks emits the <hyperlinks> entries of a sheet. External
// links take sequential rIds following the sheet relationship order —
// tables, then the drawing, then pivot tables, then hyperlinks — which
// generateUnifiedSheetRelsXML mirrors; internal links carry a location
// attribute instead and consume no relationship.
func buildHyperlinks(ws *Worksheet) *outHyperlinks {
	if len(ws.Hyperlinks) == 0 {
		return nil
	}
	base := len(ws.Tables) + len(ws.PivotTables)
	if len(ws.Pictures) > 0 || len(ws.Charts) > 0 {
		base++
	}
	links := make([]outHyperlink, 0, len(ws.Hyperlinks))
	next := 0
	for _, h := range ws.Hyperlinks {
		link := outHyperlink{Ref: h.Ref, Location: h.Location, Tooltip: h.Tooltip, Display: h.Display}
		if h.Target != "" {
			next++
			link.RID = fmt.Sprintf("rId%d", base+next)
		}
		links = append(links, link)
	}
	return &outHyperlinks{Links: links}
}

// cellToOutCell converts the in-memory Cell to the output XML form,
// determining the correct t attribute and resolving shared-string indices.
func cellToOutCell(c *Cell, ssIndex map[string]int) outCell {
	oc := outCell{
		Ref: c.Ref,
		S:   c.StyleID,
	}

	if c.Formula != "" {
		oc.F = c.Formula
	}

	switch v := c.Value.(type) {
	case nil:
		// Empty cell — emit the <c> element without a value so the cell
		// reference is preserved, but no content.
		return oc

	case bool:
		oc.T = "b"
		if v {
			oc.V = "1"
		} else {
			oc.V = "0"
		}

	case string:
		idx, ok := ssIndex[v]
		if ok {
			oc.T = "s"
			oc.V = strconv.Itoa(idx)
		} else {
			// String not in shared-strings table — write as an inline
			// formula-result string so the value survives even in this
			// shouldn't-happen case.
			oc.T = "str"
			oc.V = v
		}

	case int:
		oc.V = strconv.Itoa(v)
	case int64:
		oc.V = strconv.FormatInt(v, 10)
	case float64:
		oc.V = strconv.FormatFloat(v, 'g', -1, 64)
	case float32:
		oc.V = strconv.FormatFloat(float64(v), 'g', -1, 32)

	case ErrorValue:
		// Excel error value — stored as t="e" with the error code in <v>.
		oc.T = "e"
		oc.V = string(v)

	default:
		// Fallback: convert to string via fmt.Sprint.
		oc.T = "str"
		oc.V = fmt.Sprint(v)
	}

	return oc
}

// ---------------------------------------------------------------------------
// Shared strings
// ---------------------------------------------------------------------------

// buildSharedStrings collects the shared-string table for the workbook and
// returns both the ordered table and a lookup map from string to index.
//
// When the workbook was loaded from a file, the original table is kept as a
// prefix in its original order: cached worksheet XML references strings by
// index, so re-sorting would corrupt every sheet that is written back
// verbatim. New strings are appended after the original entries. For
// workbooks created from scratch the table is sorted for deterministic
// output.
func buildSharedStrings(wb *Workbook) ([]string, map[string]int) {
	index := make(map[string]int, len(wb.sstBase))
	table := make([]string, 0, len(wb.sstBase))
	for _, s := range wb.sstBase {
		if _, exists := index[s]; exists {
			continue // defensive: an original table should be unique
		}
		index[s] = len(table)
		table = append(table, s)
	}

	added := 0
	for _, ws := range wb.Worksheets {
		if ws.cells == nil {
			continue
		}
		for _, c := range ws.cells.All() {
			s, ok := c.Value.(string)
			if !ok {
				continue
			}
			if _, exists := index[s]; exists {
				continue
			}
			index[s] = len(table)
			table = append(table, s)
			added++
		}
	}

	if len(wb.sstBase) == 0 && added > 0 {
		// Deterministic output for workbooks created from scratch.
		sort.Strings(table)
		for i, s := range table {
			index[s] = i
		}
	}
	return table, index
}

func generateSSTXML(table []string) (string, error) {
	items := make([]outSI, len(table))
	for i, s := range table {
		items[i] = outSI{T: s}
	}
	sst := outSST{
		Count:       len(table),
		UniqueCount: len(table),
		Items:       items,
	}
	return marshalXML(sst)
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// splitRef decomposes an A1-style reference like "AA10" into its column
// ("AA") and row number (10).
func splitRef(ref string) (col string, row int) {
	ref = strings.ToUpper(strings.TrimSpace(ref))
	// Find the boundary between letters and digits.
	pos := 0
	for pos < len(ref) && ((ref[pos] >= 'A' && ref[pos] <= 'Z') || (ref[pos] >= 'a' && ref[pos] <= 'z')) {
		pos++
	}
	col = strings.ToUpper(ref[:pos])
	if pos < len(ref) {
		if n, err := strconv.Atoi(ref[pos:]); err == nil {
			row = n
		}
	}
	return
}

// marshalXML marshals v to indented XML, prepends the standard XML header,
// and returns the result as a string ready for writing into a ZIP entry.
func marshalXML(v interface{}) (string, error) {
	body, err := xml.MarshalIndent(v, "", "  ")
	if err != nil {
		return "", fmt.Errorf("xml marshal failed: %w", err)
	}
	return xml.Header + string(body) + "\n", nil
}

// writeZipString creates a ZIP entry named name and writes content into it.
func writeZipString(zw *zip.Writer, name, content string) error {
	fw, err := zw.Create(name)
	if err != nil {
		return err
	}
	_, err = fw.Write([]byte(content))
	return err
}
