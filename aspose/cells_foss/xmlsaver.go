package cells_foss

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"os"
	"sort"
	"strconv"
	"strings"
)

// ---------------------------------------------------------------------------
// Output XML types (used exclusively for marshaling, never for parsing)
// ---------------------------------------------------------------------------

type outWorkbookXML struct {
	XMLName xml.Name   `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	R       string     `xml:"xmlns:r,attr"`
	Sheets  outSheets  `xml:"sheets"`
}

type outSheets struct {
	Sheet []outSheet `xml:"sheet"`
}

type outSheet struct {
	Name    string `xml:"name,attr"`
	SheetID string `xml:"sheetId,attr"`
	RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

type outWorksheet struct {
	XMLName         xml.Name            `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	SheetData       outSheetData        `xml:"sheetData"`
	DataValidations *outDataValidations `xml:"dataValidations,omitempty"`
	TableParts      *outTableParts      `xml:"tableParts,omitempty"`
	Drawing         *outDrawingRef      `xml:"drawing,omitempty"`
}

type outDrawingRef struct {
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

type outSheetData struct {
	Rows []outRow `xml:"row"`
}

type outRow struct {
	R     int      `xml:"r,attr"`
	Cells []outCell `xml:"c"`
}

type outCell struct {
	Ref string `xml:"r,attr"`
	T   string `xml:"t,attr,omitempty"`
	S   int    `xml:"s,attr,omitempty"`
	V   string `xml:"v,omitempty"`
	F   string `xml:"f,omitempty"`
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
// SaveWorkbook
// ---------------------------------------------------------------------------

// SaveWorkbook writes the workbook to an .xlsx file at path.
//
// When the workbook was loaded from a file and has not been modified, the
// cached SourceXML is reused verbatim for each worksheet so that round-trip
// fidelity is preserved byte-for-byte on unchanged content.
//
// When the workbook is new or has been modified, all XML parts are
// regenerated: a shared-strings table is built from every string cell, sheet
// XML is marshaled, and the OPC scaffolding ([Content_Types].xml, _rels, …)
// is written fresh.
func SaveWorkbook(wb *Workbook, path string) error {
	if wb == nil {
		return fmt.Errorf("saving workbook: workbook is nil")
	}

	// Build the ZIP archive in memory.
	var zipBuf bytes.Buffer
	zw := zip.NewWriter(&zipBuf)

	if err := writeOPCScaffolding(zw, wb); err != nil {
		zw.Close()
		return fmt.Errorf("saving workbook: %w", err)
	}

	// ---- workbook.xml ----
	wbXML := generateWorkbookXML(wb.Worksheets)
	if err := writeZipString(zw, "xl/workbook.xml", wbXML); err != nil {
		zw.Close()
		return fmt.Errorf("saving workbook: %w", err)
	}

	// ---- workbook.xml.rels ----
	relsXML := generateWorkbookRelsXML(wb.Worksheets)
	if err := writeZipString(zw, "xl/_rels/workbook.xml.rels", relsXML); err != nil {
		zw.Close()
		return fmt.Errorf("saving workbook: %w", err)
	}

	// ---- styles.xml ----
	hasStyles := len(wb.styles) > 1 || wb.StylesXML != nil
	if hasStyles {
		var ssXML string
		if !wb.Modified && wb.StylesXML != nil {
			ssXML = string(wb.StylesXML)
		} else {
			ssXML = generateStylesXML(wb)
		}
		if ssXML != "" {
			if err := writeZipString(zw, "xl/styles.xml", ssXML); err != nil {
				zw.Close()
				return fmt.Errorf("saving workbook: %w", err)
			}
		}
	}

	// ---- Build shared-strings table ----
	ssTable, ssIndex := buildSharedStrings(wb.Worksheets)

	// ---- Sheet XML + tables ----
	for i, ws := range wb.Worksheets {
		sheetPath := fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)

		var sheetContent string
		if !wb.Modified && wb.SourceXML != nil && i == 0 {
			sheetContent = string(wb.SourceXML)
		} else {
			sheetContent = generateSheetXML(ws, ssIndex)
		}

		if err := writeZipString(zw, sheetPath, sheetContent); err != nil {
			zw.Close()
			return fmt.Errorf("saving workbook: %w", err)
		}

		// ---- Table XML files ----
		for tid, t := range ws.Tables {
			tablePath := fmt.Sprintf("xl/tables/%s.xml", strings.ToLower(t.Name))
			tableXML := generateTableXML(t, tid+1)
			if err := writeZipString(zw, tablePath, tableXML); err != nil {
				zw.Close()
				return fmt.Errorf("saving workbook: table %q: %w", t.Name, err)
			}
		}

		// ---- Sheet rels (unified: tables + drawing) ----
		unifiedRels := generateUnifiedSheetRelsXML(ws.Tables, ws.Pictures)
		if unifiedRels != "" {
			relsPath := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", i+1)
			if err := writeZipString(zw, relsPath, unifiedRels); err != nil {
				zw.Close()
				return fmt.Errorf("saving workbook: sheet rels: %w", err)
			}
		}

		// ---- Drawing XML + drawing rels + media ----
		if len(ws.Pictures) > 0 {
			drawingPath := fmt.Sprintf("xl/drawings/drawing%d.xml", i+1)
			drawingXML := generateDrawingXML(ws.Pictures, 0)
			if err := writeZipString(zw, drawingPath, drawingXML); err != nil {
				zw.Close()
				return fmt.Errorf("saving workbook: drawing: %w", err)
			}

			drRelsPath := fmt.Sprintf("xl/drawings/_rels/drawing%d.xml.rels", i+1)
			drRelsXML := generateDrawingRelsXML(ws.Pictures)
			if err := writeZipString(zw, drRelsPath, drRelsXML); err != nil {
				zw.Close()
				return fmt.Errorf("saving workbook: drawing rels: %w", err)
			}

			for pi, pic := range ws.Pictures {
				ext := pic.Format
				if ext == "jpeg" {
					ext = "jpg"
				}
				mediaPath := fmt.Sprintf("xl/media/image%d.%s", pi+1, ext)
				fw, err := zw.Create(mediaPath)
				if err != nil {
					zw.Close()
					return fmt.Errorf("saving workbook: media: %w", err)
				}
				if _, err := fw.Write(pic.Data); err != nil {
					zw.Close()
					return fmt.Errorf("saving workbook: media: %w", err)
				}
			}
		}
	}

	// ---- sharedStrings.xml (only when there are strings) ----
	if len(ssTable) > 0 {
		ssXML := generateSSTXML(ssTable)
		if err := writeZipString(zw, "xl/sharedStrings.xml", ssXML); err != nil {
			zw.Close()
			return fmt.Errorf("saving workbook: %w", err)
		}
	}

	if err := zw.Close(); err != nil {
		return fmt.Errorf("saving workbook: %w", err)
	}

	// ---- Write to disk (encrypt if password is set) ----
	zipBytes := zipBuf.Bytes()

	var outData []byte
	if wb.password != "" {
		infoXML, encPkg, err := encryptPackage(zipBytes, wb.password)
		if err != nil {
			return fmt.Errorf("saving workbook: encryption: %w", err)
		}
		outData = writeEncryptedFile(infoXML, encPkg)
	} else {
		outData = zipBytes
	}

	if err := os.WriteFile(path, outData, 0644); err != nil {
		return fmt.Errorf("saving workbook: %w", err)
	}

	// After a successful save the workbook is no longer "dirty" and the
	// SourceXML cache is invalidated so the next Save regenerates.
	wb.Modified = false
	wb.SourceXML = nil
	wb.FilePath = path

	return nil
}

// isWorkbookModified reports whether any cell in the workbook has been
// changed since the last load or save.
func isWorkbookModified(wb *Workbook) bool {
	if wb == nil {
		return false
	}
	return wb.Modified
}

// ---------------------------------------------------------------------------
// OPC scaffolding (content types, package relationships)
// ---------------------------------------------------------------------------

func writeOPCScaffolding(zw *zip.Writer, wb *Workbook) error {
	var ct strings.Builder
	ct.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	ct.WriteString(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` + "\n")
	ct.WriteString(`  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` + "\n")
	ct.WriteString(`  <Default Extension="xml" ContentType="application/xml"/>` + "\n")
	ct.WriteString(`  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` + "\n")
	for i := range wb.Worksheets {
		fmt.Fprintf(&ct, `  <Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`+"\n", i+1)
	}
	ct.WriteString(`  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` + "\n")
	ct.WriteString(`  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` + "\n")
	ct.WriteString(generateContentTypesForTables(wb.Worksheets))
	ct.WriteString(generateContentTypesForDrawings(wb.Worksheets))
	ct.WriteString(`</Types>` + "\n")

	if err := writeZipString(zw, "[Content_Types].xml", ct.String()); err != nil {
		return err
	}

	rels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
	return writeZipString(zw, "_rels/.rels", rels)
}

// ---------------------------------------------------------------------------
// workbook.xml generation
// ---------------------------------------------------------------------------

func generateWorkbookXML(worksheets []*Worksheet) string {
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

	return marshalXML(wb)
}

// ---------------------------------------------------------------------------
// workbook.xml.rels generation
// ---------------------------------------------------------------------------

func generateWorkbookRelsXML(worksheets []*Worksheet) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + "\n")
	nextID := 1
	for i := range worksheets {
		fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>`+"\n", nextID, i+1)
		nextID++
	}
	fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`+"\n", nextID)
	nextID++
	fmt.Fprintf(&b, `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`+"\n", nextID)
	b.WriteString(`</Relationships>` + "\n")
	return b.String()
}

// ---------------------------------------------------------------------------
// sheet XML generation
// ---------------------------------------------------------------------------

// cellGroup holds cells grouped by their row number for ordered output.
type cellGroup struct {
	row   int
	cells []*Cell
}

func generateSheetXML(ws *Worksheet, ssIndex map[string]int) string {
	// Group cells by row.
	rows := make(map[int][]*Cell)
	for _, c := range ws.cells.All() {
		_, row := splitRef(c.Ref)
		rows[row] = append(rows[row], c)
	}

	// Sort rows.
	groups := make([]cellGroup, 0, len(rows))
	for r, cells := range rows {
		groups = append(groups, cellGroup{row: r, cells: cells})
	}
	sort.Slice(groups, func(i, j int) bool { return groups[i].row < groups[j].row })

	// Build output rows with cells sorted by column within each row.
	outRows := make([]outRow, 0, len(groups))
	for _, g := range groups {
		sort.Slice(g.cells, func(i, j int) bool {
			ci, _ := splitRef(g.cells[i].Ref)
			cj, _ := splitRef(g.cells[j].Ref)
			return ci < cj
		})

		outCells := make([]outCell, 0, len(g.cells))
		for _, c := range g.cells {
			oc := cellToOutCell(c, ssIndex)
			outCells = append(outCells, oc)
		}

		outRows = append(outRows, outRow{R: g.row, Cells: outCells})
	}

	wsOut := outWorksheet{
		SheetData:       outSheetData{Rows: outRows},
		DataValidations: buildDataValidations(ws.DataValidations),
		TableParts:      buildTableParts(ws.Tables),
	}
	if len(ws.Pictures) > 0 {
		wsOut.Drawing = &outDrawingRef{RID: drawingRID(ws.Tables)}
	}
	return marshalXML(wsOut)
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
			// String not in shared-strings table — write inline.
			// A well-formed table always contains every string, so this
			// branch is a safety net.
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

	default:
		// Fallback: convert to string via fmt.Sprint.
		oc.V = fmt.Sprint(v)
	}

	return oc
}

// ---------------------------------------------------------------------------
// Shared strings
// ---------------------------------------------------------------------------

// buildSharedStrings scans all worksheets, collects unique string cell values,
// and returns both the ordered table and a lookup map from string to index.
func buildSharedStrings(worksheets []*Worksheet) ([]string, map[string]int) {
	seen := make(map[string]struct{})
	var table []string

	for _, ws := range worksheets {
		for _, c := range ws.cells.All() {
			s, ok := c.Value.(string)
			if !ok {
				continue
			}
			if _, exists := seen[s]; exists {
				continue
			}
			seen[s] = struct{}{}
			table = append(table, s)
		}
	}

	// Deterministic output: sort strings so the table is stable.
	sort.Strings(table)

	index := make(map[string]int, len(table))
	for i, s := range table {
		index[s] = i
	}
	return table, index
}

func generateSSTXML(table []string) string {
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
	for pos < len(ref) && ref[pos] >= 'A' && ref[pos] <= 'Z' {
		pos++
	}
	col = ref[:pos]
	if pos < len(ref) {
		row, _ = strconv.Atoi(ref[pos:])
	}
	return
}

// marshalXML marshals v to indented XML, prepends the standard XML header,
// and returns the result as a string ready for writing into a ZIP entry.
func marshalXML(v interface{}) string {
	body, err := xml.MarshalIndent(v, "", "  ")
	if err != nil {
		// All types passed to this function are simple structs; a marshal
		// error here is a programming mistake.
		panic(fmt.Sprintf("cells_foss: xml marshal failed: %v", err))
	}
	return xml.Header + string(body) + "\n"
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
