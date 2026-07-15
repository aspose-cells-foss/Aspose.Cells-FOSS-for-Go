package cells_foss

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// RowCallback is invoked by StreamingReader.ProcessRows once for every row
// in the worksheet.  rowIndex is the 1-based row number; cells maps A1-style
// references to their string values.  Return a non-nil error to stop
// processing early.
type RowCallback func(rowIndex int, cells map[string]string) error

// StreamingReader reads an .xlsx workbook row by row without loading the
// entire sheet XML into memory.  It is suitable for files with hundreds of
// thousands of rows where a full Workbook load would be too expensive.
type StreamingReader struct {
	path string
}

// NewStreamingReader creates a StreamingReader for the .xlsx file at path.
// The file is not opened until ProcessRows is called.
func NewStreamingReader(path string) *StreamingReader {
	return &StreamingReader{path: path}
}

// ProcessRows opens the workbook, resolves the named sheet to its XML part,
// loads the shared-strings table (if present), and then streams through the
// sheet data one row at a time, calling callback for each row.
//
// When sheetName is empty the first sheet in the workbook is used.  The
// shared-strings table is held in memory, but the sheet XML is never fully
// buffered — peak memory is proportional to the widest row, not the file size.
func (sr *StreamingReader) ProcessRows(sheetName string, callback RowCallback) error {
	if callback == nil {
		return fmt.Errorf("streaming reader: callback is nil")
	}

	// 1. Open the ZIP archive.
	r, err := zip.OpenReader(sr.path)
	if err != nil {
		return fmt.Errorf("streaming reader: %w", err)
	}
	defer r.Close()

	// 2. Load shared strings (small — fits in memory even for huge workbooks).
	sharedStrings, _ := loadSharedStrings(r)

	// 3. Resolve sheet path.
	sheetPath, err := resolveStreamSheetPath(r, sheetName)
	if err != nil {
		return fmt.Errorf("streaming reader: %w", err)
	}

	// 4. Open the sheet XML stream.
	rc, err := openZipEntry(r, sheetPath)
	if err != nil {
		return fmt.Errorf("streaming reader: cannot open %s: %w", sheetPath, err)
	}
	defer rc.Close()

	// 5. Token-based streaming parse.
	return streamSheetRows(rc, sharedStrings, callback)
}

// ---------------------------------------------------------------------------
// Token-based row streaming
// ---------------------------------------------------------------------------

// streamSheetRows reads the sheet XML token by token.  When it encounters a
// <row> element it decodes that single row (and its <c> children) into a
// map and calls the callback, then moves on to the next token.  The XML
// decoder's internal buffer is bounded by the widest row, not the file size.
func streamSheetRows(r io.Reader, ss map[int]string, callback RowCallback) error {
	decoder := xml.NewDecoder(r)

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return fmt.Errorf("streaming reader: XML error: %w", err)
		}
		if tok == nil {
			return nil
		}

		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "row" {
			continue
		}

		// Extract the 1-based row index from the r attribute.
		rowIdx := 0
		for _, attr := range start.Attr {
			if attr.Name.Local == "r" {
				rowIdx, _ = strconv.Atoi(attr.Value)
				break
			}
		}

		// Parse cells within this row.
		cells, err := parseStreamRow(decoder, start, ss)
		if err != nil {
			return fmt.Errorf("streaming reader: row %d: %w", rowIdx, err)
		}

		// Fire the callback.
		if err := callback(rowIdx, cells); err != nil {
			return err
		}
	}
}

// parseStreamRow decodes the <c> children of a single <row> element using
// the decoder's current position.  It stops when the matching </row> is seen.
func parseStreamRow(decoder *xml.Decoder, rowStart xml.StartElement, ss map[int]string) (map[string]string, error) {
	cells := make(map[string]string)

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.EndElement:
			if t.Name.Local == "row" {
				return cells, nil
			}

		case xml.StartElement:
			if t.Name.Local == "c" {
				ref, val := parseStreamCell(decoder, t, ss)
				if ref != "" {
					cells[ref] = val
				}
			}
		}
	}

	return cells, nil
}

// parseStreamCell decodes a single <c> element and its children, returning
// the A1 reference and the resolved string value.
func parseStreamCell(decoder *xml.Decoder, cellStart xml.StartElement, ss map[int]string) (ref, value string) {
	// Decode into a lightweight anonymous struct.
	var cell struct {
		Ref string `xml:"r,attr"`
		T   string `xml:"t,attr"`
		V   string `xml:"v"`
	}
	if err := decoder.DecodeElement(&cell, &cellStart); err != nil {
		return "", ""
	}

	if cell.Ref == "" {
		return "", ""
	}

	// Resolve the value.
	if cell.T == "s" && cell.V != "" {
		idx, err := strconv.Atoi(strings.TrimSpace(cell.V))
		if err == nil {
			if s, ok := ss[idx]; ok {
				return cell.Ref, s
			}
		}
		return cell.Ref, cell.V
	}

	if cell.T == "b" {
		if cell.V == "1" {
			return cell.Ref, "TRUE"
		}
		return cell.Ref, "FALSE"
	}

	return cell.Ref, cell.V
}

// ---------------------------------------------------------------------------
// Sheet path resolution
// ---------------------------------------------------------------------------

// resolveStreamSheetPath maps a sheet name (or empty string for the first
// sheet) to the ZIP entry path like "xl/worksheets/sheet1.xml".
func resolveStreamSheetPath(r *zip.ReadCloser, sheetName string) (string, error) {
	type relTarget struct {
		name   string
		target string
	}

	// Parse workbook.xml to get sheet → rId mapping.
	wbRaw, err := readZipFile(r, "xl/workbook.xml")
	if err != nil {
		return "", fmt.Errorf("cannot read workbook.xml: %w", err)
	}

	var wb struct {
		Sheets struct {
			Sheet []struct {
				Name string `xml:"name,attr"`
				RID  string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
			} `xml:"sheet"`
		} `xml:"sheets"`
	}
	if err := xml.Unmarshal(wbRaw, &wb); err != nil {
		return "", fmt.Errorf("cannot parse workbook.xml: %w", err)
	}

	if len(wb.Sheets.Sheet) == 0 {
		return "", fmt.Errorf("no sheets in workbook")
	}

	// Default to the first sheet.
	if sheetName == "" {
		sheetName = wb.Sheets.Sheet[0].Name
	}

	// Find the rId for the named sheet.
	var targetRID string
	for _, s := range wb.Sheets.Sheet {
		if strings.EqualFold(s.Name, sheetName) {
			targetRID = s.RID
			break
		}
	}
	if targetRID == "" {
		return "", fmt.Errorf("sheet %q not found", sheetName)
	}

	// Parse relationships to resolve rId → target.
	relsRaw, err := readZipFile(r, "xl/_rels/workbook.xml.rels")
	if err != nil {
		// Fall back to sequential naming.
		for i, s := range wb.Sheets.Sheet {
			if strings.EqualFold(s.Name, sheetName) {
				return fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1), nil
			}
		}
		return "", fmt.Errorf("cannot resolve sheet path for %q", sheetName)
	}

	var rels struct {
		Rels []struct {
			ID     string `xml:"Id,attr"`
			Target string `xml:"Target,attr"`
		} `xml:"Relationship"`
	}
	if err := xml.Unmarshal(relsRaw, &rels); err != nil {
		return "", fmt.Errorf("cannot parse relationships: %w", err)
	}

	for _, rel := range rels.Rels {
		if rel.ID == targetRID {
			return "xl/" + rel.Target, nil
		}
	}

	return "", fmt.Errorf("no relationship found for rId %q", targetRID)
}

// openZipEntry opens a named file from the ZIP archive for reading.
func openZipEntry(r *zip.ReadCloser, name string) (io.ReadCloser, error) {
	for _, f := range r.File {
		if f.Name == name {
			return f.Open()
		}
	}
	return nil, fmt.Errorf("entry %q not found", name)
}
