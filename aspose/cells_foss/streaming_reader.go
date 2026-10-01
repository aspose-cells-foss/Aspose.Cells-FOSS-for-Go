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

// StreamOptions configures filtering for streaming row processing.
// All fields are optional; zero values mean no filtering.
type StreamOptions struct {
	// StartRow is the 1-based row number to start processing from.
	// Rows before this are skipped. <=0 means from the beginning.
	StartRow int

	// EndRow is the 1-based row number to stop processing at (inclusive).
	// Rows after this are skipped. <=0 means to the end.
	EndRow int

	// Columns limits which columns are returned. Each entry is a column
	// letter (e.g., "A", "B", "AA"). nil or empty means all columns.
	Columns []string
}

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
	sharedStrings, _ := loadSharedStrings(&r.Reader)

	// 3. Resolve sheet path.
	sheetPath, err := resolveStreamSheetPath(&r.Reader, sheetName)
	if err != nil {
		return fmt.Errorf("streaming reader: %w", err)
	}

	// 4. Open the sheet XML stream.
	rc, err := openZipEntry(&r.Reader, sheetPath)
	if err != nil {
		return fmt.Errorf("streaming reader: cannot open %s: %w", sheetPath, err)
	}
	defer rc.Close()

	// 5. Token-based streaming parse.
	return streamSheetRows(rc, sharedStrings, callback, nil)
}

// ProcessRowsWithFilter opens the workbook and streams through the named sheet,
// applying the filters specified in opts. Only rows and columns matching the
// filter criteria are passed to the callback.
//
// When sheetName is empty the first sheet in the workbook is used.
// When opts is nil, all rows and columns are processed (equivalent to ProcessRows).
func (sr *StreamingReader) ProcessRowsWithFilter(sheetName string, opts *StreamOptions, callback RowCallback) error {
	if callback == nil {
		return fmt.Errorf("streaming reader: callback is nil")
	}

	r, err := zip.OpenReader(sr.path)
	if err != nil {
		return fmt.Errorf("streaming reader: %w", err)
	}
	defer r.Close()

	sharedStrings, _ := loadSharedStrings(&r.Reader)

	sheetPath, err := resolveStreamSheetPath(&r.Reader, sheetName)
	if err != nil {
		return fmt.Errorf("streaming reader: %w", err)
	}

	rc, err := openZipEntry(&r.Reader, sheetPath)
	if err != nil {
		return fmt.Errorf("streaming reader: cannot open %s: %w", sheetPath, err)
	}
	defer rc.Close()

	return streamSheetRows(rc, sharedStrings, callback, opts)
}

// ProcessRowsWithRange streams through rows in the specified range [startRow, endRow].
// Both bounds are 1-based and inclusive. startRow <= 0 means from the beginning;
// endRow <= 0 means to the end.
func (sr *StreamingReader) ProcessRowsWithRange(sheetName string, startRow, endRow int, callback RowCallback) error {
	return sr.ProcessRowsWithFilter(sheetName, &StreamOptions{
		StartRow: startRow,
		EndRow:   endRow,
	}, callback)
}

// ProcessRowsWithColumns streams through all rows but only returns cells from
// the specified columns. Each column is a letter string (e.g., "A", "B", "AA").
func (sr *StreamingReader) ProcessRowsWithColumns(sheetName string, columns []string, callback RowCallback) error {
	return sr.ProcessRowsWithFilter(sheetName, &StreamOptions{
		Columns: columns,
	}, callback)
}

// ---------------------------------------------------------------------------
// Token-based row streaming
// ---------------------------------------------------------------------------

// streamSheetRows reads the sheet XML token by token.  When it encounters a
// <row> element it decodes that single row (and its <c> children) into a
// map and calls the callback, then moves on to the next token.  The XML
// decoder's internal buffer is bounded by the widest row, not the file size.
// opts controls row range and column filtering; nil means no filtering.
func streamSheetRows(r io.Reader, ss []string, callback RowCallback, opts *StreamOptions) error {
	decoder := xml.NewDecoder(r)

	// Build column filter set if specified.
	colFilter := make(map[string]bool)
	if opts != nil && len(opts.Columns) > 0 {
		for _, c := range opts.Columns {
			colFilter[strings.ToUpper(c)] = true
		}
	}

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

		// Apply row range filter.
		if opts != nil {
			if opts.StartRow > 0 && rowIdx < opts.StartRow {
				// Skip this row by consuming it without processing.
				if err := skipRow(decoder); err != nil {
					return fmt.Errorf("streaming reader: row %d: %w", rowIdx, err)
				}
				continue
			}
			if opts.EndRow > 0 && rowIdx > opts.EndRow {
				// Past the end range; stop processing.
				return nil
			}
		}

		// Parse cells within this row.
		cells, err := parseStreamRow(decoder, start, ss, colFilter)
		if err != nil {
			return fmt.Errorf("streaming reader: row %d: %w", rowIdx, err)
		}

		// Fire the callback.
		if err := callback(rowIdx, cells); err != nil {
			return err
		}
	}
}

// skipRow consumes tokens until the matching </row> is found, discarding them.
func skipRow(decoder *xml.Decoder) error {
	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			return err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return nil
}

// parseStreamRow decodes the <c> children of a single <row> element using
// the decoder's current position.  It stops when the matching </row> is seen.
// colFilter limits which columns are included; empty map means all columns.
func parseStreamRow(decoder *xml.Decoder, rowStart xml.StartElement, ss []string, colFilter map[string]bool) (map[string]string, error) {
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
				ref, val := parseStreamCell(decoder, t, ss, colFilter)
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
// colFilter limits which columns are included; empty map means all columns.
func parseStreamCell(decoder *xml.Decoder, cellStart xml.StartElement, ss []string, colFilter map[string]bool) (ref, value string) {
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

	// Apply column filter.
	if len(colFilter) > 0 {
		col, _ := splitRef(cell.Ref)
		if !colFilter[strings.ToUpper(col)] {
			return "", ""
		}
	}

	// Resolve the value.
	if cell.T == "s" && cell.V != "" {
		idx, err := strconv.Atoi(strings.TrimSpace(cell.V))
		if err == nil && idx >= 0 && idx < len(ss) {
			return cell.Ref, ss[idx]
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
func resolveStreamSheetPath(zr *zip.Reader, sheetName string) (string, error) {
	sheetDefs, err := parseWorkbookXML(zr)
	if err != nil {
		return "", fmt.Errorf("cannot read workbook.xml: %w", err)
	}
	if len(sheetDefs) == 0 {
		return "", fmt.Errorf("no sheets in workbook")
	}

	// Default to the first sheet.
	if sheetName == "" {
		sheetName = sheetDefs[0].Name
	}

	// Find the rId for the named sheet.
	var targetRID string
	for _, def := range sheetDefs {
		if strings.EqualFold(def.Name, sheetName) {
			targetRID = def.RID
			break
		}
	}
	if targetRID == "" {
		return "", fmt.Errorf("sheet %q not found", sheetName)
	}

	rels, err := parseWorkbookRels(zr)
	if err != nil {
		// No relationships part — fall back to sequential part naming.
		for i, def := range sheetDefs {
			if strings.EqualFold(def.Name, sheetName) {
				return fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1), nil
			}
		}
		return "", fmt.Errorf("cannot resolve sheet path for %q", sheetName)
	}

	if target, ok := rels[targetRID]; ok {
		return "xl/" + target, nil
	}

	return "", fmt.Errorf("no relationship found for rId %q", targetRID)
}

// openZipEntry opens a named file from the ZIP archive for reading.
func openZipEntry(zr *zip.Reader, name string) (io.ReadCloser, error) {
	for _, f := range zr.File {
		if f.Name == name {
			return f.Open()
		}
	}
	return nil, fmt.Errorf("entry %q not found", name)
}
