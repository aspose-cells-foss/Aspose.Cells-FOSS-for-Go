package cells_foss

import (
	"bytes"
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
)

// utf8BOM is the UTF-8 byte-order mark. Written by ExportToCSV so Excel
// detects UTF-8 encoding; stripped by ImportFromCSV when present.
var utf8BOM = []byte{0xEF, 0xBB, 0xBF}

// ExportToCSV writes the worksheet at sheetIndex to a CSV file using the
// given delimiter (e.g. ',', '\t', ';').  Cells are ordered left→right,
// top→bottom; empty cells are written as empty fields.
func (wb *Workbook) ExportToCSV(sheetIndex int, path string, delimiter rune) error {
	if wb == nil {
		return fmt.Errorf("csv export: workbook is nil")
	}
	if sheetIndex < 0 || sheetIndex >= len(wb.Worksheets) {
		return fmt.Errorf("csv export: sheet index %d out of range [0, %d)", sheetIndex, len(wb.Worksheets))
	}

	data, err := wb.Worksheets[sheetIndex].ToCSV(delimiter)
	if err != nil {
		return fmt.Errorf("csv export: %w", err)
	}

	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("csv export: %w", err)
	}
	defer f.Close()

	// A UTF-8 BOM lets Excel detect UTF-8 so non-ASCII text is displayed
	// correctly when the file is opened directly.
	if _, err := f.Write(utf8BOM); err != nil {
		return fmt.Errorf("csv export: %w", err)
	}

	w := csv.NewWriter(f)
	w.Comma = delimiter
	w.UseCRLF = true // Windows-style line endings for better Excel compatibility.

	if err := w.WriteAll(data); err != nil {
		return fmt.Errorf("csv export: %w", err)
	}
	w.Flush()
	return w.Error()
}

// ImportFromCSV reads a CSV file and imports its contents into a new
// worksheet with the given name.  The worksheet is appended to the workbook.
func (wb *Workbook) ImportFromCSV(path string, sheetName string, delimiter rune) error {
	if wb == nil {
		return fmt.Errorf("csv import: workbook is nil")
	}

	f, err := os.Open(path)
	if err != nil {
		return fmt.Errorf("csv import: %w", err)
	}
	defer f.Close()

	// Skip a UTF-8 BOM if present; otherwise it would end up in the first
	// field of the first record.
	var head [3]byte
	if n, err := io.ReadFull(f, head[:]); err == nil && n == 3 && bytes.Equal(head[:], utf8BOM) {
		// Leave the file positioned just past the BOM.
	} else if _, err := f.Seek(0, io.SeekStart); err != nil {
		return fmt.Errorf("csv import: %w", err)
	}

	r := csv.NewReader(f)
	r.Comma = delimiter
	r.LazyQuotes = true
	r.TrimLeadingSpace = true

	data, err := r.ReadAll()
	if err != nil {
		return fmt.Errorf("csv import: %w", err)
	}

	// Create the worksheet through AddWorksheet so the name is validated
	// (Excel rejects duplicate or malformed sheet names) and the cell
	// back-references are wired.
	ws, err := wb.AddWorksheet(sheetName)
	if err != nil {
		return fmt.Errorf("csv import: %w", err)
	}

	if err := ws.FromCSV(data, delimiter); err != nil {
		// Roll back the appended sheet so a failed import leaves the
		// workbook unchanged.
		wb.Worksheets = wb.Worksheets[:len(wb.Worksheets)-1]
		return fmt.Errorf("csv import: %w", err)
	}

	return nil
}

// maxCSVCells bounds the grid ToCSV is willing to allocate. A sheet with two
// far-apart cells (e.g. A1 and XFD1048576) would otherwise span the full
// 17-billion-cell sheet area and exhaust memory.
const maxCSVCells = 16 * 1024 * 1024

// ToCSV converts the worksheet's cells into a 2D string slice suitable for
// writing with encoding/csv.  Rows and columns are ordered; empty cells
// produce empty strings.  Leading and trailing empty rows and columns are
// trimmed to the bounding box of the populated cells.
func (ws *Worksheet) ToCSV(delimiter rune) ([][]string, error) {
	all := ws.Cells().All()
	if len(all) == 0 {
		return [][]string{}, nil
	}

	// Resolve every cell's grid position and the bounding box in one pass.
	type placedCell struct {
		row, col int
		value    string
	}
	cells := make([]placedCell, 0, len(all))
	minRow, maxRow := -1, 0
	minCol, maxCol := -1, 0
	for ref, cell := range all {
		colName, row := splitRef(ref)
		col := colToNum(colName)
		cells = append(cells, placedCell{row: row, col: col, value: CellToString(cell.Value)})
		if row < minRow || minRow < 0 {
			minRow = row
		}
		if row > maxRow {
			maxRow = row
		}
		if col < minCol || minCol < 0 {
			minCol = col
		}
		if col > maxCol {
			maxCol = col
		}
	}

	numRows := maxRow - minRow + 1
	numCols := maxCol - minCol + 1
	if int64(numRows)*int64(numCols) > maxCSVCells {
		return nil, fmt.Errorf("csv export: populated cells span %dx%d = %d grid cells, exceeding the %d limit (sheet is too sparse for CSV export)",
			numRows, numCols, int64(numRows)*int64(numCols), maxCSVCells)
	}

	grid := make([][]string, numRows)
	for r := range grid {
		grid[r] = make([]string, numCols)
	}
	for _, pc := range cells {
		grid[pc.row-minRow][pc.col-minCol] = pc.value
	}
	return grid, nil
}

// FromCSV populates the worksheet with data from a 2D string slice.  Row 0
// of the slice maps to worksheet row 1, column 0 maps to column A.  Values
// that are exactly representable as numbers or booleans are stored as
// float64 or bool (see inferCSVValue); everything else stays a string.
func (ws *Worksheet) FromCSV(data [][]string, delimiter rune) error {
	_ = delimiter // reserved for future use
	for r, row := range data {
		for c, val := range row {
			ref := numToCol(c) + strconv.Itoa(r+1)
			if err := ws.Cells().Set(ref, inferCSVValue(val)); err != nil {
				return fmt.Errorf("csv import: row %d, col %d: %w", r, c, err)
			}
		}
	}
	return nil
}

// inferCSVValue converts CSV text to a typed value the way Excel does when
// opening a CSV file. Conversion is conservative: integers and floats are
// only converted when their canonical text form matches the trimmed input
// exactly, so leading zeros ("007"), trailing-zero formats ("1.50"), and
// exponent spellings ("1e5") stay strings.
func inferCSVValue(s string) interface{} {
	t := strings.TrimSpace(s)
	if t == "" {
		return s
	}
	switch strings.ToUpper(t) {
	case "TRUE":
		return true
	case "FALSE":
		return false
	}
	if n, err := strconv.ParseInt(t, 10, 64); err == nil && strconv.FormatInt(n, 10) == t {
		return float64(n)
	}
	if f, err := strconv.ParseFloat(t, 64); err == nil && strconv.FormatFloat(f, 'g', -1, 64) == t {
		return f
	}
	return s
}

// CellToString converts a cell value to its CSV string representation.
func CellToString(v interface{}) string {
	if v == nil {
		return ""
	}
	switch val := v.(type) {
	case string:
		return val
	case bool:
		if val {
			return "TRUE"
		}
		return "FALSE"
	case float64:
		return strconv.FormatFloat(val, 'g', -1, 64)
	case float32:
		return strconv.FormatFloat(float64(val), 'g', -1, 32)
	case int:
		return strconv.Itoa(val)
	case int64:
		return strconv.FormatInt(val, 10)
	default:
		return fmt.Sprint(val)
	}
}
