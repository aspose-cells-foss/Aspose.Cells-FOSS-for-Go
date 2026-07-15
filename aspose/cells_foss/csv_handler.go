package cells_foss

import (
	"encoding/csv"
	"fmt"
	"os"
	"strconv"
)

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

	r := csv.NewReader(f)
	r.Comma = delimiter
	r.LazyQuotes = true
	r.TrimLeadingSpace = true

	data, err := r.ReadAll()
	if err != nil {
		return fmt.Errorf("csv import: %w", err)
	}

	// Create a new worksheet.
	idx := len(wb.Worksheets)
	ws := &Worksheet{
		Name:  sheetName,
		Index: idx,
		cells: &Cells{},
	}
	ws.cells.setParent(wb)

	if err := ws.FromCSV(data, delimiter); err != nil {
		return fmt.Errorf("csv import: %w", err)
	}

	wb.Worksheets = append(wb.Worksheets, ws)
	wb.Modified = true
	return nil
}

// ToCSV converts the worksheet's cells into a 2D string slice suitable for
// writing with encoding/csv.  Rows and columns are ordered; empty cells
// produce empty strings.
func (ws *Worksheet) ToCSV(delimiter rune) ([][]string, error) {
	all := ws.Cells().All()
	if len(all) == 0 {
		return [][]string{}, nil
	}

	// Determine the bounding box.
	maxRow, maxCol := 0, 0
	type cellPos struct{ row, col int }
	positions := make(map[string]cellPos, len(all))

	for ref := range all {
		colStr, row := splitRef(ref)
		col := colToNum(colStr)
		positions[ref] = cellPos{row, col}
		if row > maxRow {
			maxRow = row
		}
		if col > maxCol {
			maxCol = col
		}
	}

	// Build the grid (1-based rows to 0-based slice).
	grid := make([][]string, maxRow+1)
	for r := 0; r <= maxRow; r++ {
		grid[r] = make([]string, maxCol+1)
	}

	for ref, cell := range all {
		pos := positions[ref]
		grid[pos.row][pos.col] = CellToString(cell.Value)
	}

	// Trim leading empty rows (common when data starts at row 2 or later).
	startRow := 0
	for startRow <= maxRow {
		empty := true
		for c := 0; c <= maxCol; c++ {
			if grid[startRow][c] != "" {
				empty = false
				break
			}
		}
		if !empty {
			break
		}
		startRow++
	}

	// Trim trailing empty rows.
	endRow := maxRow
	for endRow >= startRow {
		empty := true
		for c := 0; c <= maxCol; c++ {
			if grid[endRow][c] != "" {
				empty = false
				break
			}
		}
		if !empty {
			break
		}
		endRow--
	}

	if startRow > endRow {
		return [][]string{}, nil
	}

	result := make([][]string, endRow-startRow+1)
	for r := startRow; r <= endRow; r++ {
		result[r-startRow] = grid[r]
	}
	return result, nil
}

// FromCSV populates the worksheet with data from a 2D string slice.  Row 0
// of the slice maps to worksheet row 1, column 0 maps to column A.
func (ws *Worksheet) FromCSV(data [][]string, delimiter rune) error {
	_ = delimiter // reserved for future use
	for r, row := range data {
		for c, val := range row {
			ref := numToCol(c) + strconv.Itoa(r+1)
			if err := ws.Cells().Set(ref, val); err != nil {
				return fmt.Errorf("csv import: row %d, col %d: %w", r, c, err)
			}
		}
	}
	return nil
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
