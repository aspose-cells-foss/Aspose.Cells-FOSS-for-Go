package cells_foss

import (
	"fmt"
	"strconv"
	"strings"
)

// Worksheet layout operations: merged cells, column widths, row heights,
// freeze panes, and autoFilter. Hyperlinks live in hyperlink.go.
//
// Layout properties round-trip through save/load: they are parsed back
// into the Worksheet on load so regenerated sheet XML (for modified
// sheets) keeps them. Unmodified sheets reuse their cached source XML.

// Sheet-area and Excel UI limits used for validation.
const (
	maxRowCount    = 1048576
	maxColCount    = 16384
	maxColumnWidth = 255   // in Excel character units
	maxRowHeight   = 409.5 // in points
)

// AddMergedCells merges the cells of an A1 range ("A1:B2"). The range is
// normalized (upper-case, corners ordered) before it is stored. Excel
// rejects files whose merged ranges overlap, so overlapping and duplicate
// ranges are refused here as well.
func (ws *Worksheet) AddMergedCells(ref string) error {
	if !strings.Contains(ref, ":") {
		return fmt.Errorf("cells_foss: merged range %q must cover at least two cells (e.g. \"A1:B2\")", ref)
	}
	c1, r1, c2, r2, err := parseA1Range(ref)
	if err != nil {
		return err
	}
	if c1 == c2 && r1 == r2 {
		return fmt.Errorf("cells_foss: merged range %q must cover at least two cells", ref)
	}
	norm := rangeRef(c1, r1, c2, r2)
	for _, existing := range ws.MergedCells {
		e1, f1, e2, f2, err := parseA1Range(existing)
		if err != nil {
			continue // a loaded range this library cannot parse is kept verbatim
		}
		if e1 == c1 && f1 == r1 && e2 == c2 && f2 == r2 {
			return fmt.Errorf("cells_foss: range %s is already merged", norm)
		}
		if c1 <= e2 && e1 <= c2 && r1 <= f2 && f1 <= r2 {
			return fmt.Errorf("cells_foss: range %s overlaps merged range %s", norm, rangeRef(e1, f1, e2, f2))
		}
	}
	ws.MergedCells = append(ws.MergedCells, norm)
	ws.Modified = true
	return nil
}

// RemoveMergedCells un-merges a previously merged range. The reference is
// matched after normalization; an error is returned when the range is not
// merged.
func (ws *Worksheet) RemoveMergedCells(ref string) error {
	c1, r1, c2, r2, err := parseA1Range(ref)
	if err != nil {
		return err
	}
	norm := rangeRef(c1, r1, c2, r2)
	for i, existing := range ws.MergedCells {
		if existing == norm {
			ws.MergedCells = append(ws.MergedCells[:i], ws.MergedCells[i+1:]...)
			ws.Modified = true
			return nil
		}
	}
	return fmt.Errorf("cells_foss: range %s is not merged", norm)
}

// SetColumnWidth sets the width of a column, identified by its letter
// ("A", "BC"), in Excel character units (the default width is 8.43).
// Widths must be in (0, 255].
func (ws *Worksheet) SetColumnWidth(col string, width float64) error {
	col = strings.ToUpper(strings.TrimSpace(col))
	if col == "" || len(col) > 3 || !isColumnLetters(col) {
		return fmt.Errorf("cells_foss: %q is not a valid column reference (e.g. \"A\" or \"BC\")", col)
	}
	if n := colToNum(col) + 1; n < 1 || n > maxColCount {
		return fmt.Errorf("cells_foss: column %q is outside the sheet area (A..XFD)", col)
	}
	if !(width > 0) || width > maxColumnWidth {
		return fmt.Errorf("cells_foss: column width %v must be in (0, %v]", width, maxColumnWidth)
	}
	if ws.ColumnWidths == nil {
		ws.ColumnWidths = make(map[string]float64)
	}
	ws.ColumnWidths[col] = width
	ws.Modified = true
	return nil
}

// SetRowHeight sets the height of a 1-based row in points (the default
// height is 15). Heights must be in (0, 409.5].
func (ws *Worksheet) SetRowHeight(row int, height float64) error {
	if row < 1 || row > maxRowCount {
		return fmt.Errorf("cells_foss: row %d is outside the sheet area (1..%d)", row, maxRowCount)
	}
	if !(height > 0) || height > maxRowHeight {
		return fmt.Errorf("cells_foss: row height %v must be in (0, %v]", height, maxRowHeight)
	}
	if ws.RowHeights == nil {
		ws.RowHeights = make(map[int]float64)
	}
	ws.RowHeights[row] = height
	ws.Modified = true
	return nil
}

// FreezePanes freezes the top row rows and the leftmost col columns of the
// sheet, keeping them visible while scrolling. FreezePanes(0, 0) removes
// the freeze. For example FreezePanes(1, 0) freezes the header row.
func (ws *Worksheet) FreezePanes(row, col int) error {
	if row < 0 || row > maxRowCount-1 {
		return fmt.Errorf("cells_foss: frozen row count %d must be in [0, %d]", row, maxRowCount-1)
	}
	if col < 0 || col > maxColCount-1 {
		return fmt.Errorf("cells_foss: frozen column count %d must be in [0, %d]", col, maxColCount-1)
	}
	ws.FrozenRow = row
	ws.FrozenCol = col
	ws.Modified = true
	return nil
}

// SetAutoFilter enables the drop-down filter buttons over an A1 range
// ("A1:C10"); the range is normalized before it is stored. Pass "" to
// remove the filter.
func (ws *Worksheet) SetAutoFilter(ref string) error {
	if strings.TrimSpace(ref) == "" {
		ws.AutoFilterRange = ""
		ws.Modified = true
		return nil
	}
	c1, r1, c2, r2, err := parseA1Range(ref)
	if err != nil {
		return err
	}
	ws.AutoFilterRange = rangeRef(c1, r1, c2, r2)
	ws.Modified = true
	return nil
}

// ---------------------------------------------------------------------------
// Reference helpers
// ---------------------------------------------------------------------------

// parseA1Range parses "A1" or "A1:B2" into ordered column/row bounds
// (1-based, corners swapped when reversed). Column and row must lie inside
// the sheet area.
func parseA1Range(ref string) (col1, row1, col2, row2 int, err error) {
	parts := strings.Split(strings.TrimSpace(ref), ":")
	if len(parts) > 2 {
		return 0, 0, 0, 0, fmt.Errorf("cells_foss: %q is not a valid A1 range", ref)
	}
	cols := make([]int, 0, 2)
	rows := make([]int, 0, 2)
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if !isCellRefPattern(p) {
			return 0, 0, 0, 0, fmt.Errorf("cells_foss: %q is not a valid A1 cell reference (in %q)", p, ref)
		}
		colName, row := splitRef(p)
		col := colToNum(colName) + 1 // colToNum is 0-based; bounds here are 1-based
		if col < 1 || col > maxColCount || row < 1 || row > maxRowCount {
			return 0, 0, 0, 0, fmt.Errorf("cells_foss: cell %q is outside the sheet area (A1..XFD%d)", p, maxRowCount)
		}
		cols = append(cols, col)
		rows = append(rows, row)
	}
	col1, col2 = cols[0], cols[len(cols)-1]
	row1, row2 = rows[0], rows[len(rows)-1]
	if col1 > col2 {
		col1, col2 = col2, col1
	}
	if row1 > row2 {
		row1, row2 = row2, row1
	}
	return col1, row1, col2, row2, nil
}

// rangeRef renders ordered 1-based bounds back into normalized A1 form.
func rangeRef(col1, row1, col2, row2 int) string {
	return numToCol(col1-1) + strconv.Itoa(row1) + ":" + numToCol(col2-1) + strconv.Itoa(row2)
}

// isColumnLetters reports whether s is 1-3 ASCII letters.
func isColumnLetters(s string) bool {
	for i := 0; i < len(s); i++ {
		if (s[i] < 'A' || s[i] > 'Z') && (s[i] < 'a' || s[i] > 'z') {
			return false
		}
	}
	return len(s) > 0
}
