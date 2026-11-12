package cells_foss

import (
	"fmt"
	"strconv"
	"strings"
)

// InsertRow inserts an empty row at the given 1-based row number, shifting
// that row and everything below it down by one. The inserted row starts out
// blank.
//
// Cells (values, formulas, styles), merged ranges, row heights, hyperlinks,
// the autofilter range, data-validation ranges, conditional-formatting
// ranges, table and pivot-table ranges, chart data ranges, and picture and
// chart anchors all move with the data. Ranges that span the insertion point
// expand to cover the new row, matching Excel's behavior.
//
// Formula text is not rewritten: a formula such as "=SUM(A1:A5)" keeps its
// original references even when the cells they point at move. Inserting fails
// with an error when sheet content already reaches the last row (1048576),
// since the shift would push cells off the sheet area.
func (ws *Worksheet) InsertRow(row int) error {
	if row < 1 || row > maxRowCount {
		return fmt.Errorf("cells_foss: row %d is outside the sheet area (1..%d)", row, maxRowCount)
	}
	if m := ws.maxContentRow(); m == maxRowCount && m >= row {
		return fmt.Errorf("cells_foss: cannot insert row %d: sheet content already extends to the last row (%d)", row, maxRowCount)
	}
	ws.shiftRows(row, 1)
	return nil
}

// DeleteRow removes the row at the given 1-based row number, shifting
// everything below it up by one. Cells on the deleted row are dropped, and
// hyperlinks anchored to them are removed.
//
// The same set of features InsertRow shifts is shifted here. Ranges spanning
// the deleted row shrink by one row; ranges lying entirely within it
// (including single-row merged cells) are dropped.
//
// Formula text is not rewritten: formulas keep their original references even
// when the cells they point at move or disappear.
func (ws *Worksheet) DeleteRow(row int) error {
	if row < 1 || row > maxRowCount {
		return fmt.Errorf("cells_foss: row %d is outside the sheet area (1..%d)", row, maxRowCount)
	}
	ws.shiftRows(row, -1)
	return nil
}

// shiftRows moves the rows at and below the 1-based row number by delta
// (+1 for an insert, -1 for a delete), carrying every row-aware feature of
// the sheet along.
func (ws *Worksheet) shiftRows(row, delta int) {
	// Cells: rekey the collection and update each moved cell's Ref. Cells on
	// a deleted row are dropped.
	if c := ws.cells; c != nil && len(c.cells) > 0 {
		next := make(map[string]*Cell, len(c.cells))
		for ref, cell := range c.cells {
			newRef, keep := shiftCellRef(ref, row, delta)
			if !keep {
				continue
			}
			cell.Ref = newRef
			next[newRef] = cell
		}
		c.cells = next
	}

	// Row heights follow their row; a deleted row's height is dropped.
	if len(ws.RowHeights) > 0 {
		heights := make(map[int]float64, len(ws.RowHeights))
		for r, h := range ws.RowHeights {
			switch {
			case delta < 0 && r == row:
			case r >= row:
				heights[r+delta] = h
			default:
				heights[r] = h
			}
		}
		ws.RowHeights = heights
	}

	// Merged ranges expand (insert) or shrink/vanish (delete).
	ws.MergedCells = shiftRefList(ws.MergedCells, row, delta)

	// Hyperlinks move with their cell; links on a deleted row are dropped.
	if len(ws.Hyperlinks) > 0 {
		links := make([]*Hyperlink, 0, len(ws.Hyperlinks))
		for _, hl := range ws.Hyperlinks {
			newRef := shiftRangeRef(hl.Ref, row, delta)
			if newRef == "" {
				continue
			}
			hl.Ref = newRef
			links = append(links, hl)
		}
		ws.Hyperlinks = links
	}

	// Sheet-level and per-feature ranges: a range spanning the affected row
	// expands on insert and shrinks on delete. Validations, conditional
	// formats and tables are dropped when their whole range vanishes
	// (possible on deletes only).
	ws.AutoFilterRange = shiftRangeRef(ws.AutoFilterRange, row, delta)
	if len(ws.DataValidations) > 0 {
		dvs := make([]*DataValidation, 0, len(ws.DataValidations))
		for _, dv := range ws.DataValidations {
			dv.Ref = shiftRangeRef(dv.Ref, row, delta)
			if dv.Ref != "" {
				dvs = append(dvs, dv)
			}
		}
		ws.DataValidations = dvs
	}
	if len(ws.ConditionalFormattings) > 0 {
		cfs := make([]*ConditionalFormatting, 0, len(ws.ConditionalFormattings))
		for _, cf := range ws.ConditionalFormattings {
			cf.Ref = shiftRangeRef(cf.Ref, row, delta)
			if cf.Ref != "" {
				cfs = append(cfs, cf)
			}
		}
		ws.ConditionalFormattings = cfs
	}
	if len(ws.Tables) > 0 {
		tables := make([]*Table, 0, len(ws.Tables))
		for _, tb := range ws.Tables {
			tb.Range = shiftRangeRef(tb.Range, row, delta)
			if tb.Range != "" {
				tables = append(tables, tb)
			}
		}
		ws.Tables = tables
	}
	for _, pt := range ws.PivotTables {
		pt.Ref = shiftRangeRef(pt.Ref, row, delta)
		pt.SourceRef = shiftRangeRef(pt.SourceRef, row, delta)
	}

	// Pictures and charts anchor at 0-based rows; chart data ranges move
	// with the cells they point at.
	for _, pic := range ws.Pictures {
		pic.Row = shiftAnchorRow(pic.Row, row, delta)
	}
	for _, ch := range ws.Charts {
		ch.Anchor.Row = shiftAnchorRow(ch.Anchor.Row, row, delta)
		for _, s := range ch.Series {
			s.Categories = shiftRangeRef(s.Categories, row, delta)
			s.Values = shiftRangeRef(s.Values, row, delta)
		}
	}

	ws.Modified = true
	if ws.cells != nil && ws.cells.wb != nil {
		ws.cells.wb.Modified = true
	}
}

// maxContentRow returns the largest 1-based row number referenced by the
// sheet's cells, row heights, or any range-bearing feature. It is used to
// detect inserts that would push content off the sheet area.
func (ws *Worksheet) maxContentRow() int {
	m := 0
	bump := func(r int) {
		if r > m {
			m = r
		}
	}
	if c := ws.cells; c != nil {
		for ref := range c.cells {
			if _, r, ok := parseA1CellRef(ref); ok {
				bump(r)
			}
		}
	}
	for r := range ws.RowHeights {
		bump(r)
	}
	refs := []string{ws.AutoFilterRange}
	refs = append(refs, ws.MergedCells...)
	for _, dv := range ws.DataValidations {
		refs = append(refs, dv.Ref)
	}
	for _, cf := range ws.ConditionalFormattings {
		refs = append(refs, cf.Ref)
	}
	for _, tb := range ws.Tables {
		refs = append(refs, tb.Range)
	}
	for _, pt := range ws.PivotTables {
		refs = append(refs, pt.Ref, pt.SourceRef)
	}
	for _, ch := range ws.Charts {
		for _, s := range ch.Series {
			refs = append(refs, s.Categories, s.Values)
		}
	}
	for _, ref := range refs {
		for _, part := range strings.Fields(ref) {
			_, rest := splitSheetPrefix(part)
			for _, bound := range strings.Split(rest, ":") {
				if _, r, ok := parseA1CellRef(bound); ok {
					bump(r)
				}
			}
		}
	}
	return m
}

// parseA1CellRef splits an A1-style cell reference ("B12", "$B$12") into its
// column letters and 1-based row number with dollar signs stripped. It
// returns ok=false for anything that is not a plain cell reference.
func parseA1CellRef(ref string) (col string, row int, ok bool) {
	s := strings.ReplaceAll(ref, "$", "")
	i := 0
	for i < len(s) && isASCIILetter(s[i]) {
		i++
	}
	if i == 0 || i == len(s) {
		return "", 0, false
	}
	for _, d := range s[i:] {
		if d < '0' || d > '9' {
			return "", 0, false
		}
	}
	n, err := strconv.Atoi(s[i:])
	if err != nil || n < 1 {
		return "", 0, false
	}
	return s[:i], n, true
}

// splitSheetPrefix separates an optional sheet qualifier ("Sheet1!",
// "'My Sheet'!") from the cell or range reference that follows it.
func splitSheetPrefix(ref string) (prefix, rest string) {
	if strings.HasPrefix(ref, "'") {
		if i := strings.IndexByte(ref[1:], '\''); i >= 0 {
			q := i + 1 // index of the closing quote
			if q+1 < len(ref) && ref[q+1] == '!' {
				return ref[:q+2], ref[q+2:]
			}
		}
		return "", ref
	}
	if i := strings.IndexByte(ref, '!'); i >= 0 {
		return ref[:i+1], ref[i+1:]
	}
	return "", ref
}

// shiftCellRef moves a single-cell reference for an insert or delete at the
// 1-based row number. It reports keep=false when the cell lies on the deleted
// row. References that cannot be parsed are returned unchanged.
func shiftCellRef(ref string, row, delta int) (string, bool) {
	col, r, ok := parseA1CellRef(ref)
	if !ok {
		return ref, true
	}
	switch {
	case delta < 0 && r == row:
		return "", false
	case r >= row:
		return col + strconv.Itoa(r+delta), true
	default:
		return ref, true
	}
}

// shiftBound moves one endpoint of a range. On a delete, a start endpoint on
// the deleted row stays put (the row below slides into its place) while an
// end endpoint on the deleted row shrinks by one. ok=false means the endpoint
// moved off the top of the sheet.
func shiftBound(r, row, delta int, isStart bool) (int, bool) {
	switch {
	case delta < 0 && r == row:
		if isStart {
			return row, true
		}
		return row - 1, row-1 >= 1
	case r >= row:
		n := r + delta
		return n, n >= 1
	default:
		return r, true
	}
}

// shiftOneRange shifts a single cell or range reference, keeping any sheet
// qualifier. It reports keep=false when the reference vanishes: a cell on
// the deleted row, or a range lying entirely within it. Unparseable parts
// are returned unchanged.
func shiftOneRange(ref string, row, delta int) (string, bool) {
	prefix, rest := splitSheetPrefix(ref)
	if rest == "" {
		return ref, true
	}
	bounds := strings.SplitN(rest, ":", 2)
	if len(bounds) == 1 {
		shifted, keep := shiftCellRef(rest, row, delta)
		if !keep {
			return "", false
		}
		return prefix + shifted, true
	}
	c1, r1, ok1 := parseA1CellRef(bounds[0])
	c2, r2, ok2 := parseA1CellRef(bounds[1])
	if !ok1 || !ok2 {
		return ref, true
	}
	n1, ok := shiftBound(r1, row, delta, true)
	if !ok {
		return "", false
	}
	n2, ok := shiftBound(r2, row, delta, false)
	if !ok || n2 < n1 {
		return "", false
	}
	if n1 == r1 && n2 == r2 {
		return ref, true
	}
	return prefix + c1 + strconv.Itoa(n1) + ":" + c2 + strconv.Itoa(n2), true
}

// shiftRangeRef shifts the row numbers of every reference in ref, which may
// be a single cell, a range, or a space-separated list of ranges (an sqref).
// Parts that vanish on a delete are dropped; unparseable parts are left
// alone. The original string is returned untouched when nothing changes.
func shiftRangeRef(ref string, row, delta int) string {
	if strings.TrimSpace(ref) == "" {
		return ref
	}
	parts := strings.Fields(ref)
	out := make([]string, 0, len(parts))
	changed := false
	for _, p := range parts {
		shifted, keep := shiftOneRange(p, row, delta)
		if !keep {
			changed = true
			continue
		}
		if shifted != p {
			changed = true
		}
		out = append(out, shifted)
	}
	if !changed {
		return ref
	}
	return strings.Join(out, " ")
}

// shiftRefList shifts every range in a list, dropping those that vanish.
func shiftRefList(refs []string, row, delta int) []string {
	if len(refs) == 0 {
		return refs
	}
	out := make([]string, 0, len(refs))
	for _, r := range refs {
		if s := shiftRangeRef(r, row, delta); s != "" {
			out = append(out, s)
		}
	}
	return out
}

// shiftAnchorRow moves a 0-based anchor row (pictures, charts) for an insert
// or delete at the 1-based row number. An anchor on the deleted row slides up
// but never past the top of the sheet.
func shiftAnchorRow(anchorRow0, row, delta int) int {
	r := anchorRow0 + 1
	if r < row {
		return anchorRow0
	}
	n := r + delta
	if n < 1 {
		n = 1
	}
	return n - 1
}
