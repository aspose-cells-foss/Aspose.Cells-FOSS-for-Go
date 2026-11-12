package cells_foss

import (
	"fmt"
	"strings"
	"unicode/utf8"
)

// Worksheet represents a single sheet within a workbook. It holds the sheet
// metadata (name and positional index) together with the collection of cells
// and data validations that make up the sheet content.
type Worksheet struct {
	Name  string
	Index int

	// Modified is set to true when this sheet's content changes (cells,
	// data validations, tables, pictures). The saver uses this flag to
	// decide whether to regenerate this sheet's XML or reuse its cached
	// sourceXML.
	Modified bool

	// cells holds the cell collection for this worksheet. Use the Cells()
	// accessor to obtain the *Cells value.
	cells *Cells

	// DataValidations holds the data-validation rules applied to this sheet.
	DataValidations []*DataValidation

	// Tables holds the structured tables defined on this worksheet.
	Tables []*Table

	// Pictures holds the images embedded on this worksheet.
	Pictures []*Picture

	// ConditionalFormattings holds the conditional formatting rules applied
	// to this worksheet.
	ConditionalFormattings []*ConditionalFormatting

	// Charts holds the charts embedded on this worksheet.
	Charts []*Chart

	// PivotTables holds the pivot tables defined on this worksheet.
	PivotTables []*PivotTable

	// MergedCells holds merged cell ranges in normalized A1 form
	// ("A1:B2"). Use AddMergedCells / RemoveMergedCells to maintain it.
	MergedCells []string

	// ColumnWidths maps column letters ("A") to widths in Excel character
	// units. SetColumnWidth maintains it.
	ColumnWidths map[string]float64

	// RowHeights maps 1-based row numbers to heights in points.
	// SetRowHeight maintains it.
	RowHeights map[int]float64

	// FrozenRow and FrozenCol are the number of rows frozen at the top and
	// columns frozen at the left of the sheet (see FreezePanes). Both zero
	// means no freeze.
	FrozenRow int
	FrozenCol int

	// AutoFilterRange is the range carrying the sheet's autofilter buttons
	// ("A1:C10"); empty means none. See SetAutoFilter.
	AutoFilterRange string

	// Hyperlinks holds the links attached to cells of this worksheet.
	// Use AddHyperlink / AddInternalHyperlink / RemoveHyperlink to maintain it.
	Hyperlinks []*Hyperlink

	// sourceXML caches the raw XML bytes of this worksheet as read from the
	// .xlsx file. When the sheet has not been modified this payload is
	// written back verbatim to preserve round-trip fidelity.
	sourceXML []byte

	// cachedParts holds the raw bytes of every part that accompanied this
	// sheet in the package it was last loaded from or saved to: the sheet's
	// .rels file plus all parts reachable through it (drawings, media,
	// charts, tables, pivot tables and caches, comments, printer settings,
	// …). Unmodified sheets re-emit these verbatim so relationships and
	// rIds in the cached sheet XML stay valid.
	cachedParts map[string][]byte
}

// Cells returns the Cells collection for this worksheet, enabling cell-level
// read and write operations via A1-style references.
func (ws *Worksheet) Cells() *Cells {
	return ws.cells
}

// AddWorksheet appends a new worksheet with the given name to the workbook
// and returns it. The returned sheet is fully wired for cell access and
// marked modified, so Save writes it as a new sheet part.
//
// The name must satisfy Excel's worksheet naming rules: 1–31 characters,
// none of the characters []:*?/\, no leading or trailing apostrophe, not
// the reserved name "History", and unique within the workbook
// (case-insensitive). An error is returned for invalid or duplicate names.
func (wb *Workbook) AddWorksheet(name string) (*Worksheet, error) {
	if wb == nil {
		return nil, fmt.Errorf("cells_foss: workbook is nil")
	}
	if err := wb.validateSheetName(name); err != nil {
		return nil, err
	}
	ws := &Worksheet{
		Name:     name,
		Index:    len(wb.Worksheets),
		Modified: true,
		cells:    &Cells{},
	}
	ws.cells.setParent(wb)
	ws.cells.setWorksheet(ws)
	wb.Worksheets = append(wb.Worksheets, ws)
	wb.Modified = true
	return ws, nil
}

// validateSheetName checks Excel's worksheet naming rules against the name
// and the sheets currently in the workbook.
func (wb *Workbook) validateSheetName(name string) error {
	if name == "" {
		return fmt.Errorf("cells_foss: worksheet name is empty")
	}
	if utf8.RuneCountInString(name) > 31 {
		return fmt.Errorf("cells_foss: worksheet name %q is longer than 31 characters", name)
	}
	if strings.ContainsAny(name, `[]:*?/\`) {
		return fmt.Errorf(`cells_foss: worksheet name %q contains an invalid character (one of [ ] : * ? / \)`, name)
	}
	if strings.HasPrefix(name, "'") || strings.HasSuffix(name, "'") {
		return fmt.Errorf("cells_foss: worksheet name %q must not start or end with an apostrophe", name)
	}
	if strings.EqualFold(name, "History") {
		return fmt.Errorf("cells_foss: worksheet name %q is reserved by Excel", name)
	}
	for _, ws := range wb.Worksheets {
		if strings.EqualFold(ws.Name, name) {
			return fmt.Errorf("cells_foss: a worksheet named %q already exists", name)
		}
	}
	return nil
}
