package cells_foss

import "fmt"

// Workbook is the top-level object representing an Excel workbook.
// It owns the collection of worksheets and caches the original source XML
// so that unmodified parts can be written back without regeneration.
type Workbook struct {
	// Worksheets holds every sheet in the workbook in display order.
	Worksheets []*Worksheet

	// SourceXML caches the raw XML bytes of the first worksheet as read from
	// the .xlsx file. When the workbook has not been modified this payload is
	// written back verbatim to preserve round-trip fidelity.
	SourceXML []byte

	// StylesXML caches the raw bytes of xl/styles.xml from the loaded file.
	// When the workbook is unmodified this payload is reused verbatim.
	StylesXML []byte

	// styles is the ordered registry of Style records. Index 0 is always the
	// default cell style.  registerStyle / getStyle manage deduplication.
	styles []*Style

	// password, when non-empty, causes Save to encrypt the workbook using
	// ECMA-376 Agile Encryption (SHA-512 + AES-256-CBC).
	password string

	// Modified is set to true whenever cell data or styles change. The saver
	// uses this flag to decide whether to regenerate XML or reuse cached XML.
	Modified bool

	// FilePath is the on-disk path from which the workbook was loaded, or the
	// path to which it was last saved.
	FilePath string
}

// SetPassword configures an open password for the workbook.  Subsequent
// calls to Save will produce an encrypted .xlsx file that requires this
// password to open.  Pass an empty string to remove password protection.
func (wb *Workbook) SetPassword(password string) error {
	if wb == nil {
		return fmt.Errorf("cells_foss: workbook is nil")
	}
	wb.password = password
	wb.Modified = true
	return nil
}

// VerifyPassword reports whether pw matches the password that was used to
// encrypt this workbook (as set by SetPassword, or read from an encrypted
// file).  When the workbook was not loaded from an encrypted file and
// SetPassword has not been called, VerifyPassword returns true for any input.
func (wb *Workbook) VerifyPassword(pw string) bool {
	if wb.password == "" {
		return true // no password set
	}
	return wb.password == pw
}

// registerStyle adds style to the workbook's style registry (deduplicating
// identical styles) and returns its zero-based index.  This index is used as
// the cell's StyleID (the s attribute in the XML).
func (wb *Workbook) registerStyle(style *Style) int {
	for i, s := range wb.styles {
		if stylesEqual(s, style) {
			return i
		}
	}
	id := len(wb.styles)
	wb.styles = append(wb.styles, style)
	wb.Modified = true
	return id
}

// getStyle returns the Style at index id, or nil when id is out of range.
func (wb *Workbook) getStyle(id int) *Style {
	if id < 0 || id >= len(wb.styles) {
		return nil
	}
	return wb.styles[id]
}

// stylesEqual reports whether two Style values are semantically identical.
func stylesEqual(a, b *Style) bool {
	if a == nil && b == nil {
		return true
	}
	if a == nil || b == nil {
		return false
	}
	return fontEqual(a.Font, b.Font) &&
		fillEqual(a.Fill, b.Fill) &&
		alignmentEqual(a.Alignment, b.Alignment) &&
		borderEqual(a.Border, b.Border)
}

func fontEqual(a, b *Font) bool {
	if a == nil && b == nil {
		return true
	}
	if a == nil || b == nil {
		return false
	}
	return a.Name == b.Name && a.Size == b.Size &&
		a.Bold == b.Bold && a.Italic == b.Italic &&
		a.Color == b.Color
}

func fillEqual(a, b *Fill) bool {
	if a == nil && b == nil {
		return true
	}
	if a == nil || b == nil {
		return false
	}
	return a.Type == b.Type && a.Color == b.Color
}

func alignmentEqual(a, b *Alignment) bool {
	if a == nil && b == nil {
		return true
	}
	if a == nil || b == nil {
		return false
	}
	return a.Horizontal == b.Horizontal &&
		a.Vertical == b.Vertical &&
		a.WrapText == b.WrapText
}

func borderEqual(a, b *Border) bool {
	if a == nil && b == nil {
		return true
	}
	if a == nil || b == nil {
		return false
	}
	return a.Top == b.Top && a.Bottom == b.Bottom &&
		a.Left == b.Left && a.Right == b.Right &&
		a.Color == b.Color
}

// Load opens the .xlsx file at path and returns the populated Workbook.
// It is a convenience wrapper around LoadWorkbook that reads all sheets,
// resolves shared strings, and caches the raw XML of the first worksheet.
func Load(path string) (*Workbook, error) {
	return LoadWorkbook(path)
}

// NewWorkbook creates an empty Workbook containing a single blank worksheet
// named "Sheet1". The workbook is marked as modified so that the first Save
// call generates fresh XML.
func NewWorkbook() *Workbook {
	wb := &Workbook{
		Modified: true,
		styles:   []*Style{DefaultStyle()},
	}
	ws := &Worksheet{
		Name:  "Sheet1",
		Index: 0,
		cells: &Cells{},
	}
	ws.cells.setParent(wb)
	wb.Worksheets = []*Worksheet{ws}
	return wb
}

// Save writes the workbook to the given file path in .xlsx format.
// When the workbook was loaded from a file and has not been modified, the
// original sheet XML is reused verbatim. Otherwise, all XML parts are
// regenerated from the in-memory model.
func (wb *Workbook) Save(path string) error {
	if wb == nil {
		return fmt.Errorf("cells_foss: cannot save a nil Workbook")
	}
	return SaveWorkbook(wb, path)
}
