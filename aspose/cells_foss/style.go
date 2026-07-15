package cells_foss

// Alignment constants for the Horizontal and Vertical fields.
const (
	AlignLeft   = "left"
	AlignCenter = "center"
	AlignRight  = "right"
	AlignTop    = "top"
	AlignMiddle = "center"
	AlignBottom = "bottom"
)

// Fill type constants.
const (
	FillTypeNone    = "none"
	FillTypeSolid   = "solid"
	FillTypeGray125 = "gray125"
)

// ---------------------------------------------------------------------------
// Font
// ---------------------------------------------------------------------------

// Font describes the typographic properties applied to cell text.
type Font struct {
	// Name is the font family name (e.g. "Calibri", "Arial").
	Name string

	// Size is the font height in points.
	Size float64

	// Bold controls whether the text is rendered in bold weight.
	Bold bool

	// Italic controls whether the text is rendered in italic style.
	Italic bool

	// Color is the font colour as an ARGB hex string (e.g. "FF000000" for black).
	// An empty string means the default colour.
	Color string
}

// ---------------------------------------------------------------------------
// Fill
// ---------------------------------------------------------------------------

// Fill describes the background appearance of a cell.
type Fill struct {
	// Type is the fill pattern: "none", "solid", or "gray125".
	Type string

	// Color is the foreground colour for solid fills as an ARGB hex string.
	Color string
}

// ---------------------------------------------------------------------------
// Alignment
// ---------------------------------------------------------------------------

// Alignment controls how cell content is positioned within the cell bounds.
type Alignment struct {
	// Horizontal alignment: AlignLeft, AlignCenter, AlignRight, or "".
	Horizontal string

	// Vertical alignment: AlignTop, AlignMiddle, AlignBottom, or "".
	Vertical string

	// WrapText enables automatic line-wrapping when the cell content is wider
	// than the column.
	WrapText bool
}

// ---------------------------------------------------------------------------
// Border
// ---------------------------------------------------------------------------

// Border defines which sides of a cell have a visible rule and the colour of
// those rules.
type Border struct {
	Top    bool
	Bottom bool
	Left   bool
	Right  bool
	Color  string // ARGB hex; empty means default (black).
}

// ---------------------------------------------------------------------------
// Style
// ---------------------------------------------------------------------------

// Style groups font, fill, alignment, and border settings into a named
// formatting record. Multiple cells may share the same Style; the library
// automatically deduplicates identical styles during save.
type Style struct {
	Font      *Font
	Fill      *Fill
	Alignment *Alignment
	Border    *Border
}

// NewStyle returns a Style initialised with sensible defaults:
//
//	Font:      Calibri 11, black, no bold / italic
//	Fill:      none
//	Alignment: (not set)
//	Border:    (not set)
func NewStyle() *Style {
	return &Style{
		Font: &Font{
			Name:  "Calibri",
			Size:  11,
			Color: "FF000000",
		},
		Fill:   &Fill{Type: FillTypeNone},
		Border: &Border{},
	}
}

// DefaultStyle returns the implicit style applied to every cell that has not
// been assigned an explicit Style.  It matches the Excel default of Calibri 11
// with no special formatting.
func DefaultStyle() *Style {
	return NewStyle()
}
