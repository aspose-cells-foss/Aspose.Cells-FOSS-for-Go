package cells_foss

import (
	"fmt"
	"strings"
)

// Hyperlink is a link attached to a single cell. External links (Target)
// are stored as relationships in the sheet part; internal links (Location,
// e.g. "Sheet2!A1") are stored as an attribute of the <hyperlink> element.
// Exactly one of Target or Location must be set.
type Hyperlink struct {
	// Ref is the A1 reference of the cell carrying the link.
	Ref string

	// Target is an external URL or file path, e.g.
	// "https://example.com" or "reports/q1.xlsx".
	Target string

	// Location is an in-workbook destination, e.g. "Sheet2!A1".
	Location string

	// Tooltip is the optional screen tip shown on hover.
	Tooltip string

	// Display is the optional friendly name stored in the file. The
	// visible cell text is the cell value, set separately.
	Display string
}

// AddHyperlink attaches an external link (URL or file path) to a cell.
// Excel allows one link per cell; a second AddHyperlink for the same cell
// is refused. The cell's visible text is not changed — set the cell value
// separately when you want a display label.
func (ws *Worksheet) AddHyperlink(ref, target string) error {
	return ws.addHyperlink(&Hyperlink{Ref: ref, Target: target})
}

// AddInternalHyperlink attaches a link to a location inside the same
// workbook (e.g. "Sheet2!A1") to a cell.
func (ws *Worksheet) AddInternalHyperlink(ref, location string) error {
	return ws.addHyperlink(&Hyperlink{Ref: ref, Location: location})
}

// RemoveHyperlink detaches the link from a cell. An error is returned
// when the cell carries no link.
func (ws *Worksheet) RemoveHyperlink(ref string) error {
	ref = strings.ToUpper(strings.TrimSpace(ref))
	for i, h := range ws.Hyperlinks {
		if h.Ref == ref {
			ws.Hyperlinks = append(ws.Hyperlinks[:i], ws.Hyperlinks[i+1:]...)
			ws.Modified = true
			return nil
		}
	}
	return fmt.Errorf("cells_foss: cell %s has no hyperlink", ref)
}

func (ws *Worksheet) addHyperlink(h *Hyperlink) error {
	if h == nil {
		return fmt.Errorf("cells_foss: hyperlink is nil")
	}
	ref := strings.ToUpper(strings.TrimSpace(h.Ref))
	if _, _, _, _, err := parseA1Range(ref); err != nil {
		return err
	}
	if (h.Target == "") == (h.Location == "") {
		return fmt.Errorf("cells_foss: hyperlink on %s must set exactly one of Target or Location", ref)
	}
	for _, existing := range ws.Hyperlinks {
		if existing.Ref == ref {
			return fmt.Errorf("cells_foss: cell %s already has a hyperlink", ref)
		}
	}
	h.Ref = ref
	ws.Hyperlinks = append(ws.Hyperlinks, h)
	ws.Modified = true
	return nil
}
