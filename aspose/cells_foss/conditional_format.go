package cells_foss

import "fmt"

// Conditional formatting type constants.
const (
	CondTypeCellIs       = "cellIs"
	CondTypeContainsText = "containsText"
	CondTypeBeginsWith   = "beginsWith"
	CondTypeEndsWith     = "endsWith"
	CondTypeExpression   = "expression"
	CondTypeColorScale   = "colorScale"
	CondTypeDataBar      = "dataBar"
	CondTypeIconSet      = "iconSet"
	CondTypeTop10        = "top10"
)

// Threshold types for color scale, data bar, and icon set rules
// (ST_CfvoType).
const (
	CondValueMin     = "min"
	CondValueMax     = "max"
	CondValueNum     = "num"
	CondValuePercent = "percent"
	CondValueFormula = "formula"
)

// Common icon set names (ST_IconSetType). AddConditionalFormatting accepts
// any name in the enumeration, including the ones without constants here
// (3TrafficLights2, 3Symbols2, 4ArrowsGray, 4RedToBlack, 5ArrowsGray).
const (
	IconSet3Arrows        = "3Arrows"
	IconSet3ArrowsGray    = "3ArrowsGray"
	IconSet3Flags         = "3Flags"
	IconSet3TrafficLights = "3TrafficLights1"
	IconSet3Signs         = "3Signs"
	IconSet3Symbols       = "3Symbols"
	IconSet4Arrows        = "4Arrows"
	IconSet4Rating        = "4Rating"
	IconSet4TrafficLights = "4TrafficLights"
	IconSet5Arrows        = "5Arrows"
	IconSet5Rating        = "5Rating"
	IconSet5Quarters      = "5Quarters"
)

// CondThreshold is one conditional-format value object (CT_Cfvo): a color
// scale stop, a data bar endpoint, or an icon set cutoff.
type CondThreshold struct {
	// Type is one of CondValueMin, CondValueMax, CondValueNum,
	// CondValuePercent, or CondValueFormula.
	Type string

	// Val is the threshold value: a number for num, "0"–"100" for percent,
	// an expression for formula. Ignored for min/max.
	Val string
}

// ColorScale is a 2- or 3-color gradient rule (CT_ColorScale). Colors are
// ARGB strings ("FFF8696B") matching Thresholds one-for-one.
type ColorScale struct {
	// Thresholds holds 2 or 3 stops, typically min/max or
	// min/percent(50)/max.
	Thresholds []CondThreshold

	// Colors holds one ARGB color per threshold.
	Colors []string
}

// DataBar is an in-cell bar rule (CT_DataBar).
type DataBar struct {
	// Min is the value where the bar length is zero.
	Min CondThreshold

	// Max is the value where the bar fills the cell.
	Max CondThreshold

	// Color is the ARGB bar color; required.
	Color string

	// HideValue draws only the bar, suppressing the cell value
	// (showValue="0"). The zero value keeps Excel's default of showing
	// both.
	HideValue bool

	// MinLength and MaxLength bound the bar length as a percentage of the
	// cell width. Zero uses Excel's defaults (10 and 90).
	MinLength int
	MaxLength int
}

// IconSet is an icon-overlay rule (CT_IconSet). Thresholds holds one cutoff
// per icon, typically percent-typed (0/33/67 for 3-icon sets).
type IconSet struct {
	// IconSet is an ST_IconSetType name, e.g. IconSet3Arrows. The name's
	// leading digit is how many thresholds are required.
	IconSet string

	// Thresholds holds one cutoff per icon.
	Thresholds []CondThreshold

	// HideValue draws only the icon, suppressing the cell value.
	HideValue bool

	// Reverse flips the icon order.
	Reverse bool
}

// validIconSets is the full ST_IconSetType enumeration.
var validIconSets = map[string]bool{
	"3Arrows": true, "3ArrowsGray": true, "3Flags": true,
	"3TrafficLights1": true, "3TrafficLights2": true,
	"3Signs": true, "3Symbols": true, "3Symbols2": true,
	"4Arrows": true, "4ArrowsGray": true, "4RedToBlack": true,
	"4Rating": true, "4TrafficLights": true,
	"5Arrows": true, "5ArrowsGray": true, "5Rating": true, "5Quarters": true,
}

// Conditional formatting operator constants for cellIs rules.
const (
	CondOpGreaterThan        = "greaterThan"
	CondOpLessThan           = "lessThan"
	CondOpBetween            = "between"
	CondOpEqual              = "equal"
	CondOpNotEqual           = "notEqual"
	CondOpGreaterThanOrEqual = "greaterThanOrEqual"
	CondOpLessThanOrEqual    = "lessThanOrEqual"
)

// ConditionalFormatting represents a set of conditional formatting rules
// applied to a range of cells on a worksheet.
type ConditionalFormatting struct {
	// Ref is the A1-style range this conditional formatting applies to,
	// e.g. "A1:A10" or "A1:B10 D1:D20".
	Ref string

	// Rules holds the conditional formatting rules in priority order.
	Rules []*ConditionalFormattingRule
}

// ConditionalFormattingRule represents a single conditional formatting rule.
type ConditionalFormattingRule struct {
	// Type is the rule type: "cellIs", "containsText", "beginsWith",
	// "endsWith", "expression", "colorScale", "dataBar", "iconSet", or
	// "top10".
	Type string

	// Operator is the comparison operator for cellIs rules, e.g.
	// "greaterThan", "lessThan", "between". Empty for other rule types.
	Operator string

	// Formula is the condition formula or value. For cellIs rules this is
	// the comparison value; for expression rules this is the formula.
	Formula string

	// Formula2 is the second formula for "between" operators.
	Formula2 string

	// Text is the text to match for containsText, beginsWith, endsWith rules.
	Text string

	// Priority controls the evaluation order. Lower numbers have higher
	// priority. Rules are evaluated in ascending priority order.
	Priority int

	// Style is the formatting to apply when the condition is met.
	// This is used for DXF (Differential Formatting) in conditional formatting.
	Style *Style

	// StyleID is the internal DXF style index used during save. It is
	// automatically assigned and should not be set manually.
	StyleID int

	// hasDxfID records that StyleID was read from a dxfId attribute in a
	// loaded file, so the saver re-emits it even when it is 0.
	hasDxfID bool

	// StopIfTrue controls whether subsequent rules are evaluated when this
	// rule's condition is met.
	StopIfTrue bool

	// Rank is how many cells a top10 rule formats; 0 saves Excel's
	// default of 10. Ignored by other rule types.
	Rank int

	// Bottom makes a top10 rule format the lowest-ranked cells instead of
	// the highest.
	Bottom bool

	// Percent interprets a top10 rule's Rank as a percentage.
	Percent bool

	// ColorScale carries the configuration of a colorScale rule; nil for
	// all other types.
	ColorScale *ColorScale

	// DataBar carries the configuration of a dataBar rule; nil for all
	// other types.
	DataBar *DataBar

	// IconSet carries the configuration of an iconSet rule; nil for all
	// other types.
	IconSet *IconSet
}

// AddConditionalFormatting adds a conditional formatting rule set to the
// worksheet. The Ref field on the ConditionalFormatting specifies the range.
//
//	cf := &cells_foss.ConditionalFormatting{
//	    Ref: "A1:A10",
//	    Rules: []*cells_foss.ConditionalFormattingRule{
//	        {
//	            Type:     cells_foss.CondTypeCellIs,
//	            Operator: cells_foss.CondOpGreaterThan,
//	            Formula:  "100",
//	            Style:    &cells_foss.Style{Font: &cells_foss.Font{Bold: true}},
//	        },
//	    },
//	}
//	ws.AddConditionalFormatting(cf)
func (ws *Worksheet) AddConditionalFormatting(cf *ConditionalFormatting) error {
	if cf == nil {
		return fmt.Errorf("cells_foss: ConditionalFormatting is nil")
	}
	if cf.Ref == "" {
		return fmt.Errorf("cells_foss: ConditionalFormatting ref is empty")
	}
	for _, rule := range cf.Rules {
		if err := validateCFRule(rule); err != nil {
			return err
		}
	}
	ws.ConditionalFormattings = append(ws.ConditionalFormattings, cf)
	ws.Modified = true
	if ws.cells != nil && ws.cells.wb != nil {
		ws.cells.wb.Modified = true
	}
	return nil
}

// validateCFRule checks the configuration carried by the colorScale,
// dataBar, iconSet, and top10 rule types. The text and expression types
// keep their permissive behavior.
func validateCFRule(rule *ConditionalFormattingRule) error {
	switch rule.Type {
	case CondTypeColorScale:
		cs := rule.ColorScale
		if cs == nil {
			return fmt.Errorf("cells_foss: colorScale rule requires ColorScale")
		}
		if len(cs.Thresholds) < 2 || len(cs.Thresholds) > 3 {
			return fmt.Errorf("cells_foss: colorScale needs 2 or 3 thresholds, got %d", len(cs.Thresholds))
		}
		if len(cs.Colors) != len(cs.Thresholds) {
			return fmt.Errorf("cells_foss: colorScale has %d colors for %d thresholds", len(cs.Colors), len(cs.Thresholds))
		}
		for _, t := range cs.Thresholds {
			if err := validateThreshold(t); err != nil {
				return err
			}
		}
		for _, c := range cs.Colors {
			if c == "" {
				return fmt.Errorf("cells_foss: colorScale color is empty")
			}
		}
	case CondTypeDataBar:
		db := rule.DataBar
		if db == nil {
			return fmt.Errorf("cells_foss: dataBar rule requires DataBar")
		}
		if db.Color == "" {
			return fmt.Errorf("cells_foss: dataBar requires Color")
		}
		if err := validateThreshold(db.Min); err != nil {
			return err
		}
		if err := validateThreshold(db.Max); err != nil {
			return err
		}
		if db.MinLength < 0 || db.MinLength > 100 || db.MaxLength < 0 || db.MaxLength > 100 {
			return fmt.Errorf("cells_foss: dataBar lengths must be within 0-100")
		}
	case CondTypeIconSet:
		is := rule.IconSet
		if is == nil {
			return fmt.Errorf("cells_foss: iconSet rule requires IconSet")
		}
		if !validIconSets[is.IconSet] {
			return fmt.Errorf("cells_foss: unknown icon set %q", is.IconSet)
		}
		if want := int(is.IconSet[0]) - '0'; len(is.Thresholds) != want {
			return fmt.Errorf("cells_foss: icon set %q needs %d thresholds, got %d", is.IconSet, want, len(is.Thresholds))
		}
		for _, t := range is.Thresholds {
			if err := validateThreshold(t); err != nil {
				return err
			}
		}
	case CondTypeTop10:
		if rule.Rank < 0 {
			return fmt.Errorf("cells_foss: top10 rank must not be negative")
		}
	}
	return nil
}

func validateThreshold(t CondThreshold) error {
	switch t.Type {
	case CondValueMin, CondValueMax:
		return nil
	case CondValueNum, CondValuePercent, CondValueFormula:
		if t.Val == "" {
			return fmt.Errorf("cells_foss: %s threshold requires Val", t.Type)
		}
		return nil
	}
	return fmt.Errorf("cells_foss: unknown threshold type %q", t.Type)
}

// RemoveConditionalFormatting removes the first conditional formatting rule
// set whose Ref exactly matches the given ref string. An error is returned
// when no matching rule set is found.
func (ws *Worksheet) RemoveConditionalFormatting(ref string) error {
	for i, cf := range ws.ConditionalFormattings {
		if cf.Ref == ref {
			ws.ConditionalFormattings = append(ws.ConditionalFormattings[:i], ws.ConditionalFormattings[i+1:]...)
			ws.Modified = true
			if ws.cells != nil && ws.cells.wb != nil {
				ws.cells.wb.Modified = true
			}
			return nil
		}
	}
	return fmt.Errorf("cells_foss: no ConditionalFormatting found for ref %q", ref)
}
