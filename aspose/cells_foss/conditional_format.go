package cells_foss

import "fmt"

// Conditional formatting type constants.
const (
	CondTypeCellIs       = "cellIs"
	CondTypeContainsText = "containsText"
	CondTypeBeginsWith   = "beginsWith"
	CondTypeEndsWith     = "endsWith"
	CondTypeExpression   = "expression"
)

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
	// "endsWith", or "expression".
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
	Style *Style

	// StyleID is the internal style index used during save. It is
	// automatically assigned and should not be set manually.
	StyleID int

	// StopIfTrue controls whether subsequent rules are evaluated when this
	// rule's condition is met.
	StopIfTrue bool
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
	ws.ConditionalFormattings = append(ws.ConditionalFormattings, cf)
	ws.Modified = true
	if ws.cells != nil && ws.cells.wb != nil {
		ws.cells.wb.Modified = true
	}
	return nil
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
