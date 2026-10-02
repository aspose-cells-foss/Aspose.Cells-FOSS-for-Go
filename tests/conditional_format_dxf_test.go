package cells_foss_test

import (
	"os"
	"testing"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// TestConditionalFormatting_WithDXFStyle verifies that conditional formatting with DXF styles can be saved.
func TestConditionalFormatting_WithDXFStyle(t *testing.T) {
	tmpFile := "test_cf_dxf.xlsx"
	defer os.Remove(tmpFile)

	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Add some data
	ws.Cells().Set("A1", 50.0)
	ws.Cells().Set("A2", 150.0)
	ws.Cells().Set("A3", 200.0)

	// Add conditional formatting with DXF style
	cf := &cells_foss.ConditionalFormatting{
		Ref: "A1:A10",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeCellIs,
				Operator: cells_foss.CondOpGreaterThan,
				Formula:  "100",
				Priority: 1,
				Style: &cells_foss.Style{
					Font: &cells_foss.Font{
						Bold:  true,
						Color: "FFFF0000",
					},
					Fill: &cells_foss.Fill{
						Type:  "solid",
						Color: "FFFFFF00",
					},
				},
			},
		},
	}

	if err := ws.AddConditionalFormatting(cf); err != nil {
		t.Fatalf("AddConditionalFormatting failed: %v", err)
	}

	// Save workbook
	if err := wb.Save(tmpFile); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify file was created
	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("Output file was not created")
	}
}

// TestConditionalFormatting_MultipleRulesWithDXF verifies multiple conditional formatting rules with different DXF styles.
func TestConditionalFormatting_MultipleRulesWithDXF(t *testing.T) {
	tmpFile := "test_cf_multi_dxf.xlsx"
	defer os.Remove(tmpFile)

	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	// Add some data
	ws.Cells().Set("A1", 50.0)
	ws.Cells().Set("A2", 150.0)
	ws.Cells().Set("A3", 250.0)

	// Add conditional formatting with multiple rules
	cf := &cells_foss.ConditionalFormatting{
		Ref: "A1:A10",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeCellIs,
				Operator: cells_foss.CondOpGreaterThan,
				Formula:  "200",
				Priority: 1,
				Style: &cells_foss.Style{
					Font: &cells_foss.Font{Bold: true},
				},
			},
			{
				Type:     cells_foss.CondTypeCellIs,
				Operator: cells_foss.CondOpBetween,
				Formula:  "100",
				Formula2: "200",
				Priority: 2,
				Style: &cells_foss.Style{
					Fill: &cells_foss.Fill{
						Type:  "solid",
						Color: "FF00FF00",
					},
				},
			},
		},
	}

	if err := ws.AddConditionalFormatting(cf); err != nil {
		t.Fatalf("AddConditionalFormatting failed: %v", err)
	}

	// Save workbook
	if err := wb.Save(tmpFile); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify file was created
	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("Output file was not created")
	}
}
