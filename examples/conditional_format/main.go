// Conditional formatting example: highlight cells based on their values
// using DXF (differential formatting) styles, color scales, data bars,
// icon sets, and top-N rules.
//
//	cd examples && go run ./conditional_format/
package main

import (
	"fmt"
	"os"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]

	ws.Cells().Set("A1", "Student")
	ws.Cells().Set("B1", "Score")
	ws.Cells().Set("C1", "Attendance")
	ws.Cells().Set("D1", "Trend")
	ws.Cells().Set("E1", "Rating")
	students := []struct {
		name       string
		score      float64
		attendance float64
		trend      float64
		rating     float64
	}{
		{"Alice", 91, 98, 0.8, 4.5},
		{"Bob", 74, 82, 0.2, 3.1},
		{"Carol", 58, 91, -0.4, 2.4},
		{"Dan", 42, 67, -0.9, 1.2},
		{"Eve", 85, 74, 0.5, 3.9},
	}
	for i, s := range students {
		ws.Cells().Set(fmt.Sprintf("A%d", i+2), s.name)
		ws.Cells().Set(fmt.Sprintf("B%d", i+2), s.score)
		ws.Cells().Set(fmt.Sprintf("C%d", i+2), s.attendance)
		ws.Cells().Set(fmt.Sprintf("D%d", i+2), s.trend)
		ws.Cells().Set(fmt.Sprintf("E%d", i+2), s.rating)
	}

	// Score > 80: green fill with dark-green bold font (Excel's "Good" look).
	good := &cells_foss.Style{
		Font: &cells_foss.Font{Bold: true, Color: "FF006100"},
		Fill: &cells_foss.Fill{Type: "solid", Color: "FFC6EFCE"},
	}
	// Score < 50: red fill with dark-red font (Excel's "Bad" look).
	bad := &cells_foss.Style{
		Font: &cells_foss.Font{Color: "FF9C0006"},
		Fill: &cells_foss.Fill{Type: "solid", Color: "FFFFC7CE"},
	}
	// Top 2 scores: gold bold font.
	gold := &cells_foss.Style{
		Font: &cells_foss.Font{Bold: true, Color: "FF9C6500"},
	}

	scores := &cells_foss.ConditionalFormatting{
		Ref: "B2:B6",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeCellIs,
				Operator: cells_foss.CondOpGreaterThan,
				Formula:  "80",
				Priority: 1,
				Style:    good,
			},
			{
				Type:     cells_foss.CondTypeCellIs,
				Operator: cells_foss.CondOpLessThan,
				Formula:  "50",
				Priority: 2,
				Style:    bad,
			},
			{
				Type:     cells_foss.CondTypeTop10,
				Rank:     2,
				Priority: 3,
				Style:    gold,
			},
		},
	}
	if err := ws.AddConditionalFormatting(scores); err != nil {
		fmt.Fprintf(os.Stderr, "AddConditionalFormatting (scores): %v\n", err)
		os.Exit(1)
	}

	// Attendance: in-cell blue data bars scaled to 0–100.
	attendance := &cells_foss.ConditionalFormatting{
		Ref: "C2:C6",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeDataBar,
				Priority: 4,
				DataBar: &cells_foss.DataBar{
					Min:   cells_foss.CondThreshold{Type: cells_foss.CondValueNum, Val: "0"},
					Max:   cells_foss.CondThreshold{Type: cells_foss.CondValueNum, Val: "100"},
					Color: "FF638EC6",
				},
			},
		},
	}
	if err := ws.AddConditionalFormatting(attendance); err != nil {
		fmt.Fprintf(os.Stderr, "AddConditionalFormatting (attendance): %v\n", err)
		os.Exit(1)
	}

	// Trend: three arrows split by percent rank.
	trend := &cells_foss.ConditionalFormatting{
		Ref: "D2:D6",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeIconSet,
				Priority: 5,
				IconSet: &cells_foss.IconSet{
					IconSet: cells_foss.IconSet3Arrows,
					Thresholds: []cells_foss.CondThreshold{
						{Type: cells_foss.CondValuePercent, Val: "0"},
						{Type: cells_foss.CondValuePercent, Val: "33"},
						{Type: cells_foss.CondValuePercent, Val: "67"},
					},
				},
			},
		},
	}
	if err := ws.AddConditionalFormatting(trend); err != nil {
		fmt.Fprintf(os.Stderr, "AddConditionalFormatting (trend): %v\n", err)
		os.Exit(1)
	}

	// Rating: red → yellow → green gradient.
	rating := &cells_foss.ConditionalFormatting{
		Ref: "E2:E6",
		Rules: []*cells_foss.ConditionalFormattingRule{
			{
				Type:     cells_foss.CondTypeColorScale,
				Priority: 6,
				ColorScale: &cells_foss.ColorScale{
					Thresholds: []cells_foss.CondThreshold{
						{Type: cells_foss.CondValueMin},
						{Type: cells_foss.CondValuePercent, Val: "50"},
						{Type: cells_foss.CondValueMax},
					},
					Colors: []string{"FFF8696B", "FFFFEB84", "FF63BE7B"},
				},
			},
		},
	}
	if err := ws.AddConditionalFormatting(rating); err != nil {
		fmt.Fprintf(os.Stderr, "AddConditionalFormatting (rating): %v\n", err)
		os.Exit(1)
	}

	path := "outputfiles/conditional_format.xlsx"
	if err := wb.Save(path); err != nil {
		fmt.Fprintf(os.Stderr, "Save: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Saved %s with %d rule set(s)\n", path, len(ws.ConditionalFormattings))

	// Rules and their configuration survive a save/load round trip.
	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Load: %v\n", err)
		os.Exit(1)
	}
	for _, lcf := range loaded.Worksheets[0].ConditionalFormattings {
		fmt.Printf("  reloaded %s: %d rule(s)\n", lcf.Ref, len(lcf.Rules))
		for _, r := range lcf.Rules {
			switch {
			case r.ColorScale != nil:
				fmt.Printf("    - %s: %d stops, colors %v (priority %d)\n",
					r.Type, len(r.ColorScale.Thresholds), r.ColorScale.Colors, r.Priority)
			case r.DataBar != nil:
				fmt.Printf("    - %s: %s→%s color=%s hideValue=%v (priority %d)\n",
					r.Type, r.DataBar.Min.Type, r.DataBar.Max.Type, r.DataBar.Color, r.DataBar.HideValue, r.Priority)
			case r.IconSet != nil:
				fmt.Printf("    - %s: %s, %d thresholds (priority %d)\n",
					r.Type, r.IconSet.IconSet, len(r.IconSet.Thresholds), r.Priority)
			case r.Type == cells_foss.CondTypeTop10:
				fmt.Printf("    - %s: rank=%d bottom=%v (priority %d)\n",
					r.Type, r.Rank, r.Bottom, r.Priority)
			default:
				fmt.Printf("    - %s %s %s (priority %d)\n", r.Type, r.Operator, r.Formula, r.Priority)
			}
		}
	}
	fmt.Println("Open in Excel — scores are highlighted, attendance shows bars,")
	fmt.Println("trend shows arrows, and ratings fade red → yellow → green.")
}
