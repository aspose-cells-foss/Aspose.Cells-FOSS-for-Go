package cells_foss_test

import (
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// TestFormula_LiteralArguments covers numeric aggregation functions with
// literal arguments, which previously resolved to zero values and were
// silently skipped.
func TestFormula_LiteralArguments(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("B1", float64(5))
	ws.Cells().Set("B2", float64(7))

	cases := []struct {
		formula string
		want    float64
	}{
		{"SUM(1,2,3)", 6},
		{"SUM(B1,2)", 7},
		{"SUM(B1:B2,10)", 22},
		{"AVERAGE(2,4)", 3},
		{"MAX(1,B2)", 7},
		{"MIN(B1,1)", 1},
		{"COUNT(1,B1,\"text\")", 2}, // text is skipped, literals count
	}
	for _, tc := range cases {
		got, err := cells_foss.CalculateFormula(tc.formula, ws)
		if err != nil {
			t.Errorf("%s: unexpected error: %v", tc.formula, err)
			continue
		}
		if got != tc.want {
			t.Errorf("%s = %v, want %v", tc.formula, got, tc.want)
		}
	}
}

// TestFormula_ArithmeticExpressions covers formulas that are not a single
// function call. These previously failed with "missing '('" or
// "missing closing ')'".
func TestFormula_ArithmeticExpressions(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", float64(10))
	ws.Cells().Set("A2", float64(4))
	ws.Cells().Set("B1", float64(5))
	ws.Cells().Set("B2", float64(6))

	cases := []struct {
		formula string
		want    interface{}
	}{
		{"1+1", float64(2)},
		{"2+3*4", float64(14)},        // precedence
		{"(2+3)*4", float64(20)},      // parens
		{"2^3", float64(8)},           // power
		{"2^3^2", float64(512)},       // right-associative power
		{"-5+2", float64(-3)},         // unary minus
		{"A1-A2", float64(6)},         // cell refs
		{"A1/A2", float64(2.5)},       // division
		{"SUM(B1:B2)+1", float64(12)}, // call inside arithmetic
		{"10-SUM(B1:B2)", float64(-1)},
		{"IF(B1>1,SUM(B1:B2),0)", float64(11)}, // nested call as IF result
	}
	for _, tc := range cases {
		got, err := cells_foss.CalculateFormula(tc.formula, ws)
		if err != nil {
			t.Errorf("%s: unexpected error: %v", tc.formula, err)
			continue
		}
		if got != tc.want {
			t.Errorf("%s = %v (%T), want %v", tc.formula, got, got, tc.want)
		}
	}

	// Excel error-value semantics: division by zero yields #DIV/0! (a
	// value, not a Go error) and arithmetic on text yields #VALUE!.
	v, err := cells_foss.CalculateFormula("1/0", ws)
	if err != nil || v != cells_foss.ErrDiv0 {
		t.Errorf("1/0 = %v (%v), want %v", v, err, cells_foss.ErrDiv0)
	}
	ws.Cells().Set("C1", "hello")
	v, err = cells_foss.CalculateFormula("C1+1", ws)
	if err != nil || v != cells_foss.ErrValue {
		t.Errorf("text+1 = %v (%v), want %v", v, err, cells_foss.ErrValue)
	}
}

// TestFormula_ConcatOperator covers the & string concatenation operator.
func TestFormula_ConcatOperator(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "foo")

	cases := []struct {
		formula string
		want    string
	}{
		{`"a"&"b"`, "ab"},
		{`A1&"!"`, "foo!"},
		{`"n="&5`, "n=5"},
	}
	for _, tc := range cases {
		got, err := cells_foss.CalculateFormula(tc.formula, ws)
		if err != nil {
			t.Errorf("%s: unexpected error: %v", tc.formula, err)
			continue
		}
		if got != tc.want {
			t.Errorf("%s = %v, want %q", tc.formula, got, tc.want)
		}
	}
}

// TestFormula_QuotedCommaArgs covers arguments containing commas inside
// quoted strings, which previously split into extra arguments.
func TestFormula_QuotedCommaArgs(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "hello")

	got, err := cells_foss.CalculateFormula(`CONCAT(A1,"a,b")`, ws)
	if err != nil {
		t.Fatalf("CONCAT with quoted comma: %v", err)
	}
	if got != "helloa,b" {
		t.Errorf("CONCAT(A1,\"a,b\") = %q, want %q", got, "helloa,b")
	}

	// Escaped doubled quotes inside a literal.
	got, err = cells_foss.CalculateFormula(`CONCAT("say ""hi""")`, ws)
	if err != nil {
		t.Fatalf("CONCAT escaped quotes: %v", err)
	}
	if got != `say "hi"` {
		t.Errorf("escaped quotes = %q, want %q", got, `say "hi"`)
	}

	// Literals and nested calls in CONCAT.
	got, err = cells_foss.CalculateFormula(`CONCAT("n=",5)`, ws)
	if err != nil {
		t.Fatalf("CONCAT literal: %v", err)
	}
	if got != "n=5" {
		t.Errorf("CONCAT(\"n=\",5) = %q, want %q", got, "n=5")
	}
}

// TestFormula_StringFunctionsUnicode covers LEN/LEFT/RIGHT/MID on multi-byte
// text. These previously operated on bytes, splitting UTF-8 sequences.
func TestFormula_StringFunctionsUnicode(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "日本語テキスト")

	cases := []struct {
		formula string
		want    interface{}
	}{
		{"LEN(A1)", float64(7)},
		{"LEFT(A1,3)", "日本語"},
		{"RIGHT(A1,3)", "キスト"},
		{"MID(A1,2,3)", "本語テ"},
	}
	for _, tc := range cases {
		got, err := cells_foss.CalculateFormula(tc.formula, ws)
		if err != nil {
			t.Errorf("%s: unexpected error: %v", tc.formula, err)
			continue
		}
		if got != tc.want {
			t.Errorf("%s = %v, want %v", tc.formula, got, tc.want)
		}
	}
}

// TestFormula_MidInvalidArgs verifies MID rejects invalid arguments with an
// error instead of panicking with a slice bounds violation.
func TestFormula_MidInvalidArgs(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Hello World")

	// Negative character count previously panicked (slice bounds [:-1]).
	if _, err := cells_foss.CalculateFormula("MID(A1,1,-1)", ws); err == nil {
		t.Error("MID with negative count should error, not panic")
	}
	// Zero-based start is invalid in Excel semantics.
	if _, err := cells_foss.CalculateFormula("MID(A1,0,3)", ws); err == nil {
		t.Error("MID with start 0 should error")
	}
	// Sanity: valid calls still work.
	got, err := cells_foss.CalculateFormula("MID(A1,7,5)", ws)
	if err != nil || got != "World" {
		t.Errorf("MID(A1,7,5) = %v (%v), want World", got, err)
	}
	got, err = cells_foss.CalculateFormula("MID(A1,7,100)", ws)
	if err != nil || got != "World" {
		t.Errorf("MID clamped = %v (%v), want World", got, err)
	}
}
