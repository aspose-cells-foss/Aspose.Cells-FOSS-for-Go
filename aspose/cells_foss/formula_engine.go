package cells_foss

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// CalculateFormula evaluates a formula expression against the data in ws and
// returns the computed result.  Only a limited set of aggregation functions is
// supported; other formulas return an error.
//
// Supported functions:
//   - SUM(ref, …)
//   - AVERAGE(ref, …)
//   - MAX(ref, …)
//   - MIN(ref, …)
//
// References may be single cells ("A1"), ranges ("A1:A10" / "A1:C1"), or
// comma-separated combinations of both.  Non-numeric cells are silently
// ignored.  Empty ranges produce an error.
func CalculateFormula(formula string, ws *Worksheet) (interface{}, error) {
	formula = strings.TrimSpace(formula)
	if formula == "" {
		return nil, fmt.Errorf("formula: empty expression")
	}

	// Extract function name and argument text.
	paren := strings.IndexByte(formula, '(')
	if paren < 0 {
		return nil, fmt.Errorf("formula: missing '(' in %q", formula)
	}
	if !strings.HasSuffix(formula, ")") {
		return nil, fmt.Errorf("formula: missing closing ')' in %q", formula)
	}

	funcName := strings.ToUpper(strings.TrimSpace(formula[:paren]))
	argsText := strings.TrimSpace(formula[paren+1 : len(formula)-1])

	// Split arguments on commas (simple split — does not handle quoted commas
	// or nested parens; adequate for the supported functions).
	args := splitFormulaArgs(argsText)
	if len(args) == 0 {
		return nil, fmt.Errorf("formula: %s requires at least one argument", funcName)
	}

	// Expand all references to concrete values.
	var values []float64
	for _, arg := range args {
		arg = strings.TrimSpace(arg)
		if strings.Contains(arg, ":") {
			// Range reference.
			vals, err := resolveRange(arg, ws)
			if err != nil {
				return nil, fmt.Errorf("formula: %w", err)
			}
			values = append(values, vals...)
		} else {
			// Single cell reference.
			v, err := resolveCellRef(arg, ws)
			if err != nil {
				return nil, fmt.Errorf("formula: %w", err)
			}
			values = append(values, v)
		}
	}

	// Apply the function.
	switch funcName {
	case "SUM":
		return sum(values), nil
	case "AVERAGE":
		if len(values) == 0 {
			return nil, fmt.Errorf("formula: AVERAGE of empty range")
		}
		return sum(values) / float64(len(values)), nil
	case "MAX":
		if len(values) == 0 {
			return nil, fmt.Errorf("formula: MAX of empty range")
		}
		return max(values), nil
	case "MIN":
		if len(values) == 0 {
			return nil, fmt.Errorf("formula: MIN of empty range")
		}
		return min(values), nil
	default:
		return nil, fmt.Errorf("formula: unsupported function %q", funcName)
	}
}

// ---------------------------------------------------------------------------
// Reference resolution
// ---------------------------------------------------------------------------

// resolveRange expands a range like "A1:A5" or "A1:C1" into the numeric
// values of every cell in the rectangular region.
func resolveRange(rng string, ws *Worksheet) ([]float64, error) {
	parts := strings.SplitN(rng, ":", 2)
	if len(parts) != 2 {
		return nil, fmt.Errorf("invalid range %q", rng)
	}
	start := strings.TrimSpace(parts[0])
	end := strings.TrimSpace(parts[1])

	sc, sr := splitRef(start)
	ec, er := splitRef(end)

	sCol := colToNum(sc)
	eCol := colToNum(ec)

	if sCol > eCol {
		sCol, eCol = eCol, sCol
	}
	if sr > er {
		sr, er = er, sr
	}

	var values []float64
	for col := sCol; col <= eCol; col++ {
		for row := sr; row <= er; row++ {
			ref := numToCol(col) + strconv.Itoa(row)
			v, err := resolveCellRef(ref, ws)
			if err != nil {
				continue // skip empty / non-numeric cells
			}
			values = append(values, v)
		}
	}
	return values, nil
}

// resolveCellRef reads a single cell and returns its numeric value.  Empty
// cells and non-numeric cells produce an error so that callers can skip them.
func resolveCellRef(ref string, ws *Worksheet) (float64, error) {
	cell, err := ws.Cells().Get(ref)
	if err != nil {
		return 0, err
	}
	return cellToFloat(cell)
}

// cellToFloat converts a cell's Value to float64.  String values that look
// like numbers are parsed; booleans (true=1, false=0) are accepted.
func cellToFloat(cell *Cell) (float64, error) {
	switch v := cell.Value.(type) {
	case nil:
		return 0, fmt.Errorf("empty")
	case float64:
		return v, nil
	case float32:
		return float64(v), nil
	case int:
		return float64(v), nil
	case int64:
		return float64(v), nil
	case string:
		if v == "" {
			return 0, fmt.Errorf("empty")
		}
		f, err := strconv.ParseFloat(strings.TrimSpace(v), 64)
		if err != nil {
			return 0, fmt.Errorf("not numeric: %q", v)
		}
		return f, nil
	case bool:
		if v {
			return 1, nil
		}
		return 0, nil
	default:
		s := fmt.Sprint(v)
		f, err := strconv.ParseFloat(strings.TrimSpace(s), 64)
		if err != nil {
			return 0, fmt.Errorf("not numeric: %q", s)
		}
		return f, nil
	}
}

// ---------------------------------------------------------------------------
// Column conversion
// ---------------------------------------------------------------------------

// colToNum converts a column letter (or letters) to a zero-based index.
// "A" → 0, "B" → 1, …, "Z" → 25, "AA" → 26.
func colToNum(col string) int {
	col = strings.ToUpper(col)
	n := 0
	for _, ch := range col {
		n = n*26 + int(ch-'A') + 1
	}
	return n - 1
}

// numToCol converts a zero-based column index to letters.
// 0 → "A", 25 → "Z", 26 → "AA".
func numToCol(n int) string {
	var out strings.Builder
	n++ // convert to 1-based for the algorithm
	for n > 0 {
		n--
		out.WriteByte(byte('A' + n%26))
		n /= 26
	}
	// Reverse the string.
	s := out.String()
	runes := []rune(s)
	for i, j := 0, len(runes)-1; i < j; i, j = i+1, j-1 {
		runes[i], runes[j] = runes[j], runes[i]
	}
	return string(runes)
}

// ---------------------------------------------------------------------------
// Arithmetic helpers
// ---------------------------------------------------------------------------

func sum(vals []float64) float64 {
	var total float64
	for _, v := range vals {
		total += v
	}
	return total
}

func max(vals []float64) float64 {
	m := -math.MaxFloat64
	for _, v := range vals {
		if v > m {
			m = v
		}
	}
	return m
}

func min(vals []float64) float64 {
	m := math.MaxFloat64
	for _, v := range vals {
		if v < m {
			m = v
		}
	}
	return m
}

// ---------------------------------------------------------------------------
// Argument splitting (simple comma-split)
// ---------------------------------------------------------------------------

func splitFormulaArgs(s string) []string {
	var args []string
	depth := 0
	start := 0
	for i, ch := range s {
		switch ch {
		case '(':
			depth++
		case ')':
			depth--
		case ',':
			if depth == 0 {
				args = append(args, s[start:i])
				start = i + 1
			}
		}
	}
	args = append(args, s[start:])
	return args
}
