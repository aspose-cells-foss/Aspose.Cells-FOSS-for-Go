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
//   - COUNT(ref, …)
//   - CONCAT(ref, …)
//   - IF(condition, true_val, false_val)
//   - COUNTIF(range, criteria)
//   - VLOOKUP(value, range, col_index, [match])
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

	// For functions that need raw arguments (CONCAT, IF, COUNTIF, VLOOKUP),
	// pass args directly without pre-expanding to float64.
	switch funcName {
	case "CONCAT":
		return concatFunc(args, ws)
	case "IF":
		return ifFunc(args, ws)
	case "COUNTIF":
		return countIfFunc(args, ws)
	case "VLOOKUP":
		return vlookupFunc(args, ws)
	}

	// For numeric aggregation functions, expand all references to float64 values.
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
				// For numeric functions, skip non-numeric cells
				continue
			}
			values = append(values, v)
		}
	}

	// Apply the numeric function.
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
	case "COUNT":
		return float64(len(values)), nil
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

// ColToNum converts a column letter (or letters) to a zero-based index.
// "A" → 0, "B" → 1, …, "Z" → 25, "AA" → 26.
func ColToNum(col string) int {
	col = strings.ToUpper(col)
	n := 0
	for _, ch := range col {
		n = n*26 + int(ch-'A') + 1
	}
	return n - 1
}

// NumToCol converts a zero-based column index to letters.
// 0 → "A", 25 → "Z", 26 → "AA".
func NumToCol(n int) string {
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

// colToNum is the internal alias for backward compatibility.
func colToNum(col string) int {
	return ColToNum(col)
}

// numToCol is the internal alias for backward compatibility.
func numToCol(n int) string {
	return NumToCol(n)
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

// ---------------------------------------------------------------------------
// CONCAT function
// ---------------------------------------------------------------------------

func concatFunc(args []string, ws *Worksheet) (interface{}, error) {
	var result strings.Builder
	for _, arg := range args {
		arg = strings.TrimSpace(arg)
		if strings.Contains(arg, ":") {
			// Range reference - get all cells
			cells, err := resolveRangeRaw(arg, ws)
			if err != nil {
				return nil, fmt.Errorf("formula: %w", err)
			}
			for _, cell := range cells {
				result.WriteString(cellToString(cell.Value))
			}
		} else {
			// Single cell reference
			cell, err := ws.Cells().Get(arg)
			if err != nil {
				continue // skip missing cells
			}
			result.WriteString(cellToString(cell.Value))
		}
	}
	return result.String(), nil
}

func cellToString(v interface{}) string {
	if v == nil {
		return ""
	}
	switch val := v.(type) {
	case string:
		return val
	case float64, float32, int, int64:
		return fmt.Sprintf("%v", val)
	case bool:
		if val {
			return "TRUE"
		}
		return "FALSE"
	default:
		return fmt.Sprintf("%v", val)
	}
}

// ---------------------------------------------------------------------------
// IF function
// ---------------------------------------------------------------------------

func ifFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) < 2 || len(args) > 3 {
		return nil, fmt.Errorf("formula: IF requires 2 or 3 arguments")
	}

	condition := strings.TrimSpace(args[0])
	trueVal := strings.TrimSpace(args[1])
	falseVal := ""
	if len(args) == 3 {
		falseVal = strings.TrimSpace(args[2])
	}

	// Evaluate condition
	condResult, err := evaluateCondition(condition, ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	if condResult {
		return evaluateValue(trueVal, ws)
	}
	if falseVal != "" {
		return evaluateValue(falseVal, ws)
	}
	return false, nil
}

// evaluateCondition evaluates a condition expression like "A1>10" or "TRUE"
func evaluateCondition(expr string, ws *Worksheet) (bool, error) {
	expr = strings.TrimSpace(expr)

	// Check for comparison operators (order matters - check multi-char first)
	operators := []string{">=", "<=", "<>", "!=", ">", "<", "="}
	for _, op := range operators {
		idx := strings.Index(expr, op)
		if idx > 0 {
			left := strings.TrimSpace(expr[:idx])
			right := strings.TrimSpace(expr[idx+len(op):])

			leftVal, err := evaluateValue(left, ws)
			if err != nil {
				return false, err
			}
			rightVal, err := evaluateValue(right, ws)
			if err != nil {
				return false, err
			}

			return compareValues(leftVal, rightVal, op)
		}
	}

	// No operator - treat as boolean value
	val, err := evaluateValue(expr, ws)
	if err != nil {
		return false, err
	}

	switch v := val.(type) {
	case bool:
		return v, nil
	case float64:
		return v != 0, nil
	case string:
		return strings.ToUpper(v) == "TRUE", nil
	default:
		return false, fmt.Errorf("cannot convert %v to boolean", val)
	}
}

// evaluateValue evaluates a value expression (cell reference, number, or string)
func evaluateValue(expr string, ws *Worksheet) (interface{}, error) {
	expr = strings.TrimSpace(expr)

	// Check if it's a quoted string
	if len(expr) >= 2 && expr[0] == '"' && expr[len(expr)-1] == '"' {
		return expr[1 : len(expr)-1], nil
	}

	// Check if it's a boolean literal
	upper := strings.ToUpper(expr)
	if upper == "TRUE" {
		return true, nil
	}
	if upper == "FALSE" {
		return false, nil
	}

	// Check if it's a number
	if f, err := strconv.ParseFloat(expr, 64); err == nil {
		return f, nil
	}

	// Must be a cell reference (only if ws is provided)
	if ws == nil {
		return expr, nil // return as string if no worksheet
	}

	cell, err := ws.Cells().Get(expr)
	if err != nil {
		return nil, fmt.Errorf("cell %s not found", expr)
	}
	return cell.Value, nil
}

// compareValues compares two values using the given operator
func compareValues(left, right interface{}, op string) (bool, error) {
	// Try numeric comparison first
	leftNum, leftIsNum := toFloat64(left)
	rightNum, rightIsNum := toFloat64(right)

	if leftIsNum && rightIsNum {
		switch op {
		case ">":
			return leftNum > rightNum, nil
		case "<":
			return leftNum < rightNum, nil
		case ">=":
			return leftNum >= rightNum, nil
		case "<=":
			return leftNum <= rightNum, nil
		case "=", "==":
			return leftNum == rightNum, nil
		case "<>", "!=":
			return leftNum != rightNum, nil
		}
	}

	// String comparison
	leftStr := fmt.Sprintf("%v", left)
	rightStr := fmt.Sprintf("%v", right)

	switch op {
	case "=", "==":
		return leftStr == rightStr, nil
	case "<>", "!=":
		return leftStr != rightStr, nil
	case ">":
		return leftStr > rightStr, nil
	case "<":
		return leftStr < rightStr, nil
	case ">=":
		return leftStr >= rightStr, nil
	case "<=":
		return leftStr <= rightStr, nil
	}

	return false, fmt.Errorf("unknown operator: %s", op)
}

func toFloat64(v interface{}) (float64, bool) {
	switch val := v.(type) {
	case float64:
		return val, true
	case float32:
		return float64(val), true
	case int:
		return float64(val), true
	case int64:
		return float64(val), true
	case string:
		if f, err := strconv.ParseFloat(val, 64); err == nil {
			return f, true
		}
		return 0, false
	case bool:
		if val {
			return 1, true
		}
		return 0, true
	default:
		return 0, false
	}
}

// ---------------------------------------------------------------------------
// COUNTIF function
// ---------------------------------------------------------------------------

func countIfFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("formula: COUNTIF requires exactly 2 arguments")
	}

	rangeRef := strings.TrimSpace(args[0])
	criteria := strings.TrimSpace(args[1])

	// Strip quotes from criteria if present
	if len(criteria) >= 2 && criteria[0] == '"' && criteria[len(criteria)-1] == '"' {
		criteria = criteria[1 : len(criteria)-1]
	}

	// Get all cells in range
	cells, err := resolveRangeRaw(rangeRef, ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	count := 0
	for _, cell := range cells {
		if matchesCriteria(cell.Value, criteria) {
			count++
		}
	}

	return float64(count), nil
}

// matchesCriteria checks if a value matches the given criteria
func matchesCriteria(value interface{}, criteria string) bool {
	// Check if criteria is a comparison expression
	operators := []string{">=", "<=", "<>", "!=", "=", ">", "<"}
	for _, op := range operators {
		if strings.HasPrefix(criteria, op) {
			criteriaVal := strings.TrimSpace(criteria[len(op):])
			criteriaParsed, err := evaluateValue(criteriaVal, nil)
			if err != nil {
				return false
			}
			result, err := compareValues(value, criteriaParsed, op)
			if err != nil {
				return false
			}
			return result
		}
	}

	// Exact match
	criteriaParsed, err := evaluateValue(criteria, nil)
	if err != nil {
		return false
	}

	result, err := compareValues(value, criteriaParsed, "=")
	if err != nil {
		return false
	}
	return result
}

// ---------------------------------------------------------------------------
// VLOOKUP function
// ---------------------------------------------------------------------------

func vlookupFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) < 3 || len(args) > 4 {
		return nil, fmt.Errorf("formula: VLOOKUP requires 3 or 4 arguments")
	}

	lookupValue := strings.TrimSpace(args[0])
	rangeRef := strings.TrimSpace(args[1])
	colIndexStr := strings.TrimSpace(args[2])

	// Parse column index
	colIndex, err := strconv.Atoi(colIndexStr)
	if err != nil {
		return nil, fmt.Errorf("formula: invalid column index: %s", colIndexStr)
	}
	if colIndex < 1 {
		return nil, fmt.Errorf("formula: column index must be >= 1")
	}

	// Get lookup value
	lookupVal, err := evaluateValue(lookupValue, ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	// Parse range
	parts := strings.SplitN(rangeRef, ":", 2)
	if len(parts) != 2 {
		return nil, fmt.Errorf("formula: invalid range: %s", rangeRef)
	}

	startRef := strings.TrimSpace(parts[0])
	endRef := strings.TrimSpace(parts[1])

	startCol, startRow := splitRef(startRef)
	endCol, endRow := splitRef(endRef)

	startColNum := colToNum(startCol)
	endColNum := colToNum(endCol)

	if startColNum > endColNum {
		startColNum, endColNum = endColNum, startColNum
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}

	// Check if colIndex is within range
	if colIndex > (endColNum - startColNum + 1) {
		return nil, fmt.Errorf("formula: column index %d out of range", colIndex)
	}

	targetColNum := startColNum + colIndex - 1

	// Search in first column
	for row := startRow; row <= endRow; row++ {
		cellRef := numToCol(startColNum) + strconv.Itoa(row)
		cell, err := ws.Cells().Get(cellRef)
		if err != nil {
			continue
		}

		match, err := compareValues(cell.Value, lookupVal, "=")
		if err != nil {
			continue
		}

		if match {
			// Found match - return value from target column
			targetRef := numToCol(targetColNum) + strconv.Itoa(row)
			targetCell, err := ws.Cells().Get(targetRef)
			if err != nil {
				return nil, fmt.Errorf("formula: target cell %s not found", targetRef)
			}
			return targetCell.Value, nil
		}
	}

	return nil, fmt.Errorf("formula: value not found")
}

// resolveRangeRaw returns the raw Cell objects in a range (not just numeric values)
func resolveRangeRaw(rng string, ws *Worksheet) ([]*Cell, error) {
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

	var cells []*Cell
	for col := sCol; col <= eCol; col++ {
		for row := sr; row <= er; row++ {
			ref := numToCol(col) + strconv.Itoa(row)
			cell, err := ws.Cells().Get(ref)
			if err != nil {
				continue // skip missing cells
			}
			cells = append(cells, cell)
		}
	}
	return cells, nil
}
