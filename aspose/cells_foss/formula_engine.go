package cells_foss

import (
	"errors"
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode/utf8"
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
//   - ROUND(number, decimals)
//   - ABS(number)
//   - POWER(base, exponent)
//   - SQRT(number)
//   - LEN(text_or_cell)
//   - LEFT(text_or_cell, num_chars)
//   - RIGHT(text_or_cell, num_chars)
//   - MID(text_or_cell, start_num, num_chars)
//   - UPPER(text_or_cell)
//   - LOWER(text_or_cell)
//   - AND(condition1, condition2, ...)
//   - OR(condition1, condition2, ...)
//   - NOT(condition)
//
// References may be single cells ("A1"), ranges ("A1:A10" / "A1:C1"), or
// comma-separated combinations of both.  Numeric and string literals may be
// used anywhere a reference is accepted.  Non-numeric cells are silently
// ignored.  Empty ranges produce an error.
//
// Formulas that are not a single function call are evaluated as arithmetic
// expressions, e.g. "1+1", "A1*2", "SUM(A1:A3)+10", "(A1+A2)/2", "-B3".
// Supported operators: + - * / ^ (power) and & (string concatenation).
//
// References may be qualified with a sheet name ("Sheet2!A1",
// "'My Sheet'!B2:C10"); the sheet is looked up in the workbook the passed
// worksheet belongs to. A reference to a sheet that does not exist produces
// an error.
//
// Evaluation follows Excel's error-value semantics: dividing by zero yields
// the ErrorValue #DIV/0! (not a Go error), SQRT of a negative number
// yields #NUM!, arithmetic on text yields #VALUE!, and error values
// propagate through arithmetic and aggregation.
func CalculateFormula(formula string, ws *Worksheet) (interface{}, error) {
	v, err := calculateFormula(formula, ws)
	var fe ErrorValue
	if errors.As(err, &fe) {
		return fe, nil // an Excel error value is a result, not a failure
	}
	return v, err
}

func calculateFormula(formula string, ws *Worksheet) (interface{}, error) {
	formula = strings.TrimSpace(formula)
	if formula == "" {
		return nil, fmt.Errorf("formula: empty expression")
	}

	// A formula is either a single top-level function call or an arithmetic
	// expression (which may itself contain function calls as operands).
	funcName, argsText, isCall := splitTopLevelCall(formula)
	if !isCall {
		return evaluateValue(formula, ws)
	}

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
	case "ROUND":
		return roundFunc(args, ws)
	case "ABS":
		return absFunc(args, ws)
	case "POWER":
		return powerFunc(args, ws)
	case "SQRT":
		return sqrtFunc(args, ws)
	case "LEN":
		return lenFunc(args, ws)
	case "LEFT":
		return leftFunc(args, ws)
	case "RIGHT":
		return rightFunc(args, ws)
	case "MID":
		return midFunc(args, ws)
	case "UPPER":
		return upperFunc(args, ws)
	case "LOWER":
		return lowerFunc(args, ws)
	case "AND":
		return andFunc(args, ws)
	case "OR":
		return orFunc(args, ws)
	case "NOT":
		return notFunc(args, ws)
	}

	// For numeric aggregation functions, expand all arguments to float64
	// values. Arguments may be ranges, cell references, literals, arithmetic
	// expressions, or nested function calls; non-numeric values are skipped.
	var values []float64
	for _, arg := range args {
		arg = strings.TrimSpace(arg)
		if isRangeArg(arg) {
			vals, err := resolveRange(arg, ws)
			if err != nil {
				return nil, fmt.Errorf("formula: %w", err)
			}
			values = append(values, vals...)
			continue
		}
		v, err := evaluateValue(arg, ws)
		if err != nil {
			// For numeric functions, skip values that cannot be resolved.
			continue
		}
		if f, ok := toFloat64(v); ok {
			values = append(values, f)
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
	target, rest, err := resolveSheetRef(rng, ws)
	if err != nil {
		return nil, err
	}
	rng, ws = rest, target

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
			cell, err := ws.Cells().Get(ref)
			if err != nil {
				continue // skip empty cells
			}
			if fe, ok := cell.Value.(ErrorValue); ok {
				return nil, fe // error values propagate through aggregation
			}
			v, err := cellToFloat(cell)
			if err != nil {
				continue // skip non-numeric cells
			}
			values = append(values, v)
		}
	}
	return values, nil
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
// Argument splitting and call detection
// ---------------------------------------------------------------------------

// splitFormulaArgs splits on commas at parenthesis depth zero, ignoring
// commas inside double-quoted string literals ("" is an escaped quote).
func splitFormulaArgs(s string) []string {
	var args []string
	depth := 0
	start := 0
	inQuote := false
	for i := 0; i < len(s); i++ {
		ch := s[i]
		if inQuote {
			if ch == '"' {
				if i+1 < len(s) && s[i+1] == '"' {
					i++ // escaped quote
					continue
				}
				inQuote = false
			}
			continue
		}
		switch ch {
		case '"':
			inQuote = true
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

// splitTopLevelCall reports whether formula is a single function call whose
// parentheses span the whole expression, e.g. "SUM(A1,B2)". Expressions like
// "SUM(A1)+1" return ok=false so the caller can treat them as arithmetic.
func splitTopLevelCall(formula string) (name, argsText string, ok bool) {
	paren := strings.IndexByte(formula, '(')
	if paren <= 0 || !strings.HasSuffix(formula, ")") {
		return "", "", false
	}
	name = strings.ToUpper(strings.TrimSpace(formula[:paren]))
	if !isFuncName(name) {
		return "", "", false
	}
	// The '(' must match the final ')' for this to be a single call.
	depth := 0
	inQuote := false
	for i := paren; i < len(formula); i++ {
		ch := formula[i]
		if inQuote {
			if ch == '"' {
				if i+1 < len(formula) && formula[i+1] == '"' {
					i++
					continue
				}
				inQuote = false
			}
			continue
		}
		switch ch {
		case '"':
			inQuote = true
		case '(':
			depth++
		case ')':
			depth--
			if depth == 0 {
				if i == len(formula)-1 {
					return name, strings.TrimSpace(formula[paren+1 : i]), true
				}
				return "", "", false
			}
		}
	}
	return "", "", false
}

// isFuncName reports whether s looks like a spreadsheet function name
// (letters, digits, '_' or '.', starting with a letter or '_').
func isFuncName(s string) bool {
	if s == "" {
		return false
	}
	for i := 0; i < len(s); i++ {
		c := s[i]
		switch {
		case c >= 'A' && c <= 'Z', c == '_':
		case c >= '0' && c <= '9', c == '.':
			if i == 0 {
				return false
			}
		default:
			return false
		}
	}
	return true
}

// isRangeArg reports whether arg is a plain range reference like "A1:C10",
// optionally sheet-qualified ("Sheet2!A1:C10") — no double-quoted text, no
// nested calls, no arithmetic.
func isRangeArg(arg string) bool {
	if _, rest, has := stripSheetPrefix(arg); has {
		arg = rest
	}
	if !strings.Contains(arg, ":") || strings.ContainsAny(arg, `"()`) {
		return false
	}
	parts := strings.SplitN(arg, ":", 2)
	return isCellRefPattern(strings.TrimSpace(parts[0])) && isCellRefPattern(strings.TrimSpace(parts[1]))
}

// ---------------------------------------------------------------------------
// Sheet-qualified references
// ---------------------------------------------------------------------------

// workbook returns the Workbook the worksheet belongs to, or nil when it was
// created outside of a workbook context.
func (ws *Worksheet) workbook() *Workbook {
	if ws == nil || ws.cells == nil {
		return nil
	}
	return ws.cells.wb
}

// stripSheetPrefix splits a sheet-qualified reference ("Sheet2!A1",
// "'My Sheet'!B2:C10") into the sheet name and the reference that follows
// the "!". Inside a quoted name, ” is an escaped apostrophe. has is false
// when arg carries no sheet prefix.
func stripSheetPrefix(arg string) (sheet, rest string, has bool) {
	arg = strings.TrimSpace(arg)
	if strings.HasPrefix(arg, "'") {
		var b strings.Builder
		i := 1
		for i < len(arg) {
			if arg[i] == '\'' {
				if i+1 < len(arg) && arg[i+1] == '\'' {
					b.WriteByte('\'')
					i += 2
					continue
				}
				break
			}
			b.WriteByte(arg[i])
			i++
		}
		if i >= len(arg) {
			return "", "", false // unterminated quote
		}
		rest = arg[i+1:]
		if !strings.HasPrefix(rest, "!") {
			return "", "", false
		}
		return b.String(), strings.TrimSpace(rest[1:]), true
	}
	if idx := strings.IndexByte(arg, '!'); idx > 0 {
		return arg[:idx], strings.TrimSpace(arg[idx+1:]), true
	}
	return "", "", false
}

// lookupSheet finds a worksheet by name (case-insensitive, matching Excel)
// in the workbook ws belongs to.
func lookupSheet(name string, ws *Worksheet) (*Worksheet, error) {
	wb := ws.workbook()
	if wb == nil {
		return nil, fmt.Errorf("formula: cross-sheet reference to %q requires a worksheet attached to a Workbook", name)
	}
	for _, s := range wb.Worksheets {
		if strings.EqualFold(s.Name, name) {
			return s, nil
		}
	}
	return nil, &SheetNotFoundError{Name: name}
}

// resolveSheetRef maps a possibly sheet-qualified reference to the worksheet
// it points into plus the unqualified remainder. Without a "!" prefix the
// input worksheet and arg are returned unchanged.
func resolveSheetRef(arg string, ws *Worksheet) (*Worksheet, string, error) {
	sheet, rest, has := stripSheetPrefix(arg)
	if !has {
		return ws, arg, nil
	}
	target, err := lookupSheet(sheet, ws)
	if err != nil {
		return nil, "", err
	}
	return target, rest, nil
}

// isCellRefPattern reports whether s is an A1-style reference.
func isCellRefPattern(s string) bool {
	i := 0
	for i < len(s) && i < 3 && ((s[i] >= 'A' && s[i] <= 'Z') || (s[i] >= 'a' && s[i] <= 'z')) {
		i++
	}
	if i == 0 || i == len(s) {
		return false
	}
	for ; i < len(s); i++ {
		if s[i] < '0' || s[i] > '9' {
			return false
		}
	}
	return true
}

// ---------------------------------------------------------------------------
// CONCAT function
// ---------------------------------------------------------------------------

func concatFunc(args []string, ws *Worksheet) (interface{}, error) {
	var result strings.Builder
	for _, arg := range args {
		arg = strings.TrimSpace(arg)
		if isRangeArg(arg) {
			// Range reference - get all cells
			cells, err := resolveRangeRaw(arg, ws)
			if err != nil {
				return nil, fmt.Errorf("formula: %w", err)
			}
			for _, cell := range cells {
				result.WriteString(cellToString(cell.Value))
			}
			continue
		}
		// Cell reference, literal, expression, or nested call.
		v, err := evaluateValue(arg, ws)
		if err != nil {
			continue // skip missing cells
		}
		result.WriteString(cellToString(v))
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

// evaluateValue evaluates a value expression: a cell reference, number,
// quoted string, boolean literal, function call, or arithmetic expression.
func evaluateValue(expr string, ws *Worksheet) (interface{}, error) {
	expr = strings.TrimSpace(expr)

	// Check if it's a single quoted string literal ("" is an escaped quote).
	// Expressions like "a"&"b" also start and end with quotes but are not
	// literals, so verify the first literal spans the whole expression.
	if len(expr) >= 2 && expr[0] == '"' && isSingleQuotedLiteral(expr) {
		return unquoteFormulaString(expr), nil
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

	// Without a worksheet only literals can be resolved.
	if ws == nil {
		return expr, nil // return as string if no worksheet
	}

	// Operators or parens mean an arithmetic expression or function call.
	if containsExpressionSyntax(expr) {
		p := &exprParser{src: expr, ws: ws}
		if v, err := p.parseExpression(); err == nil {
			p.skipSpaces()
			if p.pos >= len(p.src) {
				return v, nil
			}
		}
		// Not a valid expression; fall through to report a missing cell.
	}

	// Sheet-qualified cell reference ("Sheet2!A1", "'My Sheet'!B2").
	if sheet, rest, has := stripSheetPrefix(expr); has {
		target, err := lookupSheet(sheet, ws)
		if err != nil {
			return nil, err
		}
		return evaluateValue(rest, target)
	}

	cell, err := ws.Cells().Get(expr)
	if err != nil {
		return nil, fmt.Errorf("cell %s not found", expr)
	}
	return cell.Value, nil
}

// isSingleQuotedLiteral reports whether s (starting with '"') is exactly one
// quoted string literal, honoring "" escapes.
func isSingleQuotedLiteral(s string) bool {
	for i := 1; i < len(s); i++ {
		if s[i] == '"' {
			if i+1 < len(s) && s[i+1] == '"' {
				i++
				continue
			}
			return i == len(s)-1
		}
	}
	return false
}

// unquoteFormulaString strips the surrounding quotes of a formula string
// literal and unescapes doubled quotes.
func unquoteFormulaString(s string) string {
	inner := s[1 : len(s)-1]
	if strings.Contains(inner, `""`) {
		return strings.ReplaceAll(inner, `""`, `"`)
	}
	return inner
}

// containsExpressionSyntax reports whether s contains an operator or paren
// outside a quoted string, i.e. whether it could be an arithmetic expression.
func containsExpressionSyntax(s string) bool {
	inQuote := false
	for i := 0; i < len(s); i++ {
		c := s[i]
		if inQuote {
			if c == '"' {
				if i+1 < len(s) && s[i+1] == '"' {
					i++
					continue
				}
				inQuote = false
			}
			continue
		}
		switch c {
		case '"':
			inQuote = true
		case '+', '-', '*', '/', '^', '&', '(':
			return true
		}
	}
	return false
}

// ---------------------------------------------------------------------------
// Arithmetic expression parser (recursive descent)
//
// Grammar:
//   expression = term (('+'|'-'|'&') term)*
//   term       = power (('*'|'/') power)*
//   power      = unary ('^' power)?          // right-associative
//   unary      = ('-'|'+') unary | primary
//   primary    = '(' expression ')' | STRING | NUMBER | TRUE | FALSE
//              | FUNCNAME '(' args ')' | CELLREF
// ---------------------------------------------------------------------------

type exprParser struct {
	src string
	pos int
	ws  *Worksheet
}

func (p *exprParser) skipSpaces() {
	for p.pos < len(p.src) && (p.src[p.pos] == ' ' || p.src[p.pos] == '\t') {
		p.pos++
	}
}

func (p *exprParser) parseExpression() (interface{}, error) {
	left, err := p.parseTerm()
	if err != nil {
		return nil, err
	}
	for {
		p.skipSpaces()
		if p.pos >= len(p.src) {
			return left, nil
		}
		op := p.src[p.pos]
		if op != '+' && op != '-' && op != '&' {
			return left, nil
		}
		p.pos++
		right, err := p.parseTerm()
		if err != nil {
			return nil, err
		}
		if fe, ok := left.(ErrorValue); ok {
			return fe, nil // error values propagate through every operator
		}
		if fe, ok := right.(ErrorValue); ok {
			return fe, nil
		}
		if op == '&' {
			left = cellToString(left) + cellToString(right)
			continue
		}
		lf, lok := toFloat64(left)
		rf, rok := toFloat64(right)
		if !lok || !rok {
			return ErrValue, nil // Excel: arithmetic on text is #VALUE!
		}
		if op == '+' {
			left = lf + rf
		} else {
			left = lf - rf
		}
	}
}

func (p *exprParser) parseTerm() (interface{}, error) {
	left, err := p.parsePower()
	if err != nil {
		return nil, err
	}
	for {
		p.skipSpaces()
		if p.pos >= len(p.src) {
			return left, nil
		}
		op := p.src[p.pos]
		if op != '*' && op != '/' {
			return left, nil
		}
		p.pos++
		right, err := p.parsePower()
		if err != nil {
			return nil, err
		}
		if fe, ok := left.(ErrorValue); ok {
			return fe, nil
		}
		if fe, ok := right.(ErrorValue); ok {
			return fe, nil
		}
		lf, lok := toFloat64(left)
		rf, rok := toFloat64(right)
		if !lok || !rok {
			return ErrValue, nil
		}
		if op == '*' {
			left = lf * rf
		} else {
			if rf == 0 {
				left = ErrDiv0 // Excel: division by zero is the #DIV/0! value
				continue
			}
			left = lf / rf
		}
	}
}

func (p *exprParser) parsePower() (interface{}, error) {
	left, err := p.parseUnary()
	if err != nil {
		return nil, err
	}
	p.skipSpaces()
	if p.pos < len(p.src) && p.src[p.pos] == '^' {
		p.pos++
		right, err := p.parsePower() // right-associative
		if err != nil {
			return nil, err
		}
		if fe, ok := left.(ErrorValue); ok {
			return fe, nil
		}
		if fe, ok := right.(ErrorValue); ok {
			return fe, nil
		}
		lf, lok := toFloat64(left)
		rf, rok := toFloat64(right)
		if !lok || !rok {
			return ErrValue, nil
		}
		return numericResult(math.Pow(lf, rf)), nil
	}
	if fe, ok := left.(ErrorValue); ok {
		return fe, nil
	}
	return left, nil
}

// numericResult converts a computed float to an Excel value: NaN and ±Inf
// results are the #NUM! error value, everything else passes through.
func numericResult(f float64) interface{} {
	if math.IsNaN(f) || math.IsInf(f, 0) {
		return ErrNum
	}
	return f
}

func (p *exprParser) parseUnary() (interface{}, error) {
	p.skipSpaces()
	if p.pos < len(p.src) {
		if p.src[p.pos] == '-' {
			p.pos++
			v, err := p.parseUnary()
			if err != nil {
				return nil, err
			}
			if fe, ok := v.(ErrorValue); ok {
				return fe, nil
			}
			f, ok := toFloat64(v)
			if !ok {
				return ErrValue, nil
			}
			return -f, nil
		}
		if p.src[p.pos] == '+' {
			p.pos++
			return p.parseUnary()
		}
	}
	return p.parsePrimary()
}

func (p *exprParser) parsePrimary() (interface{}, error) {
	p.skipSpaces()
	if p.pos >= len(p.src) {
		return nil, fmt.Errorf("formula: unexpected end of expression")
	}
	ch := p.src[p.pos]
	switch {
	case ch == '(':
		p.pos++
		v, err := p.parseExpression()
		if err != nil {
			return nil, err
		}
		p.skipSpaces()
		if p.pos >= len(p.src) || p.src[p.pos] != ')' {
			return nil, fmt.Errorf("formula: missing closing ')'")
		}
		p.pos++
		return v, nil
	case ch == '"':
		return p.parseQuotedString()
	case ch == '\'':
		return p.parseQuotedSheetRef()
	case isASCIIDigit(ch) || ch == '.':
		return p.parseNumber()
	case isASCIILetter(ch) || ch == '_':
		return p.parseIdentifier()
	}
	return nil, fmt.Errorf("formula: unexpected character %q", string(ch))
}

// parseQuotedSheetRef reads a quoted sheet name ("'My Sheet'!A1") at the
// current position and resolves the reference that follows it.
func (p *exprParser) parseQuotedSheetRef() (interface{}, error) {
	var b strings.Builder
	i := p.pos + 1
	for i < len(p.src) {
		if p.src[i] == '\'' {
			if i+1 < len(p.src) && p.src[i+1] == '\'' {
				b.WriteByte('\'')
				i += 2
				continue
			}
			break
		}
		b.WriteByte(p.src[i])
		i++
	}
	if i >= len(p.src) {
		return nil, fmt.Errorf("formula: unterminated quoted sheet name")
	}
	if i+1 >= len(p.src) || p.src[i+1] != '!' {
		return nil, fmt.Errorf("formula: expected ! after quoted sheet name %q", b.String())
	}
	p.pos = i + 2
	return p.parseSheetQualifiedRef(b.String())
}

func (p *exprParser) parseQuotedString() (interface{}, error) {
	var b strings.Builder
	p.pos++ // opening quote
	for p.pos < len(p.src) {
		c := p.src[p.pos]
		if c == '"' {
			if p.pos+1 < len(p.src) && p.src[p.pos+1] == '"' {
				b.WriteByte('"')
				p.pos += 2
				continue
			}
			p.pos++
			return b.String(), nil
		}
		b.WriteByte(c)
		p.pos++
	}
	return nil, fmt.Errorf("formula: unterminated string literal")
}

func (p *exprParser) parseNumber() (interface{}, error) {
	start := p.pos
	for p.pos < len(p.src) && (isASCIIDigit(p.src[p.pos]) || p.src[p.pos] == '.') {
		p.pos++
	}
	// Optional exponent, e.g. 1e5 or 2E-3.
	if p.pos < len(p.src) && (p.src[p.pos] == 'e' || p.src[p.pos] == 'E') {
		save := p.pos
		p.pos++
		if p.pos < len(p.src) && (p.src[p.pos] == '+' || p.src[p.pos] == '-') {
			p.pos++
		}
		if p.pos < len(p.src) && isASCIIDigit(p.src[p.pos]) {
			for p.pos < len(p.src) && isASCIIDigit(p.src[p.pos]) {
				p.pos++
			}
		} else {
			p.pos = save // not an exponent after all
		}
	}
	tok := p.src[start:p.pos]
	f, err := strconv.ParseFloat(tok, 64)
	if err != nil {
		return nil, fmt.Errorf("formula: invalid number %q", tok)
	}
	return f, nil
}

func (p *exprParser) parseIdentifier() (interface{}, error) {
	start := p.pos
	for p.pos < len(p.src) {
		c := p.src[p.pos]
		if isASCIILetter(c) || isASCIIDigit(c) || c == '_' || c == '.' {
			p.pos++
			continue
		}
		break
	}
	name := p.src[start:p.pos]

	// Function call: NAME followed by a balanced parenthesized group.
	p.skipSpaces()
	if p.pos < len(p.src) && p.src[p.pos] == '(' {
		end, ok := p.matchCallEnd(p.pos)
		if !ok {
			return nil, fmt.Errorf("formula: unbalanced parentheses in %s(", name)
		}
		call := p.src[start:end]
		p.pos = end
		return CalculateFormula(call, p.ws)
	}

	switch strings.ToUpper(name) {
	case "TRUE":
		return true, nil
	case "FALSE":
		return false, nil
	}

	// Sheet-qualified reference: Name!ref.
	if p.pos < len(p.src) && p.src[p.pos] == '!' {
		p.pos++
		return p.parseSheetQualifiedRef(name)
	}

	// A range in arithmetic context has no numeric value.
	p.skipSpaces()
	if p.pos < len(p.src) && p.src[p.pos] == ':' {
		return nil, fmt.Errorf("formula: range starting at %s is not supported in this context", name)
	}

	cell, err := p.ws.Cells().Get(name)
	if err != nil {
		return nil, fmt.Errorf("cell %s not found", name)
	}
	return cell.Value, nil
}

// parseSheetQualifiedRef reads the cell reference following "SheetName!" (or
// "'Sheet Name'!") at the current position and resolves it against that
// sheet. Ranges are rejected — aggregation functions take them through
// resolveRange instead of the expression parser.
func (p *exprParser) parseSheetQualifiedRef(sheet string) (interface{}, error) {
	start := p.pos
	for p.pos < len(p.src) && (isASCIILetter(p.src[p.pos]) || isASCIIDigit(p.src[p.pos])) {
		p.pos++
	}
	ref := p.src[start:p.pos]
	if !isCellRefPattern(ref) {
		return nil, fmt.Errorf("formula: invalid cell reference %q after sheet %q", ref, sheet)
	}
	p.skipSpaces()
	if p.pos < len(p.src) && p.src[p.pos] == ':' {
		return nil, fmt.Errorf("formula: range starting at %s!%s is not supported in this context", sheet, ref)
	}
	target, err := lookupSheet(sheet, p.ws)
	if err != nil {
		return nil, err
	}
	cell, err := target.Cells().Get(ref)
	if err != nil {
		return nil, fmt.Errorf("cell %s!%s not found", sheet, ref)
	}
	return cell.Value, nil
}

// matchCallEnd finds the index just past the ')' matching the '(' at open.
func (p *exprParser) matchCallEnd(open int) (int, bool) {
	depth := 0
	inQuote := false
	for i := open; i < len(p.src); i++ {
		c := p.src[i]
		if inQuote {
			if c == '"' {
				if i+1 < len(p.src) && p.src[i+1] == '"' {
					i++
					continue
				}
				inQuote = false
			}
			continue
		}
		switch c {
		case '"':
			inQuote = true
		case '(':
			depth++
		case ')':
			depth--
			if depth == 0 {
				return i + 1, true
			}
		}
	}
	return 0, false
}

func isASCIIDigit(c byte) bool  { return c >= '0' && c <= '9' }
func isASCIILetter(c byte) bool { return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') }

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
	target, rest, err := resolveSheetRef(rng, ws)
	if err != nil {
		return nil, err
	}
	rng, ws = rest, target

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

// ---------------------------------------------------------------------------
// ROUND function
// ---------------------------------------------------------------------------

func roundFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("formula: ROUND requires exactly 2 arguments")
	}

	numVal, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	decVal, err := evaluateValue(strings.TrimSpace(args[1]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	num, ok := toFloat64(numVal)
	if !ok {
		return nil, fmt.Errorf("formula: first argument must be numeric")
	}

	dec, ok := toFloat64(decVal)
	if !ok {
		return nil, fmt.Errorf("formula: second argument must be numeric")
	}

	multiplier := math.Pow(10, dec)
	return math.Round(num*multiplier) / multiplier, nil
}

// ---------------------------------------------------------------------------
// ABS function
// ---------------------------------------------------------------------------

func absFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("formula: ABS requires exactly 1 argument")
	}

	val, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	num, ok := toFloat64(val)
	if !ok {
		return nil, fmt.Errorf("formula: argument must be numeric")
	}

	return math.Abs(num), nil
}

// ---------------------------------------------------------------------------
// POWER function
// ---------------------------------------------------------------------------

func powerFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("formula: POWER requires exactly 2 arguments")
	}

	baseVal, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	expVal, err := evaluateValue(strings.TrimSpace(args[1]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	if fe, ok := baseVal.(ErrorValue); ok {
		return fe, nil
	}
	if fe, ok := expVal.(ErrorValue); ok {
		return fe, nil
	}
	base, ok := toFloat64(baseVal)
	if !ok {
		return nil, fmt.Errorf("formula: first argument must be numeric")
	}

	exp, ok := toFloat64(expVal)
	if !ok {
		return nil, fmt.Errorf("formula: second argument must be numeric")
	}

	return numericResult(math.Pow(base, exp)), nil
}

// ---------------------------------------------------------------------------
// SQRT function
// ---------------------------------------------------------------------------

func sqrtFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("formula: SQRT requires exactly 1 argument")
	}

	val, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	if fe, ok := val.(ErrorValue); ok {
		return fe, nil
	}
	num, ok := toFloat64(val)
	if !ok {
		return nil, fmt.Errorf("formula: argument must be numeric")
	}

	if num < 0 {
		return ErrNum, nil // Excel: SQRT of a negative number is #NUM!
	}

	return math.Sqrt(num), nil
}

// ---------------------------------------------------------------------------
// LEN function
// ---------------------------------------------------------------------------

func lenFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("formula: LEN requires exactly 1 argument")
	}

	val, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	text := fmt.Sprintf("%v", val)
	return float64(utf8.RuneCountInString(text)), nil
}

// ---------------------------------------------------------------------------
// LEFT function
// ---------------------------------------------------------------------------

func leftFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("formula: LEFT requires exactly 2 arguments")
	}

	val, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	numCharsVal, err := evaluateValue(strings.TrimSpace(args[1]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	text := fmt.Sprintf("%v", val)
	numChars, ok := toFloat64(numCharsVal)
	if !ok {
		return nil, fmt.Errorf("formula: second argument must be numeric")
	}

	runes := []rune(text)
	n := int(numChars)
	if n < 0 {
		return nil, fmt.Errorf("formula: number of characters cannot be negative")
	}
	if n > len(runes) {
		n = len(runes)
	}

	return string(runes[:n]), nil
}

// ---------------------------------------------------------------------------
// RIGHT function
// ---------------------------------------------------------------------------

func rightFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("formula: RIGHT requires exactly 2 arguments")
	}

	val, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	numCharsVal, err := evaluateValue(strings.TrimSpace(args[1]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	text := fmt.Sprintf("%v", val)
	numChars, ok := toFloat64(numCharsVal)
	if !ok {
		return nil, fmt.Errorf("formula: second argument must be numeric")
	}

	runes := []rune(text)
	n := int(numChars)
	if n < 0 {
		return nil, fmt.Errorf("formula: number of characters cannot be negative")
	}
	if n > len(runes) {
		n = len(runes)
	}

	return string(runes[len(runes)-n:]), nil
}

// ---------------------------------------------------------------------------
// MID function
// ---------------------------------------------------------------------------

func midFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 3 {
		return nil, fmt.Errorf("formula: MID requires exactly 3 arguments")
	}

	val, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	startVal, err := evaluateValue(strings.TrimSpace(args[1]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	numCharsVal, err := evaluateValue(strings.TrimSpace(args[2]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	text := fmt.Sprintf("%v", val)
	startPos, ok := toFloat64(startVal)
	if !ok {
		return nil, fmt.Errorf("formula: second argument must be numeric")
	}
	numChars, ok := toFloat64(numCharsVal)
	if !ok {
		return nil, fmt.Errorf("formula: third argument must be numeric")
	}

	start := int(startPos) // 1-based
	n := int(numChars)

	if start < 1 {
		return nil, fmt.Errorf("formula: start position must be >= 1")
	}
	if n < 0 {
		return nil, fmt.Errorf("formula: number of characters cannot be negative")
	}

	runes := []rune(text)
	if start-1 >= len(runes) {
		return "", nil
	}
	end := start - 1 + n
	if end > len(runes) {
		end = len(runes)
	}

	return string(runes[start-1 : end]), nil
}

// ---------------------------------------------------------------------------
// UPPER function
// ---------------------------------------------------------------------------

func upperFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("formula: UPPER requires exactly 1 argument")
	}

	val, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	return strings.ToUpper(fmt.Sprintf("%v", val)), nil
}

// ---------------------------------------------------------------------------
// LOWER function
// ---------------------------------------------------------------------------

func lowerFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("formula: LOWER requires exactly 1 argument")
	}

	val, err := evaluateValue(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}

	return strings.ToLower(fmt.Sprintf("%v", val)), nil
}

// ---------------------------------------------------------------------------
// AND function
// ---------------------------------------------------------------------------

func andFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) < 1 {
		return nil, fmt.Errorf("formula: AND requires at least 1 argument")
	}

	for _, arg := range args {
		result, err := evaluateCondition(strings.TrimSpace(arg), ws)
		if err != nil {
			return nil, fmt.Errorf("formula: %w", err)
		}
		if !result {
			return false, nil
		}
	}
	return true, nil
}

// ---------------------------------------------------------------------------
// OR function
// ---------------------------------------------------------------------------

func orFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) < 1 {
		return nil, fmt.Errorf("formula: OR requires at least 1 argument")
	}

	for _, arg := range args {
		result, err := evaluateCondition(strings.TrimSpace(arg), ws)
		if err != nil {
			return nil, fmt.Errorf("formula: %w", err)
		}
		if result {
			return true, nil
		}
	}
	return false, nil
}

// ---------------------------------------------------------------------------
// NOT function
// ---------------------------------------------------------------------------

func notFunc(args []string, ws *Worksheet) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("formula: NOT requires exactly 1 argument")
	}

	result, err := evaluateCondition(strings.TrimSpace(args[0]), ws)
	if err != nil {
		return nil, fmt.Errorf("formula: %w", err)
	}
	return !result, nil
}
