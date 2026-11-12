package cells_foss

// ErrorValue represents an Excel error value such as #DIV/0!. Error values
// are produced by formula evaluation (e.g. division by zero) and can be
// stored in cells like any other value:
//
//	v, err := cells_foss.CalculateFormula("1/0", ws) // v == cells_foss.ErrDiv0, err == nil
//	ws.Cells().Set("A1", v)                           // saved as <c t="e"><v>#DIV/0!</v></c>
//
// A cell holding an error value reports it through AsString; AsFloat64 and
// AsInt return ok=false. In formulas, error values propagate: an operation
// with an error operand yields that error, matching Excel.
//
// (Not to be confused with the FormulaError struct, which reports Go-level
// formula parse/evaluation failures.)
type ErrorValue string

// The seven error values defined by Excel.
const (
	ErrDiv0  ErrorValue = "#DIV/0!" // division by zero
	ErrNA    ErrorValue = "#N/A"    // value not available
	ErrName  ErrorValue = "#NAME?"  // unrecognized name
	ErrNull  ErrorValue = "#NULL!"  // empty intersection
	ErrNum   ErrorValue = "#NUM!"   // invalid numeric result
	ErrRef   ErrorValue = "#REF!"   // invalid cell reference
	ErrValue ErrorValue = "#VALUE!" // wrong type of argument
)

// Error implements the error interface so an ErrorValue can travel through
// error-returning evaluation paths; errors.As recovers it as a value.
func (e ErrorValue) Error() string { return string(e) }

// IsErrorValue reports whether s is one of Excel's seven error codes.
func IsErrorValue(s string) bool {
	switch ErrorValue(s) {
	case ErrDiv0, ErrNA, ErrName, ErrNull, ErrNum, ErrRef, ErrValue:
		return true
	}
	return false
}
