package cells_foss

import "fmt"

// DataValidationType enumerates the kinds of data validation rules supported
// by Excel. The string values match the OOXML ST_DataValidationType names.
const (
	DataValidationTypeNone       = "none"
	DataValidationTypeWhole      = "whole"
	DataValidationTypeDecimal    = "decimal"
	DataValidationTypeList       = "list"
	DataValidationTypeDate       = "date"
	DataValidationTypeTime       = "time"
	DataValidationTypeTextLength = "textLength"
	DataValidationTypeCustom     = "custom"
)

// Error style constants for data validation alerts.
const (
	ErrorStyleStop        = "stop"
	ErrorStyleWarning     = "warning"
	ErrorStyleInformation = "information"
)

// DataValidation represents a single data-validation rule applied to a range
// of cells on a worksheet.
type DataValidation struct {
	// Type is the validation type (e.g. "list", "whole", "decimal").
	Type string

	// Ref is the A1-style range (or space-separated list of ranges) this
	// validation rule applies to, e.g. "A1:A10".
	Ref string

	// Formula1 is the primary formula or value used by the validation rule.
	// For list validations this can be a range reference ("$A$1:$A$10") or
	// a double-quoted comma-separated list (`"Yes,No"`).
	Formula1 string

	// Formula2 is the secondary formula (used for "between" / "notBetween").
	Formula2 string

	// AllowBlank controls whether empty cells pass validation.
	AllowBlank bool

	// ShowErrorMessage controls whether an error alert is shown on invalid input.
	ShowErrorMessage bool

	// ErrorTitle is the caption of the error dialog.
	ErrorTitle string

	// ErrorMessage is the body text of the error dialog.
	ErrorMessage string

	// ErrorStyle controls the severity of the validation alert:
	// "stop", "warning", or "information".
	ErrorStyle string
}

// AddDataValidation appends a data-validation rule to the worksheet.  The
// ref parameter (e.g. "A1:A10") is stored on the DataValidation so it can
// be serialised as the sqref attribute.
//
//	dv := &DataValidation{
//	    Type:     cells_foss.DataValidationTypeList,
//	    Formula1: `"Yes,No"`,
//	}
//	ws.AddDataValidation("A2:A10", dv)
func (ws *Worksheet) AddDataValidation(ref string, dv *DataValidation) error {
	if dv == nil {
		return fmt.Errorf("cells_foss: DataValidation is nil")
	}
	if ref == "" {
		return fmt.Errorf("cells_foss: DataValidation ref is empty")
	}
	dv.Ref = ref
	ws.DataValidations = append(ws.DataValidations, dv)
	if ws.cells != nil && ws.cells.wb != nil {
		ws.cells.wb.Modified = true
	}
	return nil
}

// RemoveDataValidation removes the first data-validation rule whose Ref
// exactly matches the given ref string.  An error is returned when no
// matching rule is found.
func (ws *Worksheet) RemoveDataValidation(ref string) error {
	for i, dv := range ws.DataValidations {
		if dv.Ref == ref {
			ws.DataValidations = append(ws.DataValidations[:i], ws.DataValidations[i+1:]...)
			if ws.cells != nil && ws.cells.wb != nil {
				ws.cells.wb.Modified = true
			}
			return nil
		}
	}
	return fmt.Errorf("cells_foss: no DataValidation found for ref %q", ref)
}
