package cells_foss_test

import (
	"testing"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// TestFormula_ROUND tests the ROUND function
func TestFormula_ROUND(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", 3.14159)

	result, err := cells_foss.CalculateFormula("ROUND(A1, 2)", ws)
	if err != nil {
		t.Fatalf("ROUND failed: %v", err)
	}
	if result.(float64) != 3.14 {
		t.Errorf("Expected 3.14, got %v", result)
	}
}

// TestFormula_ABS tests the ABS function
func TestFormula_ABS(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", -5.5)

	result, err := cells_foss.CalculateFormula("ABS(A1)", ws)
	if err != nil {
		t.Fatalf("ABS failed: %v", err)
	}
	if result.(float64) != 5.5 {
		t.Errorf("Expected 5.5, got %v", result)
	}
}

// TestFormula_POWER tests the POWER function
func TestFormula_POWER(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", 2)
	ws.Cells().Set("A2", 3)

	result, err := cells_foss.CalculateFormula("POWER(A1, A2)", ws)
	if err != nil {
		t.Fatalf("POWER failed: %v", err)
	}
	if result.(float64) != 8 {
		t.Errorf("Expected 8, got %v", result)
	}
}

// TestFormula_SQRT tests the SQRT function
func TestFormula_SQRT(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", 16)

	result, err := cells_foss.CalculateFormula("SQRT(A1)", ws)
	if err != nil {
		t.Fatalf("SQRT failed: %v", err)
	}
	if result.(float64) != 4 {
		t.Errorf("Expected 4, got %v", result)
	}
}

// TestFormula_LEN tests the LEN function
func TestFormula_LEN(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Hello")

	result, err := cells_foss.CalculateFormula("LEN(A1)", ws)
	if err != nil {
		t.Fatalf("LEN failed: %v", err)
	}
	if result.(float64) != 5 {
		t.Errorf("Expected 5, got %v", result)
	}
}

// TestFormula_LEFT tests the LEFT function
func TestFormula_LEFT(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Hello World")

	result, err := cells_foss.CalculateFormula("LEFT(A1, 5)", ws)
	if err != nil {
		t.Fatalf("LEFT failed: %v", err)
	}
	if result.(string) != "Hello" {
		t.Errorf("Expected 'Hello', got %v", result)
	}
}

// TestFormula_RIGHT tests the RIGHT function
func TestFormula_RIGHT(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Hello World")

	result, err := cells_foss.CalculateFormula("RIGHT(A1, 5)", ws)
	if err != nil {
		t.Fatalf("RIGHT failed: %v", err)
	}
	if result.(string) != "World" {
		t.Errorf("Expected 'World', got %v", result)
	}
}

// TestFormula_MID tests the MID function
func TestFormula_MID(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Hello World")

	result, err := cells_foss.CalculateFormula("MID(A1, 7, 5)", ws)
	if err != nil {
		t.Fatalf("MID failed: %v", err)
	}
	if result.(string) != "World" {
		t.Errorf("Expected 'World', got %v", result)
	}
}

// TestFormula_UPPER tests the UPPER function
func TestFormula_UPPER(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "hello")

	result, err := cells_foss.CalculateFormula("UPPER(A1)", ws)
	if err != nil {
		t.Fatalf("UPPER failed: %v", err)
	}
	if result.(string) != "HELLO" {
		t.Errorf("Expected 'HELLO', got %v", result)
	}
}

// TestFormula_LOWER tests the LOWER function
func TestFormula_LOWER(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "HELLO")

	result, err := cells_foss.CalculateFormula("LOWER(A1)", ws)
	if err != nil {
		t.Fatalf("LOWER failed: %v", err)
	}
	if result.(string) != "hello" {
		t.Errorf("Expected 'hello', got %v", result)
	}
}

// TestFormula_AND tests the AND function
func TestFormula_AND(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", 10)
	ws.Cells().Set("A2", 20)

	result, err := cells_foss.CalculateFormula("AND(A1>5, A2>15)", ws)
	if err != nil {
		t.Fatalf("AND failed: %v", err)
	}
	if result.(bool) != true {
		t.Errorf("Expected true, got %v", result)
	}
}

// TestFormula_OR tests the OR function
func TestFormula_OR(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", 10)
	ws.Cells().Set("A2", 5)

	result, err := cells_foss.CalculateFormula("OR(A1>15, A2>3)", ws)
	if err != nil {
		t.Fatalf("OR failed: %v", err)
	}
	if result.(bool) != true {
		t.Errorf("Expected true, got %v", result)
	}
}

// TestFormula_NOT tests the NOT function
func TestFormula_NOT(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", 10)

	result, err := cells_foss.CalculateFormula("NOT(A1>15)", ws)
	if err != nil {
		t.Fatalf("NOT failed: %v", err)
	}
	if result.(bool) != true {
		t.Errorf("Expected true, got %v", result)
	}
}
