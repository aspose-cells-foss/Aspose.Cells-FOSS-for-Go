package cells_foss_test

import (
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestVBAProject_SetAndGet(t *testing.T) {
	wb := cells_foss.NewWorkbook()

	// Initially no VBA project.
	if wb.HasVBAProject() {
		t.Error("new workbook should not have VBA project")
	}

	// Set VBA project.
	data := []byte{0x01, 0x02, 0x03, 0x04} // Dummy binary data
	if err := wb.SetVBAProject(data); err != nil {
		t.Fatalf("SetVBAProject: %v", err)
	}

	if !wb.HasVBAProject() {
		t.Error("workbook should have VBA project after SetVBAProject")
	}

	vba := wb.GetVBAProject()
	if vba == nil {
		t.Fatal("GetVBAProject returned nil")
	}
	if len(vba.Data) != 4 {
		t.Errorf("VBA data length = %d, want 4", len(vba.Data))
	}
}

func TestVBAProject_EmptyData(t *testing.T) {
	wb := cells_foss.NewWorkbook()

	if err := wb.SetVBAProject([]byte{}); err == nil {
		t.Error("empty VBA data should error")
	}

	if err := wb.SetVBAProject(nil); err == nil {
		t.Error("nil VBA data should error")
	}
}

func TestVBAProject_Clear(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.SetVBAProject([]byte{0x01, 0x02})

	if !wb.HasVBAProject() {
		t.Error("should have VBA project")
	}

	wb.ClearVBAProject()

	if wb.HasVBAProject() {
		t.Error("should not have VBA project after ClearVBAProject")
	}

	if wb.GetVBAProject() != nil {
		t.Error("GetVBAProject should return nil after ClearVBAProject")
	}
}

func TestVBAProject_MarksModified(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Modified = false

	wb.SetVBAProject([]byte{0x01})
	if !wb.Modified {
		t.Error("SetVBAProject should mark workbook as modified")
	}

	wb.Modified = false
	wb.ClearVBAProject()
	if !wb.Modified {
		t.Error("ClearVBAProject should mark workbook as modified")
	}
}
