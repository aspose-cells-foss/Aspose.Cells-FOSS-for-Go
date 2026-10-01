package cells_foss

import "fmt"

// VBAProject holds the binary data of a VBA project.
//
// IMPORTANT: This library does NOT implement a VBA compiler or execution
// engine. It can only store and retrieve pre-compiled VBA project binary
// data. To create or edit VBA macros, use Microsoft Excel or another VBA
// development environment, then attach the resulting .bin file using
// Workbook.SetVBAProject.
//
// When saving a workbook with a VBA project, the file should use the .xlsm
// extension (macro-enabled workbook) to indicate it contains macros.
type VBAProject struct {
	// Data holds the raw VBA project binary data (vbaProject.bin content).
	Data []byte
}

// SetVBAProject attaches a VBA project to the workbook. The data parameter
// should be the raw bytes of a vbaProject.bin file from an existing .xlsm
// workbook.
//
// **Limitation**: This library cannot create or modify VBA code. It only
// stores the binary data for round-trip preservation.
//
//	data, err := os.ReadFile("path/to/vbaProject.bin")
//	if err != nil {
//	    log.Fatal(err)
//	}
//	wb.SetVBAProject(data)
//	wb.Save("output.xlsm") // Use .xlsm extension for macro-enabled workbooks
func (wb *Workbook) SetVBAProject(data []byte) error {
	if len(data) == 0 {
		return fmt.Errorf("cells_foss: VBA project data is empty")
	}
	wb.vbaProject = &VBAProject{Data: data}
	wb.Modified = true
	return nil
}

// GetVBAProject returns the VBA project data attached to the workbook, or
// nil if no VBA project is present.
func (wb *Workbook) GetVBAProject() *VBAProject {
	return wb.vbaProject
}

// HasVBAProject reports whether the workbook has an attached VBA project.
func (wb *Workbook) HasVBAProject() bool {
	return wb.vbaProject != nil && len(wb.vbaProject.Data) > 0
}

// ClearVBAProject removes the VBA project from the workbook.
func (wb *Workbook) ClearVBAProject() {
	wb.vbaProject = nil
	wb.Modified = true
}
