// check_open_xlsx validates the internal structure of an .xlsx file.
//
// Usage:
//
//	go run verify/check_open_xlsx.go <file.xlsx>
//
// The tool opens the file as a ZIP archive and checks that every required
// OPC / OOXML part is present and well-formed XML. It also reports basic
// statistics such as sheet count, row count, and cell count.
package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Fprintf(os.Stderr, "Usage: go run verify/check_open_xlsx.go <file.xlsx>\n")
		os.Exit(1)
	}

	path := os.Args[1]
	if err := checkXLSX(path); err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	fmt.Println("PASS — all checks passed.")
}

// ---------------------------------------------------------------------------
// XML parsing helpers (mirror the structures in the cells_foss package)
// ---------------------------------------------------------------------------

type checkWB struct {
	Sheets struct {
		Sheet []struct {
			Name    string `xml:"name,attr"`
			SheetID string `xml:"sheetId,attr"`
		} `xml:"sheet"`
	} `xml:"sheets"`
}

type checkSheet struct {
	SheetData struct {
		Rows []struct {
			R     int `xml:"r,attr"`
			Cells []struct {
				Ref string `xml:"r,attr"`
				T   string `xml:"t,attr,omitempty"`
				V   string `xml:"v"`
			} `xml:"c"`
		} `xml:"row"`
	} `xml:"sheetData"`
}

type checkSST struct {
	Items []struct {
		T string `xml:"t"`
	} `xml:"si"`
}

// ---------------------------------------------------------------------------
// checkXLSX
// ---------------------------------------------------------------------------

func checkXLSX(path string) error {
	abs, err := filepath.Abs(path)
	if err != nil {
		abs = path
	}
	fmt.Printf("Checking %s\n", abs)

	data, err := os.ReadFile(path)
	if err != nil {
		return fmt.Errorf("cannot read file: %w", err)
	}

	zr2, err := newZipReader(data)
	if err != nil {
		return fmt.Errorf("not a valid ZIP / XLSX file: %w", err)
	}

	required := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"xl/workbook.xml",
		"xl/_rels/workbook.xml.rels",
	}
	found := make(map[string]bool)
	sheetFiles := []string{}

	for _, f := range zr2.File {
		found[f.Name] = true
		if dir, _ := filepath.Split(f.Name); dir == "xl/worksheets/" {
			sheetFiles = append(sheetFiles, f.Name)
		}
	}

	// ---- Required entries ----
	for _, name := range required {
		if !found[name] {
			return fmt.Errorf("missing required entry: %q", name)
		}
		fmt.Printf("  ✓ %s\n", name)
	}

	if len(sheetFiles) == 0 {
		return fmt.Errorf("no worksheet files found in xl/worksheets/")
	}
	for _, sf := range sheetFiles {
		fmt.Printf("  ✓ %s\n", sf)
	}

	// ---- Parse workbook.xml ----
	wbRaw, _ := readZipEntry(zr2, "xl/workbook.xml")
	var wb checkWB
	if err := xml.Unmarshal(wbRaw, &wb); err != nil {
		return fmt.Errorf("workbook.xml is not valid XML: %w", err)
	}
	fmt.Printf("  ✓ workbook.xml: %d sheet(s)\n", len(wb.Sheets.Sheet))
	for _, s := range wb.Sheets.Sheet {
		fmt.Printf("    - %q (sheetId=%s)\n", s.Name, s.SheetID)
	}

	// ---- Parse each sheet ----
	totalRows, totalCells := 0, 0
	for _, sf := range sheetFiles {
		raw, err := readZipEntry(zr2, sf)
		if err != nil {
			return fmt.Errorf("cannot read %s: %w", sf, err)
		}
		var sheet checkSheet
		if err := xml.Unmarshal(raw, &sheet); err != nil {
			return fmt.Errorf("%s is not valid XML: %w", sf, err)
		}
		rows := len(sheet.SheetData.Rows)
		cells := 0
		for _, r := range sheet.SheetData.Rows {
			cells += len(r.Cells)
		}
		fmt.Printf("  ✓ %s: %d row(s), %d cell(s)\n", sf, rows, cells)
		totalRows += rows
		totalCells += cells
	}

	// ---- Shared strings (optional) ----
	if found["xl/sharedStrings.xml"] {
		ssRaw, _ := readZipEntry(zr2, "xl/sharedStrings.xml")
		var sst checkSST
		if err := xml.Unmarshal(ssRaw, &sst); err != nil {
			return fmt.Errorf("sharedStrings.xml is not valid XML: %w", err)
		}
		fmt.Printf("  ✓ sharedStrings.xml: %d string(s)\n", len(sst.Items))
	} else {
		fmt.Println("  ℹ sharedStrings.xml not present (inline values only)")
	}

	// ---- Summary ----
	fmt.Println()
	fmt.Printf("Summary: %d sheet(s), %d row(s), %d cell(s)\n",
		len(wb.Sheets.Sheet), totalRows, totalCells)

	return nil
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

func newZipReader(data []byte) (*zip.Reader, error) {
	// We need an io.ReaderAt. Wrap the data in a bytes.Reader-compatible type.
	return zip.NewReader(&byteReader{data: data}, int64(len(data)))
}

type byteReader struct {
	data []byte
}

func (r *byteReader) ReadAt(p []byte, off int64) (int, error) {
	if off >= int64(len(r.data)) {
		return 0, io.EOF
	}
	n := copy(p, r.data[off:])
	if n < len(p) {
		return n, io.EOF
	}
	return n, nil
}

func readZipEntry(zr *zip.Reader, name string) ([]byte, error) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("entry %q not found", name)
}
