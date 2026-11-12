// check_open_xlsx validates the internal structure of an .xlsx file.
//
// Usage:
//
//	go run verify/check_open_xlsx.go <file.xlsx>
//
// The tool opens the file as a ZIP archive and checks that every required
// OPC / OOXML part is present, that all XML parts are well-formed, that
// relationship targets and content-type declarations cover every part, and
// that shared-string indices used by cells are in bounds. It also reports
// basic statistics such as sheet count, row count, and cell count.
package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Fprintf(os.Stderr, "Usage: go run verify/check_open_xlsx.go <file.xlsx>\n")
		os.Exit(1)
	}

	pathArg := os.Args[1]
	if err := checkXLSX(pathArg); err != nil {
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
			RID     string `xml:"id,attr"`
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

type checkContentTypes struct {
	Defaults []struct {
		Extension   string `xml:"Extension,attr"`
		ContentType string `xml:"ContentType,attr"`
	} `xml:"Default"`
	Overrides []struct {
		PartName    string `xml:"PartName,attr"`
		ContentType string `xml:"ContentType,attr"`
	} `xml:"Override"`
}

type checkRels struct {
	Rel []struct {
		ID         string `xml:"Id,attr"`
		Type       string `xml:"Type,attr"`
		Target     string `xml:"Target,attr"`
		TargetMode string `xml:"TargetMode,attr"`
	} `xml:"Relationship"`
}

const worksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

// ---------------------------------------------------------------------------
// checkXLSX
// ---------------------------------------------------------------------------

func checkXLSX(pathArg string) error {
	abs, err := filepath.Abs(pathArg)
	if err != nil {
		abs = pathArg
	}
	fmt.Printf("Checking %s\n", abs)

	data, err := os.ReadFile(pathArg)
	if err != nil {
		return fmt.Errorf("cannot read file: %w", err)
	}

	// Password-protected workbooks are OLE/CFB containers (or legacy ECRX
	// files), not ZIP packages. Report them clearly instead of a ZIP parse
	// error.
	cfbMagic := []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
	if len(data) >= 8 && bytes.Equal(data[:8], cfbMagic) {
		return fmt.Errorf("file is an encrypted OLE/CFB container (password-protected workbook); decrypt it first, e.g. with cells_foss.LoadWithPassword")
	}
	if len(data) >= 4 && bytes.Equal(data[:4], []byte("ECRX")) {
		return fmt.Errorf("file is an encrypted legacy ECRX container; decrypt it first, e.g. with cells_foss.LoadWithPassword")
	}

	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return fmt.Errorf("not a valid ZIP / XLSX file: %w", err)
	}

	entries := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		if f.FileInfo().IsDir() {
			continue
		}
		raw, err := readZipEntry(f)
		if err != nil {
			return fmt.Errorf("cannot read %s: %w", f.Name, err)
		}
		entries[f.Name] = raw
	}

	// ---- Required entries ----
	for _, name := range []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"xl/workbook.xml",
		"xl/_rels/workbook.xml.rels",
	} {
		if _, ok := entries[name]; !ok {
			return fmt.Errorf("missing required entry: %q", name)
		}
		fmt.Printf("  ✓ %s\n", name)
	}

	// ---- Well-formedness of every XML part ----
	xmlParts := make([]string, 0, len(entries))
	for name := range entries {
		if strings.HasSuffix(name, ".xml") || strings.HasSuffix(name, ".rels") {
			xmlParts = append(xmlParts, name)
		}
	}
	sort.Strings(xmlParts)
	for _, name := range xmlParts {
		if err := checkWellFormed(entries[name]); err != nil {
			return fmt.Errorf("%s is not well-formed XML: %w", name, err)
		}
	}
	fmt.Printf("  ✓ %d XML part(s) well-formed\n", len(xmlParts))

	// ---- Content types cover every part ----
	var ct checkContentTypes
	if err := xml.Unmarshal(entries["[Content_Types].xml"], &ct); err != nil {
		return fmt.Errorf("[Content_Types].xml parse error: %w", err)
	}
	defaults := make(map[string]bool)
	for _, d := range ct.Defaults {
		defaults[strings.ToLower(d.Extension)] = true
	}
	overrides := make(map[string]bool)
	for _, o := range ct.Overrides {
		overrides[strings.TrimPrefix(o.PartName, "/")] = true
	}
	for name := range entries {
		if name == "[Content_Types].xml" {
			continue
		}
		ext := strings.ToLower(path.Ext(name))
		if !overrides[name] && !defaults[strings.TrimPrefix(ext, ".")] {
			return fmt.Errorf("part %q has no content type (no Default for %q, no Override)", name, ext)
		}
	}
	fmt.Printf("  ✓ [Content_Types].xml covers all %d part(s)\n", len(entries)-1)

	// ---- Relationship targets exist ----
	for _, relsName := range []string{"_rels/.rels", "xl/_rels/workbook.xml.rels"} {
		// "_rels/.rels" describes the package root; "xl/_rels/workbook.xml.rels"
		// describes "xl/workbook.xml", so targets resolve against "xl/".
		dir, base := path.Dir(relsName), path.Base(relsName)
		parent := strings.TrimSuffix(strings.TrimSuffix(dir, "_rels"), "/")
		base = strings.TrimSuffix(base, ".rels") // "" for the root .rels
		if parent != "" {
			base = parent + "/" + base
		}
		base = path.Dir(base)
		if base == "." {
			base = ""
		}
		var rels checkRels
		if err := xml.Unmarshal(entries[relsName], &rels); err != nil {
			return fmt.Errorf("%s parse error: %w", relsName, err)
		}
		for _, rel := range rels.Rel {
			if rel.TargetMode == "External" || strings.HasPrefix(rel.Target, "http") {
				continue
			}
			target := path.Clean(path.Join(base, rel.Target))
			if _, ok := entries[target]; !ok {
				return fmt.Errorf("%s: relationship %s target %q not found in package", relsName, rel.ID, target)
			}
		}
		fmt.Printf("  ✓ %s: all targets resolve\n", relsName)
	}

	// ---- Parse workbook.xml and resolve worksheet parts via rels ----
	var wb checkWB
	if err := xml.Unmarshal(entries["xl/workbook.xml"], &wb); err != nil {
		return fmt.Errorf("workbook.xml parse error: %w", err)
	}
	var wbRels checkRels
	if err := xml.Unmarshal(entries["xl/_rels/workbook.xml.rels"], &wbRels); err != nil {
		return fmt.Errorf("workbook.xml.rels parse error: %w", err)
	}
	relByID := make(map[string]string, len(wbRels.Rel))
	sheetTargets := make(map[string]bool)
	for _, rel := range wbRels.Rel {
		relByID[rel.ID] = path.Clean(path.Join("xl", rel.Target))
		if rel.Type == worksheetRelType {
			sheetTargets[path.Clean(path.Join("xl", rel.Target))] = true
		}
	}
	fmt.Printf("  ✓ workbook.xml: %d sheet(s)\n", len(wb.Sheets.Sheet))
	sheetFiles := make([]string, 0, len(wb.Sheets.Sheet))
	for _, s := range wb.Sheets.Sheet {
		target, ok := relByID[s.RID]
		if !ok {
			return fmt.Errorf("workbook.xml: sheet %q references unknown relationship %s", s.Name, s.RID)
		}
		if _, ok := entries[target]; !ok {
			return fmt.Errorf("workbook.xml: sheet %q points at missing part %q", s.Name, target)
		}
		fmt.Printf("    - %q (sheetId=%s) → %s\n", s.Name, s.SheetID, target)
		sheetFiles = append(sheetFiles, target)
	}
	for target := range sheetTargets {
		if _, ok := entries[target]; !ok {
			return fmt.Errorf("worksheet relationship points at missing part %q", target)
		}
	}

	// ---- Shared strings (optional) ----
	var sst *checkSST
	if raw, ok := entries["xl/sharedStrings.xml"]; ok {
		sst = new(checkSST)
		if err := xml.Unmarshal(raw, sst); err != nil {
			return fmt.Errorf("sharedStrings.xml parse error: %w", err)
		}
		fmt.Printf("  ✓ sharedStrings.xml: %d string(s)\n", len(sst.Items))
	} else {
		fmt.Println("  ℹ sharedStrings.xml not present (inline values only)")
	}

	// ---- Parse each sheet, check SST indices ----
	totalRows, totalCells := 0, 0
	for _, sf := range sheetFiles {
		var sheet checkSheet
		if err := xml.Unmarshal(entries[sf], &sheet); err != nil {
			return fmt.Errorf("%s parse error: %w", sf, err)
		}
		rows := len(sheet.SheetData.Rows)
		cells := 0
		for _, r := range sheet.SheetData.Rows {
			for _, c := range r.Cells {
				cells++
				if c.T == "s" && sst != nil {
					idx, err := strconv.Atoi(strings.TrimSpace(c.V))
					if err != nil || idx < 0 || idx >= len(sst.Items) {
						return fmt.Errorf("%s: cell %s has out-of-range shared string index %q (%d strings available)", sf, c.Ref, c.V, len(sst.Items))
					}
				}
			}
		}
		fmt.Printf("  ✓ %s: %d row(s), %d cell(s)\n", sf, rows, cells)
		totalRows += rows
		totalCells += cells
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

// checkWellFormed reports whether data is well-formed XML by decoding it
// token by token (no schema validation).
func checkWellFormed(data []byte) error {
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		_, err := dec.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
	}
}

func readZipEntry(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}
