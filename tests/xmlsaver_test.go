package cells_foss_test

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"sort"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestSaveNewWorkbook_RoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", "Name")
	c.Set("B1", "Score")
	c.Set("A2", "Alice")
	c.Set("B2", float64(95.5))
	c.Set("A3", "Bob")
	c.Set("B3", float64(87))

	dir := t.TempDir()
	p := filepath.Join(dir, "rt.xlsx")
	if err := wb.Save(p); err != nil {
		t.Fatalf("Save: %v", err)
	}

	loaded, _ := cells_foss.LoadWorkbook(p)
	lc := loaded.Worksheets[0].Cells()
	checks := map[string]string{"A1": "Name", "B1": "Score", "A2": "Alice", "B2": "95.5", "A3": "Bob", "B3": "87"}
	for ref, want := range checks {
		c, _ := lc.Get(ref)
		if got := cells_foss.CellToString(c.Value); got != want {
			t.Errorf("%s = %q, want %q", ref, got, want)
		}
	}
}

func TestSaveUnmodified_RawSheetBytesIdentical(t *testing.T) {
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>Hello</v></c></row></sheetData>
</worksheet>`

	inPath := makeNumericXLSX(t, sheetXML)
	originalSheet := readRawSheetXML(t, inPath)

	wb, _ := cells_foss.LoadWorkbook(inPath)
	dir := t.TempDir()
	outPath := filepath.Join(dir, "unmodified.xlsx")
	wb.Save(outPath)

	savedSheet := readRawSheetXML(t, outPath)
	if !bytes.Equal(originalSheet, savedSheet) {
		t.Errorf("unmodified sheet XML changed:\noriginal: %s\nsaved:   %s", originalSheet, savedSheet)
	}
}

func TestSaveModified_RegeneratesSheetXML(t *testing.T) {
	sheetXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>
</worksheet>`

	inPath := makeNumericXLSX(t, sheetXML)
	originalSheet := readRawSheetXML(t, inPath)

	wb, _ := cells_foss.LoadWorkbook(inPath)
	wb.Worksheets[0].Cells().Set("A2", float64(99))

	dir := t.TempDir()
	outPath := filepath.Join(dir, "modified.xlsx")
	wb.Save(outPath)

	savedSheet := readRawSheetXML(t, outPath)
	if bytes.Equal(originalSheet, savedSheet) {
		t.Error("modified sheet should differ from original")
	}
}

func TestSave_BooleanRoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", true)
	c.Set("A2", false)

	dir := t.TempDir()
	p := filepath.Join(dir, "bool.xlsx")
	wb.Save(p)
	loaded, _ := cells_foss.LoadWorkbook(p)
	lc := loaded.Worksheets[0].Cells()

	c1, _ := lc.Get("A1")
	if c1.Value != true {
		t.Errorf("A1 = %v", c1.Value)
	}
	c2, _ := lc.Get("A2")
	if c2.Value != false {
		t.Errorf("A2 = %v", c2.Value)
	}
}

func TestSave_UpdatesFilePath(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	dir := t.TempDir()
	p := filepath.Join(dir, "fp.xlsx")
	wb.Save(p)
	if wb.FilePath != p {
		t.Errorf("FilePath = %q, want %q", wb.FilePath, p)
	}
}

func TestSave_ResetsModifiedFlag(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	if !wb.Modified {
		t.Error("new workbook should be modified")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "mf.xlsx")
	wb.Save(p)
	if wb.Modified {
		t.Error("modified flag should be false after save")
	}
}

// cellRefsInRow returns the A1 references of every <c> element in the given
// row, in the order they appear in the saved sheet XML.
func cellRefsInRow(t *testing.T, sheetXML []byte, row int) []string {
	t.Helper()

	rowPattern := regexp.MustCompile(fmt.Sprintf(`(?s)<row r="%d">(.*?)</row>`, row))
	m := rowPattern.FindSubmatch(sheetXML)
	if m == nil {
		t.Fatalf("row %d not found in sheet XML: %s", row, sheetXML)
	}

	var refs []string
	for _, c := range regexp.MustCompile(`<c r="([A-Z]+[0-9]+)"`).FindAllSubmatch(m[1], -1) {
		refs = append(refs, string(c[1]))
	}
	return refs
}

// ECMA-376 expects cells within a row in ascending column order. Columns past
// Z are where lexicographic comparison of the column letters ("AA" < "B")
// diverges from numeric order, so cover that boundary explicitly.
func TestSave_CellsAreInAscendingColumnOrder(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	for _, ref := range []string{"AB1", "B1", "AA1", "A1", "Z1"} {
		if err := c.Set(ref, ref); err != nil {
			t.Fatalf("Set(%q): %v", ref, err)
		}
	}

	dir := t.TempDir()
	p := filepath.Join(dir, "order.xlsx")
	if err := wb.Save(p); err != nil {
		t.Fatalf("Save: %v", err)
	}

	got := cellRefsInRow(t, readRawSheetXML(t, p), 1)
	want := []string{"A1", "B1", "Z1", "AA1", "AB1"}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("row 1 cells = %v, want %v", got, want)
	}
}

func TestSave_NilAndRepeated(t *testing.T) {
	var nilWB *cells_foss.Workbook
	if err := nilWB.Save("/tmp/x.xlsx"); err == nil {
		t.Error("nil save should error")
	}

	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "persistent")
	dir := t.TempDir()
	wb.Save(filepath.Join(dir, "s1.xlsx"))
	wb.Save(filepath.Join(dir, "s2.xlsx")) // repeated save without modifications
}

// TestSave_PreservesUnmodeledWorkbookParts verifies that workbook-scoped
// parts the library does not model (docProps, theme) survive modified saves
// byte-for-byte, together with the relationships that reach them.
func TestSave_PreservesUnmodeledWorkbookParts(t *testing.T) {
	const (
		coreXML  = `<cp:coreProperties xmlns:cp="ns-core">MARKER-CORE</cp:coreProperties>`
		appXML   = `<Properties xmlns="ns-ext">MARKER-APP</Properties>`
		themeXML = `<a:theme xmlns:a="ns-theme" name="MARKER-THEME"/>`
	)

	dir := t.TempDir()
	src := filepath.Join(dir, "input.xlsx")
	writeWorkbookPartsFile(t, src, map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
		"xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
		"xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`,
		"xl/worksheets/sheet1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		"docProps/core.xml":   coreXML,
		"docProps/app.xml":    appXML,
		"xl/theme/theme1.xml": themeXML,
	})

	wb, err := cells_foss.LoadWorkbook(src)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}

	// Two consecutive modified saves: the second one works from the package
	// bytes produced by the first, so preservation must be self-sustaining.
	prev := src
	for i, ref := range []string{"B2", "C3"} {
		wb.Worksheets[0].Cells().Set(ref, i+2)
		out := filepath.Join(dir, fmt.Sprintf("out%d.xlsx", i+1))
		if err := wb.Save(out); err != nil {
			t.Fatalf("Save %d: %v", i+1, err)
		}

		if got := readZipEntry(t, out, "docProps/core.xml"); string(got) != coreXML {
			t.Errorf("save %d: core.xml = %q, want %q", i+1, got, coreXML)
		}
		if got := readZipEntry(t, out, "docProps/app.xml"); string(got) != appXML {
			t.Errorf("save %d: app.xml = %q, want %q", i+1, got, appXML)
		}
		if got := readZipEntry(t, out, "xl/theme/theme1.xml"); string(got) != themeXML {
			t.Errorf("save %d: theme1.xml = %q, want %q", i+1, got, themeXML)
		}

		rootRels := readZipEntry(t, out, "_rels/.rels")
		if !bytes.Contains(rootRels, []byte(`Target="docProps/core.xml"`)) ||
			!bytes.Contains(rootRels, []byte(`Target="docProps/app.xml"`)) {
			t.Errorf("save %d: root rels lost docProps relationships: %s", i+1, rootRels)
		}
		wbRels := readZipEntry(t, out, "xl/_rels/workbook.xml.rels")
		if !bytes.Contains(wbRels, []byte(`Target="theme/theme1.xml"`)) {
			t.Errorf("save %d: workbook rels lost theme relationship: %s", i+1, wbRels)
		}
		ct := readZipEntry(t, out, "[Content_Types].xml")
		for _, part := range []string{"/docProps/core.xml", "/docProps/app.xml", "/xl/theme/theme1.xml"} {
			if !bytes.Contains(ct, []byte(`PartName="`+part+`"`)) {
				t.Errorf("save %d: content types lost declaration for %s", i+1, part)
			}
		}
		prev = out
	}

	// The edited cells must have round-tripped alongside the preserved parts.
	reloaded, err := cells_foss.LoadWorkbook(prev)
	if err != nil {
		t.Fatalf("reload: %v", err)
	}
	c, err := reloaded.Worksheets[0].Cells().Get("C3")
	if err != nil {
		t.Fatalf("Get C3: %v", err)
	}
	got, ok := c.AsInt()
	if !ok || got != 3 {
		t.Errorf("C3 = %d (ok=%v), want 3", got, ok)
	}
}

// writeWorkbookPartsFile writes the given parts, in sorted name order, into
// a zip package at path.
func writeWorkbookPartsFile(t *testing.T, path string, parts map[string]string) {
	t.Helper()
	f, err := os.Create(path)
	if err != nil {
		t.Fatalf("create %s: %v", path, err)
	}
	defer f.Close()

	zw := zip.NewWriter(f)
	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		fw, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip create %s: %v", name, err)
		}
		if _, err := fw.Write([]byte(parts[name])); err != nil {
			t.Fatalf("zip write %s: %v", name, err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatalf("zip close: %v", err)
	}
}
