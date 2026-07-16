package cells_foss_test

import (
	"os"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose/cells_foss/v26/aspose/cells_foss"
)

func writeStringFile(path, content string) error {
	return os.WriteFile(path, []byte(content), 0644)
}

// makeNumericXLSX creates a minimal .xlsx with the given sheet XML.
func makeNumericXLSX(t *testing.T, sheetXML string) string {
	t.Helper()
	dir := t.TempDir()
	p := filepath.Join(dir, "input.xlsx")
	if err := cells_foss.WriteTestXLSX(p, sheetXML, ""); err != nil {
		t.Fatalf("WriteTestXLSX: %v", err)
	}
	return p
}

// readRawSheetXML returns raw bytes of sheet1.xml from an .xlsx.
func readRawSheetXML(t *testing.T, path string) []byte {
	t.Helper()
	raw, err := cells_foss.ReadTestZipEntry(path, "xl/worksheets/sheet1.xml")
	if err != nil {
		t.Fatalf("ReadTestZipEntry: %v", err)
	}
	return raw
}

// readZipEntry returns raw bytes of the named entry from an .xlsx.
func readZipEntry(t *testing.T, path, name string) []byte {
	t.Helper()
	raw, err := cells_foss.ReadTestZipEntry(path, name)
	if err != nil {
		t.Fatalf("ReadTestZipEntry %q: %v", name, err)
	}
	return raw
}
