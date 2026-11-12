package cells_foss

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// TestSaveTo_LoadFrom_RoundTrip saves a workbook to an in-memory stream and
// loads it back, checking cell content and that neither call invents a file
// path.
func TestSaveTo_LoadFrom_RoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "streamed")
	ws.Cells().Set("B2", 42.0)

	var buf bytes.Buffer
	if err := wb.SaveTo(&buf); err != nil {
		t.Fatalf("SaveTo: %v", err)
	}
	if buf.Len() == 0 {
		t.Fatal("SaveTo wrote nothing")
	}
	if wb.FilePath != "" {
		t.Fatalf("SaveTo set FilePath = %q, want empty", wb.FilePath)
	}

	loaded, err := cells_foss.LoadFrom(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("LoadFrom: %v", err)
	}
	if loaded.FilePath != "" {
		t.Fatalf("LoadFrom set FilePath = %q, want empty", loaded.FilePath)
	}
	c, err := loaded.Worksheets[0].Cells().Get("A1")
	if err != nil {
		t.Fatalf("Get A1: %v", err)
	}
	if c.AsString() != "streamed" {
		t.Fatalf("A1 = %q, want %q", c.AsString(), "streamed")
	}
	c2, err := loaded.Worksheets[0].Cells().Get("B2")
	if err != nil {
		t.Fatalf("Get B2: %v", err)
	}
	if v, ok := c2.AsFloat64(); !ok || v != 42 {
		t.Fatalf("B2 = %v (ok=%v), want 42", v, ok)
	}
}

// TestSaveTo_UnmodifiedIsByteIdentical verifies the verbatim round-trip fast
// path also holds for stream saves: an unmodified loaded workbook streams
// byte-for-byte what is on disk, and FilePath stays pointing at the file.
func TestSaveTo_UnmodifiedIsByteIdentical(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "verbatim")
	path := filepath.Join(t.TempDir(), "stream.xlsx")
	if err := wb.Save(path); err != nil {
		t.Fatalf("Save: %v", err)
	}

	loaded, err := cells_foss.LoadWorkbook(path)
	if err != nil {
		t.Fatalf("LoadWorkbook: %v", err)
	}
	var buf bytes.Buffer
	if err := loaded.SaveTo(&buf); err != nil {
		t.Fatalf("SaveTo: %v", err)
	}
	onDisk, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}
	if !bytes.Equal(buf.Bytes(), onDisk) {
		t.Fatalf("streamed bytes differ from file bytes (%d vs %d)", buf.Len(), len(onDisk))
	}
	if loaded.FilePath != path {
		t.Fatalf("SaveTo changed FilePath to %q, want %q", loaded.FilePath, path)
	}
}

// TestLoadFrom_RejectsEncryptedStream checks that an encrypted package on a
// stream is rejected with guidance, and that LoadFromWithPassword decrypts
// it, keeps the password for re-saving, and rejects a wrong password.
func TestLoadFrom_RejectsEncryptedStream(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "secret")
	if err := wb.SetPassword("hunter2"); err != nil {
		t.Fatalf("SetPassword: %v", err)
	}
	var buf bytes.Buffer
	if err := wb.SaveTo(&buf); err != nil {
		t.Fatalf("SaveTo: %v", err)
	}

	if _, err := cells_foss.LoadFrom(bytes.NewReader(buf.Bytes())); err == nil {
		t.Fatal("LoadFrom accepted an encrypted stream")
	} else if !strings.Contains(err.Error(), "LoadFromWithPassword") {
		t.Fatalf("error %q should point at LoadFromWithPassword", err)
	}

	if _, err := cells_foss.LoadFromWithPassword(bytes.NewReader(buf.Bytes()), "wrong"); err == nil {
		t.Fatal("LoadFromWithPassword accepted a wrong password")
	}

	loaded, err := cells_foss.LoadFromWithPassword(bytes.NewReader(buf.Bytes()), "hunter2")
	if err != nil {
		t.Fatalf("LoadFromWithPassword: %v", err)
	}
	c, err := loaded.Worksheets[0].Cells().Get("A1")
	if err != nil {
		t.Fatalf("Get: %v", err)
	}
	if c.AsString() != "secret" {
		t.Fatalf("A1 = %q, want %q", c.AsString(), "secret")
	}

	// The password survives the stream load: re-streaming encrypts again.
	var buf2 bytes.Buffer
	if err := loaded.SaveTo(&buf2); err != nil {
		t.Fatalf("SaveTo (encrypted): %v", err)
	}
	if _, err := cells_foss.LoadFrom(bytes.NewReader(buf2.Bytes())); err == nil {
		t.Fatal("re-saved stream lost its encryption")
	}
	again, err := cells_foss.LoadFromWithPassword(bytes.NewReader(buf2.Bytes()), "hunter2")
	if err != nil {
		t.Fatalf("LoadFromWithPassword (2nd): %v", err)
	}
	c2, err := again.Worksheets[0].Cells().Get("A1")
	if err != nil {
		t.Fatalf("Get (2nd): %v", err)
	}
	if c2.AsString() != "secret" {
		t.Fatalf("A1 = %q, want %q", c2.AsString(), "secret")
	}
}

// TestLoadFromWithPassword_PlainStream checks that an unencrypted stream
// loads through the password variant without complaint.
func TestLoadFromWithPassword_PlainStream(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "plain")
	var buf bytes.Buffer
	if err := wb.SaveTo(&buf); err != nil {
		t.Fatalf("SaveTo: %v", err)
	}
	loaded, err := cells_foss.LoadFromWithPassword(bytes.NewReader(buf.Bytes()), "ignored")
	if err != nil {
		t.Fatalf("LoadFromWithPassword: %v", err)
	}
	c, err := loaded.Worksheets[0].Cells().Get("A1")
	if err != nil {
		t.Fatalf("Get: %v", err)
	}
	if c.AsString() != "plain" {
		t.Fatalf("A1 = %q, want %q", c.AsString(), "plain")
	}
}

// TestSaveTo_NilGuards checks the error paths for nil receivers and writers.
func TestSaveTo_NilGuards(t *testing.T) {
	var buf bytes.Buffer
	var wb *cells_foss.Workbook
	if err := wb.SaveTo(&buf); err == nil {
		t.Fatal("SaveTo on a nil Workbook should fail")
	}
	wb = cells_foss.NewWorkbook()
	if err := wb.SaveTo(nil); err == nil {
		t.Fatal("SaveTo with a nil writer should fail")
	}
}
