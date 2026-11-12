// Stream I/O example: save a workbook to an in-memory stream and read it
// back without touching the filesystem. SaveTo/LoadFrom work with any
// io.Writer/io.Reader — an HTTP response, a zip entry, a network connection.
//
//	cd examples && go run ./stream_io/
package main

import (
	"bytes"
	"fmt"
	"os"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Item")
	ws.Cells().Set("B1", "Price")
	ws.Cells().Set("A2", "Coffee")
	ws.Cells().Set("B2", 4.5)

	// SaveTo writes the exact same .xlsx bytes Save would, to any writer.
	var buf bytes.Buffer
	if err := wb.SaveTo(&buf); err != nil {
		fmt.Fprintf(os.Stderr, "SaveTo: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Streamed %d bytes of .xlsx data\n", buf.Len())

	// LoadFrom reads a package back from any reader.
	loaded, err := cells_foss.LoadFrom(bytes.NewReader(buf.Bytes()))
	if err != nil {
		fmt.Fprintf(os.Stderr, "LoadFrom: %v\n", err)
		os.Exit(1)
	}
	cell, err := loaded.Worksheets[0].Cells().Get("A2")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Get: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Reloaded A2 from stream: %s\n", cell.AsString())

	// Encrypted workbooks stream too. The password stays on the workbook,
	// so SaveTo re-encrypts and LoadFromWithPassword reads it back.
	if err := loaded.SetPassword("secret"); err != nil {
		fmt.Fprintf(os.Stderr, "SetPassword: %v\n", err)
		os.Exit(1)
	}
	var enc bytes.Buffer
	if err := loaded.SaveTo(&enc); err != nil {
		fmt.Fprintf(os.Stderr, "SaveTo (encrypted): %v\n", err)
		os.Exit(1)
	}
	if _, err := cells_foss.LoadFrom(bytes.NewReader(enc.Bytes())); err == nil {
		fmt.Fprintln(os.Stderr, "LoadFrom should reject an encrypted stream")
		os.Exit(1)
	}
	secure, err := cells_foss.LoadFromWithPassword(bytes.NewReader(enc.Bytes()), "secret")
	if err != nil {
		fmt.Fprintf(os.Stderr, "LoadFromWithPassword: %v\n", err)
		os.Exit(1)
	}
	cell, err = secure.Worksheets[0].Cells().Get("A2")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Get (encrypted): %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Reloaded A2 from encrypted stream: %s\n", cell.AsString())
}
