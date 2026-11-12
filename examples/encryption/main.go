// Encryption example: password-protect a workbook with ECMA-376 Agile
// Encryption (the same scheme Excel uses), then load it back and remove the
// password.
//
//	cd examples && go run ./encryption/
package main

import (
	"fmt"
	"os"

	"github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func main() {
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	ws.Cells().Set("A1", "Confidential")
	ws.Cells().Set("B1", 42)

	const password = "S3cret!"
	if err := wb.SetPassword(password); err != nil {
		fmt.Fprintf(os.Stderr, "SetPassword: %v\n", err)
		os.Exit(1)
	}

	path := "outputfiles/encrypted.xlsx"
	if err := wb.Save(path); err != nil {
		fmt.Fprintf(os.Stderr, "Save: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Saved %s (OLE/CFB container with Agile Encryption)\n", path)

	// A plain load fails: the file is no longer a ZIP package.
	if _, err := cells_foss.LoadWorkbook(path); err == nil {
		fmt.Fprintln(os.Stderr, "unexpected: encrypted file loaded without password")
		os.Exit(1)
	} else {
		fmt.Println("LoadWorkbook without password correctly fails")
	}

	// Wrong passwords are rejected.
	if _, err := cells_foss.LoadWithPassword(path, "wrong-password"); err == nil {
		fmt.Fprintln(os.Stderr, "unexpected: wrong password accepted")
		os.Exit(1)
	}
	fmt.Println("Wrong password correctly rejected")

	// The correct password decrypts the workbook.
	loaded, err := cells_foss.LoadWithPassword(path, password)
	if err != nil {
		fmt.Fprintf(os.Stderr, "LoadWithPassword: %v\n", err)
		os.Exit(1)
	}
	a1, _ := loaded.Worksheets[0].Cells().Get("A1")
	fmt.Printf("Decrypted with password: A1=%q\n", a1.AsString())

	// CheckPassword is the safe verification API: it returns false when no
	// password is set and compares in constant time. (VerifyPassword is
	// deprecated.)
	fmt.Printf("CheckPassword(correct)=%v CheckPassword(wrong)=%v\n",
		loaded.CheckPassword(password), loaded.CheckPassword("nope"))

	// Remove the password and save a plain workbook again.
	if err := loaded.SetPassword(""); err != nil {
		fmt.Fprintf(os.Stderr, "SetPassword(\"\"): %v\n", err)
		os.Exit(1)
	}
	plain := "outputfiles/decrypted.xlsx"
	if err := loaded.Save(plain); err != nil {
		fmt.Fprintf(os.Stderr, "Save plain: %v\n", err)
		os.Exit(1)
	}
	reopened, err := cells_foss.LoadWorkbook(plain)
	if err != nil {
		fmt.Fprintf(os.Stderr, "LoadWorkbook plain: %v\n", err)
		os.Exit(1)
	}
	b1, _ := reopened.Worksheets[0].Cells().Get("B1")
	b1val, _ := b1.AsFloat64()
	fmt.Printf("Password removed, %s loads without a password: B1=%v\n", plain, b1val)
}
