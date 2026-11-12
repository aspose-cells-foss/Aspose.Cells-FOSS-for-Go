package cells_foss_test

import (
	"bytes"
	"crypto/aes"
	"crypto/cipher"
	"crypto/rand"
	"crypto/sha512"
	"encoding/base64"
	"encoding/binary"
	"fmt"
	"io"
	mrand "math/rand"
	"os"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

func TestEncryption_RoundTrip(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	c := wb.Worksheets[0].Cells()
	c.Set("A1", "Secret")
	c.Set("B1", float64(42))

	wb.SetPassword("test123")
	dir := t.TempDir()
	p := filepath.Join(dir, "enc.xlsx")
	wb.Save(p)

	// Normal load should fail.
	if _, err := cells_foss.LoadWorkbook(p); err == nil {
		t.Error("LoadWorkbook on encrypted file should fail")
	}

	// Wrong password should fail.
	if _, err := cells_foss.LoadWithPassword(p, "wrong"); err == nil {
		t.Error("wrong password should fail")
	}

	// Correct password.
	loaded, err := cells_foss.LoadWithPassword(p, "test123")
	if err != nil {
		t.Fatalf("LoadWithPassword: %v", err)
	}

	lc := loaded.Worksheets[0].Cells()
	ca1, _ := lc.Get("A1")
	if ca1.Value != "Secret" {
		t.Errorf("A1 = %v", ca1.Value)
	}
	cb1, _ := lc.Get("B1")
	if cells_foss.CellToString(cb1.Value) != "42" {
		t.Errorf("B1 = %v", cb1.Value)
	}

	// Password preserved.
	if !loaded.VerifyPassword("test123") {
		t.Error("VerifyPassword failed")
	}
	if loaded.VerifyPassword("wrong") {
		t.Error("VerifyPassword should return false")
	}

	// Remove password, re-save.
	loaded.Modified = true
	loaded.SetPassword("")
	p2 := filepath.Join(dir, "plain.xlsx")
	loaded.Save(p2)

	wb2, _ := cells_foss.LoadWorkbook(p2)
	c2, _ := wb2.Worksheets[0].Cells().Get("A1")
	if c2.Value != "Secret" {
		t.Errorf("after password removal: A1 = %v", c2.Value)
	}
}

func TestEncryption_EmptyAndNil(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "plain")

	dir := t.TempDir()
	p := filepath.Join(dir, "plain.xlsx")
	wb.Save(p)

	// LoadWithPassword on plain file should work.
	loaded, err := cells_foss.LoadWithPassword(p, "any")
	if err != nil {
		t.Fatalf("LoadWithPassword on plain: %v", err)
	}
	if !loaded.VerifyPassword("any") {
		t.Error("VerifyPassword on plain should return true")
	}

	var nilWB *cells_foss.Workbook
	if err := nilWB.SetPassword("x"); err == nil {
		t.Error("nil SetPassword should error")
	}
}

func TestEncryption_LoadWithPasswordErrors(t *testing.T) {
	if _, err := cells_foss.LoadWithPassword("nonexistent.xlsx", "pw"); err == nil {
		t.Error("nonexistent file should error")
	}
}

func TestEncryption_CFBStructure(t *testing.T) {
	// Verify that encrypted files use the standard OLE/CFB format.
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "test")
	wb.SetPassword("secret")

	dir := t.TempDir()
	p := filepath.Join(dir, "enc.xlsx")
	wb.Save(p)

	// Read the file and verify it starts with OLE/CFB magic.
	data, err := os.ReadFile(p)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}

	// OLE/CFB magic: D0 CF 11 E0 A1 B1 1A E1
	magic := []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
	if len(data) < 8 {
		t.Fatal("file too short")
	}
	for i, b := range magic {
		if data[i] != b {
			t.Errorf("byte %d: got 0x%02X, want 0x%02X", i, data[i], b)
		}
	}

	// Verify the file can be loaded back.
	loaded, err := cells_foss.LoadWithPassword(p, "secret")
	if err != nil {
		t.Fatalf("LoadWithPassword: %v", err)
	}
	c, _ := loaded.Worksheets[0].Cells().Get("A1")
	if c.Value != "test" {
		t.Errorf("A1 = %v, want 'test'", c.Value)
	}
}

// TestEncryption_AgileSpecStructure verifies the EncryptionInfo XML stored
// in the CFB container matches the [MS-OFFCRYPTO] Agile layout Excel
// writes: namespaced root, password keyEncryptor URI, and the
// verifier-hash-input/value attribute pair. The XML lives verbatim in the
// mini-stream, so a byte scan of the container finds it.
func TestEncryption_AgileSpecStructure(t *testing.T) {
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "spec")
	wb.SetPassword("pw")

	p := filepath.Join(t.TempDir(), "enc.xlsx")
	if err := wb.Save(p); err != nil {
		t.Fatalf("Save: %v", err)
	}
	data, err := os.ReadFile(p)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}

	want := []string{
		`<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"`,
		`<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">`,
		`<p:encryptedKey `,
		`spinCount="100000"`,
		`encryptedVerifierHashInput="`,
		`encryptedVerifierHashValue="`,
		`encryptedKeyValue="`,
		`<dataIntegrity encryptedHmacKey="`,
		`encryptedHmacValue="`,
		`cipherAlgorithm="AES"`,
		`hashAlgorithm="SHA512"`,
	}
	for _, s := range want {
		if !bytes.Contains(data, []byte(s)) {
			t.Errorf("EncryptionInfo missing %q", s)
		}
	}
	// The legacy single-verifier attribute must not appear in new files.
	if bytes.Contains(data, []byte(`encryptedVerifierValue=`)) {
		t.Error("new files must not use the legacy encryptedVerifierValue attribute")
	}

	// 8-byte EncryptionInfo header: version 0x00040004 + reserved 0x40,
	// immediately followed by the XML declaration.
	idx := bytes.Index(data, []byte("<?xml"))
	if idx < 8 {
		t.Fatalf("XML declaration not found in container")
	}
	if v := binary.LittleEndian.Uint32(data[idx-8 : idx-4]); v != 0x00040004 {
		t.Errorf("version header = 0x%08X, want 0x00040004", v)
	}
	if v := binary.LittleEndian.Uint32(data[idx-4 : idx]); v != 0x40 {
		t.Errorf("reserved header = 0x%08X, want 0x00000040", v)
	}
}

// TestEncryption_LargeFileRoundTrip exercises the CFB DIFAT chain: a
// package above ~7 MiB needs more than 109 FAT sectors, which the header
// DIFAT array alone cannot index. Earlier versions panicked reading such
// files.
func TestEncryption_LargeFileRoundTrip(t *testing.T) {
	if testing.Short() {
		t.Skip("large-file encryption round-trip")
	}

	wb := cells_foss.NewWorkbook()
	cells := wb.Worksheets[0].Cells()
	cells.Set("A1", "BIG-START")
	// ~12.8 MiB of near-incompressible random printable text across 100k
	// cells; the deflated package (~10 MiB) forces the CFB container past
	// the 109-FAT-sector boundary where the DIFAT chain kicks in.
	rng := mrand.New(mrand.NewSource(42))
	for row := 2; row <= 100001; row++ {
		buf := make([]byte, 128)
		for k := range buf {
			buf[k] = byte(33 + rng.Intn(94))
		}
		cells.Set(fmt.Sprintf("B%d", row), string(buf))
	}
	cells.Set("C100002", "BIG-END")

	wb.SetPassword("bigpw")
	p := filepath.Join(t.TempDir(), "big.xlsx")
	if err := wb.Save(p); err != nil {
		t.Fatalf("Save: %v", err)
	}
	if fi, err := os.Stat(p); err != nil {
		t.Fatalf("Stat: %v", err)
	} else if fi.Size() < 7*1024*1024 {
		t.Errorf("container only %d bytes; expected > 7 MiB to trigger DIFAT chain", fi.Size())
	}

	// Wrong password must fail cleanly (no panic).
	if _, err := cells_foss.LoadWithPassword(p, "nope"); err == nil {
		t.Error("wrong password should fail")
	}

	loaded, err := cells_foss.LoadWithPassword(p, "bigpw")
	if err != nil {
		t.Fatalf("LoadWithPassword: %v", err)
	}
	lc := loaded.Worksheets[0].Cells()
	if c, err := lc.Get("A1"); err != nil || c.AsString() != "BIG-START" {
		t.Errorf("A1 = %v (%v)", c.AsString(), err)
	}
	if c, err := lc.Get("C100002"); err != nil || c.AsString() != "BIG-END" {
		t.Errorf("C100002 = %v (%v)", c.AsString(), err)
	}
	if c, err := lc.Get("B30000"); err != nil {
		t.Errorf("B30000 missing: %v", err)
	} else if len(c.AsString()) != 128 {
		t.Errorf("B30000 length = %d, want 128", len(c.AsString()))
	}
}

// TestEncryption_LegacyFileCompat builds a file with the pre-spec scheme
// older versions of this library wrote (ECRX container, big-endian hash
// iterator, single-key CBC with prefixed IV, encryptedVerifierValue
// attribute) and verifies it still loads.
func TestEncryption_LegacyFileCompat(t *testing.T) {
	// Plaintext: an ordinary saved workbook.
	wb := cells_foss.NewWorkbook()
	wb.Worksheets[0].Cells().Set("A1", "legacy-payload")
	plain := filepath.Join(t.TempDir(), "plain.xlsx")
	if err := wb.Save(plain); err != nil {
		t.Fatalf("Save: %v", err)
	}
	zipBytes, err := os.ReadFile(plain)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}

	const password = "legacypw"
	salt := make([]byte, 16)
	if _, err := io.ReadFull(rand.Reader, salt); err != nil {
		t.Fatalf("salt: %v", err)
	}

	// Legacy key derivation: UTF-16LE password, SHA-512, big-endian
	// iterator, direct truncation (no block key).
	pwd := make([]byte, 0, 2*len(password))
	for _, r := range password {
		var b [2]byte
		binary.LittleEndian.PutUint16(b[:], uint16(r))
		pwd = append(pwd, b[0], b[1])
	}
	h := sha512.New()
	h.Write(salt)
	h.Write(pwd)
	digest := h.Sum(nil)
	for i := uint32(0); i < 100000; i++ {
		var ib [4]byte
		binary.BigEndian.PutUint32(ib[:], i)
		h.Reset()
		h.Write(ib[:])
		h.Write(digest)
		digest = h.Sum(nil)
	}
	derived := digest[:32]

	// Legacy AES-CBC: random IV prefixed to output, PKCS#7 padding.
	legacyEncrypt := func(key, plaintext []byte) []byte {
		block, err := aes.NewCipher(key)
		if err != nil {
			t.Fatalf("aes: %v", err)
		}
		iv := make([]byte, 16)
		if _, err := io.ReadFull(rand.Reader, iv); err != nil {
			t.Fatalf("iv: %v", err)
		}
		pad := 16 - len(plaintext)%16
		padded := make([]byte, len(plaintext)+pad)
		copy(padded, plaintext)
		for i := len(plaintext); i < len(padded); i++ {
			padded[i] = byte(pad)
		}
		out := make([]byte, 16+len(padded))
		copy(out[:16], iv)
		cipher.NewCBCEncrypter(block, iv).CryptBlocks(out[16:], padded)
		return out
	}

	encKey := make([]byte, 32)
	if _, err := io.ReadFull(rand.Reader, encKey); err != nil {
		t.Fatalf("key: %v", err)
	}
	encVerifier := legacyEncrypt(derived, []byte("legacy-verifier"))
	encKeyValue := legacyEncrypt(derived, encKey)
	ciphertext := legacyEncrypt(encKey, zipBytes)

	b64 := base64.StdEncoding.EncodeToString
	infoXML := fmt.Sprintf(
		`<encryption><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="%s"/>`+
			`<keyEncryptors><keyEncryptor uri="{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}">`+
			`<encryptedKey saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="%s" encryptedKeyValue="%s" encryptedVerifierValue="%s"/>`+
			`</keyEncryptor></keyEncryptors></encryption>`,
		b64(salt), b64(salt), b64(encKeyValue), b64(encVerifier))

	// ECRX container: magic + info length + info XML + encrypted package.
	var buf bytes.Buffer
	binary.Write(&buf, binary.LittleEndian, uint32(0x58524345))
	binary.Write(&buf, binary.LittleEndian, uint32(len(infoXML)))
	buf.WriteString(infoXML)
	buf.Write(ciphertext)

	p := filepath.Join(t.TempDir(), "legacy.xlsx")
	if err := os.WriteFile(p, buf.Bytes(), 0o644); err != nil {
		t.Fatalf("WriteFile: %v", err)
	}

	if _, err := cells_foss.LoadWithPassword(p, "wrong"); err == nil {
		t.Error("legacy file: wrong password should fail")
	}
	loaded, err := cells_foss.LoadWithPassword(p, password)
	if err != nil {
		t.Fatalf("legacy LoadWithPassword: %v", err)
	}
	c, err := loaded.Worksheets[0].Cells().Get("A1")
	if err != nil || c.AsString() != "legacy-payload" {
		t.Errorf("A1 = %v (%v), want legacy-payload", c.AsString(), err)
	}
}
