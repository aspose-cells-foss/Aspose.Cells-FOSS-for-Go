package cells_foss

import (
	"bytes"
	"crypto/aes"
	"crypto/cipher"
	"crypto/rand"
	"crypto/sha512"
	"encoding/base64"
	"encoding/binary"
	"encoding/xml"
	"fmt"
	"io"
)

// ---------------------------------------------------------------------------
// ECMA-376 Agile Encryption constants
// ---------------------------------------------------------------------------

const (
	spinCount      = 100000 // SHA-512 iteration count
	saltSize       = 16     // bytes
	keyBits        = 256    // AES-256
	blockSize      = 16     // AES block size (128 bits)
	hashSize       = 64     // SHA-512 output size (bytes)
	encryptionInfoGUID = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}"
)

// ---------------------------------------------------------------------------
// EncryptionInfo XML types (Agile Encryption)
// ---------------------------------------------------------------------------

type xmlEncryptionInfo struct {
	XMLName        xml.Name          `xml:"encryption"`
	KeyData        xmlKeyData        `xml:"keyData"`
	KeyEncryptors  xmlKeyEncryptors  `xml:"keyEncryptors"`
}

type xmlKeyData struct {
	SaltSize        int    `xml:"saltSize,attr"`
	BlockSize       int    `xml:"blockSize,attr"`
	KeyBits         int    `xml:"keyBits,attr"`
	HashSize        int    `xml:"hashSize,attr"`
	CipherAlgorithm string `xml:"cipherAlgorithm,attr"`
	CipherChaining  string `xml:"cipherChaining,attr"`
	HashAlgorithm   string `xml:"hashAlgorithm,attr"`
	SaltValue       string `xml:"saltValue,attr"`
}

type xmlKeyEncryptors struct {
	KeyEncryptor xmlKeyEncryptor `xml:"keyEncryptor"`
}

type xmlKeyEncryptor struct {
	URI              string              `xml:"uri,attr"`
	EncryptedKey     xmlEncryptedKey     `xml:"encryptedKey"`
}

type xmlEncryptedKey struct {
	SpinCount            string `xml:"spinCount,attr"`
	SaltSize             string `xml:"saltSize,attr"`
	BlockSize            string `xml:"blockSize,attr"`
	KeyBits              string `xml:"keyBits,attr"`
	HashSize             string `xml:"hashSize,attr"`
	CipherAlgorithm      string `xml:"cipherAlgorithm,attr"`
	CipherChaining       string `xml:"cipherChaining,attr"`
	HashAlgorithm        string `xml:"hashAlgorithm,attr"`
	SaltValue            string `xml:"saltValue,attr"`
	EncryptedKeyValue    string `xml:"encryptedKeyValue,attr"`
	EncryptedVerifierValue string `xml:"encryptedVerifierValue,attr"`
}

// ---------------------------------------------------------------------------
// Key derivation (ECMA-376 §3.2.1)
// ---------------------------------------------------------------------------

// deriveKey hashes the password with the given salt using the ECMA-376
// iterative SHA-512 scheme.  It returns enough derived bytes to match keyBits.
func deriveKey(password string, salt []byte, keyBits int) []byte {
	// Convert password to UTF-16LE bytes.
	pwdBytes := encodeUTF16LE(password)

	// H₀ = SHA512(salt || password)
	h := sha512.New()
	h.Write(salt)
	h.Write(pwdBytes)
	hash := h.Sum(nil)

	// Iterate.
	for i := uint32(0); i < spinCount; i++ {
		var ibuf [4]byte
		binary.BigEndian.PutUint32(ibuf[:], i)
		h.Reset()
		h.Write(ibuf[:])
		h.Write(hash)
		hash = h.Sum(nil)
	}

	// If the hash is shorter than needed, continue with a different scheme.
	// For 256-bit keys, SHA-512 (64 bytes) is more than enough.
	need := keyBits / 8
	if need > len(hash) {
		// Append more bytes using the scheme:
		// Hn = SHA512(Hn-1 || salt || password)
		extra := make([]byte, 0, need)
		extra = append(extra, hash...)
		for len(extra) < need {
			h.Reset()
			h.Write(extra[len(extra)-hashSize:])
			h.Write(salt)
			h.Write(pwdBytes)
			extra = append(extra, h.Sum(nil)...)
		}
		return extra[:need]
	}

	return hash[:need]
}

// encodeUTF16LE converts a Go string to little-endian UTF-16.
func encodeUTF16LE(s string) []byte {
	var buf bytes.Buffer
	for _, r := range s {
		if r <= 0xFFFF {
			binary.Write(&buf, binary.LittleEndian, uint16(r))
		} else {
			r -= 0x10000
			binary.Write(&buf, binary.LittleEndian, uint16(0xD800|(r>>10)))
			binary.Write(&buf, binary.LittleEndian, uint16(0xDC00|(r&0x3FF)))
		}
	}
	return buf.Bytes()
}

// ---------------------------------------------------------------------------
// Encrypt / decrypt package
// ---------------------------------------------------------------------------

// encryptPackage encrypts plaintext using the Agile Encryption scheme.
// Returns the EncryptionInfo XML and the encrypted data.
func encryptPackage(plaintext []byte, password string) ([]byte, []byte, error) {
	// 1. Generate salt.
	salt := make([]byte, saltSize)
	if _, err := io.ReadFull(rand.Reader, salt); err != nil {
		return nil, nil, fmt.Errorf("encryption: salt: %w", err)
	}

	// 2. Derive key from password.
	derived := deriveKey(password, salt, keyBits)

	// 3. Generate random intermediate key and verifier.
	encKey := make([]byte, keyBits/8)
	if _, err := io.ReadFull(rand.Reader, encKey); err != nil {
		return nil, nil, fmt.Errorf("encryption: encKey: %w", err)
	}
	verifier := make([]byte, 16)
	if _, err := io.ReadFull(rand.Reader, verifier); err != nil {
		return nil, nil, fmt.Errorf("encryption: verifier: %w", err)
	}

	// 4. Encrypt the intermediate key with the derived key.
	encryptedKeyValue, err := aesCBCEncrypt(derived, encKey)
	if err != nil {
		return nil, nil, fmt.Errorf("encryption: encrypt key: %w", err)
	}
	encryptedVerifier, err := aesCBCEncrypt(derived, verifier)
	if err != nil {
		return nil, nil, fmt.Errorf("encryption: encrypt verifier: %w", err)
	}

	// 5. Encrypt the package data with the intermediate key.
	encryptedData, err := aesCBCEncrypt(encKey, plaintext)
	if err != nil {
		return nil, nil, fmt.Errorf("encryption: encrypt package: %w", err)
	}

	// 6. Build EncryptionInfo XML.
	infoXML := buildEncryptionInfoXML(salt, encryptedKeyValue, encryptedVerifier)

	return infoXML, encryptedData, nil
}

// decryptPackage decrypts ciphertext using the password and the
// EncryptionInfo XML.  Returns the decrypted plaintext.
func decryptPackage(infoXML, ciphertext []byte, password string) ([]byte, error) {
	// 1. Parse EncryptionInfo.
	var info xmlEncryptionInfo
	if err := xml.Unmarshal(infoXML, &info); err != nil {
		return nil, fmt.Errorf("decryption: cannot parse EncryptionInfo: %w", err)
	}

	// 2. Extract salt.
	salt, err := base64Decode(info.KeyEncryptors.KeyEncryptor.EncryptedKey.SaltValue)
	if err != nil {
		return nil, fmt.Errorf("decryption: salt: %w", err)
	}

	// 3. Derive key.
	derived := deriveKey(password, salt, keyBits)

	// 4. Decrypt the intermediate key.
	encKeyB64 := info.KeyEncryptors.KeyEncryptor.EncryptedKey.EncryptedKeyValue
	encKeyBytes, err := base64Decode(encKeyB64)
	if err != nil {
		return nil, fmt.Errorf("decryption: encryptedKeyValue: %w", err)
	}

	// 5. Decrypt and verify the verifier.
	verifierB64 := info.KeyEncryptors.KeyEncryptor.EncryptedKey.EncryptedVerifierValue
	verifierBytes, err := base64Decode(verifierB64)
	if err != nil {
		return nil, fmt.Errorf("decryption: verifier: %w", err)
	}
	_, err = aesCBCDecrypt(derived, verifierBytes)
	if err != nil {
		return nil, fmt.Errorf("decryption: invalid password")
	}

	// 6. Decrypt intermediate key.
	encKey, err := aesCBCDecrypt(derived, encKeyBytes)
	if err != nil {
		return nil, fmt.Errorf("decryption: cannot decrypt key: %w", err)
	}

	// 7. Decrypt package.
	plaintext, err := aesCBCDecrypt(encKey, ciphertext)
	if err != nil {
		return nil, fmt.Errorf("decryption: cannot decrypt package: %w", err)
	}

	return plaintext, nil
}

// ---------------------------------------------------------------------------
// AES-256-CBC helpers
// ---------------------------------------------------------------------------

func aesCBCEncrypt(key, plaintext []byte) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}

	// Generate random IV (stored as first block of output).
	iv := make([]byte, blockSize)
	if _, err := io.ReadFull(rand.Reader, iv); err != nil {
		return nil, err
	}

	// PKCS#7 padding.
	padded := pkcs7Pad(plaintext, blockSize)

	ciphertext := make([]byte, len(iv)+len(padded))
	copy(ciphertext[:blockSize], iv)

	mode := cipher.NewCBCEncrypter(block, iv)
	mode.CryptBlocks(ciphertext[blockSize:], padded)

	return ciphertext, nil
}

func aesCBCDecrypt(key, ciphertext []byte) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}

	if len(ciphertext) < blockSize {
		return nil, fmt.Errorf("ciphertext too short")
	}

	iv := ciphertext[:blockSize]
	data := ciphertext[blockSize:]

	if len(data)%blockSize != 0 {
		return nil, fmt.Errorf("ciphertext not block-aligned")
	}

	plain := make([]byte, len(data))
	mode := cipher.NewCBCDecrypter(block, iv)
	mode.CryptBlocks(plain, data)

	return pkcs7Unpad(plain)
}

func pkcs7Pad(data []byte, blockSize int) []byte {
	padLen := blockSize - len(data)%blockSize
	padded := make([]byte, len(data)+padLen)
	copy(padded, data)
	for i := len(data); i < len(padded); i++ {
		padded[i] = byte(padLen)
	}
	return padded
}

func pkcs7Unpad(data []byte) ([]byte, error) {
	if len(data) == 0 {
		return nil, fmt.Errorf("empty data")
	}
	padLen := int(data[len(data)-1])
	if padLen > len(data) || padLen > blockSize || padLen == 0 {
		return nil, fmt.Errorf("invalid padding")
	}
	for i := len(data) - padLen; i < len(data); i++ {
		if data[i] != byte(padLen) {
			return nil, fmt.Errorf("invalid padding")
		}
	}
	return data[:len(data)-padLen], nil
}

// ---------------------------------------------------------------------------
// EncryptionInfo XML builder
// ---------------------------------------------------------------------------

func buildEncryptionInfoXML(salt, encryptedKey, encryptedVerifier []byte) []byte {
	info := xmlEncryptionInfo{
		KeyData: xmlKeyData{
			SaltSize:        saltSize,
			BlockSize:       blockSize,
			KeyBits:         keyBits,
			HashSize:        hashSize,
			CipherAlgorithm: "AES",
			CipherChaining:  "ChainingModeCBC",
			HashAlgorithm:   "SHA512",
			SaltValue:       base64Encode(salt),
		},
		KeyEncryptors: xmlKeyEncryptors{
			KeyEncryptor: xmlKeyEncryptor{
				URI: encryptionInfoGUID,
				EncryptedKey: xmlEncryptedKey{
					SpinCount:              fmt.Sprintf("%d", spinCount),
					SaltSize:               fmt.Sprintf("%d", saltSize),
					BlockSize:              fmt.Sprintf("%d", blockSize),
					KeyBits:                fmt.Sprintf("%d", keyBits),
					HashSize:               fmt.Sprintf("%d", hashSize),
					CipherAlgorithm:        "AES",
					CipherChaining:         "ChainingModeCBC",
					HashAlgorithm:          "SHA512",
					SaltValue:              base64Encode(salt),
					EncryptedKeyValue:      base64Encode(encryptedKey),
					EncryptedVerifierValue: base64Encode(encryptedVerifier),
				},
			},
		},
	}

	body, err := xml.MarshalIndent(info, "", "  ")
	if err != nil {
		// Should never happen.
		panic(fmt.Sprintf("encryption: marshal EncryptionInfo: %v", err))
	}
	out := make([]byte, 0, len(xml.Header)+len(body)+1)
	out = append(out, xml.Header...)
	out = append(out, body...)
	out = append(out, '\n')
	return out
}

// ---------------------------------------------------------------------------
// Base64 helpers (standard encoding)
// ---------------------------------------------------------------------------

func base64Encode(data []byte) string {
	return base64.StdEncoding.EncodeToString(data)
}

func base64Decode(s string) ([]byte, error) {
	return base64.StdEncoding.DecodeString(s)
}

// ---------------------------------------------------------------------------
// Simple binary container for encrypted files
// ---------------------------------------------------------------------------
// We use a minimal header so that LoadWorkbook can detect encrypted files:
//
//	Offset  Size  Field
//	  0      4    Magic: "ECRX" (0x58524345 LE)
//	  4      4    EncryptionInfo length (uint32 LE)
//	  8      N    EncryptionInfo XML bytes
//	8+N     M    Encrypted package bytes
// ---------------------------------------------------------------------------

const encryptedFileMagic = 0x58524345 // "ECRX" in little-endian

func isEncryptedFile(data []byte) bool {
	if len(data) < 8 {
		return false
	}
	return binary.LittleEndian.Uint32(data[:4]) == encryptedFileMagic
}

func readEncryptedFile(data []byte) (infoXML []byte, encryptedPackage []byte, err error) {
	if !isEncryptedFile(data) {
		return nil, nil, fmt.Errorf("not an encrypted file")
	}
	infoLen := binary.LittleEndian.Uint32(data[4:8])
	start := int(8)
	end := start + int(infoLen)
	if end > len(data) {
		return nil, nil, fmt.Errorf("truncated encrypted file")
	}
	infoXML = data[start:end]
	encryptedPackage = data[end:]
	return infoXML, encryptedPackage, nil
}

func writeEncryptedFile(infoXML, encryptedPackage []byte) []byte {
	var buf bytes.Buffer
	var magic [4]byte
	binary.LittleEndian.PutUint32(magic[:], encryptedFileMagic)
	buf.Write(magic[:])

	var lenBuf [4]byte
	binary.LittleEndian.PutUint32(lenBuf[:], uint32(len(infoXML)))
	buf.Write(lenBuf[:])

	buf.Write(infoXML)
	buf.Write(encryptedPackage)
	return buf.Bytes()
}
