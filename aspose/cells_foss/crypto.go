package cells_foss

import (
	"bytes"
	"crypto/aes"
	"crypto/cipher"
	"crypto/hmac"
	"crypto/rand"
	"crypto/sha1"
	"crypto/sha256"
	"crypto/sha512"
	"crypto/subtle"
	"encoding/base64"
	"encoding/binary"
	"encoding/xml"
	"fmt"
	"hash"
	"io"
	"strconv"
	"strings"
)

// ---------------------------------------------------------------------------
// ECMA-376 / MS-OFFCRYPTO Agile Encryption
//
// Files produced by SetPassword use the Agile scheme exactly as Microsoft
// Excel writes it, so they can be opened by Excel:
//
//   - EncryptionInfo stream: 4-byte version (0x00040004) + 4-byte reserved
//     + namespaced XML describing keyData and one password keyEncryptor.
//   - EncryptedPackage stream: 8-byte decrypted-size header + ciphertext
//     produced in 4096-byte segments, each with a key and IV derived from
//     the segment index.
//
// Files written by earlier versions of this library used a simplified,
// non-spec scheme (single-key CBC with a prefixed random IV and a
// big-endian hash iterator). Loading those files is still supported; the
// two schemes are told apart by the EncryptionInfo attributes.
//
// Reference: [MS-OFFCRYPTO] 2.3.4 "Agile Encryption"
// ---------------------------------------------------------------------------

const (
	spinCount               = 100000 // SHA-512 iteration count
	saltSize                = 16     // bytes
	keyBits                 = 256    // AES-256
	blockSize               = 16     // AES block size (128 bits)
	hashSize                = 64     // SHA-512 output size (bytes)
	agileSegmentSize        = 4096   // package encryption segment size
	agileVersion            = 0x00040004
	passwordKeyEncryptorURI = "http://schemas.microsoft.com/office/2006/keyEncryptor/password"
)

// Agile block keys ([MS-OFFCRYPTO] 2.3.4.11–2.3.4.16). The spec states them
// as 64-bit constants; on the wire they appear most-significant byte first,
// matching the reference implementations Excel interops with.
var (
	blockKeyVerifierHashInput = []byte{0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79}
	blockKeyVerifierHashValue = []byte{0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e}
	blockKeyEncryptedKeyValue = []byte{0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6}
	blockKeyDataIntegrityKey  = []byte{0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6}
	blockKeyDataIntegrityVal  = []byte{0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33}
)

// ---------------------------------------------------------------------------
// EncryptionInfo XML types (parsing; namespace-less tags match by local name
// so both namespaced spec files and legacy unnamespaced files decode)
// ---------------------------------------------------------------------------

type xmlEncryptionInfo struct {
	XMLName       xml.Name         `xml:"encryption"`
	KeyData       xmlKeyData       `xml:"keyData"`
	DataIntegrity xmlDataIntegrity `xml:"dataIntegrity"`
	KeyEncryptors xmlKeyEncryptors `xml:"keyEncryptors"`
}

type xmlDataIntegrity struct {
	EncryptedHmacKey   string `xml:"encryptedHmacKey,attr"`
	EncryptedHmacValue string `xml:"encryptedHmacValue,attr"`
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
	URI          string          `xml:"uri,attr"`
	EncryptedKey xmlEncryptedKey `xml:"encryptedKey"`
}

type xmlEncryptedKey struct {
	SpinCount                string `xml:"spinCount,attr"`
	SaltSize                 string `xml:"saltSize,attr"`
	BlockSize                string `xml:"blockSize,attr"`
	KeyBits                  string `xml:"keyBits,attr"`
	HashSize                 string `xml:"hashSize,attr"`
	CipherAlgorithm          string `xml:"cipherAlgorithm,attr"`
	CipherChaining           string `xml:"cipherChaining,attr"`
	HashAlgorithm            string `xml:"hashAlgorithm,attr"`
	SaltValue                string `xml:"saltValue,attr"`
	EncryptedKeyValue        string `xml:"encryptedKeyValue,attr"`
	EncryptedVerifierValue   string `xml:"encryptedVerifierValue,attr"`     // legacy scheme only
	EncryptedVerifierInput   string `xml:"encryptedVerifierHashInput,attr"` // Agile scheme
	EncryptedVerifierHashVal string `xml:"encryptedVerifierHashValue,attr"` // Agile scheme
}

// ---------------------------------------------------------------------------
// Hash helpers
// ---------------------------------------------------------------------------

// agileHash resolves the hashAlgorithm attribute to a constructor. An empty
// name defaults to SHA-512, what this library has always written.
func agileHash(name string) (func() hash.Hash, error) {
	switch strings.ToUpper(strings.TrimSpace(name)) {
	case "", "SHA512":
		return sha512.New, nil
	case "SHA384":
		return sha512.New384, nil
	case "SHA256":
		return sha256.New, nil
	case "SHA1":
		return sha1.New, nil
	}
	return nil, fmt.Errorf("unsupported hash algorithm %q", name)
}

// agileSum hashes the concatenation of parts.
func agileSum(hf func() hash.Hash, parts ...[]byte) []byte {
	h := hf()
	for _, p := range parts {
		h.Write(p)
	}
	return h.Sum(nil)
}

// fitKey truncates digest to size bytes, or pads it with 0x36 when shorter
// ([MS-OFFCRYPTO] 2.3.4.11 key derivation final step).
func fitKey(digest []byte, size int) []byte {
	if len(digest) >= size {
		return digest[:size]
	}
	out := make([]byte, size)
	copy(out, digest)
	for i := len(digest); i < size; i++ {
		out[i] = 0x36
	}
	return out
}

// le32 returns i as 4 little-endian bytes (Agile segment/iterator block keys).
func le32(i int) []byte {
	b := make([]byte, 4)
	binary.LittleEndian.PutUint32(b, uint32(i))
	return b
}

// ---------------------------------------------------------------------------
// Key derivation
// ---------------------------------------------------------------------------

// deriveKeyAgile implements the [MS-OFFCRYPTO] 2.3.4.11 password key
// derivation: an iterative little-endian hash of the password followed by a
// block-key-specific final hash.
func deriveKeyAgile(password string, salt []byte, spin int, hf func() hash.Hash, keyBytes int, blockKey []byte) []byte {
	h := hf()
	h.Write(salt)
	h.Write(encodeUTF16LE(password))
	digest := h.Sum(nil)

	for i := 0; i < spin; i++ {
		h.Reset()
		h.Write(le32(i))
		h.Write(digest)
		digest = h.Sum(nil)
	}

	return fitKey(agileSum(hf, digest, blockKey), keyBytes)
}

// deriveKey is the LEGACY (pre-spec) derivation used by files written by
// earlier versions of this library: big-endian iterator, no block key.
// Kept only so those files remain loadable.
func deriveKey(password string, salt []byte, keyBits int) []byte {
	pwdBytes := encodeUTF16LE(password)

	h := sha512.New()
	h.Write(salt)
	h.Write(pwdBytes)
	hashv := h.Sum(nil)

	for i := uint32(0); i < spinCount; i++ {
		var ibuf [4]byte
		binary.BigEndian.PutUint32(ibuf[:], i)
		h.Reset()
		h.Write(ibuf[:])
		h.Write(hashv)
		hashv = h.Sum(nil)
	}

	return hashv[:keyBits/8]
}

// encodeUTF16LE converts a Go string to little-endian UTF-16 without a
// null terminator.
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
// AES-CBC helpers
// ---------------------------------------------------------------------------

// cbcEncrypt encrypts data with an explicit IV. data must already be a
// multiple of the AES block size; Agile pads with random bytes rather than
// PKCS#7.
func cbcEncrypt(key, iv, data []byte) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	if len(data)%aes.BlockSize != 0 {
		return nil, fmt.Errorf("plaintext not block-aligned")
	}
	out := make([]byte, len(data))
	cipher.NewCBCEncrypter(block, iv).CryptBlocks(out, data)
	return out, nil
}

// cbcDecrypt decrypts data with an explicit IV without stripping padding.
func cbcDecrypt(key, iv, data []byte) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	if len(data) == 0 || len(data)%aes.BlockSize != 0 {
		return nil, fmt.Errorf("ciphertext not block-aligned")
	}
	out := make([]byte, len(data))
	cipher.NewCBCDecrypter(block, iv).CryptBlocks(out, data)
	return out, nil
}

// padRandom appends random bytes so the result is a multiple of blockSize.
// Data already aligned is returned with a full block of padding NOT added —
// Agile keeps aligned plaintexts as-is.
func padRandom(data []byte, size int) ([]byte, error) {
	rem := len(data) % size
	if rem == 0 {
		return data, nil
	}
	out := make([]byte, len(data)+size-rem)
	copy(out, data)
	if _, err := io.ReadFull(rand.Reader, out[len(data):]); err != nil {
		return nil, err
	}
	return out, nil
}

// aesCBCEncrypt is the LEGACY scheme: random IV prefixed to the output plus
// PKCS#7 padding. Kept for reading files written by earlier versions.
func aesCBCEncrypt(key, plaintext []byte) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	iv := make([]byte, blockSize)
	if _, err := io.ReadFull(rand.Reader, iv); err != nil {
		return nil, err
	}
	padded := pkcs7Pad(plaintext, blockSize)
	ciphertext := make([]byte, len(iv)+len(padded))
	copy(ciphertext[:blockSize], iv)
	cipher.NewCBCEncrypter(block, iv).CryptBlocks(ciphertext[blockSize:], padded)
	return ciphertext, nil
}

// aesCBCDecrypt is the LEGACY counterpart of aesCBCEncrypt.
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
	cipher.NewCBCDecrypter(block, iv).CryptBlocks(plain, data)
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
// Package encryption / decryption
// ---------------------------------------------------------------------------

// encryptPackage encrypts plaintext with the Agile scheme. It returns the
// EncryptionInfo XML, the segmented ciphertext, and the original size (which
// the container stores so decryption can strip the final segment's padding).
func encryptPackage(plaintext []byte, password string) ([]byte, []byte, uint64, error) {
	hf := sha512.New

	// keyData salt drives the per-segment IVs; the keyEncryptor salt drives
	// the password-derived keys. Excel uses two independent salts.
	pkgSalt := make([]byte, saltSize)
	pwSalt := make([]byte, saltSize)
	if _, err := io.ReadFull(rand.Reader, pkgSalt); err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: salt: %w", err)
	}
	if _, err := io.ReadFull(rand.Reader, pwSalt); err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: salt: %w", err)
	}

	// Random secret key encrypts the package; the password only wraps it.
	secretKey := make([]byte, keyBits/8)
	if _, err := io.ReadFull(rand.Reader, secretKey); err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: secret key: %w", err)
	}

	// Password verifier: random input, its hash, both encrypted with
	// block-key-specific derivations of the password.
	verifierInput := make([]byte, saltSize)
	if _, err := io.ReadFull(rand.Reader, verifierInput); err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: verifier: %w", err)
	}
	verifierHash := agileSum(hf, verifierInput)

	keyVerIn := deriveKeyAgile(password, pwSalt, spinCount, hf, keyBits/8, blockKeyVerifierHashInput)
	keyVerVal := deriveKeyAgile(password, pwSalt, spinCount, hf, keyBits/8, blockKeyVerifierHashValue)
	keyKeyVal := deriveKeyAgile(password, pwSalt, spinCount, hf, keyBits/8, blockKeyEncryptedKeyValue)

	encVerIn, err := cbcEncrypt(keyVerIn, pwSalt, verifierInput)
	if err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: verifier input: %w", err)
	}
	encVerValPadded, err := padRandom(verifierHash, blockSize)
	if err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: verifier value: %w", err)
	}
	encVerVal, err := cbcEncrypt(keyVerVal, pwSalt, encVerValPadded)
	if err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: verifier value: %w", err)
	}
	encKeyValue, err := cbcEncrypt(keyKeyVal, pwSalt, secretKey)
	if err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: key value: %w", err)
	}

	// Segment-wise package encryption ([MS-OFFCRYPTO] 2.3.4.15).
	ciphertext, err := encryptSegments(plaintext, secretKey, pkgSalt, hf)
	if err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: package: %w", err)
	}

	// Data integrity values ([MS-OFFCRYPTO] 2.3.4.16): an HMAC over the
	// complete EncryptedPackage stream (8-byte size header + ciphertext),
	// with the HMAC key and value themselves encrypted under the secret key.
	encHmacKey, encHmacVal, err := buildDataIntegrity(plaintext, ciphertext, secretKey, pkgSalt, hf)
	if err != nil {
		return nil, nil, 0, fmt.Errorf("encryption: integrity: %w", err)
	}

	infoXML := buildEncryptionInfoXML(pkgSalt, pwSalt, encVerIn, encVerVal, encKeyValue, encHmacKey, encHmacVal)
	return infoXML, ciphertext, uint64(len(plaintext)), nil
}

// buildDataIntegrity computes the encryptedHmacKey/encryptedHmacValue pair.
// The IVs are H(keyDataSalt||blockKey) truncated to the block size; both
// values are CBC-encrypted with the secret key.
func buildDataIntegrity(plaintext, ciphertext, secretKey, pkgSalt []byte, hf func() hash.Hash) (encHmacKey, encHmacVal []byte, err error) {
	hmacKey := make([]byte, hashSize)
	if _, err := io.ReadFull(rand.Reader, hmacKey); err != nil {
		return nil, nil, err
	}
	iv1 := fitKey(agileSum(hf, pkgSalt, blockKeyDataIntegrityKey), blockSize)
	iv2 := fitKey(agileSum(hf, pkgSalt, blockKeyDataIntegrityVal), blockSize)

	encHmacKey, err = cbcEncrypt(secretKey, iv1, hmacKey)
	if err != nil {
		return nil, nil, err
	}

	mac := hmac.New(hf, hmacKey)
	var sizeHdr [8]byte
	binary.LittleEndian.PutUint64(sizeHdr[:], uint64(len(plaintext)))
	mac.Write(sizeHdr[:])
	mac.Write(ciphertext)
	hmacVal := mac.Sum(nil)

	encHmacVal, err = cbcEncrypt(secretKey, iv2, hmacVal)
	if err != nil {
		return nil, nil, err
	}
	return encHmacKey, encHmacVal, nil
}

// verifyDataIntegrity recomputes the HMAC over the EncryptedPackage stream
// and compares it with the stored value.
func verifyDataIntegrity(di xmlDataIntegrity, origSize uint64, ciphertext, secretKey, pkgSalt []byte, hf func() hash.Hash) error {
	encHmacKey, err := base64Decode(di.EncryptedHmacKey)
	if err != nil {
		return fmt.Errorf("encryptedHmacKey: %w", err)
	}
	encHmacVal, err := base64Decode(di.EncryptedHmacValue)
	if err != nil {
		return fmt.Errorf("encryptedHmacValue: %w", err)
	}
	iv1 := fitKey(agileSum(hf, pkgSalt, blockKeyDataIntegrityKey), blockSize)
	iv2 := fitKey(agileSum(hf, pkgSalt, blockKeyDataIntegrityVal), blockSize)

	hmacKey, err := cbcDecrypt(secretKey, iv1, encHmacKey)
	if err != nil {
		return err
	}
	hmacVal, err := cbcDecrypt(secretKey, iv2, encHmacVal)
	if err != nil {
		return err
	}

	mac := hmac.New(hf, hmacKey)
	var sizeHdr [8]byte
	binary.LittleEndian.PutUint64(sizeHdr[:], origSize)
	mac.Write(sizeHdr[:])
	mac.Write(ciphertext)
	if !hmac.Equal(mac.Sum(nil), hmacVal) {
		return fmt.Errorf("HMAC mismatch")
	}
	return nil
}

// encryptSegments encrypts plaintext in 4096-byte segments
// ([MS-OFFCRYPTO] 2.3.4.15). Every segment uses the secret key directly;
// only the IV varies: IV(i) = H(pkgSalt||LE32(i)) truncated to the block
// size. The final segment is padded to the AES block size.
func encryptSegments(plaintext, secretKey, pkgSalt []byte, hf func() hash.Hash) ([]byte, error) {
	var out []byte
	for i, off := 0, 0; off < len(plaintext) || i == 0; i++ {
		end := off + agileSegmentSize
		if end > len(plaintext) {
			end = len(plaintext)
		}
		seg, err := padRandom(plaintext[off:end], blockSize)
		if err != nil {
			return nil, err
		}
		iv := fitKey(agileSum(hf, pkgSalt, le32(i)), blockSize)
		enc, err := cbcEncrypt(secretKey, iv, seg)
		if err != nil {
			return nil, err
		}
		out = append(out, enc...)
		if end >= len(plaintext) {
			break
		}
		off = end
	}
	return out, nil
}

// decryptPackage decrypts an EncryptedPackage payload using the password and
// the EncryptionInfo XML. origSize is the declared plaintext size from the
// container (Agile scheme); it may be 0 for legacy containers, which carry
// their exact length in PKCS#7 padding instead.
func decryptPackage(infoXML, ciphertext []byte, password string, origSize uint64) ([]byte, error) {
	var info xmlEncryptionInfo
	if err := xml.Unmarshal(infoXML, &info); err != nil {
		return nil, fmt.Errorf("decryption: cannot parse EncryptionInfo: %w", err)
	}
	ek := info.KeyEncryptors.KeyEncryptor.EncryptedKey
	if ek.EncryptedVerifierInput != "" && ek.EncryptedVerifierHashVal != "" {
		return decryptPackageAgile(info, ciphertext, password, origSize)
	}
	return decryptPackageLegacy(ek, ciphertext, password)
}

// decryptPackageAgile implements the [MS-OFFCRYPTO] Agile decryption,
// including the password check via the verifier hash pair.
func decryptPackageAgile(info xmlEncryptionInfo, ciphertext []byte, password string, origSize uint64) ([]byte, error) {
	kd, ek := info.KeyData, info.KeyEncryptors.KeyEncryptor.EncryptedKey
	if !strings.EqualFold(strings.TrimSpace(kd.CipherAlgorithm), "AES") {
		return nil, fmt.Errorf("decryption: unsupported cipher %q", kd.CipherAlgorithm)
	}
	if kd.CipherChaining != "" && kd.CipherChaining != "ChainingModeCBC" {
		return nil, fmt.Errorf("decryption: unsupported chaining %q", kd.CipherChaining)
	}

	hf, err := agileHash(kd.HashAlgorithm)
	if err != nil {
		return nil, fmt.Errorf("decryption: %w", err)
	}
	ekHF, err := agileHash(ek.HashAlgorithm)
	if err != nil {
		ekHF = hf
	}

	pkgSalt, err := base64Decode(kd.SaltValue)
	if err != nil {
		return nil, fmt.Errorf("decryption: keyData salt: %w", err)
	}
	pwSalt, err := base64Decode(ek.SaltValue)
	if err != nil {
		return nil, fmt.Errorf("decryption: keyEncryptor salt: %w", err)
	}

	spin, err := strconv.Atoi(strings.TrimSpace(ek.SpinCount))
	if err != nil || spin <= 0 {
		spin = spinCount
	}
	ekKeyBytes, err := strconv.Atoi(strings.TrimSpace(ek.KeyBits))
	if err != nil || ekKeyBytes <= 0 {
		ekKeyBytes = keyBits
	}
	ekKeyBytes /= 8

	pkgKeyBits := kd.KeyBits
	if pkgKeyBits <= 0 {
		pkgKeyBits = keyBits
	}

	iv := fitKey(pwSalt, blockSize)
	keyVerIn := deriveKeyAgile(password, pwSalt, spin, ekHF, ekKeyBytes, blockKeyVerifierHashInput)
	keyVerVal := deriveKeyAgile(password, pwSalt, spin, ekHF, ekKeyBytes, blockKeyVerifierHashValue)
	keyKeyVal := deriveKeyAgile(password, pwSalt, spin, ekHF, ekKeyBytes, blockKeyEncryptedKeyValue)

	encVerIn, err := base64Decode(ek.EncryptedVerifierInput)
	if err != nil {
		return nil, fmt.Errorf("decryption: verifier input: %w", err)
	}
	encVerVal, err := base64Decode(ek.EncryptedVerifierHashVal)
	if err != nil {
		return nil, fmt.Errorf("decryption: verifier value: %w", err)
	}
	verIn, err := cbcDecrypt(keyVerIn, iv, encVerIn)
	if err != nil {
		return nil, fmt.Errorf("decryption: invalid password")
	}
	verVal, err := cbcDecrypt(keyVerVal, iv, encVerVal)
	if err != nil {
		return nil, fmt.Errorf("decryption: invalid password")
	}

	// Password check: hash the decrypted verifier input and compare with the
	// decrypted verifier hash (constant time). A wrong password yields
	// garbage in both and fails here rather than producing random plaintext.
	inputLen := saltSize
	if kd.SaltSize > 0 {
		inputLen = kd.SaltSize
	}
	if len(verIn) < inputLen {
		return nil, fmt.Errorf("decryption: invalid password")
	}
	wantLen := ekHF().Size()
	expected := agileSum(ekHF, verIn[:inputLen])
	if len(verVal) < wantLen || subtle.ConstantTimeCompare(expected[:wantLen], verVal[:wantLen]) != 1 {
		return nil, fmt.Errorf("decryption: invalid password")
	}

	encKeyValue, err := base64Decode(ek.EncryptedKeyValue)
	if err != nil {
		return nil, fmt.Errorf("decryption: encryptedKeyValue: %w", err)
	}
	secretKey, err := cbcDecrypt(keyKeyVal, iv, encKeyValue)
	if err != nil {
		return nil, fmt.Errorf("decryption: cannot decrypt key: %w", err)
	}
	keyLen := pkgKeyBits / 8
	if keyLen > len(secretKey) {
		keyLen = len(secretKey)
	}
	if keyLen != 16 && keyLen != 24 && keyLen != 32 {
		return nil, fmt.Errorf("decryption: invalid key size %d bytes", keyLen)
	}
	secretKey = secretKey[:keyLen]

	// Payload integrity check when the file carries dataIntegrity values.
	if info.DataIntegrity.EncryptedHmacKey != "" && info.DataIntegrity.EncryptedHmacValue != "" && origSize > 0 {
		if err := verifyDataIntegrity(info.DataIntegrity, origSize, ciphertext, secretKey, pkgSalt, hf); err != nil {
			return nil, fmt.Errorf("decryption: integrity check failed: %w", err)
		}
	}

	plaintext, err := decryptSegments(ciphertext, secretKey, pkgSalt, hf)
	if err != nil {
		return nil, fmt.Errorf("decryption: %w", err)
	}
	if origSize > 0 && uint64(len(plaintext)) >= origSize {
		plaintext = plaintext[:origSize]
	}
	return plaintext, nil
}

// decryptSegments reverses encryptSegments.
func decryptSegments(ciphertext, secretKey, pkgSalt []byte, hf func() hash.Hash) ([]byte, error) {
	var out []byte
	for i, off := 0, 0; off < len(ciphertext); i++ {
		end := off + agileSegmentSize
		if end > len(ciphertext) {
			end = len(ciphertext)
		}
		seg := ciphertext[off:end]
		if len(seg)%blockSize != 0 {
			return nil, fmt.Errorf("segment %d not block-aligned", i)
		}
		iv := fitKey(agileSum(hf, pkgSalt, le32(i)), blockSize)
		plain, err := cbcDecrypt(secretKey, iv, seg)
		if err != nil {
			return nil, err
		}
		out = append(out, plain...)
		off = end
	}
	return out, nil
}

// decryptPackageLegacy decrypts files written by earlier versions of this
// library: big-endian key derivation, single-key CBC with a prefixed IV and
// PKCS#7 padding.
func decryptPackageLegacy(ek xmlEncryptedKey, ciphertext []byte, password string) ([]byte, error) {
	salt, err := base64Decode(ek.SaltValue)
	if err != nil {
		return nil, fmt.Errorf("decryption: salt: %w", err)
	}

	derived := deriveKey(password, salt, keyBits)

	encKeyBytes, err := base64Decode(ek.EncryptedKeyValue)
	if err != nil {
		return nil, fmt.Errorf("decryption: encryptedKeyValue: %w", err)
	}

	// The verifier decrypt must unpad cleanly; a wrong password virtually
	// always fails here.
	if _, err := aesCBCDecrypt(derived, mustDecodeOrEmpty(ek.EncryptedVerifierValue)); err != nil {
		return nil, fmt.Errorf("decryption: invalid password")
	}

	encKey, err := aesCBCDecrypt(derived, encKeyBytes)
	if err != nil {
		return nil, fmt.Errorf("decryption: cannot decrypt key: %w", err)
	}

	plaintext, err := aesCBCDecrypt(encKey, ciphertext)
	if err != nil {
		return nil, fmt.Errorf("decryption: cannot decrypt package: %w", err)
	}
	return plaintext, nil
}

func mustDecodeOrEmpty(s string) []byte {
	b, err := base64Decode(s)
	if err != nil {
		return nil
	}
	return b
}

// ---------------------------------------------------------------------------
// EncryptionInfo XML builder (spec layout with namespaces, as Excel writes)
// ---------------------------------------------------------------------------

func buildEncryptionInfoXML(pkgSalt, pwSalt, encVerIn, encVerVal, encKeyValue, encHmacKey, encHmacVal []byte) []byte {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ` +
		`xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate" ` +
		`xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` + "\n")
	fmt.Fprintf(&b, `  <keyData saltSize="%d" blockSize="%d" keyBits="%d" hashSize="%d" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="%s"/>`+"\n",
		saltSize, blockSize, keyBits, hashSize, base64Encode(pkgSalt))
	fmt.Fprintf(&b, `  <dataIntegrity encryptedHmacKey="%s" encryptedHmacValue="%s"/>`+"\n",
		base64Encode(encHmacKey), base64Encode(encHmacVal))
	b.WriteString(`  <keyEncryptors>` + "\n")
	fmt.Fprintf(&b, `    <keyEncryptor uri="%s">`+"\n", passwordKeyEncryptorURI)
	fmt.Fprintf(&b, `      <p:encryptedKey saltSize="%d" blockSize="%d" keyBits="%d" hashSize="%d" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="%s" spinCount="%d" encryptedVerifierHashInput="%s" encryptedVerifierHashValue="%s" encryptedKeyValue="%s"/>`+"\n",
		saltSize, blockSize, keyBits, hashSize, base64Encode(pwSalt), spinCount,
		base64Encode(encVerIn), base64Encode(encVerVal), base64Encode(encKeyValue))
	b.WriteString(`    </keyEncryptor>` + "\n")
	b.WriteString(`  </keyEncryptors>` + "\n")
	b.WriteString(`</encryption>` + "\n")
	return []byte(b.String())
}

// ---------------------------------------------------------------------------
// Base64 helpers (standard encoding)
// ---------------------------------------------------------------------------

func base64Encode(data []byte) string {
	return base64.StdEncoding.EncodeToString(data)
}

func base64Decode(s string) ([]byte, error) {
	return base64.StdEncoding.DecodeString(strings.TrimSpace(s))
}

// ---------------------------------------------------------------------------
// Simple binary container for legacy encrypted files
// ---------------------------------------------------------------------------
// Earlier versions wrapped Agile data in a minimal header so that
// LoadWorkbook could detect encrypted files:
//
//	Offset  Size  Field
//	  0      4    Magic: "ECRX" (0x58524345 LE)
//	  4      4    EncryptionInfo length (uint32 LE)
//	  8      N    EncryptionInfo XML bytes
//	8+N     M    Encrypted package bytes
//
// Such files can still be loaded; new files always use the OLE/CFB
// container (see olecfb.go).
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
