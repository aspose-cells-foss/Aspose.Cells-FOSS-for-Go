package cells_foss

import (
	"bytes"
	"encoding/binary"
	"fmt"
)

// ---------------------------------------------------------------------------
// Minimal OLE/CFB (Compound Binary File) writer
//
// This implements just enough of the OLE/CFB format (MS-CFB) to produce
// files that Microsoft Excel can read as Agile-Encryption wrappers.
//
// Reference: [MS-CFB] Compound File Binary File Format
// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb
//
// The writer produces a version-3 compound file (sector size 512,
// mini-sector size 64) containing:
//   - Header (512 bytes)
//   - FAT sector(s)
//   - Directory sector(s)
//   - Mini-stream FAT sector(s)
//   - Mini-stream data (for the small EncryptionInfo stream)
//   - Regular data sectors (for the large EncryptedPackage stream)
// ---------------------------------------------------------------------------

const (
	cfbMagic               = 0xB1A1CF11E0A1F0CE // not used directly; see cfbMagicBytes
	cfbVersion3            = 3
	cfbSectorSize3         = 512
	cfbMiniSectorSize      = 64
	cfbMiniStreamCutoff    = 4096
	cfbDirEntrySize        = 128
	cfbFATEntriesPerSector = cfbSectorSize3 / 4 // 128

	// Special FAT values.
	cfbFreeSector  = 0xFFFFFFFF // 0xFFFFFFFE is actually free; see below
	cfbEndOfChain  = 0xFFFFFFFE
	cfbFATSector   = 0xFFFFFFFD
	cfbDIFATSector = 0xFFFFFFFC

	// Actually the MS-CFB spec uses:
	// FREESEC     = 0xFFFFFFFE
	// ENDOFCHAIN  = 0xFFFFFFFE
	// Let me use the correct values from the spec.
	cfbFree     = 0xFFFFFFFE
	cfbEnd      = 0xFFFFFFFE // same as free for end-of-chain
	cfbFATVal   = 0xFFFFFFFD
	cfbDIFATVal = 0xFFFFFFFC
)

// OLE/CFB magic bytes: D0 CF 11 E0 A1 B1 1A E1
var cfbMagicBytes = [8]byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

// cfbHeader is the 512-byte compound file header.
type cfbHeader struct {
	Magic              [8]byte
	ClsID              [16]byte // Usually zeros.
	MinorVersion       uint16
	MajorVersion       uint16
	ByteOrder          uint16
	SectorShift        uint16
	MiniSectorShift    uint16
	Reserved           [6]byte
	NumDirSectors      uint32
	NumFATSectors      uint32
	FirstDirSector     uint32
	TransactionSig     uint32
	MiniStreamCutoff   uint32
	FirstMiniFATSector uint32
	NumMiniFATSectors  uint32
	FirstDIFATSector   uint32
	NumDIFATSectors    uint32
	DIFAT              [109]uint32
}

// cfbDirEntry is a 128-byte directory entry.
type cfbDirEntry struct {
	Name           [64]byte // UTF-16LE name, null-terminated
	NameSize       uint16   // length in bytes including null terminator
	Type           byte     // 0=unknown, 1=storage, 2=stream, 5=root
	Color          byte     // 0=red, 1=black
	LeftSiblingID  uint32
	RightSiblingID uint32
	ChildID        uint32
	ClsID          [16]byte
	State          uint32
	CreateTime     uint64
	ModifyTime     uint64
	StartSector    uint32
	Size           uint64
}

// writeEncryptedCFB wraps Agile Encryption data in an OLE/CFB container
// that Microsoft Excel can open.
//
// The container has:
//   - "EncryptionInfo" stream: 4-byte version header + XML
//   - "EncryptedPackage" stream: 8-byte size + encrypted bytes
func writeEncryptedCFB(encInfoXML, encryptedPkg []byte, origSize uint64) ([]byte, error) {
	// Build stream contents.
	infoStream := buildEncryptionInfoStream(encInfoXML)
	pkgStream := buildEncryptedPackageStream(encryptedPkg, origSize)

	return writeCFBContainer(infoStream, pkgStream)
}

// buildEncryptionInfoStream prepends the 4-byte Agile version header.
func buildEncryptionInfoStream(infoXML []byte) []byte {
	buf := make([]byte, 4+len(infoXML))
	// Agile Encryption version: 0x00040004 (major=4, minor=4).
	binary.LittleEndian.PutUint32(buf[:4], 0x00040004)
	copy(buf[4:], infoXML)
	return buf
}

// buildEncryptedPackageStream prepends the 8-byte original size.
func buildEncryptedPackageStream(encPkg []byte, origSize uint64) []byte {
	buf := make([]byte, 8+len(encPkg))
	binary.LittleEndian.PutUint64(buf[:8], origSize)
	copy(buf[8:], encPkg)
	return buf
}

// writeCFBContainer writes a minimal OLE/CFB file containing two streams.
//
// Layout:
//
//	Sector 0:    FAT sector(s)
//	Sector 1+:   Directory sector(s)
//	Then:        Mini-FAT sector(s)
//	Then:        Mini-stream data (EncryptionInfo, < 4096 bytes)
//	Then:        EncryptedPackage data (regular sectors)
func writeCFBContainer(infoStream, pkgStream []byte) ([]byte, error) {
	// Calculate sector counts.
	infoBytes := infoStream
	pkgBytes := pkgStream

	// Determine which streams go in the mini-stream (< 4096 bytes).
	infoInMiniStream := len(infoBytes) < cfbMiniStreamCutoff
	pkgInMiniStream := len(pkgBytes) < cfbMiniStreamCutoff

	// Mini-stream: holds small streams.
	var infoMiniSectors, pkgMiniSectors int
	var miniStreamSize int

	if infoInMiniStream {
		infoMiniSectors = (len(infoBytes) + cfbMiniSectorSize - 1) / cfbMiniSectorSize
		if infoMiniSectors == 0 {
			infoMiniSectors = 1
		}
	}
	if pkgInMiniStream {
		pkgMiniSectors = (len(pkgBytes) + cfbMiniSectorSize - 1) / cfbMiniSectorSize
		if pkgMiniSectors == 0 {
			pkgMiniSectors = 1
		}
	}

	totalMiniSectors := infoMiniSectors + pkgMiniSectors
	if totalMiniSectors > 0 {
		miniStreamSize = totalMiniSectors * cfbMiniSectorSize
	}

	// Mini-stream data sectors in the compound file.
	miniStreamSectors := (miniStreamSize + cfbSectorSize3 - 1) / cfbSectorSize3

	// Mini-FAT: one entry per mini-sector used.
	miniFATEntries := totalMiniSectors
	miniFATSectors := 0
	if miniFATEntries > 0 {
		miniFATSectors = (miniFATEntries*4 + cfbSectorSize3 - 1) / cfbSectorSize3
	}

	// Regular sectors for EncryptedPackage (if not in mini-stream).
	pkgSectors := 0
	if !pkgInMiniStream {
		pkgSectors = (len(pkgBytes) + cfbSectorSize3 - 1) / cfbSectorSize3
		if pkgSectors == 0 {
			pkgSectors = 1
		}
	}

	// Directory: 1 sector (4 entries, we use 3).
	dirSectors := 1

	// FAT: we need enough FAT sectors to cover all sectors.
	// Total sectors (excluding header): FAT + Dir + MiniFAT + MiniStream + Package
	// Start with 1 FAT sector and grow if needed.
	fatSectors := 1
	totalDataSectors := dirSectors + miniFATSectors + miniStreamSectors + pkgSectors
	for {
		fatCapacity := fatSectors * cfbFATEntriesPerSector
		// Each FAT sector itself uses a FAT entry, so subtract fatSectors.
		usable := fatCapacity - fatSectors
		if usable >= totalDataSectors+fatSectors {
			break
		}
		fatSectors++
	}

	// Sector allocation:
	//   Sectors [0 .. fatSectors-1]: FAT
	//   Sectors [fatSectors .. fatSectors+dirSectors-1]: Directory
	//   Then: Mini-FAT sectors
	//   Then: Mini-stream data
	//   Then: Package data
	sectorIdx := 0
	fatStart := sectorIdx
	sectorIdx += fatSectors
	dirStart := sectorIdx
	sectorIdx += dirSectors
	miniFATStart := sectorIdx
	sectorIdx += miniFATSectors
	miniStreamStart := sectorIdx
	sectorIdx += miniStreamSectors
	pkgStart := sectorIdx
	sectorIdx += pkgSectors
	totalSectors := sectorIdx

	// Build FAT.
	fat := make([]uint32, fatSectors*cfbFATEntriesPerSector)
	for i := range fat {
		fat[i] = cfbFree
	}
	// FAT sectors themselves.
	for i := 0; i < fatSectors; i++ {
		fat[fatStart+i] = cfbFATVal
	}
	// Directory chain.
	for i := 0; i < dirSectors; i++ {
		if i < dirSectors-1 {
			fat[dirStart+i] = uint32(dirStart + i + 1)
		} else {
			fat[dirStart+i] = cfbEndOfChain
		}
	}
	// Mini-FAT chain.
	for i := 0; i < miniFATSectors; i++ {
		if i < miniFATSectors-1 {
			fat[miniFATStart+i] = uint32(miniFATStart + i + 1)
		} else {
			fat[miniFATStart+i] = cfbEndOfChain
		}
	}
	// Mini-stream data chain.
	for i := 0; i < miniStreamSectors; i++ {
		if i < miniStreamSectors-1 {
			fat[miniStreamStart+i] = uint32(miniStreamStart + i + 1)
		} else {
			fat[miniStreamStart+i] = cfbEndOfChain
		}
	}
	// Package data chain.
	for i := 0; i < pkgSectors; i++ {
		if i < pkgSectors-1 {
			fat[pkgStart+i] = uint32(pkgStart + i + 1)
		} else {
			fat[pkgStart+i] = cfbEndOfChain
		}
	}

	// Build Mini-FAT: chains for streams within the mini-stream.
	miniFAT := make([]uint32, miniFATSectors*cfbFATEntriesPerSector)
	for i := range miniFAT {
		miniFAT[i] = cfbFree
	}
	// EncryptionInfo chain.
	pkgMiniStart := infoMiniSectors // package starts after info in mini-stream
	if infoInMiniStream {
		for i := 0; i < infoMiniSectors; i++ {
			if i < infoMiniSectors-1 {
				miniFAT[i] = uint32(i + 1)
			} else {
				miniFAT[i] = cfbEndOfChain
			}
		}
	}
	// EncryptedPackage chain (if in mini-stream).
	if pkgInMiniStream {
		for i := 0; i < pkgMiniSectors; i++ {
			idx := pkgMiniStart + i
			if i < pkgMiniSectors-1 {
				miniFAT[idx] = uint32(idx + 1)
			} else {
				miniFAT[idx] = cfbEndOfChain
			}
		}
	}

	// Build directory entries.
	dirEntries := make([]cfbDirEntry, 4) // root + 2 streams + 1 unused

	// Entry 0: Root Entry.
	writeDirName(&dirEntries[0], "Root Entry")
	dirEntries[0].NameSize = uint16(len("Root Entry")*2 + 2)
	dirEntries[0].Type = 5  // root storage
	dirEntries[0].Color = 1 // black
	dirEntries[0].LeftSiblingID = 0xFFFFFFFF
	dirEntries[0].RightSiblingID = 0xFFFFFFFF
	dirEntries[0].ChildID = 1 // first child is EncryptionInfo
	dirEntries[0].StartSector = uint32(miniStreamStart)
	dirEntries[0].Size = uint64(miniStreamSize)

	// Entry 1: EncryptionInfo stream.
	writeDirName(&dirEntries[1], "EncryptionInfo")
	dirEntries[1].NameSize = uint16(len("EncryptionInfo")*2 + 2)
	dirEntries[1].Type = 2  // stream
	dirEntries[1].Color = 1 // black
	dirEntries[1].LeftSiblingID = 0xFFFFFFFF
	dirEntries[1].RightSiblingID = 2
	dirEntries[1].ChildID = 0xFFFFFFFF
	dirEntries[1].StartSector = 0 // first mini-sector
	dirEntries[1].Size = uint64(len(infoBytes))

	// Entry 2: EncryptedPackage stream.
	writeDirName(&dirEntries[2], "EncryptedPackage")
	dirEntries[2].NameSize = uint16(len("EncryptedPackage")*2 + 2)
	dirEntries[2].Type = 2  // stream
	dirEntries[2].Color = 1 // black
	dirEntries[2].LeftSiblingID = 0xFFFFFFFF
	dirEntries[2].RightSiblingID = 0xFFFFFFFF
	dirEntries[2].ChildID = 0xFFFFFFFF
	if pkgInMiniStream {
		dirEntries[2].StartSector = uint32(pkgMiniStart) // mini-sector
	} else {
		dirEntries[2].StartSector = uint32(pkgStart) // regular sector
	}
	dirEntries[2].Size = uint64(len(pkgBytes))

	// Entry 3: unused.
	dirEntries[3].Type = 0
	dirEntries[3].LeftSiblingID = 0xFFFFFFFF
	dirEntries[3].RightSiblingID = 0xFFFFFFFF
	dirEntries[3].ChildID = 0xFFFFFFFF

	// Build header.
	var hdr cfbHeader
	copy(hdr.Magic[:], cfbMagicBytes[:])
	hdr.MinorVersion = 0x003E
	hdr.MajorVersion = cfbVersion3
	hdr.ByteOrder = 0xFFFE
	hdr.SectorShift = 9     // 2^9 = 512
	hdr.MiniSectorShift = 6 // 2^6 = 64
	hdr.NumDirSectors = 0   // must be 0 for v3
	hdr.NumFATSectors = uint32(fatSectors)
	hdr.FirstDirSector = uint32(dirStart)
	hdr.MiniStreamCutoff = cfbMiniStreamCutoff
	hdr.FirstMiniFATSector = uint32(miniFATStart)
	hdr.NumMiniFATSectors = uint32(miniFATSectors)
	hdr.FirstDIFATSector = 0xFFFFFFFE // none
	hdr.NumDIFATSectors = 0

	// Fill DIFAT array in header (up to 109 entries).
	for i := 0; i < 109; i++ {
		if i < fatSectors {
			hdr.DIFAT[i] = uint32(fatStart + i)
		} else {
			hdr.DIFAT[i] = 0xFFFFFFFF // free
		}
	}

	// Assemble the file.
	fileSize := 512 + totalSectors*cfbSectorSize3
	out := make([]byte, fileSize)

	// Write header.
	var hdrBuf bytes.Buffer
	binary.Write(&hdrBuf, binary.LittleEndian, &hdr)
	copy(out[:512], hdrBuf.Bytes())

	// Write FAT sector(s).
	for i := 0; i < fatSectors; i++ {
		off := 512 + (fatStart+i)*cfbSectorSize3
		for j := 0; j < cfbFATEntriesPerSector; j++ {
			binary.LittleEndian.PutUint32(out[off+j*4:], fat[i*cfbFATEntriesPerSector+j])
		}
	}

	// Write directory sector(s).
	dirOff := 512 + dirStart*cfbSectorSize3
	for i, entry := range dirEntries {
		off := dirOff + i*cfbDirEntrySize
		writeDirEntry(out[off:], &entry)
	}

	// Write Mini-FAT sector(s).
	mfOff := 512 + miniFATStart*cfbSectorSize3
	for i := 0; i < miniFATEntries; i++ {
		binary.LittleEndian.PutUint32(out[mfOff+i*4:], miniFAT[i])
	}

	// Write mini-stream data.
	msOff := 512 + miniStreamStart*cfbSectorSize3
	if infoInMiniStream {
		copy(out[msOff:], infoBytes)
	}
	if pkgInMiniStream {
		pkgMsOff := msOff + infoMiniSectors*cfbMiniSectorSize
		copy(out[pkgMsOff:], pkgBytes)
	}

	// Write package data (EncryptedPackage) if not in mini-stream.
	if !pkgInMiniStream {
		pkgOff := 512 + pkgStart*cfbSectorSize3
		copy(out[pkgOff:], pkgBytes)
	}

	return out, nil
}

// writeDirName encodes an ASCII name as UTF-16LE into a directory entry.
func writeDirName(entry *cfbDirEntry, name string) {
	for i, ch := range name {
		if i*2+1 >= 64 {
			break
		}
		entry.Name[i*2] = byte(ch)
		entry.Name[i*2+1] = 0
	}
}

// writeDirEntry writes a 128-byte directory entry to buf.
func writeDirEntry(buf []byte, e *cfbDirEntry) {
	copy(buf[0:64], e.Name[:])
	binary.LittleEndian.PutUint16(buf[64:], e.NameSize)
	buf[66] = e.Type
	buf[67] = e.Color
	binary.LittleEndian.PutUint32(buf[68:], e.LeftSiblingID)
	binary.LittleEndian.PutUint32(buf[72:], e.RightSiblingID)
	binary.LittleEndian.PutUint32(buf[76:], e.ChildID)
	copy(buf[80:96], e.ClsID[:])
	binary.LittleEndian.PutUint32(buf[96:], e.State)
	binary.LittleEndian.PutUint64(buf[100:], e.CreateTime)
	binary.LittleEndian.PutUint64(buf[108:], e.ModifyTime)
	binary.LittleEndian.PutUint32(buf[116:], e.StartSector)
	binary.LittleEndian.PutUint64(buf[120:], e.Size)
}

// ---------------------------------------------------------------------------
// OLE/CFB reader (minimal — just enough to load Agile Encryption files)
// ---------------------------------------------------------------------------

// readCFBStreams reads an OLE/CFB file and returns the named streams.
func readCFBStreams(data []byte) (map[string][]byte, error) {
	if len(data) < 512 {
		return nil, fmt.Errorf("file too short for CFB header")
	}

	// Verify magic.
	for i, b := range cfbMagicBytes {
		if data[i] != b {
			return nil, fmt.Errorf("not a CFB file (bad magic)")
		}
	}

	// With CLSID field, offsets are:
	// 30: Sector shift
	sectorShift := binary.LittleEndian.Uint16(data[30:])
	sectorSize := 1 << sectorShift

	// CFB header offsets (from MS-CFB spec):
	// 44: Number of FAT sectors
	// 48: First directory sector
	// 60: First mini FAT sector
	// 64: Number of mini FAT sectors
	numFATSectors := binary.LittleEndian.Uint32(data[44:])
	firstDirSector := binary.LittleEndian.Uint32(data[48:])
	firstMiniFATSector := binary.LittleEndian.Uint32(data[60:])
	numMiniFATSectors := binary.LittleEndian.Uint32(data[64:])

	// Build sector offset helper.
	sectorOff := func(idx uint32) int {
		return 512 + int(idx)*sectorSize
	}

	// Read DIFAT from header.
	var fatSectorIDs []uint32
	for i := 0; i < 109; i++ {
		v := binary.LittleEndian.Uint32(data[76+i*4:])
		if v >= 0xFFFFFFFE {
			break
		}
		fatSectorIDs = append(fatSectorIDs, v)
	}
	if uint32(len(fatSectorIDs)) < numFATSectors {
		// TODO: follow DIFAT chain for large files. For our use case
		// (small encryption wrappers) 109 FAT entries is more than enough.
		fatSectorIDs = fatSectorIDs[:numFATSectors]
	}

	// Read FAT.
	totalFATEntries := len(fatSectorIDs) * (sectorSize / 4)
	fat := make([]uint32, totalFATEntries)
	for i, sid := range fatSectorIDs {
		off := sectorOff(sid)
		for j := 0; j < sectorSize/4; j++ {
			fat[i*(sectorSize/4)+j] = binary.LittleEndian.Uint32(data[off+j*4:])
		}
	}

	// Read directory entries.
	type dirInfo struct {
		name        string
		streamType  byte
		startSector uint32
		size        uint64
	}
	var entries []dirInfo
	dirOff := sectorOff(firstDirSector)
	for i := 0; i < sectorSize/cfbDirEntrySize; i++ {
		off := dirOff + i*cfbDirEntrySize
		entryType := data[off+66]
		if entryType == 0 {
			continue
		}
		// Read UTF-16LE name.
		nameSize := int(binary.LittleEndian.Uint16(data[off+64:]))
		var name []byte
		for j := 0; j < nameSize-2; j += 2 {
			name = append(name, data[off+j])
		}
		startSec := binary.LittleEndian.Uint32(data[off+116:])
		size := binary.LittleEndian.Uint64(data[off+120:])
		entries = append(entries, dirInfo{
			name:        string(name),
			streamType:  entryType,
			startSector: startSec,
			size:        size,
		})
	}

	// Read streams.
	readStream := func(startSector uint32, size uint64, isMiniStream bool, miniStreamData []byte) []byte {
		if isMiniStream {
			out := make([]byte, size)
			// Read mini-FAT.
			miniFAT := make([]uint32, int(numMiniFATSectors)*(sectorSize/4))
			for i := uint32(0); i < numMiniFATSectors; i++ {
				off := sectorOff(firstMiniFATSector + i)
				for j := 0; j < sectorSize/4; j++ {
					miniFAT[int(i)*(sectorSize/4)+j] = binary.LittleEndian.Uint32(data[off+j*4:])
				}
			}
			// Follow mini-sector chain.
			pos := 0
			ms := int(startSector)
			for pos < int(size) && ms >= 0 && uint32(ms) != 0xFFFFFFFE {
				srcOff := ms * cfbMiniSectorSize
				n := cfbMiniSectorSize
				if pos+n > int(size) {
					n = int(size) - pos
				}
				if srcOff+n <= len(miniStreamData) {
					copy(out[pos:], miniStreamData[srcOff:srcOff+n])
				}
				pos += n
				if uint32(ms) < uint32(len(miniFAT)) {
					ms = int(miniFAT[ms])
				} else {
					break
				}
			}
			return out
		}
		// Regular stream: follow sector chain.
		out := make([]byte, size)
		pos := 0
		s := startSector
		for pos < int(size) && s < 0xFFFFFFFE {
			off := sectorOff(s)
			n := sectorSize
			if pos+n > int(size) {
				n = int(size) - pos
			}
			copy(out[pos:], data[off:off+n])
			pos += n
			s = fat[s]
		}
		return out
	}

	// Find the mini-stream from the root entry.
	var miniStreamData []byte
	for _, e := range entries {
		if e.streamType == 5 { // root storage
			miniStreamData = readStream(e.startSector, e.size, false, nil)
			break
		}
	}

	// Read named streams.
	result := make(map[string][]byte)
	for _, e := range entries {
		if e.streamType != 2 {
			continue
		}
		isMini := e.size < cfbMiniStreamCutoff
		result[e.name] = readStream(e.startSector, e.size, isMini, miniStreamData)
	}

	return result, nil
}

// ---------------------------------------------------------------------------
// High-level helpers used by xmlloader.go
// ---------------------------------------------------------------------------

// isCFBFile returns true when data starts with the OLE/CFB magic bytes.
func isCFBFile(data []byte) bool {
	if len(data) < 8 {
		return false
	}
	for i, b := range cfbMagicBytes {
		if data[i] != b {
			return false
		}
	}
	return true
}

// readCFBEncrypted extracts the EncryptionInfo XML and EncryptedPackage
// bytes from an OLE/CFB container produced by writeEncryptedCFB.
//
// The EncryptionInfo stream starts with a 4-byte version header (0x00040004
// for Agile) followed by the XML. The EncryptedPackage stream starts with
// an 8-byte original-size header followed by the encrypted data.
func readCFBEncrypted(data []byte) (infoXML, encPkg []byte, err error) {
	streams, err := readCFBStreams(data)
	if err != nil {
		return nil, nil, fmt.Errorf("reading CFB: %w", err)
	}

	infoStream, ok := streams["EncryptionInfo"]
	if !ok {
		return nil, nil, fmt.Errorf("CFB: missing EncryptionInfo stream")
	}
	pkgStream, ok := streams["EncryptedPackage"]
	if !ok {
		return nil, nil, fmt.Errorf("CFB: missing EncryptedPackage stream")
	}

	// EncryptionInfo: skip 4-byte version header.
	if len(infoStream) < 4 {
		return nil, nil, fmt.Errorf("CFB: EncryptionInfo too short")
	}
	infoXML = infoStream[4:]

	// EncryptedPackage: skip 8-byte size header.
	if len(pkgStream) < 8 {
		return nil, nil, fmt.Errorf("CFB: EncryptedPackage too short")
	}
	encPkg = pkgStream[8:]

	return infoXML, encPkg, nil
}
