package cells_foss

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"unicode/utf16"
)

// ---------------------------------------------------------------------------
// Minimal OLE/CFB (Compound Binary File) writer and reader
//
// This implements just enough of the OLE/CFB format (MS-CFB) to produce and
// read files that Microsoft Excel accepts as Agile-Encryption wrappers.
//
// Reference: [MS-CFB] Compound File Binary File Format
// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb
//
// The writer produces a version-3 compound file (sector size 512,
// mini-sector size 64) containing:
//   - Header (512 bytes, including the 109-entry DIFAT array)
//   - FAT sector(s)
//   - DIFAT sector(s) (only when more than 109 FAT sectors are needed)
//   - Directory sector(s)
//   - Mini-FAT sector(s)
//   - Mini-stream data (for the small EncryptionInfo stream)
//   - Regular data sectors (for the large EncryptedPackage stream)
// ---------------------------------------------------------------------------

const (
	cfbVersion3              = 3
	cfbSectorSize3           = 512
	cfbMiniSectorSize        = 64
	cfbMiniStreamCutoff      = 4096
	cfbDirEntrySize          = 128
	cfbHeaderSize            = 512
	cfbFATEntriesPerSector   = cfbSectorSize3 / 4         // 128
	cfbHeaderDIFATEntries    = 109                        // DIFAT slots in the header
	cfbDIFATEntriesPerSector = cfbFATEntriesPerSector - 1 // 127 IDs + next-sector pointer

	// Special FAT values ([MS-CFB] 2.1).
	cfbFreeSect   = 0xFFFFFFFF // unallocated sector
	cfbEndOfChain = 0xFFFFFFFE // end of a sector chain
	cfbFatSect    = 0xFFFFFFFD // sector holds part of the FAT
	cfbDifSect    = 0xFFFFFFFC // sector holds part of the DIFAT chain
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
	DIFAT              [cfbHeaderDIFATEntries]uint32
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
//   - "EncryptionInfo" stream: 4-byte version + 4-byte reserved + XML
//   - "EncryptedPackage" stream: 8-byte size + encrypted bytes
func writeEncryptedCFB(encInfoXML, encryptedPkg []byte, origSize uint64) ([]byte, error) {
	infoStream := buildEncryptionInfoStream(encInfoXML)
	pkgStream := buildEncryptedPackageStream(encryptedPkg, origSize)

	return writeCFBContainer(infoStream, pkgStream)
}

// buildEncryptionInfoStream prepends the 8-byte Agile header: version
// 0x00040004 (major=4, minor=4) followed by the 4-byte reserved field,
// which the spec fixes at 0x00000040 ([MS-OFFCRYPTO] 2.3.4.5).
func buildEncryptionInfoStream(infoXML []byte) []byte {
	buf := make([]byte, 8+len(infoXML))
	binary.LittleEndian.PutUint32(buf[:4], agileVersion)
	binary.LittleEndian.PutUint32(buf[4:8], 0x40)
	copy(buf[8:], infoXML)
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
//	Sectors [0..f):      FAT sectors
//	Sectors [f..f+d):    DIFAT sectors (only when f > 109)
//	Then:                Directory sector(s)
//	Then:                Mini-FAT sector(s)
//	Then:                Mini-stream data (EncryptionInfo, < 4096 bytes)
//	Then:                EncryptedPackage data (regular sectors)
func writeCFBContainer(infoStream, pkgStream []byte) ([]byte, error) {
	// Determine which streams go in the mini-stream (< 4096 bytes).
	infoInMiniStream := len(infoStream) < cfbMiniStreamCutoff
	pkgInMiniStream := len(pkgStream) < cfbMiniStreamCutoff

	infoMiniSectors := miniSectorCount(len(infoStream), infoInMiniStream)
	pkgMiniSectors := miniSectorCount(len(pkgStream), pkgInMiniStream)

	totalMiniSectors := infoMiniSectors + pkgMiniSectors
	miniStreamSize := totalMiniSectors * cfbMiniSectorSize
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
		pkgSectors = (len(pkgStream) + cfbSectorSize3 - 1) / cfbSectorSize3
	}

	// Directory: 1 sector (4 entries, we use 3).
	dirSectors := 1

	totalDataSectors := dirSectors + miniFATSectors + miniStreamSectors + pkgSectors

	// FAT sizing is a fixed point: FAT sectors cover every sector in the
	// file (including FAT and DIFAT sectors), and once more than 109 FAT
	// sectors exist, DIFAT sectors are needed to index them, which in turn
	// adds sectors the FAT must cover.
	fatSectors, difatSectors := 1, 0
	for iter := 0; iter < 64; iter++ {
		d := 0
		if fatSectors > cfbHeaderDIFATEntries {
			d = (fatSectors - cfbHeaderDIFATEntries + cfbDIFATEntriesPerSector - 1) / cfbDIFATEntriesPerSector
		}
		total := fatSectors + d + totalDataSectors
		need := (total + cfbFATEntriesPerSector - 1) / cfbFATEntriesPerSector
		if need < 1 {
			need = 1
		}
		if need == fatSectors && d == difatSectors {
			break
		}
		fatSectors, difatSectors = need, d
	}

	// Sector allocation.
	sectorIdx := 0
	fatStart := sectorIdx
	sectorIdx += fatSectors
	difatStart := sectorIdx
	sectorIdx += difatSectors
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
		fat[i] = cfbFreeSect
	}
	for i := 0; i < fatSectors; i++ {
		fat[fatStart+i] = cfbFatSect
	}
	for i := 0; i < difatSectors; i++ {
		fat[difatStart+i] = cfbDifSect
	}
	writeChain(fat, dirStart, dirSectors)
	writeChain(fat, miniFATStart, miniFATSectors)
	writeChain(fat, miniStreamStart, miniStreamSectors)
	writeChain(fat, pkgStart, pkgSectors)

	// Build Mini-FAT: chains for streams within the mini-stream.
	miniFAT := make([]uint32, miniFATSectors*cfbFATEntriesPerSector)
	for i := range miniFAT {
		miniFAT[i] = cfbFreeSect
	}
	pkgMiniStart := infoMiniSectors // package starts after info in mini-stream
	writeChain(miniFAT, 0, infoMiniSectors)
	writeChain(miniFAT, pkgMiniStart, pkgMiniSectors)

	// Build directory entries.
	dirEntries := make([]cfbDirEntry, 4) // root + 2 streams + 1 unused

	// Entry 0: Root Entry. Its stream is the mini-stream container.
	writeDirName(&dirEntries[0], "Root Entry")
	dirEntries[0].NameSize = uint16(len("Root Entry")*2 + 2)
	dirEntries[0].Type = 5 // root storage
	dirEntries[0].Color = 1
	dirEntries[0].LeftSiblingID = cfbFreeSect
	dirEntries[0].RightSiblingID = cfbFreeSect
	dirEntries[0].ChildID = 1 // first child is EncryptionInfo
	if totalMiniSectors > 0 {
		dirEntries[0].StartSector = uint32(miniStreamStart)
		dirEntries[0].Size = uint64(miniStreamSize)
	} else {
		dirEntries[0].StartSector = cfbEndOfChain
	}

	// Entry 1: EncryptionInfo stream.
	writeDirName(&dirEntries[1], "EncryptionInfo")
	dirEntries[1].NameSize = uint16(len("EncryptionInfo")*2 + 2)
	dirEntries[1].Type = 2 // stream
	dirEntries[1].Color = 1
	dirEntries[1].LeftSiblingID = cfbFreeSect
	dirEntries[1].RightSiblingID = 2
	dirEntries[1].ChildID = cfbFreeSect
	dirEntries[1].StartSector = 0 // first mini-sector
	dirEntries[1].Size = uint64(len(infoStream))

	// Entry 2: EncryptedPackage stream.
	writeDirName(&dirEntries[2], "EncryptedPackage")
	dirEntries[2].NameSize = uint16(len("EncryptedPackage")*2 + 2)
	dirEntries[2].Type = 2 // stream
	dirEntries[2].Color = 1
	dirEntries[2].LeftSiblingID = cfbFreeSect
	dirEntries[2].RightSiblingID = cfbFreeSect
	dirEntries[2].ChildID = cfbFreeSect
	if pkgInMiniStream {
		dirEntries[2].StartSector = uint32(pkgMiniStart) // mini-sector
	} else {
		dirEntries[2].StartSector = uint32(pkgStart) // regular sector
	}
	dirEntries[2].Size = uint64(len(pkgStream))

	// Entry 3: unused.
	dirEntries[3].Type = 0
	dirEntries[3].LeftSiblingID = cfbFreeSect
	dirEntries[3].RightSiblingID = cfbFreeSect
	dirEntries[3].ChildID = cfbFreeSect

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
	if miniFATSectors > 0 {
		hdr.FirstMiniFATSector = uint32(miniFATStart)
	} else {
		hdr.FirstMiniFATSector = cfbEndOfChain
	}
	hdr.NumMiniFATSectors = uint32(miniFATSectors)
	if difatSectors > 0 {
		hdr.FirstDIFATSector = uint32(difatStart)
	} else {
		hdr.FirstDIFATSector = cfbEndOfChain
	}
	hdr.NumDIFATSectors = uint32(difatSectors)

	// Header DIFAT array: first min(109, fatSectors) FAT sector IDs.
	headerFATIDs := fatSectors
	if headerFATIDs > cfbHeaderDIFATEntries {
		headerFATIDs = cfbHeaderDIFATEntries
	}
	for i := 0; i < cfbHeaderDIFATEntries; i++ {
		if i < headerFATIDs {
			hdr.DIFAT[i] = uint32(fatStart + i)
		} else {
			hdr.DIFAT[i] = cfbFreeSect
		}
	}

	// Assemble the file.
	fileSize := cfbHeaderSize + totalSectors*cfbSectorSize3
	out := make([]byte, fileSize)

	// Write header.
	var hdrBuf bytes.Buffer
	binary.Write(&hdrBuf, binary.LittleEndian, &hdr)
	copy(out[:cfbHeaderSize], hdrBuf.Bytes())

	// Write FAT sector(s).
	for i := 0; i < fatSectors; i++ {
		off := cfbHeaderSize + (fatStart+i)*cfbSectorSize3
		for j := 0; j < cfbFATEntriesPerSector; j++ {
			binary.LittleEndian.PutUint32(out[off+j*4:], fat[i*cfbFATEntriesPerSector+j])
		}
	}

	// Write DIFAT sector(s): 127 FAT sector IDs each, last slot is the
	// next DIFAT sector (or ENDOFCHAIN).
	for i := 0; i < difatSectors; i++ {
		off := cfbHeaderSize + (difatStart+i)*cfbSectorSize3
		for j := 0; j < cfbDIFATEntriesPerSector; j++ {
			id := cfbHeaderDIFATEntries + i*cfbDIFATEntriesPerSector + j
			v := uint32(cfbFreeSect)
			if id < fatSectors {
				v = uint32(fatStart + id)
			}
			binary.LittleEndian.PutUint32(out[off+j*4:], v)
		}
		next := uint32(cfbEndOfChain)
		if i < difatSectors-1 {
			next = uint32(difatStart + i + 1)
		}
		binary.LittleEndian.PutUint32(out[off+cfbDIFATEntriesPerSector*4:], next)
	}

	// Write directory sector(s).
	dirOff := cfbHeaderSize + dirStart*cfbSectorSize3
	for i := range dirEntries {
		writeDirEntry(out[dirOff+i*cfbDirEntrySize:], &dirEntries[i])
	}

	// Write Mini-FAT sector(s).
	if miniFATEntries > 0 {
		mfOff := cfbHeaderSize + miniFATStart*cfbSectorSize3
		for i := 0; i < miniFATEntries; i++ {
			binary.LittleEndian.PutUint32(out[mfOff+i*4:], miniFAT[i])
		}
	}

	// Write mini-stream data.
	if totalMiniSectors > 0 {
		msOff := cfbHeaderSize + miniStreamStart*cfbSectorSize3
		if infoInMiniStream {
			copy(out[msOff:], infoStream)
		}
		if pkgInMiniStream {
			copy(out[msOff+infoMiniSectors*cfbMiniSectorSize:], pkgStream)
		}
	}

	// Write package data (EncryptedPackage) if not in mini-stream.
	if !pkgInMiniStream && pkgSectors > 0 {
		copy(out[cfbHeaderSize+pkgStart*cfbSectorSize3:], pkgStream)
	}

	return out, nil
}

// miniSectorCount returns how many 64-byte mini-sectors n bytes occupy.
func miniSectorCount(n int, inMini bool) int {
	if !inMini {
		return 0
	}
	c := (n + cfbMiniSectorSize - 1) / cfbMiniSectorSize
	if c == 0 {
		c = 1
	}
	return c
}

// writeChain marks count consecutive fat entries starting at start as a
// single sector chain terminated by ENDOFCHAIN. count==0 is a no-op.
func writeChain(fat []uint32, start, count int) {
	for i := 0; i < count; i++ {
		if i < count-1 {
			fat[start+i] = uint32(start + i + 1)
		} else {
			fat[start+i] = cfbEndOfChain
		}
	}
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
//
// All reads are bounds-checked; malformed containers produce errors instead
// of panics.
// ---------------------------------------------------------------------------

type cfbReader struct {
	data         []byte
	sectorSize   int
	miniSectorSz int
	miniCutoff   uint64
	totalSectors int
	fat          []uint32
	firstMiniFAT uint32
	numMiniFAT   uint32
}

// sectorOff returns the file offset of a sector, or -1 when out of range.
func (r *cfbReader) sectorOff(idx uint32) int {
	if int(idx) >= r.totalSectors {
		return -1
	}
	return cfbHeaderSize + int(idx)*r.sectorSize
}

// chain follows a FAT chain from start, guarding against cycles and
// out-of-range sectors. Special values (ENDOFCHAIN/FATSECT/DIFSECT/
// FREESECT, all >= 0xFFFFFFFA) terminate the chain.
func (r *cfbReader) chain(start uint32) ([]uint32, error) {
	if start >= 0xFFFFFFFA {
		return nil, nil
	}
	var out []uint32
	seen := make(map[uint32]bool)
	s := start
	for s < 0xFFFFFFFA {
		if seen[s] {
			return nil, fmt.Errorf("CFB: cycle in sector chain at %d", s)
		}
		if int(s) >= len(r.fat) {
			return nil, fmt.Errorf("CFB: sector %d outside FAT", s)
		}
		seen[s] = true
		out = append(out, s)
		s = r.fat[s]
	}
	return out, nil
}

// readChainData concatenates the sectors of a chain, clipped to size bytes.
func (r *cfbReader) readChainData(start uint32, size uint64) ([]byte, error) {
	if size > uint64(len(r.data)) {
		return nil, fmt.Errorf("CFB: stream size %d exceeds file size", size)
	}
	sectors, err := r.chain(start)
	if err != nil {
		return nil, err
	}
	out := make([]byte, size)
	pos := 0
	for _, sid := range sectors {
		if pos >= int(size) {
			break
		}
		off := r.sectorOff(sid)
		if off < 0 || off >= len(r.data) {
			return nil, fmt.Errorf("CFB: sector %d outside file", sid)
		}
		n := r.sectorSize
		if off+n > len(r.data) {
			n = len(r.data) - off
		}
		if pos+n > int(size) {
			n = int(size) - pos
		}
		copy(out[pos:pos+n], r.data[off:off+n])
		pos += n
	}
	if pos < int(size) {
		return nil, fmt.Errorf("CFB: chain for %d-byte stream ended after %d bytes", size, pos)
	}
	return out, nil
}

// readCFBStreams reads an OLE/CFB file and returns the named streams.
func readCFBStreams(data []byte) (map[string][]byte, error) {
	if len(data) < cfbHeaderSize {
		return nil, fmt.Errorf("file too short for CFB header")
	}
	for i, b := range cfbMagicBytes {
		if data[i] != b {
			return nil, fmt.Errorf("not a CFB file (bad magic)")
		}
	}

	r := &cfbReader{data: data}

	sectorShift := binary.LittleEndian.Uint16(data[30:])
	miniSectorShift := binary.LittleEndian.Uint16(data[32:])
	if sectorShift < 7 || sectorShift > 16 {
		return nil, fmt.Errorf("CFB: invalid sector shift %d", sectorShift)
	}
	if miniSectorShift < 2 || miniSectorShift > 12 {
		return nil, fmt.Errorf("CFB: invalid mini sector shift %d", miniSectorShift)
	}
	r.sectorSize = 1 << sectorShift
	r.miniSectorSz = 1 << miniSectorShift
	if r.sectorSize > len(data) {
		return nil, fmt.Errorf("CFB: sector size %d exceeds file size", r.sectorSize)
	}
	r.totalSectors = (len(data) - cfbHeaderSize) / r.sectorSize

	numFATSectors := binary.LittleEndian.Uint32(data[44:])
	firstDirSector := binary.LittleEndian.Uint32(data[48:])
	r.miniCutoff = uint64(binary.LittleEndian.Uint32(data[56:]))
	if r.miniCutoff == 0 {
		r.miniCutoff = cfbMiniStreamCutoff
	}
	r.firstMiniFAT = binary.LittleEndian.Uint32(data[60:])
	r.numMiniFAT = binary.LittleEndian.Uint32(data[64:])
	firstDIFATSector := binary.LittleEndian.Uint32(data[68:])
	numDIFATSectors := binary.LittleEndian.Uint32(data[72:])

	if numFATSectors == 0 || numFATSectors > uint32(r.totalSectors) {
		return nil, fmt.Errorf("CFB: invalid FAT sector count %d", numFATSectors)
	}

	// Collect FAT sector IDs: header DIFAT array first, then the DIFAT
	// chain for files with more than 109 FAT sectors.
	fatSectorIDs := make([]uint32, 0, numFATSectors)
	for i := 0; i < cfbHeaderDIFATEntries && uint32(len(fatSectorIDs)) < numFATSectors; i++ {
		v := binary.LittleEndian.Uint32(data[76+i*4:])
		if v >= 0xFFFFFFFA { // FREESECT marks unused slots
			break
		}
		if int(v) >= r.totalSectors {
			return nil, fmt.Errorf("CFB: DIFAT entry %d outside file", v)
		}
		fatSectorIDs = append(fatSectorIDs, v)
	}
	if uint32(len(fatSectorIDs)) < numFATSectors {
		// Follow the DIFAT chain. Each DIFAT sector holds
		// sectorSize/4-1 FAT IDs plus the next-DIFAT sector pointer.
		idsPerSector := r.sectorSize/4 - 1
		next := firstDIFATSector
		seen := make(map[uint32]bool)
		for i := uint32(0); i < numDIFATSectors && uint32(len(fatSectorIDs)) < numFATSectors; i++ {
			if next >= 0xFFFFFFFA || seen[next] {
				break
			}
			seen[next] = true
			off := r.sectorOff(next)
			if off < 0 || off+r.sectorSize > len(data) {
				return nil, fmt.Errorf("CFB: DIFAT sector %d outside file", next)
			}
			for j := 0; j < idsPerSector && uint32(len(fatSectorIDs)) < numFATSectors; j++ {
				v := binary.LittleEndian.Uint32(data[off+j*4:])
				if v >= 0xFFFFFFFA {
					continue
				}
				if int(v) >= r.totalSectors {
					return nil, fmt.Errorf("CFB: DIFAT entry %d outside file", v)
				}
				fatSectorIDs = append(fatSectorIDs, v)
			}
			next = binary.LittleEndian.Uint32(data[off+idsPerSector*4:])
		}
	}
	if uint32(len(fatSectorIDs)) < numFATSectors {
		return nil, fmt.Errorf("CFB: found %d FAT sectors, header declares %d", len(fatSectorIDs), numFATSectors)
	}

	// Build the FAT.
	entriesPerSector := r.sectorSize / 4
	r.fat = make([]uint32, len(fatSectorIDs)*entriesPerSector)
	for i, sid := range fatSectorIDs {
		off := r.sectorOff(sid)
		if off < 0 || off+r.sectorSize > len(data) {
			return nil, fmt.Errorf("CFB: FAT sector %d outside file", sid)
		}
		for j := 0; j < entriesPerSector; j++ {
			r.fat[i*entriesPerSector+j] = binary.LittleEndian.Uint32(data[off+j*4:])
		}
	}

	// Read directory entries (following the directory chain).
	type dirInfo struct {
		name        string
		streamType  byte
		startSector uint32
		size        uint64
	}
	dirData, err := readChainBounded(r, firstDirSector)
	if err != nil {
		return nil, fmt.Errorf("CFB: directory: %w", err)
	}
	var entries []dirInfo
	var root dirInfo
	rootFound := false
	for off := 0; off+cfbDirEntrySize <= len(dirData); off += cfbDirEntrySize {
		e := dirData[off : off+cfbDirEntrySize]
		entryType := e[66]
		if entryType == 0 {
			continue
		}
		// nameSize counts bytes including the UTF-16 null terminator.
		nameSize := int(binary.LittleEndian.Uint16(e[64:]))
		charBytes := nameSize - 2
		if charBytes < 0 || charBytes > 62 {
			charBytes = 62
		}
		units := make([]uint16, 0, charBytes/2)
		for j := 0; j+1 < charBytes; j += 2 {
			units = append(units, binary.LittleEndian.Uint16(e[j:]))
		}
		if n := len(units); n > 0 && units[n-1] == 0 {
			units = units[:n-1]
		}
		info := dirInfo{
			name:        string(utf16.Decode(units)),
			streamType:  entryType,
			startSector: binary.LittleEndian.Uint32(e[116:]),
			size:        binary.LittleEndian.Uint64(e[120:]),
		}
		entries = append(entries, info)
		if entryType == 5 && !rootFound {
			root, rootFound = info, true
		}
	}
	if !rootFound {
		return nil, fmt.Errorf("CFB: no root directory entry")
	}

	// Mini-stream container (held by the root entry in regular sectors).
	var miniStreamData []byte
	if root.size > 0 {
		miniStreamData, err = r.readChainData(root.startSector, root.size)
		if err != nil {
			return nil, fmt.Errorf("CFB: mini stream: %w", err)
		}
	}

	// Mini-FAT (chained through the regular FAT).
	var miniFAT []uint32
	if r.numMiniFAT > 0 && r.firstMiniFAT < 0xFFFFFFFA {
		mfSectors, err := r.chain(r.firstMiniFAT)
		if err != nil {
			return nil, fmt.Errorf("CFB: mini FAT: %w", err)
		}
		miniFAT = make([]uint32, len(mfSectors)*entriesPerSector)
		for i, sid := range mfSectors {
			off := r.sectorOff(sid)
			if off < 0 || off+r.sectorSize > len(data) {
				return nil, fmt.Errorf("CFB: mini FAT sector %d outside file", sid)
			}
			for j := 0; j < entriesPerSector; j++ {
				miniFAT[i*entriesPerSector+j] = binary.LittleEndian.Uint32(data[off+j*4:])
			}
		}
	}

	// readMiniStream assembles a stream stored in mini-sectors.
	readMiniStream := func(startSector uint32, size uint64) ([]byte, error) {
		if size > uint64(len(miniStreamData)) {
			return nil, fmt.Errorf("CFB: mini stream size %d exceeds container %d", size, len(miniStreamData))
		}
		out := make([]byte, size)
		pos := 0
		s := startSector
		seen := make(map[uint32]bool)
		for pos < int(size) && s < 0xFFFFFFFA {
			if seen[s] {
				return nil, fmt.Errorf("CFB: cycle in mini chain at %d", s)
			}
			seen[s] = true
			if int(s) >= len(miniFAT) {
				return nil, fmt.Errorf("CFB: mini sector %d outside mini FAT", s)
			}
			srcOff := int(s) * r.miniSectorSz
			n := r.miniSectorSz
			if pos+n > int(size) {
				n = int(size) - pos
			}
			if srcOff+n > len(miniStreamData) {
				return nil, fmt.Errorf("CFB: mini sector %d outside mini stream", s)
			}
			copy(out[pos:pos+n], miniStreamData[srcOff:srcOff+n])
			pos += n
			s = miniFAT[s]
		}
		if pos < int(size) {
			return nil, fmt.Errorf("CFB: mini chain ended after %d of %d bytes", pos, size)
		}
		return out, nil
	}

	// Read named streams.
	result := make(map[string][]byte)
	for _, e := range entries {
		if e.streamType != 2 {
			continue
		}
		var payload []byte
		if e.size < r.miniCutoff {
			payload, err = readMiniStream(e.startSector, e.size)
		} else {
			payload, err = r.readChainData(e.startSector, e.size)
		}
		if err != nil {
			return nil, fmt.Errorf("CFB: stream %q: %w", e.name, err)
		}
		result[e.name] = payload
	}

	return result, nil
}

// readChainBounded reads a FAT chain without knowing its byte size up front
// (used for the directory, whose chain may span several sectors).
func readChainBounded(r *cfbReader, start uint32) ([]byte, error) {
	sectors, err := r.chain(start)
	if err != nil {
		return nil, err
	}
	if len(sectors) == 0 {
		return nil, fmt.Errorf("empty chain")
	}
	out := make([]byte, 0, len(sectors)*r.sectorSize)
	for _, sid := range sectors {
		off := r.sectorOff(sid)
		if off < 0 || off >= len(r.data) {
			return nil, fmt.Errorf("CFB: sector %d outside file", sid)
		}
		n := r.sectorSize
		if off+n > len(r.data) {
			n = len(r.data) - off
		}
		out = append(out, r.data[off:off+n]...)
	}
	return out, nil
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

// readCFBEncrypted extracts the EncryptionInfo XML, EncryptedPackage bytes,
// and declared original size from an OLE/CFB container.
//
// Spec-compliant EncryptionInfo streams start with an 8-byte header
// (version 0x00040004 + 4 reserved zero bytes) followed by the XML. Files
// written by earlier versions of this library used only the 4-byte version;
// both are detected. The EncryptedPackage stream always starts with an
// 8-byte original-size header.
func readCFBEncrypted(data []byte) (infoXML, encPkg []byte, origSize uint64, err error) {
	streams, err := readCFBStreams(data)
	if err != nil {
		return nil, nil, 0, fmt.Errorf("reading CFB: %w", err)
	}

	infoStream, ok := streams["EncryptionInfo"]
	if !ok {
		return nil, nil, 0, fmt.Errorf("CFB: missing EncryptionInfo stream")
	}
	pkgStream, ok := streams["EncryptedPackage"]
	if !ok {
		return nil, nil, 0, fmt.Errorf("CFB: missing EncryptedPackage stream")
	}

	// EncryptionInfo: spec files carry an 8-byte header (version + reserved
	// 0x40) before the XML; files from earlier versions of this library had
	// only the 4-byte version, so the XML starts with '<' at offset 4.
	skip := 8
	if infoStream[4] == '<' {
		skip = 4
	}
	if len(infoStream) <= skip {
		return nil, nil, 0, fmt.Errorf("CFB: EncryptionInfo too short")
	}
	infoXML = infoStream[skip:]

	// EncryptedPackage: 8-byte size header.
	if len(pkgStream) < 8 {
		return nil, nil, 0, fmt.Errorf("CFB: EncryptedPackage too short")
	}
	origSize = binary.LittleEndian.Uint64(pkgStream[:8])
	encPkg = pkgStream[8:]

	return infoXML, encPkg, origSize, nil
}
