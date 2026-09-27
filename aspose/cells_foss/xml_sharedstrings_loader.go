package cells_foss

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"strings"
)

// ---------------------------------------------------------------------------
// Internal XML types for decoding sharedStrings.xml
// ---------------------------------------------------------------------------

// xmlSST represents the top-level <sst> element in xl/sharedStrings.xml.
type xmlSST struct {
	XMLName     xml.Name `xml:"sst"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`
	Items       []xmlSI  `xml:"si"`
}

// xmlSI mirrors a <si> (string item) element. It supports the simple <t>
// text form as well as the <r> rich-text runs form.
type xmlSI struct {
	Text string   `xml:"t"`
	Runs []xmlRun `xml:"r"`
}

// xmlRun represents a single rich-text run <r> containing a <t> element.
type xmlRun struct {
	Text string `xml:"t"`
}

// ---------------------------------------------------------------------------
// loadSharedStrings
// ---------------------------------------------------------------------------

// loadSharedStrings reads xl/sharedStrings.xml from the ZIP archive and
// returns a map from 0-based index to the resolved string value.
// It returns an empty map (and no error) when the file is absent, which is
// valid for workbooks that use inline values only.
func loadSharedStrings(zr *zip.Reader) (map[int]string, error) {
	raw, err := readZipFile(zr, "xl/sharedStrings.xml")
	if err != nil {
		return map[int]string{}, nil
	}

	var sst xmlSST
	if err := xml.Unmarshal(raw, &sst); err != nil {
		return nil, fmt.Errorf("shared strings: %w", err)
	}

	out := make(map[int]string, len(sst.Items))
	for i, si := range sst.Items {
		out[i] = resolveSIText(si)
	}
	return out, nil
}

// resolveSIText extracts the plain-text content from a shared string item,
// preferring the simple <t> form and falling back to concatenating <r><t>
// runs for rich-text entries.
func resolveSIText(si xmlSI) string {
	if si.Text != "" {
		return si.Text
	}
	if len(si.Runs) == 0 {
		return ""
	}
	var buf strings.Builder
	for _, run := range si.Runs {
		buf.WriteString(run.Text)
	}
	return buf.String()
}
