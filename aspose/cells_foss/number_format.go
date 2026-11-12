package cells_foss

import (
	"fmt"
	"math"
	"strings"
	"time"
)

// ======================================================================
// Number format codes
// ======================================================================

// Style.NumberFormat accepts any Excel format code. The constants below
// cover the most common ones. Codes that Excel knows as built-ins (e.g.
// "0.00" or "0%") are saved under their standard numFmtId; every other
// code is written as a custom <numFmt> entry in styles.xml with an id
// starting at 164, as required by ECMA-376.
const (
	NumFmtGeneral  = "General"
	NumFmtInt      = "0"
	NumFmtFloat2   = "0.00"
	NumFmtThousand = "#,##0"
	NumFmtMoney    = "#,##0.00"
	NumFmtPercent  = "0%"
	NumFmtPercent2 = "0.00%"
	NumFmtDate     = "yyyy-mm-dd"
	NumFmtDateTime = "yyyy-mm-dd hh:mm:ss"
	NumFmtTime     = "hh:mm:ss"
	NumFmtText     = "@"
)

// builtinNumFmtCodes maps the standard Excel numFmtIds to their format
// codes (ECMA-376 part 1, 18.8.30). Only the commonly used ids are
// listed; ids outside this table are resolved through the file's own
// <numFmts> entries when loading.
var builtinNumFmtCodes = map[int]string{
	0:  "General",
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	9:  "0%",
	10: "0.00%",
	14: "mm-dd-yy",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
	49: "@",
}

var builtinNumFmtIDs = map[string]int{}

func init() {
	for id, code := range builtinNumFmtCodes {
		builtinNumFmtIDs[code] = id
	}
}

// numFmtAllocator assigns ids to custom format codes while a styles part
// is generated. Ids start at 164; built-in codes reuse the standard ids
// and are not allocated.
type numFmtAllocator struct {
	ids   map[string]int
	codes []string // in first-use order, i.e. ascending id
}

func newNumFmtAllocator() *numFmtAllocator {
	return &numFmtAllocator{ids: make(map[string]int)}
}

// idFor returns the numFmtId for a format code, registering custom codes
// on first use. Empty and "General" both mean the default format (id 0).
func (a *numFmtAllocator) idFor(code string) int {
	code = strings.TrimSpace(code)
	if code == "" || strings.EqualFold(code, NumFmtGeneral) {
		return 0
	}
	if id, ok := builtinNumFmtIDs[code]; ok {
		return id
	}
	if id, ok := a.ids[code]; ok {
		return id
	}
	id := 164 + len(a.codes)
	a.ids[code] = id
	a.codes = append(a.codes, code)
	return id
}

// entries returns the <numFmt> records for all allocated custom codes.
func (a *numFmtAllocator) entries() []outNumFmt {
	out := make([]outNumFmt, 0, len(a.codes))
	for i, code := range a.codes {
		out = append(out, outNumFmt{NumFmtId: 164 + i, FormatCode: code})
	}
	return out
}

// codeForID resolves a numFmtId back to its canonical format code: the
// built-in code for standard ids, the registered code for custom ids.
// Returns "" for id 0 (General).
func (a *numFmtAllocator) codeForID(id int) string {
	if id == 0 {
		return ""
	}
	if code, ok := builtinNumFmtCodes[id]; ok {
		return code
	}
	for code, cid := range a.ids {
		if cid == id {
			return code
		}
	}
	return ""
}

// ======================================================================
// Excel date serial numbers (1900 date system)
// ======================================================================

// Serial 1 corresponds to 1900-01-01, but Excel (inheriting the bug from
// Lotus 1-2-3) wrongly treats 1900 as a leap year, so serials from 61
// (1900-03-01) onward match an epoch of 1899-12-30. The helpers below
// refuse the ambiguous low range, which includes the nonexistent
// 1900-02-29 (serial 60).
var excelEpoch = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)

var minExcelDate = time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC)

// maxExcelSerial is the serial of 9999-12-31, the last date representable
// in Excel.
const maxExcelSerial = 2958465

// daysSinceEpoch returns the whole number of days from 1899-12-30 to the
// given wall-clock date, using integer Julian Day Number arithmetic.
// (time.Sub cannot be used: time.Duration saturates at ±292 years, far
// short of Excel's year-9999 horizon.)
func daysSinceEpoch(year int, month time.Month, day int) int64 {
	// Fliegel & Van Flandern's formula for the Julian Day Number of a
	// proleptic Gregorian calendar date.
	jdn := func(y, m, d int64) int64 {
		return (1461*(y+4800+(m-14)/12))/4 + (367*(m-2-12*((m-14)/12)))/12 -
			(3*((y+4900+(m-14)/12)/100))/4 + d - 32075
	}
	return jdn(int64(year), int64(month), int64(day)) - jdn(1899, 12, 30)
}

// TimeToSerial converts a time to an Excel serial date number (1900 date
// system): whole days since 1899-12-30 plus a fraction for the time of
// day. Only the wall-clock fields of t are used; its location is ignored,
// and the result is UTC-neutral. Times before 1900-03-01 (serial 61) or
// after 9999-12-31 (serial 2958465) are rejected.
//
// Store the result in a cell and apply a date Style (e.g. NumberFormat
// NumFmtDate) so Excel renders it as a date.
func TimeToSerial(t time.Time) (float64, error) {
	wall := time.Date(t.Year(), t.Month(), t.Day(),
		t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), time.UTC)
	if wall.Before(minExcelDate) {
		return 0, fmt.Errorf("cells_foss: date %s is before 1900-03-01; Excel serials below 61 are ambiguous due to the 1900 leap-year bug",
			t.Format("2006-01-02"))
	}
	days := daysSinceEpoch(wall.Year(), wall.Month(), wall.Day())
	if days > maxExcelSerial {
		return 0, fmt.Errorf("cells_foss: date %s is beyond the maximum Excel date 9999-12-31 (serial %d)",
			t.Format("2006-01-02"), maxExcelSerial)
	}
	secs := wall.Hour()*3600 + wall.Minute()*60 + wall.Second()
	return float64(days) + (float64(secs)+float64(wall.Nanosecond())/1e9)/86400, nil
}

// SerialToTime converts an Excel serial date number (1900 date system) to
// a time.Time in UTC. The fraction of the serial is the time of day and is
// resolved to the nearest second. Serials must be >= 61 (1900-03-01) and
// <= 2958465 (9999-12-31).
func SerialToTime(serial float64) (time.Time, error) {
	if !(serial >= 61) { // also rejects NaN
		return time.Time{}, fmt.Errorf("cells_foss: serial %v is below 61 (1900-03-01); lower serials are ambiguous due to Excel's 1900 leap-year bug", serial)
	}
	days := int64(math.Floor(serial))
	if days > maxExcelSerial {
		return time.Time{}, fmt.Errorf("cells_foss: serial %v is beyond the maximum Excel date 9999-12-31 (serial %d)", serial, maxExcelSerial)
	}
	secs := int64(math.Round((serial - float64(days)) * 86400))
	return excelEpoch.AddDate(0, 0, int(days)).Add(time.Duration(secs) * time.Second), nil
}
