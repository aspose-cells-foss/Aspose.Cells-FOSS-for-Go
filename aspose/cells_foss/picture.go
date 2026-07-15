package cells_foss

import (
	"fmt"
	"strings"
)

// emuPerPixel converts pixels to EMUs (English Metric Units) at 96 DPI.
// 1 px = 9525 EMU.
const emuPerPixel int64 = 9525

// Picture represents an image embedded in a worksheet.  The image data is
// stored in the Data field and written to xl/media/ during save.
type Picture struct {
	// Data holds the raw image bytes (PNG or JPEG).
	Data []byte

	// Format is the lower-case image format: "png" or "jpeg".
	Format string

	// Row is the 0-based row index where the top-left corner of the picture
	// is anchored.
	Row int

	// Col is the 0-based column index where the top-left corner of the
	// picture is anchored.
	Col int

	// RowOff is the vertical offset from the top of the anchor row, in EMUs.
	RowOff int64

	// ColOff is the horizontal offset from the left of the anchor column, in EMUs.
	ColOff int64

	// Width is the display width of the picture in pixels.
	Width int

	// Height is the display height of the picture in pixels.
	Height int

	// Name is the display name used in the drawing XML (e.g. "Picture 1").
	Name string
}

// NewPicture creates a Picture value from raw image bytes.  The format
// parameter must be "png" or "jpeg".  Callers must attach the picture to a
// worksheet with Worksheet.AddPicture before saving.
func NewPicture(data []byte, format string) *Picture {
	format = strings.ToLower(format)
	if format == "jpg" {
		format = "jpeg"
	}
	return &Picture{
		Data:    data,
		Format:  format,
		Name:    fmt.Sprintf("Picture %d", len(data)), // temporary; AddPicture reassigns
		RowOff:  0,
		ColOff:  0,
	}
}

// SetAnchor positions the picture at the given 0-based row and column.
func (p *Picture) SetAnchor(row, col int) {
	p.Row = row
	p.Col = col
}

// AddPicture attaches pic to the worksheet, assigns it a unique name, and
// marks the workbook as modified.  The picture data must be PNG or JPEG.
func (ws *Worksheet) AddPicture(pic *Picture) error {
	if pic == nil {
		return fmt.Errorf("cells_foss: cannot add nil Picture")
	}
	if len(pic.Data) == 0 {
		return fmt.Errorf("cells_foss: Picture.Data is empty")
	}
	if pic.Format != "png" && pic.Format != "jpeg" {
		return fmt.Errorf("cells_foss: unsupported picture format %q (must be png or jpeg)", pic.Format)
	}

	// Auto-assign a unique name.
	pic.Name = fmt.Sprintf("Picture %d", len(ws.Pictures)+1)

	ws.Pictures = append(ws.Pictures, pic)
	if ws.cells != nil && ws.cells.wb != nil {
		ws.cells.wb.Modified = true
	}
	return nil
}
