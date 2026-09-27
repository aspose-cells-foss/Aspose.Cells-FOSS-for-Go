package cells_foss_test

import (
	"fmt"
	"path/filepath"
	"testing"

	cells_foss "github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Go/v26/aspose/cells_foss"
)

// BenchmarkNewWorkbook_Save measures creating a new workbook, populating
// it with data, and saving to disk.
func BenchmarkNewWorkbook_Save(b *testing.B) {
	dir := b.TempDir()
	for i := 0; i < b.N; i++ {
		wb := cells_foss.NewWorkbook()
		ws := wb.Worksheets[0]
		c := ws.Cells()
		for row := 1; row <= 100; row++ {
			c.Set(fmt.Sprintf("A%d", row), fmt.Sprintf("Row%d", row))
			c.Set(fmt.Sprintf("B%d", row), float64(row))
		}
		outPath := filepath.Join(dir, fmt.Sprintf("bench_new_%d.xlsx", i))
		if err := wb.Save(outPath); err != nil {
			b.Fatal(err)
		}
	}
}

// BenchmarkLoadWorkbook measures loading a workbook from disk.
func BenchmarkLoadWorkbook(b *testing.B) {
	// Create a test file first.
	dir := b.TempDir()
	testFile := filepath.Join(dir, "bench_load.xlsx")
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	c := ws.Cells()
	for row := 1; row <= 100; row++ {
		c.Set(fmt.Sprintf("A%d", row), fmt.Sprintf("Row%d", row))
		c.Set(fmt.Sprintf("B%d", row), float64(row))
	}
	wb.Save(testFile)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := cells_foss.LoadWorkbook(testFile)
		if err != nil {
			b.Fatal(err)
		}
	}
}

// BenchmarkLargeSheet_Save measures saving a workbook with a large sheet
// (1000 rows × 10 columns = 10,000 cells).
func BenchmarkLargeSheet_Save(b *testing.B) {
	dir := b.TempDir()
	for i := 0; i < b.N; i++ {
		wb := cells_foss.NewWorkbook()
		ws := wb.Worksheets[0]
		c := ws.Cells()

		// Header row.
		for col := 0; col < 10; col++ {
			ref := cells_foss.NumToCol(col) + "1"
			c.Set(ref, fmt.Sprintf("Col%d", col))
		}

		// Data rows.
		for row := 2; row <= 1000; row++ {
			for col := 0; col < 10; col++ {
				ref := cells_foss.NumToCol(col) + fmt.Sprintf("%d", row)
				if col%2 == 0 {
					c.Set(ref, fmt.Sprintf("Data%d_%d", row, col))
				} else {
					c.Set(ref, float64(row*10+col))
				}
			}
		}

		outPath := filepath.Join(dir, fmt.Sprintf("bench_large_%d.xlsx", i))
		if err := wb.Save(outPath); err != nil {
			b.Fatal(err)
		}
	}
}

// BenchmarkStreamingReader measures row-by-row streaming through a large file.
func BenchmarkStreamingReader(b *testing.B) {
	// Create a test file first.
	dir := b.TempDir()
	testFile := filepath.Join(dir, "bench_stream.xlsx")
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	c := ws.Cells()
	for row := 1; row <= 1000; row++ {
		c.Set(fmt.Sprintf("A%d", row), fmt.Sprintf("Row%d", row))
		c.Set(fmt.Sprintf("B%d", row), float64(row))
		c.Set(fmt.Sprintf("C%d", row), fmt.Sprintf("Data%d", row))
	}
	wb.Save(testFile)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		sr := cells_foss.NewStreamingReader(testFile)
		rowCount := 0
		err := sr.ProcessRows("", func(rowIdx int, cells map[string]string) error {
			rowCount++
			return nil
		})
		if err != nil {
			b.Fatal(err)
		}
		if rowCount != 1000 {
			b.Errorf("expected 1000 rows, got %d", rowCount)
		}
	}
}

// BenchmarkUnmodifiedSave_RoundTrip measures saving an unmodified loaded
// workbook (should reuse cached source XML).
func BenchmarkUnmodifiedSave_RoundTrip(b *testing.B) {
	dir := b.TempDir()
	testFile := filepath.Join(dir, "bench_rt.xlsx")
	wb := cells_foss.NewWorkbook()
	ws := wb.Worksheets[0]
	c := ws.Cells()
	for row := 1; row <= 100; row++ {
		c.Set(fmt.Sprintf("A%d", row), fmt.Sprintf("Row%d", row))
		c.Set(fmt.Sprintf("B%d", row), float64(row))
	}
	wb.Save(testFile)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		loaded, err := cells_foss.LoadWorkbook(testFile)
		if err != nil {
			b.Fatal(err)
		}
		outPath := filepath.Join(dir, fmt.Sprintf("bench_rt_out_%d.xlsx", i))
		if err := loaded.Save(outPath); err != nil {
			b.Fatal(err)
		}
	}
}
