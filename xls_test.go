package xls

import (
	"testing"
)

func TestOpen(t *testing.T) {
	// Open test XLS file
	wb, err := Open("testdata/bigtable.xls")
	if err != nil {
		t.Fatalf("failed to open XLS file: %v", err)
	}

	// Verify first sheet exists
	sheet := wb.GetSheet(0)
	if sheet == nil {
		t.Fatal("expected non-nil sheet at index 0")
	}

	t.Logf("Opened sheet: %q with %d rows", sheet.Name, sheet.MaxRow)

	// Inspect specific rows (example rows: 265 to 267)
	for i := 265; i <= 267; i++ {
		row := sheet.Row(i)
		if row == nil {
			t.Logf("row %d is nil", i)
			continue
		}
		t.Logf("row %d: firstCol=%d, lastCol=%d", i, row.FirstCol(), row.LastCol())
		for col := row.FirstCol(); col < row.LastCol(); col++ {
			val := row.Col(col)
			t.Logf("  col %d: %q", col, val)
		}
	}
}
