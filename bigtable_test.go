package xls

import (
	"fmt"
	"testing"
	"time"
)

// TestBigTable verifies that values in specific columns match expected patterns
// across thousands of rows. It checks for content and formatting correctness.
func TestBigTable(t *testing.T) {
	const filePath = "testdata/bigtable.xls"

	// Attempt to open the test XLS file
	xlFile, err := Open(filePath)
	if err != nil {
		t.Fatalf("failed to open XLS file %q: %v", filePath, err)
	}

	// Get the first worksheet
	sheet := xlFile.GetSheet(0)
	if sheet == nil {
		t.Fatal("failed to get sheet at index 0")
	}

	// Initialize counter and start dates
	cnt1, cnt2, cnt3 := 1, 10000, 20000
	date1 := mustParseDate("2015-01-01")
	date2 := mustParseDate("2016-01-01")
	date3 := mustParseDate("2017-01-01")

	// Define how many rows we expect to validate
	const rowCount = 4999

	for i := 1; i <= rowCount; i++ {
		row := sheet.Row(i)
		if row == nil {
			t.Logf("row %d is nil, skipping", i)
			continue
		}

		expectedCol2 := fmt.Sprintf("%d от %s", cnt1, date1.Format("02.01.2006"))
		expectedCol5 := fmt.Sprintf("%d от %s", cnt2, date2.Format("02.01.2006"))
		expectedCol8 := fmt.Sprintf("%d от %s", cnt3, date3.Format("02.01.2006"))

		actualCol2 := row.Col(2)
		actualCol5 := row.Col(5)
		actualCol8 := row.Col(8)

		if actualCol2 != expectedCol2 {
			t.Errorf("row %d: col 2 mismatch: got %q, want %q", i, actualCol2, expectedCol2)
		}
		if actualCol5 != expectedCol5 {
			t.Errorf("row %d: col 5 mismatch: got %q, want %q", i, actualCol5, expectedCol5)
		}
		if actualCol8 != expectedCol8 {
			t.Errorf("row %d: col 8 mismatch: got %q, want %q", i, actualCol8, expectedCol8)
		}

		// Advance counters and dates
		cnt1++
		cnt2++
		cnt3++
		date1 = date1.AddDate(0, 0, 1)
		date2 = date2.AddDate(0, 0, 1)
		date3 = date3.AddDate(0, 0, 1)
	}
}

// mustParseDate parses a date in "YYYY-MM-DD" format and panics if it fails.
// Used to keep test code concise and clean.
func mustParseDate(value string) time.Time {
	date, err := time.Parse("2006-01-02", value)
	if err != nil {
		panic(fmt.Sprintf("invalid test date %q: %v", value, err))
	}
	return date
}
