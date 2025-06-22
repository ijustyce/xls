package xls

import (
	"fmt"
	"math"
	"strconv"
	"time"

	"github.com/tealeg/xlsx"
)

// CompareXlsXlsx compares the content of an XLS file against an XLSX file.
// It returns an empty string if the files are considered equivalent,
// or a string describing the first encountered mismatch.
func CompareXlsXlsx(xlsFilePath, xlsxFilePath string) string {
	// Open .xls and .xlsx files
	xlsFile, err := Open(xlsFilePath)
	if err != nil {
		return fmt.Sprintf("Cannot open XLS file: %s", err)
	}
	xlsxFile, err := xlsx.OpenFile(xlsxFilePath)
	if err != nil {
		return fmt.Sprintf("Cannot open XLSX file: %s", err)
	}

	for sheetIdx, xlsxSheet := range xlsxFile.Sheets {
		xlsSheet := xlsFile.GetSheet(sheetIdx)
		if xlsSheet == nil {
			return fmt.Sprintf("Missing XLS sheet at index %d", sheetIdx)
		}

		for rowIdx, xlsxRow := range xlsxSheet.Rows {
			xlsRow := xlsSheet.Row(rowIdx)
			if xlsRow == nil {
				continue // sparse XLS row
			}

			for colIdx, xlsxCell := range xlsxRow.Cells {
				// Prefer formatted XLSX value, fallback to raw string
				xlsxRaw := xlsxCell.String()
				if val, err := xlsxCell.FormattedValue(); err == nil {
					xlsxRaw = val
				}
				xlsRaw := xlsRow.Col(colIdx)

				// Try normalizing both to comparable format (e.g., date/time)
				normXlsx := normalizeExcelCell(xlsxRaw)
				normXls := normalizeExcelCell(xlsRaw)

				if normXlsx == normXls {
					continue // exact match after normalization
				}

				// Try comparing as float (e.g., numeric Excel serials)
				xlsFloat, xlsErr := strconv.ParseFloat(xlsRaw, 64)
				xlsxFloat, xlsxErr := strconv.ParseFloat(xlsxRaw, 64)
				if xlsErr == nil && xlsxErr == nil {
					if math.Abs(xlsFloat-xlsxFloat) < 1e-7 {
						continue // numerically equal
					}
					xlsTime := excelEpoch.Add(time.Duration(xlsFloat * 24 * float64(time.Hour)))
					xlsxTime := excelEpoch.Add(time.Duration(xlsxFloat * 24 * float64(time.Hour)))
					if xlsTime.Truncate(time.Second).Equal(xlsxTime.Truncate(time.Second)) {
						continue // time-equal
					}
					return fmt.Sprintf(
						"Sheet %q, row %d, col %d: numeric mismatch — xls: %f (%s), xlsx: %f (%s)",
						xlsxSheet.Name, rowIdx, colIdx,
						xlsFloat, xlsTime.Format("2006-01-02 15:04:05"),
						xlsxFloat, xlsxTime.Format("2006-01-02 15:04:05"),
					)
				}

				// Final fallback: literal string mismatch
				return fmt.Sprintf(
					"Sheet %q, row %d, col %d: mismatch — xls: %q → %q, xlsx: %q → %q",
					xlsxSheet.Name, rowIdx, colIdx,
					xlsRaw, normXls,
					xlsxRaw, normXlsx,
				)
			}
		}
	}

	return "" // no mismatch found
}

// Excel's epoch for serial date calculation is 1899-12-30
var excelEpoch = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)

// normalizeExcelCell attempts to parse and format an Excel float serial as a human-readable string.
// It falls back to the original string if it's not numeric.
func normalizeExcelCell(val string) string {
	f, err := strconv.ParseFloat(val, 64)
	if err == nil {
		t := excelEpoch.Add(time.Duration(f * 24 * float64(time.Hour)))
		return t.Format("2006-01-02") // ISO date
	}

	// Attempt to parse common localized date strings
	for _, layout := range []string{"02.01.2006", "2.1.2006"} {
		if t, err := time.Parse(layout, val); err == nil {
			return t.Format("2006-01-02")
		}
	}

	return val // return as-is if not convertible
}
