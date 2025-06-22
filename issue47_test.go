package xls

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// TestIssue47 checks that each .xls file in testdata/ matches its corresponding .xlsx file.
// It uses the CompareXlsXlsx function to compare file contents.
// This test was originally created to verify issue #47 but now serves as a regression test for .xls/.xlsx parity.
func TestIssue47(t *testing.T) {
	t.Parallel()
	testdataPath := "testdata"

	// Read the directory contents
	entries, err := os.ReadDir(testdataPath)
	if err != nil {
		t.Fatalf("cannot read testdata directory: %v", err)
	}

	// Loop over all files in the testdata directory
	for _, entry := range entries {
		// Skip directories and non-.xls files
		if entry.IsDir() || filepath.Ext(entry.Name()) != ".xls" {
			continue
		}

		// Build the full path to the .xls file
		xlsFile := filepath.Join(testdataPath, entry.Name())

		// Derive the expected .xlsx filename from the .xls filename
		baseName := strings.TrimSuffix(entry.Name(), ".xls")
		xlsxFile := filepath.Join(testdataPath, baseName+".xlsx")

		// Run each file pair as a subtest
		t.Run(entry.Name(), func(t *testing.T) {
			// Skip if .xlsx file does not exist
			if _, err := os.Stat(xlsxFile); err != nil {
				t.Skipf("Skipping test for %q: XLSX file %q not found", xlsFile, xlsxFile)
			}

			// Compare files and log any differences
			if diff := CompareXlsXlsx(xlsFile, xlsxFile); diff != "" {
				t.Errorf("Mismatch between %q and %q:\n%s", xlsFile, xlsxFile, diff)
			}
		})
	}
}
