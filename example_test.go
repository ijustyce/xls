package xls

import (
	"fmt"
)

// ExampleOpen demonstrates how to open an XLS file and access basic metadata.
func ExampleOpen() {
	xlFile, err := Open("testdata/Table.xls")
	if err != nil {
		fmt.Println("failed to open XLS:", err)
		return
	}

	// Print workbook author metadata
	fmt.Println("Author:", xlFile.Author)
}

// ExampleWorkBook_NumberSheets shows how to list all sheet names in the workbook.
func ExampleWorkBook_NumSheets() {
	xlFile, err := Open("testdata/Table.xls")
	if err != nil {
		fmt.Println("failed to open XLS:", err)
		return
	}

	// Iterate over all sheets and print their names
	for i := 0; i < xlFile.NumSheets(); i++ {
		sheet := xlFile.GetSheet(i)
		fmt.Println("Sheet:", sheet.Name)
	}
}

// ExampleWorkBook_GetSheet reads the first sheet and prints the first two columns of each row.
func ExampleWorkBook_GetSheet() {
	xlFile, err := Open("testdata/Table.xls")
	if err != nil {
		fmt.Println("failed to open XLS:", err)
		return
	}

	sheet := xlFile.GetSheet(0)
	if sheet == nil {
		fmt.Println("sheet not found")
		return
	}

	fmt.Printf("Total Lines: %d (%s)", sheet.MaxRow, sheet.Name)

	// Iterate over all rows and print values from the first two columns
	for i := 0; i <= int(sheet.MaxRow); i++ {
		row := sheet.Row(i)
		col1 := row.Col(0)
		col2 := row.Col(1)
		fmt.Printf("\n%s, %s", col1, col2)
	}
}
