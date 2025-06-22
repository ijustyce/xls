//nolint:mnd
package xls

import (
	"encoding/binary"
	"io"
	"os"
	"unicode/utf16"

	"golang.org/x/text/encoding/charmap"
)

// xls workbook type
type WorkBook struct {
	Is5ver   bool
	Type     uint16
	Codepage uint16
	Xfs      []st_xf_data
	Fonts    []Font
	Formats  map[uint16]*Format
	// All the sheets from the workbook
	sheets        []*WorkSheet
	Author        string
	rs            io.ReadSeeker
	sstParser     *sstParser
	sst           []string
	continueUtf16 uint16
	continueRich  uint16
	continueApsb  uint32
	dateMode      uint16
}

// read workbook from ole2 file
func newWorkBookFromOle2(readSeeker io.ReadSeeker) *WorkBook {
	workBook := new(WorkBook)
	workBook.Formats = make(map[uint16]*Format)
	// wb.bts = bts
	workBook.rs = readSeeker
	workBook.sheets = make([]*WorkSheet, 0)
	workBook.Parse(readSeeker)

	return workBook
}

func (wb *WorkBook) Parse(buf io.ReadSeeker) {
	currentBOF := new(bof)
	previousBOF := new(bof)
	sstOffset := 0

	for {
		if err := binary.Read(buf, binary.LittleEndian, currentBOF); err == nil {
			sstOffset, previousBOF, currentBOF = wb.parseBof(buf, currentBOF, previousBOF, sstOffset)
		} else {
			break
		}
	}
}

func (wb *WorkBook) addXf(xf st_xf_data) {
	wb.Xfs = append(wb.Xfs, xf)
}

func (wb *WorkBook) addFont(font *FontInfo, buf io.ReadSeeker) {
	name, _ := wb.getString(buf, uint16(font.NameB))
	wb.Fonts = append(wb.Fonts, Font{Info: font, Name: name})
}

func (wb *WorkBook) addFormat(format *Format) {
	if wb.Formats == nil {
		os.Exit(1)
	}

	wb.Formats[format.Head.Index] = format
}

func decodeWindows1251(enc []byte) string {
	dec := charmap.Windows1251.NewDecoder()
	out, _ := dec.Bytes(enc)

	return string(out)
}

func (wb *WorkBook) getString(buf io.ReadSeeker, size uint16) (res string, err error) {
	if wb.Is5ver {
		bts := make([]byte, size)
		_, err = buf.Read(bts)
		res = decodeWindows1251(bts)
		// res = string(bts)
	} else {
		richtextNum := uint16(0)
		phoneticSize := uint32(0)
		var flag byte
		err = binary.Read(buf, binary.LittleEndian, &flag)

		if flag&0x8 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &richtextNum)
		} else if wb.continueRich > 0 {
			richtextNum = wb.continueRich
			wb.continueRich = 0
		}

		if flag&0x4 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &phoneticSize)
		} else if wb.continueApsb > 0 {
			phoneticSize = wb.continueApsb
			wb.continueApsb = 0
		}

		if flag&0x1 != 0 {
			bts := make([]uint16, size)
			i := uint16(0)

			for ; i < size && err == nil; i++ {
				err = binary.Read(buf, binary.LittleEndian, &bts[i])
			}

			// when eof found, we dont want to append last element
			var runes []rune
			if err == io.EOF {
				i--
			}

			runes = utf16.Decode(bts[:i])
			res = string(runes)

			if i < size {
				wb.continueUtf16 = size - i
			}
		} else {
			bts := make([]byte, size)
			var n int
			n, err = buf.Read(bts)
			if uint16(n) < size {
				wb.continueUtf16 = size - uint16(n)
				err = io.EOF
			}

			bts1 := make([]uint16, n)
			for k, v := range bts[:n] {
				bts1[k] = uint16(v)
			}

			runes := utf16.Decode(bts1)
			res = string(runes)
		}

		if richtextNum > 0 {
			var bts []byte
			var seekSize int64

			if wb.Is5ver {
				seekSize = int64(2 * richtextNum)
			} else {
				seekSize = int64(4 * richtextNum)
			}

			bts = make([]byte, seekSize)
			err = binary.Read(buf, binary.LittleEndian, bts)

			if err == io.EOF {
				wb.continueRich = richtextNum
			}
		}

		if phoneticSize > 0 {
			bts := make([]byte, phoneticSize)
			err = binary.Read(buf, binary.LittleEndian, bts)

			if err == io.EOF {
				wb.continueApsb = phoneticSize
			}
		}
	}
	return
}

func (wb *WorkBook) addSheet(sheet *boundsheet, buf io.ReadSeeker) {
	name, _ := wb.getString(buf, uint16(sheet.Name))
	wb.sheets = append(wb.sheets, &WorkSheet{bs: sheet, Name: name, wb: wb, Visibility: TWorkSheetVisibility(sheet.Visible)})
}

// reading a sheet from the compress file to memory, you should call this before you try to get anything from sheet
func (wb *WorkBook) prepareSheet(sheet *WorkSheet) {
	_, err := wb.rs.Seek(int64(sheet.bs.Filepos), 0)
	if err != nil {
		panic("xls: prepareSheet: " + err.Error())
	}

	sheet.parse(wb.rs)
}

// Get one sheet by its number
func (wb *WorkBook) GetSheet(num int) *WorkSheet {
	if num < len(wb.sheets) {
		s := wb.sheets[num]
		if !s.parsed {
			wb.prepareSheet(s)
		}

		return s
	}

	return nil
}

// Get the number of all sheets, look into example
func (wb *WorkBook) NumSheets() int {
	return len(wb.sheets)
}

func (wb *WorkBook) GetSheetByName(sheetName string) *WorkSheet {
	for _, sheet := range wb.sheets {
		if sheet.Name == sheetName {
			if !sheet.parsed {
				wb.prepareSheet(sheet)
			}

			return sheet
		}
	}

	return nil
}

func (wb *WorkBook) GetFirstSheet() *WorkSheet {
	sheet := wb.sheets[0]

	if !sheet.parsed {
		wb.prepareSheet(sheet)
	}

	return sheet
}

// ReadAllCells reads all cell data from the workbook up to a maximum number of rows.
// Note: This may consume significant memory for large files.
func (wb *WorkBook) ReadAllCells(maxRowsTotal int) [][]string {
	var res [][]string

	for _, sheet := range wb.sheets {
		if len(res) >= maxRowsTotal {
			break
		}

		wb.prepareSheet(sheet)

		if sheet.MaxRow == 0 {
			continue
		}

		remaining := maxRowsTotal - len(res)
		rowCount := sheet.MaxRow + 1

		if remaining < int(rowCount) {
			rowCount = uint16(remaining)
		}

		temp := make([][]string, rowCount)

		for rowIndex, row := range sheet.rows {
			if rowIndex >= rowCount {
				break
			}

			if len(row.cols) == 0 {
				continue
			}

			// Pre-allocate a reasonably sized slice
			data := make([]string, 0, 10)

			for _, col := range row.cols {
				str := col.String(wb)
				last := col.LastCol()
				first := col.FirstCol()

				// Extend slice if needed
				if len(data) <= int(last) {
					newData := make([]string, int(last)+1)
					copy(newData, data)
					data = newData
				}

				for j, s := range str {
					data[int(first)+j] = s
				}
			}

			temp[rowIndex] = data
		}

		res = append(res, temp...)
	}

	return res
}
