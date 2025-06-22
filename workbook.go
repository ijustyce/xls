//nolint:mnd
package xls

import (
	"bytes"
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
	b := new(bof)
	bofPre := new(bof)
	offset := 0

	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err == nil {
			bofPre, b, offset = wb.parseBof(buf, b, bofPre, offset)
		} else {
			break
		}
	}
}

func (wb *WorkBook) addXf(xf st_xf_data) {
	wb.Xfs = append(wb.Xfs, xf)
}

func (wb *WorkBook) addFont(font *FontInfo, buf io.ReadSeeker) {
	name, _ := wb.get_string(buf, uint16(font.NameB))
	wb.Fonts = append(wb.Fonts, Font{Info: font, Name: name})
}

func (wb *WorkBook) addFormat(format *Format) {
	if wb.Formats == nil {
		os.Exit(1)
	}

	wb.Formats[format.Head.Index] = format
}

func (wb *WorkBook) parseBof(buf io.ReadSeeker, b *bof, pre *bof, offsetPre int) (after *bof, afterUsing *bof, offset int) {
	after = b
	afterUsing = pre
	bts := make([]byte, b.Size)
	binary.Read(buf, binary.LittleEndian, bts)
	bufItem := bytes.NewReader(bts)

	switch b.ID {
	case 0x809:
		bif := new(biffHeader)
		binary.Read(bufItem, binary.LittleEndian, bif)

		if bif.Ver != 0x600 {
			wb.Is5ver = true
		}

		wb.Type = bif.Type
	case 0x042: // CODEPAGE
		binary.Read(bufItem, binary.LittleEndian, &wb.Codepage)
	case 0x3c: // CONTINUE
		if pre.ID == 0xfc {
			var size uint16
			var err error

			if wb.continueUtf16 >= 1 {
				size = wb.continueUtf16
				wb.continueUtf16 = 0
			} else {
				err = binary.Read(bufItem, binary.LittleEndian, &size)
			}

			for err == nil && offsetPre < len(wb.sst) {
				var str string
				str, err = wb.get_string(bufItem, size)
				wb.sst[offsetPre] = wb.sst[offsetPre] + str

				if err == io.EOF {
					break
				}

				offsetPre++
				err = binary.Read(bufItem, binary.LittleEndian, &size)
			}
		}

		offset = offsetPre
		after = pre
		afterUsing = b
	case 0xfc: // SST
		info := new(SstInfo)
		binary.Read(bufItem, binary.LittleEndian, info)
		wb.sst = make([]string, info.Count)
		var size uint16
		i := 0
		// dont forget to initialize offset
		offset = 0

		for ; i < int(info.Count); i++ {
			var err error
			err = binary.Read(bufItem, binary.LittleEndian, &size)

			if err == nil {
				var str string
				str, err = wb.get_string(bufItem, size)
				wb.sst[i] = wb.sst[i] + str
			}

			if err == io.EOF {
				break
			}
		}

		offset = i
	case 0x85: // boundsheet
		bs := new(boundsheet)
		binary.Read(bufItem, binary.LittleEndian, bs)
		// different for BIFF5 and BIFF8
		wb.addSheet(bs, bufItem)
	case 0x0e0: // XF
		if wb.Is5ver {
			xf := new(Xf5)
			binary.Read(bufItem, binary.LittleEndian, xf)
			wb.addXf(xf)
		} else {
			xf := new(Xf8)
			binary.Read(bufItem, binary.LittleEndian, xf)
			wb.addXf(xf)
		}
	case 0x031: // FONT
		f := new(FontInfo)
		binary.Read(bufItem, binary.LittleEndian, f)
		wb.addFont(f, bufItem)
	case 0x41E: // FORMAT
		font := new(Format)
		binary.Read(bufItem, binary.LittleEndian, &font.Head)
		font.str, _ = wb.get_string(bufItem, font.Head.Size)
		wb.addFormat(font)
	case 0x22: // DATEMODE
		binary.Read(bufItem, binary.LittleEndian, &wb.dateMode)
	}

	return
}

func decodeWindows1251(enc []byte) string {
	dec := charmap.Windows1251.NewDecoder()
	out, _ := dec.Bytes(enc)

	return string(out)
}

func (wb *WorkBook) get_string(buf io.ReadSeeker, size uint16) (res string, err error) {
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
				i = i - 1
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
			var bts []byte
			bts = make([]byte, phoneticSize)
			err = binary.Read(buf, binary.LittleEndian, bts)

			if err == io.EOF {
				wb.continueApsb = phoneticSize
			}
		}
	}
	return
}

func (wb *WorkBook) addSheet(sheet *boundsheet, buf io.ReadSeeker) {
	name, _ := wb.get_string(buf, uint16(sheet.Name))
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

// helper function to read all cells from file
// Notice: the max value is the limit of the max capacity of lines.
// Warning: the helper function will need big memeory if file is large.
func (wb *WorkBook) ReadAllCells(max int) (res [][]string) {
	res = make([][]string, 0)
	for _, sheet := range wb.sheets {
		if len(res) < max {
			max = max - len(res)

			wb.prepareSheet(sheet)

			if sheet.MaxRow != 0 {
				leng := int(sheet.MaxRow) + 1
				if max < leng {
					leng = max
				}

				temp := make([][]string, leng)
				for k, row := range sheet.rows {
					data := make([]string, 0)

					if len(row.cols) > 0 {
						for _, col := range row.cols {
							if uint16(len(data)) <= col.LastCol() {
								data = append(data, make([]string, col.LastCol()-uint16(len(data))+1)...)
							}

							str := col.String(wb)

							for i := uint16(0); i < col.LastCol()-col.FirstCol()+1; i++ {
								data[col.FirstCol()+i] = str[i]
							}
						}

						if leng > int(k) {
							temp[k] = data
						}
					}
				}

				res = append(res, temp...)
			}
		}
	}
	return
}
