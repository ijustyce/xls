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

func (w *WorkBook) Parse(buf io.ReadSeeker) {
	b := new(bof)
	bofPre := new(bof)
	// buf := bytes.NewReader(bts)
	offset := 0

	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err == nil {
			bofPre, b, offset = w.parseBof(buf, b, bofPre, offset)
		} else {
			break
		}
	}
}

func (w *WorkBook) addXf(xf st_xf_data) {
	w.Xfs = append(w.Xfs, xf)
}

func (w *WorkBook) addFont(font *FontInfo, buf io.ReadSeeker) {
	name, _ := w.get_string(buf, uint16(font.NameB))
	w.Fonts = append(w.Fonts, Font{Info: font, Name: name})
}

func (w *WorkBook) addFormat(format *Format) {
	if w.Formats == nil {
		os.Exit(1)
	}
	w.Formats[format.Head.Index] = format
}

func (wb *WorkBook) parseBof(buf io.ReadSeeker, b *bof, pre *bof, offset_pre int) (after *bof, after_using *bof, offset int) {
	after = b
	after_using = pre
	bts := make([]byte, b.Size)
	binary.Read(buf, binary.LittleEndian, bts)
	buf_item := bytes.NewReader(bts)
	switch b.ID {
	case 0x809:
		bif := new(biffHeader)
		binary.Read(buf_item, binary.LittleEndian, bif)
		if bif.Ver != 0x600 {
			wb.Is5ver = true
		}
		wb.Type = bif.Type
	case 0x042: // CODEPAGE
		binary.Read(buf_item, binary.LittleEndian, &wb.Codepage)
	case 0x3c: // CONTINUE
		if pre.ID == 0xfc {
			var size uint16
			var err error
			if wb.continueUtf16 >= 1 {
				size = wb.continueUtf16
				wb.continueUtf16 = 0
			} else {
				err = binary.Read(buf_item, binary.LittleEndian, &size)
			}

			for err == nil && offset_pre < len(wb.sst) {
				var str string
				str, err = wb.get_string(buf_item, size)
				wb.sst[offset_pre] = wb.sst[offset_pre] + str

				if err == io.EOF {
					break
				}

				offset_pre++
				err = binary.Read(buf_item, binary.LittleEndian, &size)
			}
		}
		offset = offset_pre
		after = pre
		after_using = b
	case 0xfc: // SST
		info := new(SstInfo)
		binary.Read(buf_item, binary.LittleEndian, info)
		wb.sst = make([]string, info.Count)
		var size uint16
		i := 0
		// dont forget to initialize offset
		offset = 0

		for ; i < int(info.Count); i++ {
			var err error
			err = binary.Read(buf_item, binary.LittleEndian, &size)
			if err == nil {
				var str string
				str, err = wb.get_string(buf_item, size)
				wb.sst[i] = wb.sst[i] + str
			}

			if err == io.EOF {
				break
			}
		}
		offset = i
	case 0x85: // boundsheet
		bs := new(boundsheet)
		binary.Read(buf_item, binary.LittleEndian, bs)
		// different for BIFF5 and BIFF8
		wb.addSheet(bs, buf_item)
	case 0x0e0: // XF
		if wb.Is5ver {
			xf := new(Xf5)
			binary.Read(buf_item, binary.LittleEndian, xf)
			wb.addXf(xf)
		} else {
			xf := new(Xf8)
			binary.Read(buf_item, binary.LittleEndian, xf)
			wb.addXf(xf)
		}
	case 0x031: // FONT
		f := new(FontInfo)
		binary.Read(buf_item, binary.LittleEndian, f)
		wb.addFont(f, buf_item)
	case 0x41E: // FORMAT
		font := new(Format)
		binary.Read(buf_item, binary.LittleEndian, &font.Head)
		font.str, _ = wb.get_string(buf_item, font.Head.Size)
		wb.addFormat(font)
	case 0x22: // DATEMODE
		binary.Read(buf_item, binary.LittleEndian, &wb.dateMode)
	}
	return
}

func decodeWindows1251(enc []byte) string {
	dec := charmap.Windows1251.NewDecoder()
	out, _ := dec.Bytes(enc)
	return string(out)
}

func (w *WorkBook) get_string(buf io.ReadSeeker, size uint16) (res string, err error) {
	if w.Is5ver {
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
		} else if w.continueRich > 0 {
			richtextNum = w.continueRich
			w.continueRich = 0
		}

		if flag&0x4 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &phoneticSize)
		} else if w.continueApsb > 0 {
			phoneticSize = w.continueApsb
			w.continueApsb = 0
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
				w.continueUtf16 = size - i
			}
		} else {
			bts := make([]byte, size)
			var n int
			n, err = buf.Read(bts)
			if uint16(n) < size {
				w.continueUtf16 = size - uint16(n)
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
			if w.Is5ver {
				seekSize = int64(2 * richtextNum)
			} else {
				seekSize = int64(4 * richtextNum)
			}
			bts = make([]byte, seekSize)
			err = binary.Read(buf, binary.LittleEndian, bts)
			if err == io.EOF {
				w.continueRich = richtextNum
			}
		}

		if phoneticSize > 0 {
			var bts []byte
			bts = make([]byte, phoneticSize)
			err = binary.Read(buf, binary.LittleEndian, bts)
			if err == io.EOF {
				w.continueApsb = phoneticSize
			}
		}
	}
	return
}

func (w *WorkBook) addSheet(sheet *boundsheet, buf io.ReadSeeker) {
	name, _ := w.get_string(buf, uint16(sheet.Name))
	w.sheets = append(w.sheets, &WorkSheet{bs: sheet, Name: name, wb: w, Visibility: TWorkSheetVisibility(sheet.Visible)})
}

// reading a sheet from the compress file to memory, you should call this before you try to get anything from sheet
func (w *WorkBook) prepareSheet(sheet *WorkSheet) {
	w.rs.Seek(int64(sheet.bs.Filepos), 0)
	sheet.parse(w.rs)
}

// Get one sheet by its number
func (w *WorkBook) GetSheet(num int) *WorkSheet {
	if num < len(w.sheets) {
		s := w.sheets[num]
		if !s.parsed {
			w.prepareSheet(s)
		}
		return s
	}

	return nil
}

// Get the number of all sheets, look into example
func (w *WorkBook) NumSheets() int {
	return len(w.sheets)
}

func (w *WorkBook) GetSheetByName(sheetName string) *WorkSheet {
	for _, sheet := range w.sheets {
		if sheet.Name == sheetName {
			if !sheet.parsed {
				w.prepareSheet(sheet)
			}

			return sheet
		}
	}

	return nil
}

func (w *WorkBook) GetFirstSheet() *WorkSheet {
	sheet := w.sheets[0]

	if !sheet.parsed {
		w.prepareSheet(sheet)
	}

	return sheet
}

// helper function to read all cells from file
// Notice: the max value is the limit of the max capacity of lines.
// Warning: the helper function will need big memeory if file is large.
func (w *WorkBook) ReadAllCells(max int) (res [][]string) {
	res = make([][]string, 0)
	for _, sheet := range w.sheets {
		if len(res) < max {
			max = max - len(res)

			w.prepareSheet(sheet)
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
							str := col.String(w)

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
