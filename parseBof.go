package xls

import (
	"bytes"
	"encoding/binary"
	"io"
)

// Handler function type for BIFF records.
type recordHandler func(wb *WorkBook, data []byte, prevBOF *bof, offsetPre int) (offset int, after *bof, afterUsing *bof)

var recordHandlers = map[uint16]recordHandler{
	0x809: handleBOF,
	0x042: handleCodepage,
	0x85:  handleBoundSheet,
	0x0e0: handleXF,
	0x031: handleFont,
	0x41E: handleFormat,
	0x22:  handleDateMode,
}

type sstParser struct {
	wb       *WorkBook
	reader   *bytes.Reader
	sstIndex int
	strCount int
}

func (wb *WorkBook) parseBof(buf io.ReadSeeker, currentBOF *bof, previousBOF *bof, offsetPre int) (after *bof, afterUsing *bof, offset int) {
	data := make([]byte, currentBOF.Size)
	binary.Read(buf, binary.LittleEndian, data)

	after = currentBOF
	afterUsing = previousBOF
	offset = offsetPre

	switch currentBOF.ID {
	case 0xfc: // SST
		parser := &sstParser{wb: wb}
		parser.parseSST(data)
		wb.sstParser = parser
		offset = parser.sstIndex
	case 0x3c: // CONTINUE
		if previousBOF.ID == 0xfc && wb.sstParser != nil {
			wb.sstParser.parseContinue(data)
			offset = wb.sstParser.sstIndex
			after = previousBOF
			afterUsing = currentBOF
		}
	default:
		handler := recordHandlers[currentBOF.ID]
		if handler != nil {
			newOffset, newAfter, newAfterUsing := handler(wb, data, previousBOF, offsetPre)
			if newAfter != nil {
				after = newAfter
			}

			if newAfterUsing != nil {
				afterUsing = newAfterUsing
			}

			offset = newOffset
		}
	}

	return
}

func handleBOF(workBook *WorkBook, data []byte, _ *bof, offsetPre int) (int, *bof, *bof) {
	buf := bytes.NewReader(data)
	bif := new(biffHeader)
	binary.Read(buf, binary.LittleEndian, bif)

	if bif.Ver != 0x600 {
		workBook.Is5ver = true
	}

	workBook.Type = bif.Type

	return offsetPre, nil, nil
}

func handleCodepage(wb *WorkBook, data []byte, _ *bof, offsetPre int) (int, *bof, *bof) {
	buf := bytes.NewReader(data)
	binary.Read(buf, binary.LittleEndian, &wb.Codepage)

	return offsetPre, nil, nil
}

func handleBoundSheet(wb *WorkBook, data []byte, _ *bof, offsetPre int) (int, *bof, *bof) {
	buf := bytes.NewReader(data)
	bs := new(boundsheet)
	binary.Read(buf, binary.LittleEndian, bs)
	wb.addSheet(bs, buf)

	return offsetPre, nil, nil
}

func handleXF(workBook *WorkBook, data []byte, _ *bof, offsetPre int) (int, *bof, *bof) {
	buf := bytes.NewReader(data)

	if workBook.Is5ver {
		xf := new(Xf5)
		binary.Read(buf, binary.LittleEndian, xf)
		workBook.addXf(xf)
	} else {
		xf := new(Xf8)
		binary.Read(buf, binary.LittleEndian, xf)
		workBook.addXf(xf)
	}

	return offsetPre, nil, nil
}

func handleFont(workBook *WorkBook, data []byte, _ *bof, offsetPre int) (int, *bof, *bof) {
	buf := bytes.NewReader(data)
	f := new(FontInfo)
	binary.Read(buf, binary.LittleEndian, f)
	workBook.addFont(f, buf)

	return offsetPre, nil, nil
}

func handleFormat(workBook *WorkBook, data []byte, _ *bof, offsetPre int) (int, *bof, *bof) {
	buf := bytes.NewReader(data)
	format := new(Format)
	binary.Read(buf, binary.LittleEndian, &format.Head)
	format.str, _ = workBook.getString(buf, format.Head.Size)
	workBook.addFormat(format)

	return offsetPre, nil, nil
}

func handleDateMode(workBook *WorkBook, data []byte, _ *bof, offsetPre int) (int, *bof, *bof) {
	buf := bytes.NewReader(data)
	binary.Read(buf, binary.LittleEndian, &workBook.dateMode)

	return offsetPre, nil, nil
}

func (p *sstParser) parseSST(data []byte) {
	p.reader = bytes.NewReader(data)

	info := new(SstInfo)
	binary.Read(p.reader, binary.LittleEndian, info)

	p.wb.sst = make([]string, info.Count)
	p.strCount = int(info.Count)

	p.parseStrings()
}

func (p *sstParser) parseContinue(data []byte) {
	p.reader = bytes.NewReader(data)
	p.parseStrings()
}

func (p *sstParser) parseStrings() {
	for p.sstIndex < p.strCount {
		var size uint16
		var err error

		if p.wb.continueUtf16 > 0 {
			size = p.wb.continueUtf16
			p.wb.continueUtf16 = 0
		} else {
			err = binary.Read(p.reader, binary.LittleEndian, &size)
			if err != nil {
				break
			}
		}

		str, err := p.wb.getString(p.reader, size)
		p.wb.sst[p.sstIndex] += str

		if err == io.EOF {
			break
		}

		p.sstIndex++
	}
}
