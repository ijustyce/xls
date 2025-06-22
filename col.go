package xls

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// content type
type contentHandler interface {
	String(*WorkBook) []string
	FirstCol() uint16
	LastCol() uint16
}

type Col struct {
	RowB      uint16
	FirstColB uint16
}

type Coler interface {
	Row() uint16
}

// Row returns the row index (0-based) where this column's data is located.
// This is used to identify the row number in the Excel sheet.
func (c *Col) Row() uint16 {
	return c.RowB
}

// FirstCol returns the starting column index (0-based) for this cell or cell range.
// In most cases, FirstCol == LastCol for a single-cell value.
func (c *Col) FirstCol() uint16 {
	return c.FirstColB
}

// LastCol returns the ending column index for this cell or cell range.
// By default, it equals FirstCol unless overridden in a derived type.
func (c *Col) LastCol() uint16 {
	return c.FirstColB
}

// String returns a string slice representation of the column's contents.
// This default implementation returns a single placeholder value and is
// meant to be overridden by concrete column types (e.g. NumberCol, RkCol).
func (c *Col) String(workBook *WorkBook) []string {
	return []string{"default"}
}

type XfRk struct {
	Index uint16
	Rk    RK
}

// String converts the RK value to its formatted string representation,
// depending on the associated cell format (Xf) and number format definition.
func (xf *XfRk) String(workBook *WorkBook) string {
	idx := int(xf.Index)
	if idx >= len(workBook.Xfs) {
		return xf.Rk.String() // fallback: no format info
	}

	formatNo := workBook.Xfs[idx].formatNo()

	// If format number is user-defined
	if formatNo >= 164 {
		return xf.renderCustomFormat(workBook, formatNo)
	}

	// Built-in date/time formats (based on OpenOffice Excel format spec)
	if isBuiltinDateFormat(formatNo) {
		return xf.renderDate(workBook)
	}

	return xf.Rk.String() // fallback: plain number
}

// renderCustomFormat handles user-defined Excel formats (formatNo >= 164).
func (xf *XfRk) renderCustomFormat(workBook *WorkBook, formatNo uint16) string {
	formatter := workBook.Formats[formatNo]
	if formatter == nil {
		return xf.Rk.String()
	}

	formatStr := strings.ToLower(formatter.str)

	// Treat as numeric if it looks like a general or number format
	if isNumericFormat(formatStr) {
		return xf.Rk.String()
	}

	// Otherwise treat as a date
	return xf.renderDate(workBook)
}

// renderDate extracts the underlying float value and renders it as a date
func (xf *XfRk) renderDate(workBook *WorkBook) string {
	intVal, floatVal, isFloat := xf.Rk.number()
	if !isFloat {
		floatVal = float64(intVal)
	}

	t := timeFromExcelTime(floatVal, workBook.dateMode == 1)

	// Use a general format for user-defined dates
	return t.Format("02.01.2006")
}

// isNumericFormat returns true if the format string appears to represent a number
func isNumericFormat(format string) bool {
	return (format == "general" ||
		strings.Contains(format, "#") ||
		strings.Contains(format, ".00")) &&
		!strings.Contains(format, "m/y") &&
		!strings.Contains(format, "d/y") &&
		!strings.Contains(format, "m.y") &&
		!strings.Contains(format, "d.y") &&
		!strings.Contains(format, "h:") &&
		!strings.Contains(format, "д.г")
}

// isBuiltinDateFormat checks if a format number is one of Excel's standard date formats
func isBuiltinDateFormat(fNo uint16) bool {
	return (14 <= fNo && fNo <= 17) || fNo == 22 ||
		(27 <= fNo && fNo <= 36) || (50 <= fNo && fNo <= 58)
}

type RK uint32

// number decodes the RK-encoded value into either an integer or a float.
// The RK format is a compact representation used in BIFF records.
func (rk RK) number() (intNum int64, floatNum float64, isFloat bool) {
	// Bit 0: value is multiplied by 100
	multiplied := rk&1 != 0

	// Bit 1: 0 = IEEE 754 float, 1 = signed 30-bit int
	isInteger := rk&2 != 0

	raw := int32(rk) >> 2

	if !isInteger {
		isFloat = true
		floatNum = math.Float64frombits(uint64(raw) << 34)

		if multiplied {
			floatNum *= 0.01
		}

		return
	}

	if multiplied {
		isFloat = true
		floatNum = float64(raw) * 0.01

		return
	}

	return int64(raw), 0, false
}

// String returns the RK value formatted as either a float or int string.
func (rk RK) String() string {
	i, f, isFloat := rk.number()
	if isFloat {
		return strconv.FormatFloat(f, 'f', -1, 64)
	}

	return strconv.FormatInt(i, 10)
}

var ErrIsInt = fmt.Errorf("is int")

// Float attempts to return the RK value as a float64.
//
// The RK format can encode either an integer or a float.
// If the RK value represents an integer, this method returns an error (ErrIsInt).
func (rk RK) Float() (float64, error) {
	_, f, isFloat := rk.number()
	if !isFloat {
		return 0, ErrIsInt
	}

	return f, nil
}

type MulrkCol struct {
	Col
	Xfrks    []XfRk
	LastColB uint16
}

// LastCol returns the last column index represented by this MulrkCol.
// This allows the caller to know how many adjacent cells are included.
func (c *MulrkCol) LastCol() uint16 {
	return c.LastColB
}

// String returns a string slice with formatted values for each cell in the multi-column RK group.
//
// Each entry in Xfrks corresponds to a cell, and the XfRk's String method is used
// to format its value according to the workbook's formatting rules.
func (c *MulrkCol) String(wb *WorkBook) []string {
	res := make([]string, len(c.Xfrks))
	for i := 0; i < len(c.Xfrks); i++ {
		xfrk := c.Xfrks[i]
		res[i] = xfrk.String(wb)
	}

	return res
}

type MulBlankCol struct {
	Col
	Xfs      []uint16
	LastColB uint16
}

// LastCol returns the last column index represented by this MulBlankCol.
// This is used when a row contains a sequence of adjacent blank cells,
// each with its own formatting (XF) index.
func (c *MulBlankCol) LastCol() uint16 {
	return c.LastColB
}

// String returns a slice of empty strings, one for each blank cell in the group.
//
// Even though these cells are visually empty, they may have distinct formatting
// information stored in the XF index (available via c.Xfs).
func (c *MulBlankCol) String(_ *WorkBook) []string {
	return make([]string, len(c.Xfs))
}

type NumberCol struct {
	Col
	Index uint16
	Float float64
}

// String returns the floating-point value of the NumberCol as a string.
//
// This corresponds to the BIFF `NUMBER` record, which stores an IEEE 754 float.
func (c *NumberCol) String(_ *WorkBook) []string {
	return []string{strconv.FormatFloat(c.Float, 'f', -1, 64)}
}

// FormulaStringCol represents a formula whose result is a string literal.
// The result has already been rendered and is stored in RenderedValue.
type FormulaStringCol struct {
	Col
	RenderedValue string
}

// String returns the already-rendered string result of a formula cell.
func (c *FormulaStringCol) String(_ *WorkBook) []string {
	return []string{c.RenderedValue}
}

// FormulaCol represents a cell that contains a formula,
// but whose result is not a string and must be interpreted from raw bytes.
//
// The Result field may contain the precomputed value as 8 bytes,
// but decoding it properly is left for future implementation (TODO).
type FormulaCol struct {
	Header struct {
		Col
		IndexXf uint16  // Format index (XF)
		Result  [8]byte // Raw result of formula evaluation
		Flags   uint16  // Evaluation flags (e.g. result type)
		_       uint32  // Unused/reserved
	}
	Bts []byte // Additional payload or expression data (currently unused)
}

// String returns a placeholder indicating that formula result parsing is not yet implemented.
func (c *FormulaCol) String(wb *WorkBook) []string {
	return []string{"FormulaCol"}
}

// RkCol represents a single cell that stores a value in RK format (compact int/float).
type RkCol struct {
	Col
	Xfrk XfRk
}

// String returns the formatted string representation of the RK value,
// respecting any number/date formatting rules from the workbook.
func (c *RkCol) String(wb *WorkBook) []string {
	return []string{c.Xfrk.String(wb)}
}

// LabelsstCol represents a cell that refers to the shared string table (SST).
type LabelsstCol struct {
	Col
	Xf  uint16 // Format index (XF)
	Sst uint32 // Index into the shared string table (wb.sst)
}

// String returns the resolved string from the SST at the given index.
func (c *LabelsstCol) String(wb *WorkBook) []string {
	return []string{wb.sst[int(c.Sst)]}
}

// labelCol represents a legacy LABEL record containing a plain string,
// stored directly in the structure rather than the SST.
//
// This record type is largely deprecated in later Excel BIFF versions.
type labelCol struct {
	BlankCol
	Str string
}

// String returns the stored string from a LABEL record.
func (c *labelCol) String(wb *WorkBook) []string {
	return []string{c.Str}
}

// BlankCol represents a cell that is visually empty but may still have formatting applied.
// This is common in sparse worksheets or cells with just borders/colors and no content.
type BlankCol struct {
	Col
	Xf uint16 // Format index (XF)
}

// String returns an empty string for a blank cell.
func (c *BlankCol) String(wb *WorkBook) []string {
	return []string{""}
}
