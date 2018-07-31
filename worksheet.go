package goxlsxwriter

import (
	"errors"
	"time"
	"unsafe"
)

/*
#cgo LDFLAGS: -L. -lxlsxwriter
#include <xlsxwriter.h>
*/
import "C"

// Worksheet represents an Excel worksheet.
type Worksheet struct {
	CWorksheet *C.struct_lxw_worksheet
	Workbook   *Workbook
}

// ImageOptions contains options to be set when inserting an image into a
// worksheet.
type ImageOptions struct {
	XOffset int
	YOffset int
	XScale  float64
	YScale  float64
}

// NewWorksheet creates and returns a new Worksheet.
func NewWorksheet(workbook *Workbook, sheetName string) *Worksheet {
	cSheetName := C.CString(sheetName)
	defer C.free(unsafe.Pointer(cSheetName))

	cWorksheet := C.workbook_add_worksheet(workbook.CWorkbook, cSheetName)

	worksheet := &Worksheet{
		CWorksheet: cWorksheet,
		Workbook:   workbook,
	}

	return worksheet
}

// WriteInt writes an integer value at the specified row and column and applies
// an optional format.
func (w *Worksheet) WriteInt(row int, col int, value int, format *Format) error {
	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)
	cValue := (C.double)(value)

	var cFormat *C.struct_lxw_format
	if format != nil {
		cFormat = format.CFormat
	}

	err := C.worksheet_write_number(w.CWorksheet, cRow, cCol, cValue, cFormat)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}

// WriteFloat64 writes a float64 value at the specified row and column and applies
// an optional format.
func (w *Worksheet) WriteFloat64(row int, col int, value float64, format *Format) error {
	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)
	cValue := (C.double)(value)

	var cFormat *C.struct_lxw_format
	if format != nil {
		cFormat = format.CFormat
	}

	err := C.worksheet_write_number(w.CWorksheet, cRow, cCol, cValue, cFormat)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}

// WriteString writes a string value at the specified row and column and applies
// an optional format.
func (w *Worksheet) WriteString(row int, col int, value string, format *Format) error {
	cValue := C.CString(value)
	defer C.free(unsafe.Pointer(cValue))

	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)

	var cFormat *C.struct_lxw_format
	if format != nil {
		cFormat = format.CFormat
	}

	err := C.worksheet_write_string(w.CWorksheet, cRow, cCol, cValue, cFormat)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}

// WriteFormula writes a formula value at the specified row and column and
// applies an optional format.
func (w *Worksheet) WriteFormula(row int, col int, formula string, format *Format) error {
	cValue := C.CString(formula)
	defer C.free(unsafe.Pointer(cValue))

	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)

	var cFormat *C.struct_lxw_format
	if format != nil {
		cFormat = format.CFormat
	}

	err := C.worksheet_write_formula(w.CWorksheet, cRow, cCol, cValue, cFormat)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}

// WriteTime writes a time.Time value at the specified row and column and
// applies an optional format.
func (w *Worksheet) WriteTime(row int, col int, val time.Time, format *Format) error {
	cValue := &C.lxw_datetime{
		year:  C.int(val.Year()),
		month: C.int(val.Month()),
		day:   C.int(val.Day()),
		hour:  C.int(val.Hour()),
		min:   C.int(val.Minute()),
		sec:   C.double(val.Second()),
	}

	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)

	var cFormat *C.struct_lxw_format
	if format != nil {
		cFormat = format.CFormat
	}

	err := C.worksheet_write_datetime(w.CWorksheet, cRow, cCol, cValue, cFormat)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}

// WriteUrl writes a URL and a display string at the specified row and column
// and applies an optional format to the display string. If the specified
// display string is empty, the URL will be used.
func (w *Worksheet) WriteUrl(row int, col int, url string, display string, format *Format) error {
	cUrl := C.CString(url)
	defer C.free(unsafe.Pointer(cUrl))

	cDisplay := C.CString(display)
	defer C.free(unsafe.Pointer(cDisplay))

	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)

	var cFormat *C.struct_lxw_format
	if format != nil {
		cFormat = format.CFormat
	}

	err := C.worksheet_write_url(w.CWorksheet, cRow, cCol, cUrl, cFormat)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	// If the display string is not empty, write it to the same row and
	// column as the URL.
	if len(display) > 0 {
		err := C.worksheet_write_string(w.CWorksheet, cRow, cCol, cDisplay, nil)
		if err != C.LXW_NO_ERROR {
			return errors.New(C.GoString(C.lxw_strerror(err)))
		}
	}

	return nil
}

// WriteBool writes a boolean value at the specified row and column and
// applies an optional format.
func (w *Worksheet) WriteBool(row int, col int, value bool, format *Format) error {
	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)

	var cFormat *C.struct_lxw_format
	if format != nil {
		cFormat = format.CFormat
	}

	// Get the int value from the specified boolean.
	var intValue int
	if value {
		intValue = 1
	}

	cInt := (C.int)(intValue)

	err := C.worksheet_write_boolean(w.CWorksheet, cRow, cCol, cInt, cFormat)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}

// WriteBlank writes a "blank" cell at the specified row and column and
// applies an optional format. Excel differentiates between an empty cell
// and a blank cell. An empty cell is a cell which doesn't contain data or
// formatting. A blank cell doesn't contain data but does contain formatting.
func (w *Worksheet) WriteBlank(row int, col int, format *Format) error {
	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)

	var cFormat *C.struct_lxw_format
	if format != nil {
		cFormat = format.CFormat
	}

	err := C.worksheet_write_blank(w.CWorksheet, cRow, cCol, cFormat)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}

// InsertImage inserts an image at the specified row and column and applies
// options.
func (w *Worksheet) InsertImage(row int, col int, filename string, options *ImageOptions) error {
	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)

	cFilename := C.CString(filename)
	defer C.free(unsafe.Pointer(cFilename))

	var cOptions *C.lxw_image_options
	if options != nil {
		cOptions = &C.lxw_image_options{
			x_offset: (C.int32_t)(options.XOffset),
			y_offset: (C.int32_t)(options.YOffset),
			x_scale:  (C.double)(options.XScale),
			y_scale:  (C.double)(options.YScale),
		}
	}

	err := C.worksheet_insert_image_opt(w.CWorksheet, cRow, cCol, cFilename, cOptions)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}

// InsertChart inserts a chart at the specified row and column and applies
// options.
func (w *Worksheet) InsertChart(row int, col int, chart *Chart, options *ImageOptions) error {
	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)

	if chart != nil {
		return errors.New("chart cannot be nil")
	}

	var cOptions *C.lxw_image_options
	if options != nil {
		cOptions = &C.lxw_image_options{
			x_offset: (C.int32_t)(options.XOffset),
			y_offset: (C.int32_t)(options.YOffset),
			x_scale:  (C.double)(options.XScale),
			y_scale:  (C.double)(options.YScale),
		}
	}

	err := C.worksheet_insert_chart_opt(w.CWorksheet, cRow, cCol, chart.CChart, cOptions)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}
