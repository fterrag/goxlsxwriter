package goxlsxwriter

import (
	"errors"
)

/*
#cgo LDFLAGS: -L. -lxlsxwriter
#include <xlsxwriter.h>
*/
import "C"
import "unsafe"

// Worksheet represents an Excel worksheet.
type Worksheet struct {
	CWorksheet *C.struct_lxw_worksheet
	Workbook   *Workbook
}

// ImageOptions contains options to be set when inserting an image into a worksheet.
type ImageOptions struct {
	XOffset int
	YOffset int
	XScale  float64
	YScale  float64
}

// NewWorksheet creates and returns a new instance of Worksheet.
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

// WriteString writes a string value at the specified row and column and applies an optional format.
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

// WriteFloat writes a float value at the specified row and column and applies an optional format.
func (w *Worksheet) WriteFloat(row int, col int, value float64, format *Format) error {
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

// InsertImage inserts an image at the specified row and column and applies options.
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
