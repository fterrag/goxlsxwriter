package xlsxwriter

/*
#cgo LDFLAGS: -L. -lxlsxwriter
#include "include/xlsxwriter.h"
*/
import "C"
import "unsafe"

type Worksheet struct {
	CWorksheet *C.struct_lxw_worksheet
	Workbook   *Workbook
}

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

func (w *Worksheet) WriteString(row int, col int, value string, format *Format) {
	cValue := C.CString(value)
	defer C.free(unsafe.Pointer(cValue))

	cRow := (C.lxw_row_t)(row)
	cCol := (C.lxw_col_t)(col)

	var cFormat *C.struct_lxw_format
	if format != nil {
		cFormat = format.CFormat
	}

	C.worksheet_write_string(w.CWorksheet, cRow, cCol, cValue, cFormat)
}
