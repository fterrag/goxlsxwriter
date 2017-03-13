package xlsxwriter

/*
#cgo LDFLAGS: -L. -lxlsxwriter
#include "include/xlsxwriter.h"
*/
import "C"
import "unsafe"

type Workbook struct {
	CWorkbook *C.struct_lxw_workbook
}

func NewWorkbook(filename string) *Workbook {
	cFilename := C.CString(filename)
	defer C.free(unsafe.Pointer(cFilename))

	cWorkbook := C.new_workbook(cFilename)

	workbook := &Workbook{
		CWorkbook: cWorkbook,
	}

	return workbook
}

func (w *Workbook) AddFormat() *Format {
	cFormat := C.workbook_add_format(w.CWorkbook)

	format := &Format{
		CFormat: cFormat,
	}

	return format
}

func (w *Workbook) Close() {
	C.workbook_close(w.CWorkbook)
}
