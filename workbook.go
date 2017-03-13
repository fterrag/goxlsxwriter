package xlsxwriter

import (
	"errors"
)

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

func (w *Workbook) Close() error {
	err := C.workbook_close(w.CWorkbook)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}
