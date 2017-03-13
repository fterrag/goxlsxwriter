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

type WorkbookOptions struct {
	ConstantMemory int
	TmpDir         string
}

func NewWorkbook(filename string, options *WorkbookOptions) *Workbook {
	cFilename := C.CString(filename)
	defer C.free(unsafe.Pointer(cFilename))

	var cOptions *C.lxw_workbook_options
	if options != nil {
		cTmpDir := C.CString(options.TmpDir)
		defer C.free(unsafe.Pointer(cTmpDir))

		cOptions = &C.lxw_workbook_options{
			constant_memory: (C.uint8_t)(options.ConstantMemory),
			tmpdir:          cTmpDir,
		}
	}

	cWorkbook := C.new_workbook_opt(cFilename, cOptions)

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
