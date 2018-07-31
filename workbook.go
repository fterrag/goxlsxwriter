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

// Workbook represents an Excel workbook.
type Workbook struct {
	CWorkbook *C.struct_lxw_workbook
}

// WorkbookOptions contains options to be set when creating a new Workbook.
type WorkbookOptions struct {
	ConstantMemory int
	TmpDir         string
}

// NewWorkbook create and returns a new Workbook.
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

// Close closes the workbook and writes the XLSX file to disk.
func (w *Workbook) Close() error {
	err := C.workbook_close(w.CWorkbook)
	if err != C.LXW_NO_ERROR {
		return errors.New(C.GoString(C.lxw_strerror(err)))
	}

	return nil
}
