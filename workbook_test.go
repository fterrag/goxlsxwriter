package xlsxwriter

import (
	"os"
)

func MakeTestWorkbook() *Workbook {
	file := "test.xlsx"

	workbook := NewWorkbook(file, nil)
	defer os.Remove(file)

	return workbook
}
