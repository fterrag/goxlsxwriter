package xlsxwriter

import (
	"testing"
)

func TestWriteString(t *testing.T) {
	workbook := MakeTestWorkbook()
	worksheet := NewWorksheet(workbook, "Sheet 1")

	row := 2
	col := 4
	value := "Hello World!"

	worksheet.WriteString(row, col, value, nil)
}
