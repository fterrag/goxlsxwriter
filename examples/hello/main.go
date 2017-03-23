package main

import (
	"github.com/fterrag/goxlsxwriter"
)

func main() {
	workbook := goxlsxwriter.NewWorkbook("hello.xlsx", nil)
	worksheet := goxlsxwriter.NewWorksheet(workbook, "Sheet 1")

	worksheet.WriteString(0, 0, "Hello", nil)
	worksheet.WriteInt(1, 0, 123, nil)

	workbook.Close()
}
