package xlsxwriter

import (
	"os"
	"testing"
)

func TestSetFontColor(t *testing.T) {
	expectedPath := "resources/xlsx/SetFontColor.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	format := NewFormat(workbook)
	format.SetFontColor(0x009900)

	worksheet.WriteString(0, 0, "Hello", format)
	worksheet.WriteString(1, 0, "World!", nil)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

func TestSetBackgroundColor(t *testing.T) {
	expectedPath := "resources/xlsx/SetBackgroundColor.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	format := NewFormat(workbook)
	format.SetBackgroundColor(0xFBD787)

	worksheet.WriteString(0, 0, "Hello", format)
	worksheet.WriteString(1, 0, "World!", nil)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}
