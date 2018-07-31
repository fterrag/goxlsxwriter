package goxlsxwriter

import (
	"os"
	"testing"
)

func TestSetFontName(t *testing.T) {
	expectedPath := "resources/xlsx/SetFontName.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	format := NewFormat(workbook)
	format.SetFontName("Verdana")

	worksheet.WriteString(0, 0, "Hello", format)
	worksheet.WriteString(1, 0, "World!", nil)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

func TestSetFontSize(t *testing.T) {
	expectedPath := "resources/xlsx/SetFontSize.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	format := NewFormat(workbook)
	format.SetFontSize(14)

	worksheet.WriteString(0, 0, "Hello", format)
	worksheet.WriteString(1, 0, "World!", nil)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

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

func TestSetBoldItalicUnderline(t *testing.T) {
	expectedPath := "resources/xlsx/SetBoldItalicUnderline.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	formatBold := NewFormat(workbook)
	formatBold.SetBold()

	formatItalic := NewFormat(workbook)
	formatItalic.SetItalic()

	formatUnderline := NewFormat(workbook)
	formatUnderline.SetUnderline(UnderlineSingle)

	worksheet.WriteString(0, 0, "Bold", formatBold)
	worksheet.WriteString(1, 1, "Italic", formatItalic)
	worksheet.WriteString(2, 2, "Underline", formatUnderline)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

func TestSetPattern(t *testing.T) {
	expectedPath := "resources/xlsx/SetPattern.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	format := NewFormat(workbook)
	format.SetPattern(PatternSolid)

	format2 := NewFormat(workbook)
	format2.SetPattern(PatternGray125)

	worksheet.WriteString(0, 0, "PATTERN_SOLID", format)
	worksheet.WriteString(1, 0, "PATTERN_GRAY_125", format2)

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

func TestSetNumericalFormat(t *testing.T) {
	expectedPath := "resources/xlsx/SetNumericalFormat.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	format := NewFormat(workbook)
	format.SetNumericalFormat("0 \"dollars and\" .00 \"cents\"")

	worksheet.WriteFloat64(0, 0, 5.50, format)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}
