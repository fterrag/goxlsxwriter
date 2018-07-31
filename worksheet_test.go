package goxlsxwriter

import (
	"os"
	"testing"
)

func TestWriteString(t *testing.T) {
	expectedPath := "resources/xlsx/WriteString.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	row := 2
	col := 4
	value := "Hello World!"

	worksheet.WriteString(row, col, value, nil)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

func TestWriteFloatInt(t *testing.T) {
	expectedPath := "resources/xlsx/WriteFloatInt.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	worksheet.WriteFloat64(2, 4, 3.14159265358, nil)
	worksheet.WriteInt(6, 7, 200, nil)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

func TestWriteFormula(t *testing.T) {
	expectedPath := "resources/xlsx/WriteFormula.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	worksheet.WriteInt(0, 0, 1, nil)
	worksheet.WriteFloat64(1, 0, 2.5, nil)
	worksheet.WriteInt(2, 0, 3, nil)

	worksheet.WriteFormula(3, 0, "=SUM(A1:A3)", nil)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

func TestWriteUrl(t *testing.T) {
	expectedPath := "resources/xlsx/WriteUrl.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	row := 1
	col := 4
	url := "https://www.google.com"
	display := "Google"
	worksheet.WriteUrl(row, col, url, display, nil)

	row = 4
	col = 1
	url = "https://www.github.com"
	display = ""
	format := NewFormat(workbook)
	format.SetFontName("Verdana")
	worksheet.WriteUrl(row, col, url, display, format)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

func TestWriteBool(t *testing.T) {
	expectedPath := "resources/xlsx/WriteBool.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	row := 1
	col := 1
	value := true
	worksheet.WriteBool(row, col, value, nil)

	row = 1
	col = 2
	value = false
	worksheet.WriteBool(row, col, value, nil)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

func TestWriteBlank(t *testing.T) {
	expectedPath := "resources/xlsx/WriteBlank.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	row := 1
	col := 1
	value := "Hello"
	worksheet.WriteString(row, col, value, nil)

	row = 1
	col = 2
	format := NewFormat(workbook)
	format.SetBackgroundColor(0xFBD787)
	worksheet.WriteBlank(row, col, format)

	row = 1
	col = 3
	value = "World!"
	worksheet.WriteString(row, col, value, nil)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}

func TestInsertImage(t *testing.T) {
	expectedPath := "resources/xlsx/InsertImage.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	worksheet.WriteString(0, 0, "Hello World!", nil)

	options := &ImageOptions{
		XScale: 2.5,
		YScale: 2.5,
	}
	worksheet.InsertImage(2, 2, "resources/gopher.png", options)

	workbook.Close()

	CompareXlsxFiles(t, expectedPath, generatedPath)
}
