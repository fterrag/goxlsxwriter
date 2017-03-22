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

func TestWriteFloat(t *testing.T) {
	expectedPath := "resources/xlsx/WriteFloat.xlsx"

	workbook, generatedPath := MakeTestWorkbook()
	defer os.Remove(generatedPath)

	worksheet := NewWorksheet(workbook, "Sheet 1")

	row := 2
	col := 4
	value := 3.14159265358

	worksheet.WriteFloat(row, col, value, nil)

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
