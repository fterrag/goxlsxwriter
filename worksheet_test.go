package goxlsxwriter

import (
	// "fmt"
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

	worksheet.WriteFloat(2, 4, 3.14159265358, nil)
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
	worksheet.WriteFloat(1, 0, 2.5, nil)
	worksheet.WriteInt(2, 0, 3, nil)

	worksheet.WriteFormula(3, 0, "=SUM(A1:A3)", nil)

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
