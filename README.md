xlsxwriter
==========

xlsxwriter provides Go bindings for the [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter) C library.

## Sample Use

```
package main

import (
    "github.com/fterrag/xlsxwriter"
)

func main() {
    workbook := xlsxwriter.NewWorkbook("example.xlsx")
    worksheet := xlsxwriter.NewWorksheet(workbook, "Sheet 1")

    format := workbook.AddFormat()

    format.SetFontName("Verdana")
    format.SetFontSize(8)
    format.SetFontColor(0x008000)

    worksheet.WriteString(0, 0, "Hello from A1!", format)
    worksheet.WriteString(4, 1, "This cell is B5", nil)

    workbook.Close()
}
```
