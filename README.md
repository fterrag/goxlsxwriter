xlsxwriter
==========

xlsxwriter provides Go bindings for the [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter) C library.

## Sample Use

![](https://cloud.githubusercontent.com/assets/22901700/23842694/75b0b3c2-078c-11e7-8ef6-5ae9489971b6.png)

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

    options := &xlsxwriter.ImageOptions{
        XScale: 0.5,
        YScale: 0.5,
    }
    worksheet.InsertImage(1, 3, "resources/gopher.png", options)

    workbook.Close()
}
```
