xlsxwriter
==========

[![Build Status](https://travis-ci.org/fterrag/xlsxwriter.svg?branch=master)](https://travis-ci.org/fterrag/xlsxwriter) [![Go Report Card](https://goreportcard.com/badge/github.com/fterrag/xlsxwriter)](https://goreportcard.com/report/github.com/fterrag/xlsxwriter)

xlsxwriter provides Go bindings for the [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter) C library.

## Install

xlsxwriter requires the libxslxwriter library to be installed. To build from source via Git:

```
$ git clone https://github.com/jmcnamara/libxlsxwriter.git
$ cd libxlsxwriter
$ make
$ make install
```

Visit [http://libxlsxwriter.github.io/getting_started.html](http://libxlsxwriter.github.io/getting_started.html) for more information on installing libxlsxwriter.

## Todo

- [ ] Increase test coverage
- [ ] Documentation
- [ ] Expand libxlsxwriter API coverage

## Sample Use

![](https://cloud.githubusercontent.com/assets/22901700/23842694/75b0b3c2-078c-11e7-8ef6-5ae9489971b6.png)

```go
package main

import (
    "github.com/fterrag/xlsxwriter"
)

func main() {
    workbook := xlsxwriter.NewWorkbook("example.xlsx", nil)
    worksheet := xlsxwriter.NewWorksheet(workbook, "Sheet 1")

    format := xlsxwriter.NewFormat(workbook)

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

## Contributing

* Submit a PR (tests and documentation included)
* Add or improve documentation
* Report issues
* Suggest new features or enhancements
