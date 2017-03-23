package main

import (
	"github.com/fterrag/goxlsxwriter"
)

func main() {
	workbook := goxlsxwriter.NewWorkbook("formula-sum.xlsx", nil)
	worksheet := goxlsxwriter.NewWorksheet(workbook, "Sheet 1")

	type expense struct {
		item string
		cost int
	}

	var data []*expense
	data = append(data, &expense{item: "Rent", cost: 1000})
	data = append(data, &expense{item: "Gas", cost: 100})
	data = append(data, &expense{item: "Food", cost: 300})
	data = append(data, &expense{item: "Gym", cost: 50})

	for i := 0; i < len(data); i++ {
		worksheet.WriteString(i, 0, data[i].item, nil)
		worksheet.WriteInt(i, 1, data[i].cost, nil)
	}

	worksheet.WriteString(len(data), 0, "Total", nil)

	// The sum of B1:B4 may display as 0 in non-Excel spreadsheet applications (see http://libxlsxwriter.github.io/bugs.html).
	worksheet.WriteFormula(len(data), 1, "=SUM(B1:B4)", nil)

	workbook.Close()
}
