package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"

	"path/filepath"
)

// func comparecheck()

// func datecheck()

// func readExcel(xlfile string) {
//	dataMap := make(map[string]string)

func main() {
	excelfilepath := "B:/auto.xlsx"
	normalpath := filepath.Clean(excelfilepath)
	dataframe, err := excelize.OpenFile(normalpath)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := dataframe.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	cell, err := dataframe.GetCellValue("Sheet", "A1")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(cell)

	rows, err := dataframe.GetRows("Sheet")
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}
}
