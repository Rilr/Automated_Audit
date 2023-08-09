package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
	//"path/filepath"
)

// func comparecheck()

// func datecheck()

// func readExcel(xlfile string) {
//	dataMap := make(map[string]string)

func main() {
	excelfilepath := "B:\\auto.xlsx"
	//normalpath := filepath.Clean(excelfilepath)
	dataframe, err := excelize.OpenFile(excelfilepath)
	if err != nil {
		fmt.Println("Couldn't open the file", err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := dataframe.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	rows, err := dataframe.GetRows("Sheet")
	if err != nil {
		fmt.Println("Couldn't retrieve rows", err)
		return
	}

	var colnameindex = map[string]int{}
	headerrow := rows[0]
	for colindex, colvalue := range headerrow {
		colnameindex[colvalue] = colindex
	}

	config_name := "Name"
	config_date := "Last Contact"

	for _, row := range rows {
		nameindex, found := colnameindex[config_name]
		dateindex, found := colnameindex[config_date]

		if !found {
			fmt.Printf(config_name, " header not found")
			fmt.Println(config_date, " header not found")
		}

		if len(row) > nameindex && len(row) > dateindex {
			fmt.Printf("%s\t%s\n", row[nameindex], row[dateindex])
		}
	}
}
