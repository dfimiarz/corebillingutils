package main

import (
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

func main(){
	fmt.Println("Core Billing Item Extractor")



	var filename, sheetname string = "test.xlsx", "report"

	//Add color to the console
	fmt.Printf("Opening file \033[32m%s\033[0m and sheet \033[32m%s\033[0m\n", filename, sheetname)

	f, err := os.Open(filename)

	if err != nil {
		fmt.Printf("\033[31m%s\033[0m\n", err)
		return
	}

	//Close the file when the function ends
	defer func(){
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	file, err := excelize.OpenReader(f)

	if err != nil {
		fmt.Printf("\033[31m%s\033[0m\n", err)
		return
	}

	//Close the stream when the function ends
	defer func ()  {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	
	rows, err := file.Rows(sheetname)

	if err != nil {
		fmt.Printf("\033[31m%s\033[0m\n", err)
	}
	//Close the rows when the function ends
	defer rows.Close()

	
	for rows.Next() {
		cells, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
		}

		for _, cell := range cells {
			fmt.Print(cell, "\t")
		}
		fmt.Println()
	}

}
