// golang program to test gdrive
// author: prr, azul software
// created: 6/2/2023
// copyright 2023 prr, Peter Riemenschneider, Azul Software
//
//

package main

import (
        "fmt"
        "os"
        gsh "google/gsheets/examples/gsheetsLib"
		"google.golang.org/api/sheets/v4"
)


func main() {

	var newSpsheet sheets.Spreadsheet
	var prop sheets.SpreadsheetProperties


    numArgs := len(os.Args)
    if numArgs < 2 {
        fmt.Println("error - no comand line arguments!")
        fmt.Println("gsheetCreate usage is: \"gsheetCreate <title>\"\n")
        os.Exit(1)
    }

//	spSheetTitle := os.Args[1]

	fmt.Printf("Creating new spreadsheet with title: %s\n", os.Args[1])

	gsheet,err := gsh.InitGSheets()
	if err != nil {
		fmt.Printf("error InitGsheets: %v\n", err)
		os.Exit(-1)
	}

	gsheets := make([]*sheets.Sheet, 3)
	for i:=0; i< 3; i++ {
		gsheets[i] = new(sheets.Sheet)
	}
	prop.Title = os.Args[1]

	newSpsheet.Properties = &prop
	newSpsheet.Sheets = gsheets

	err = gsheet.CreateSpreadsheet(&newSpsheet)
	if err != nil {
		fmt.Printf("error CreateSpreadSheet: %v\n", err)
		os.Exit(-1)
	}
//	gsh.PrintValueRange(valObj)
//	fmt.Printf("gsheet: %v\n", gsheet)

//	gsh.PrintSheetValues(gsheet.GspSheet)

	gsh.PrintSheetInfo(gsheet.GspSheet)

    fmt.Println("Success!")
    os.Exit(0)
}
