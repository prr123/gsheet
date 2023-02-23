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
//      "io"
        gsh "google/gsheets/examples/gsheetsLib"
)


func main() {


    numArgs := len(os.Args)
    if numArgs < 2 {
        fmt.Println("error - no comand line arguments!")
        fmt.Println("gsheetTest usage is: \"gsheetTest sheetId\"\n")
        os.Exit(1)
    }

	sheetId := os.Args[1]

//	fmt.Printf("sheet Id: %s\n", sheetId)

	gsheet,err := gsh.InitGSheets()
	if err != nil {
		fmt.Printf("error InitGsheets: %v\n", err)
		os.Exit(-1)
	}

	err = gsheet.ReadGrid(sheetId)
	if err != nil {
		fmt.Printf("error ReadGrid: %v\n", err)
		os.Exit(-1)
	}
/*
	rang:= "Sheet1!A1"
	valObj, err := gsheet.ReadCells(sheetId, rang)
	if err != nil {
		fmt.Printf("error ReadCells: %v\n", err)
		os.Exit(-1)
	}
*/
//	gsh.PrintValueRange(valObj)
//	fmt.Printf("gsheet: %v\n", gsheet)

	gsh.PrintSheetValues(gsheet.GspSheet)

//	gsh.PrintSheetInfo(gsheet.GspSheet)

    fmt.Println("Success!")
    os.Exit(0)
}
