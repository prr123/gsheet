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

//	var newSpsheet sheets.Spreadsheet
//	var prop sheets.SpreadsheetProperties


    numArgs := len(os.Args)
    if numArgs < 2 {
        fmt.Println("error - no comand line arguments!")
        fmt.Println("gsheetUpdCell usage is: \"gsheetUpdCell <id>\"\n")
        os.Exit(1)
    }

	spshId := os.Args[1]

	fmt.Printf("Creating new spreadsheet with title: %s\n", os.Args[1])

	gsheet,err := gsh.InitGSheets()
	if err != nil {
		fmt.Printf("error InitGsheets: %v\n", err)
		os.Exit(-1)
	}

	var vr sheets.ValueRange

	myval := []interface{}{"One"}
	vr.Values = append(vr.Values, myval)

	wr := "Sheet1!A2"

	cellNum, err := gsheet.UpdData(spshId, wr, &vr)
	if err != nil {
		fmt.Printf("error UpdData: %v\n", err)
		os.Exit(-1)
	}

	if cellNum != 1 {fmt.Printf(" success %d cells updated\n", cellNum)} else {fmt.Printf(" success one cell updated\n")}

    fmt.Println("Success!")
    os.Exit(0)
}
