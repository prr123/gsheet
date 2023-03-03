//
// sheetLib.go
// author PRR
// created 25/2/2022
// major rev 18/02/2023
//
// copyright 2022-2023 prr, azulsoftware
//

package gsheetsLib

import (
	"fmt"
	"context"
	"encoding/json"
// 	"io/ioutil"
//	"net/http"
	"os"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/googleapi"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/sheets/v4"

)

type GSheetsObj  struct {
    Ctx context.Context
    GdSvc *drive.Service
    GshSvc *sheets.Service
	GspSheet *sheets.Spreadsheet
	GspSheetData bool
}

type cred struct {
    Installed credItems `json:"installed"`
    Web credItems `json:"web"`
}

type credItems struct {
    ClientId string `json:"client_id"`
    ProjectId string `json:"project_id"`
    AuthUri string `json:"auth_uri"`
    TokenUri string `json:"token_uri"`
//  Auth_provider_x509_cert_url string `json:"auth_provider_x509_cert_url"`
    ClientSecret string `json:"client_secret"`
    RedirectUris []string `json:"redirect_uris"`
}


// Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
        f, err := os.Open(file)
        if err != nil {
                return nil, err
        }
        defer f.Close()
        tok := &oauth2.Token{}
        err = json.NewDecoder(f).Decode(tok)
        return tok, err
}

// method that initializes the GSheetsObj and creates services for gdrive and sheets
func InitGSheets() (gsh *GSheetsObj, err error){

    var cred cred
    var config oauth2.Config
    var gshObj GSheetsObj

    ctx := context.Background()
    gshObj.Ctx = ctx

    credFilNam := "/home/peter/go/src/google/gdoc/loginCred.json"
    credbuf, err := os.ReadFile(credFilNam)
    if err != nil {return nil, fmt.Errorf("os.Read %s: %v!", credFilNam, err)}

    err = json.Unmarshal(credbuf,&cred)
    if err != nil {return nil, fmt.Errorf("json.UnMarshal credbuf: %v\n", err)}

    if len(cred.Installed.ClientId) > 0 {
        config.ClientID = cred.Installed.ClientId
        config.ClientSecret = cred.Installed.ClientSecret
    }
    if len(cred.Web.ClientId) > 0 {
        config.ClientID = cred.Web.ClientId
        config.ClientSecret = cred.Web.ClientSecret
    }


    config.Scopes = make([]string,2)
    config.Scopes[0] = "https://www.googleapis.com/auth/drive"
    config.Scopes[1] = "https://www.googleapis.com/auth/sheets"

    config.Endpoint = google.Endpoint

    tokFile := "/home/peter/go/src/google/gdoc/tokNew.json"
    tok, err := tokenFromFile(tokFile)
    if err != nil {return nil, fmt.Errorf("tokenFromFile: %v!", err)}

    client := config.Client(context.Background(), tok)

    gdsvc, err := drive.NewService(ctx, option.WithHTTPClient(client))
    if err != nil {return nil, fmt.Errorf("Unable to create Drive Service: %v!", err)}

    gshObj.GdSvc = gdsvc

    gsheetsvc, err := sheets.NewService(ctx, option.WithHTTPClient(client))
    if err != nil {return nil, fmt.Errorf("Unable to create Sheets Service: %v!", err)}

    gshObj.GshSvc = gsheetsvc

    return &gshObj, nil
}

// method that returns a spreadsheet without the grid data
// todo: combine with ReadGrid
func (gs *GSheetsObj) GetSpreadsheet(spSheetId string) (err error){

	svc := gs.GshSvc

    spSheet, err := svc.Spreadsheets.Get(spSheetId).Do()
	if err != nil {return fmt.Errorf("could not open spreadsheet!")}
	gs.GspSheet = spSheet
	gs.GspSheetData = false
	return nil
}


// method that returns a spreadsheet with all grid data
func (gs *GSheetsObj) ReadGrid(spSheetId string) (err error){

	svc := gs.GshSvc

    spSheet, err := svc.Spreadsheets.Get(spSheetId).IncludeGridData(true).Do()
    if err != nil {return fmt.Errorf("could not open spreadsheet!")}

	gs.GspSheet = spSheet
	gs.GspSheetData = true

	return nil
}

// method that reads a gridrange specified by the cellRanges
func (gs *GSheetsObj) ReadGridRange(spSheetId string, cellRange *[]string) (err error){

	svc := gs.GshSvc

	fields := []googleapi.Field{"spreadsheetId","properties.title","sheets.properties","sheets.data"}

    spSheet, err := svc.Spreadsheets.Get(spSheetId).Fields(fields...).Ranges(*cellRange...).Do()
    if err != nil {return fmt.Errorf("could not get spreadsheet: %v!", err)}

    gs.GspSheet = spSheet
    gs.GspSheetData = true

    return nil
}

// methods that reads the content of all cells of the range 'cellRange' in a spreadsheet with id 'spSheetId'.
//
func (gs *GSheetsObj) ReadCells(spSheetId string, cellRange string) (valObj *sheets.ValueRange, err error){

	svc := gs.GshSvc

//	rang:= "Sheet1!A1"
    valObj, err = svc.Spreadsheets.Values.Get(spSheetId, cellRange).Do()
    if err != nil {return nil, fmt.Errorf("could not open spreadsheet!")}

	return valObj, nil
}

// method that creates a  new spreadsheet
// todo: move file into a specified directory
func (gs *GSheetsObj) CreateSpreadsheet(nspSheet *sheets.Spreadsheet) (err error){

	svc := gs.GshSvc
	ctx := gs.Ctx

    spSheet, err := svc.Spreadsheets.Create(nspSheet).Context(ctx).Do()
    if err != nil {return fmt.Errorf("could not create spreadsheet: %v!", err)}

//	id := spSheet.SpreadsheetId
	gs.GspSheet = spSheet

//	gs.GspSheetData = true

	return nil
}

// method that copies a spreadsheet to a new spreadsheet
// todo: implement
func (gs *GSheetsObj) CopySpreadsheet(dirId string) (err error){

//	svc := gs.GshSvc
//	ctx := gs.Ctx

//    spSheet, err := svc.Spreadsheets.Create(nspSheet).Context(ctx).Do()
//    if err != nil {return fmt.Errorf("could not create spreadsheet: %v!", err)}

	return nil
}


func (gs *GSheetsObj) UpdSheet(spshId string, updReq *sheets.BatchUpdateValuesRequest) (err error){

	svc := gs.GshSvc
	ctx := gs.Ctx

	updResp, err := svc.Spreadsheets.Values.BatchUpdate(spshId, updReq).Context(ctx).Do()
    if err != nil {return fmt.Errorf("could not update spreadsheet: %v!", err)}

	for i:=0; i< len(updResp.Responses); i++ {
		resp := updResp.Responses[i]
		fmt.Printf(" resp[i]: %v/n",i, resp)
	}

	return nil
}

func (gs *GSheetsObj) UpdData(spshId string, wr string, valRang *sheets.ValueRange) (cellNum int, err error){

//    var updReq sheets.BatchUpdateValuesRequest
    svc := gs.GshSvc
    ctx := gs.Ctx

    updValResp, err := svc.Spreadsheets.Values.Update(spshId, wr, valRang).Context(ctx).ValueInputOption("RAW").Do()
    if err != nil {return 0, fmt.Errorf("could not update values in spreadsheet: %v!", err)}

	PrintUpdValResp(updValResp)

    return int(updValResp.UpdatedCells), nil
}

func PrintUpdValResp(updValResp *sheets.UpdateValuesResponse) {

	fmt.Printf("Spredsheet Id: %s\n", updValResp.SpreadsheetId)
	fmt.Printf("  updated cells: %d\n", int(updValResp.UpdatedCells))
	fmt.Printf("  updated range: %s\n", updValResp.UpdatedRange)

	valAr := updValResp.UpdatedData
	fmt.Printf("  updated values: %v\n",valAr)

//	for i:= 0; i< len(updValResp.UpdatedData.Values); i++ {
//		val := updValResp.UpdatedData.Values[i]
//	}
}

func (gs *GSheetsObj) WriteData(spshId string, updReq *sheets.BatchUpdateValuesRequest) (cellNum int, err error){

//    var updReq sheets.BatchUpdateValuesRequest
/*
    svc := gs.GshSvc
    ctx := gs.Ctx

    updResp, err := svc.Spreadsheets.Values.BatchUpdate(spshId, &updReq).Context(ctx).ValueInputOption("RAW").Do()
    if err != nil {return 0, fmt.Errorf("could not update values in spreadsheet: %v!", err)}

	PrintUpdResp(updResp)
*/
    return 0, nil

}

func PrintUpdResp(updResp *sheets.BatchUpdateValuesResponse) {

	fmt.Printf("*** update responses [%d] ***\n", int(updResp.TotalUpdatedCells))
	fmt.Printf("Spredsheet Id: %s\n", updResp.SpreadsheetId)
    for i:=0; i< len(updResp.Responses); i++ {
        resp := updResp.Responses[i]
		fmt.Printf("  response[%d]: updated cells: %d\n", i, int(resp.UpdatedCells))
    }
}

// method that fills the cells specified by ValueRange to a spredsheet
func (gs *GSheetsObj) WriteCells(cellRange *sheets.ValueRange) (err error){

//	svc := gs.GshSvc

	return nil
}



// function that prints the contents of the cells specified in the value range
// todo: make it into a method
func PrintValueRange(valObj *sheets.ValueRange) {

	fmt.Printf("Range: %s\n", valObj.Range)
	fmt.Printf("values: %v len outer: %d\n", valObj.Values, len(valObj.Values))

	for i:=0; i<len(valObj.Values); i++ {
		cellVal := valObj.Values[i]
		fmt.Printf("val [%d]: %d\n", i, len(cellVal))
		for j:=0; j< len(cellVal); j++ {
			fmt.Printf("value[%d][%d]: %s\n", i, j, cellVal[j])
		}
	}

}

// function that prints the content of a spreadsheet
func PrintSheetValues(spSheet *sheets.Spreadsheet) {

	prop:= spSheet.Properties
	fmt.Println("\n*** PrintSheetValues ***")
	fmt.Printf("Title:  %s\n", prop.Title)


	for ish:=0; ish < len(spSheet.Sheets); ish++ {
		sheet := spSheet.Sheets[ish]
		prop := sheet.Properties
		if prop.GridProperties == nil {
			fmt.Printf("sheet[%d]: no grid properties!", ish)
			continue
		}
		fmt.Printf("sheet[%d]: rows: %d cols: %d \n", ish, prop.GridProperties.RowCount, prop.GridProperties.ColumnCount)

		fmt.Printf("data items: %d\n", len(sheet.Data))
		for i:=0; i< len(sheet.Data); i++ {
			rows := sheet.Data[i]
			fmt.Printf("  row[%d]: row %d col: %d num: %d\n", i, rows.StartRow, rows.StartColumn, len(rows.RowData))

			for j:=0; j< len(rows.RowData); j++ {
				rowDat := rows.RowData[j]
				fmt.Printf("cellrow[%d-%d]: %d\n", i, j, len(rowDat.Values))

				for k:=0; k< len(rowDat.Values); k++ {
					cell := rowDat.Values[k]
					cellVal := cell.EffectiveValue
					fmt.Printf("cell [%d]: %s: ", k, cell.FormattedValue)
					if cellVal.NumberValue != nil {
						fmt.Printf("num: %f", k, *(cellVal.NumberValue))
					}
					if cellVal.StringValue != nil {
						fmt.Printf("str: %s", *(cellVal.StringValue))
					}
					if cellVal.BoolValue != nil {
						fmt.Printf("bool: %t", *(cellVal.BoolValue))
					}
					fmt.Println()
				}
			}
		}
	}
}

// function that prints the property information of a spreadsheet
func PrintSheetInfo(spSheet *sheets.Spreadsheet) {

	prop:= spSheet.Properties
	fmt.Printf("Title:  %s\n", prop.Title)
	fmt.Printf("Id:     %s\n", spSheet.SpreadsheetId)
	fmt.Printf("sheets: %d\n", len(spSheet.Sheets))

	fmt.Printf("\nSpreadsheet Properties\n")
	fmt.Printf("  theme font:  %s\n",prop.SpreadsheetTheme.PrimaryFontFamily)
	fmt.Printf("  theme colors: %d\n", len(prop.SpreadsheetTheme.ThemeColors))
	for i:=0; i< len(prop.SpreadsheetTheme.ThemeColors); i++ {
		fmt.Printf("theme colors [%d]: \n", i)
		colPair:= prop.SpreadsheetTheme.ThemeColors[i]
		rgb:= colPair.Color.RgbColor
		fmt.Printf("    type: %s\n",colPair.ColorType)
		fmt.Printf("    style: %s\n", colPair.Color.ThemeColor)
		fmt.Printf("    color: alpha %.1f red %.1f green %.1f blue %.1f\n",rgb.Alpha, rgb.Red, rgb.Green, rgb.Blue)
	}


	for i:=0; i< len(spSheet.Sheets); i++ {
		sheet := spSheet.Sheets[i]
		shProp := sheet.Properties
		fmt.Printf("\n*** Sheet[%d]:\n", i+1)
		fmt.Printf("  Name:  %s\n", shProp.Title)
		fmt.Printf("  Id:    %d\n", shProp.SheetId)
		fmt.Printf("  Type:  %s\n", shProp.SheetType)
		fmt.Printf("  Index: %d\n", shProp.Index)
	}

	fmt.Println()
}

