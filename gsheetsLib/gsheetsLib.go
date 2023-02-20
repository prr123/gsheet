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
//	"google.golang.org/api/googleapi"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/sheets/v4"

)

type GSheetsObj  struct {
    Ctx context.Context
    GdSvc *drive.Service
    GshSvc *sheets.Service
	GspSheet *sheets.Spreadsheet
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
// Retrieve a token, saves the token, then returns the generated client.

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

//    gdocsvc, err := docs.NewService(ctx, option.WithHTTPClient(client))
//    if err != nil {return nil, fmt.Errorf("Unable to create Doc Service: %v!", err)}

//    gdobj.GdocSvc = gdocsvc

    gsheetsvc, err := sheets.NewService(ctx, option.WithHTTPClient(client))
    if err != nil {return nil, fmt.Errorf("Unable to create Sheets Service: %v!", err)}

    gshObj.GshSvc = gsheetsvc

    return &gshObj, nil
}


func (gs *GSheetsObj) GetSpreadsheet(spSheetId string) (err error){

	svc := gs.GshSvc

    spSheet, err := svc.Spreadsheets.Get(spSheetId).Do()
	if err != nil {return fmt.Errorf("could not open spreadsheet!")}
	gs.GspSheet = spSheet
	return nil
}

func (gs *GSheetsObj) ReadSpreadSheet() (err error){

//	spSheet := gs.GspSheet
	return nil
}

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

