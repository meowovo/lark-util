package lark_util

import (
	"encoding/json"
	"net/http"
	"net/url"

	"github.com/pkg/errors"
)

func (l *LarkU) CreateExcel(folderToken, title string) (spreadsheetToken string, err error) {
	httpCode, respBody, err := l.LarkPost("/open-apis/sheets/v3/spreadsheets", map[string]interface{}{
		"folder_token": folderToken,
		"title":        title,
	})
	if err != nil {
		err = errors.Errorf("http error: %+v", err)
		return
	}
	if httpCode != http.StatusOK {
		err = errors.Errorf("http error: code= %d | %+v", httpCode, respBody)
		return
	}
	type CreateSpreadsheetResp struct {
		Code int32  `json:"code,omitempty"`
		Msg  string `json:"msg,omitempty"`
		Data struct {
			SpreadSheet struct {
				SpreadsheetToken string `json:"spreadsheet_token,omitempty"`
			}
		}
	}
	m := new(CreateSpreadsheetResp)
	_ = json.Unmarshal(respBody, &m)
	if m.Code != 0 {
		err = errors.Errorf("remote service error: code = %d | %s", m.Code, m.Msg)
		return
	}
	spreadsheetToken = m.Data.SpreadSheet.SpreadsheetToken
	return
}

type ExcelInfo struct {
	Properties *struct {
		SheetCount int `json:"sheetCount"`
	} `json:"properties"`
	Sheets []*struct {
		SheetId        string `json:"sheetId"`
		Title          string `json:"title"`
		Index          int    `json:"index"`
		RowCount       int    `json:"rowCount"`
		ColumnCount    int    `json:"columnCount"`
		FrozenColCount int    `json:"frozenColCount"`
		FrozenRowCount int    `json:"frozenRowCount"`
		// 合并单元格的信息
		Merges []*struct {
			ColumnCount      int `json:"columnCount"`
			RowCount         int `json:"rowCount"`
			StartColumnIndex int `json:"startColumnIndex"`
			StartRowIndex    int `json:"startRowIndex"`
		} `json:"merges,omitempty"`
	} `json:"sheets"`
}

type ExcelInfoResp struct {
	Code int        `json:"code"`
	Msg  string     `json:"msg"`
	Data *ExcelInfo `json:"data"`
}

func (l *LarkU) GetExcelInfo(excelToken string) (info *ExcelInfo, err error) {
	httpCode, respBody, err := l.LarkGet("/open-apis/sheets/v2/spreadsheets/"+excelToken+"/metainfo", url.Values{})
	if err != nil {
		err = errors.Errorf("http error: %+v", err)
		return
	}
	if httpCode != http.StatusOK {
		err = errors.Errorf("http error: code= %d | %+v", httpCode, respBody)
		return
	}
	m := new(ExcelInfoResp)
	_ = json.Unmarshal(respBody, &m)
	if m.Code != 0 {
		err = errors.Errorf("remote service error: code = %d | %s", m.Code, m.Msg)
		return
	}
	info = m.Data
	return
}

func (l *LarkU) UpdateSheetTitle(excelToken, sheetId, title string) (err error) {
	type UpdateSheetReq struct {
		Requests []struct {
			UpdateSheet struct {
				Properties struct {
					SheetId string `json:"sheetId"`
					Title   string `json:"title"`
				} `json:"properties"`
			} `json:"updateSheet"`
		} `json:"requests"`
	}
	httpCode, respBody, err := l.LarkPost("/open-apis/sheets/v2/spreadsheets/"+excelToken+"/sheets_batch_update", map[string]interface{}{
		"requests": []interface{}{
			map[string]interface{}{
				"updateSheet": map[string]interface{}{
					"properties": map[string]interface{}{
						"sheetId": sheetId,
						"title":   title,
					},
				},
			},
		},
	})
	if err != nil {
		err = errors.Errorf("http error: %+v", err)
		return
	}
	if httpCode != http.StatusOK {
		err = errors.Errorf("http error: code= %d | %+v", httpCode, respBody)
	}
	return
}

const (
	MergeCellTypeAll     = "MERGE_ALL"
	MergeCellTypeRows    = "MERGE_ROWS"
	MergeCellTypeCOLUMNS = "MERGE_COLUMNS"
)

func (l *LarkU) MergeCells(excelToken, sheetId, cellRange, mergeType string) (err error) {
	if mergeType == "" {
		mergeType = MergeCellTypeAll
	}
	httpCode, respBody, err := l.LarkPost("/open-apis/sheets/v2/spreadsheets/"+excelToken+"/merge_cells", map[string]interface{}{
		"range":     sheetId + "!" + cellRange,
		"mergeType": mergeType,
	})
	if err != nil {
		err = errors.Errorf("http error: %+v", err)
		return
	}
	if httpCode != http.StatusOK {
		err = errors.Errorf("http error: code= %d | %+v", httpCode, respBody)
	}
	return
}

func insertValueToCell() {

}
