package lark_util

import (
	"encoding/json"
	"net/http"
	"net/url"

	"github.com/pkg/errors"
)

// CreateExcel 创建Excel folderToken是文件夹的唯一标识.用浏览器打开飞书文档,可以通过url看到标识信息.
// 如https://rg975ojk5z.feishu.cn/drive/folder/fldcnKM2eYBLu75CuyuN3Mv6iqg 中的fldcnKM2eYBLu75CuyuN3Mv6iqg就是folderToken
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
		err = errors.Errorf("http error: code= %d | %s", httpCode, string(respBody))
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

type (
	// ExcelInfo 表格元数据
	ExcelInfo struct {
		Properties       Properties `json:"properties"`
		Sheets           []Sheets   `json:"sheets"`
		SpreadsheetToken string     `json:"spreadsheetToken"`
	}
	Properties struct {
		Title       string `json:"title"`       // 表格的标题
		OwnerUserId int    `json:"ownerUserId"` // 所有者的 id，取决于user_id_type的值，仅user_id_type不为空是返回该值
		SheetCount  int    `json:"sheetCount"`  // sheet 数
		Revision    int    `json:"revision"`    // 该sheet的版本
	}
	Merges struct {
		ColumnCount      int `json:"columnCount"`      // 合并单元格范围的列数量
		RowCount         int `json:"rowCount"`         // 合并单元格范围的行数量
		StartColumnIndex int `json:"startColumnIndex"` // 合并单元格范围的开始列下标，index 从 0 开始
		StartRowIndex    int `json:"startRowIndex"`    // 合并单元格范围的开始行下标，index 从 0 开始
	}
	Dimension struct {
		EndIndex       int    `json:"endIndex"`       // 保护行列的结束位置，位置从1开始
		MajorDimension string `json:"majorDimension"` // 若为ROWS，则为保护行；为COLUMNS，则为保护列
		SheetID        string `json:"sheetId"`        // 保护范围所在工作表 ID
		StartIndex     int    `json:"startIndex"`     // 保护行列的起始位置，位置从1开始
	}
	ProtectedRange struct {
		Dimension Dimension `json:"dimension"` // 保护行列的信息，如果为保护工作表，则该字段为空
		ProtectID string    `json:"protectId"` // 保护范围ID
		SheetID   string    `json:"sheetId"`   // 保护工作表ID
		LockInfo  string    `json:"lockInfo"`  // 保护说明
	}
	BlockInfo struct {
		BlockToken string `json:"blockToken"` // block的token
		BlockType  string `json:"blockType"`  // block的类型
	}
	Sheets struct {
		SheetId        string           `json:"sheetId"`                  // sheet id
		Title          string           `json:"title"`                    // sheet 标题
		Index          int              `json:"index"`                    // sheet 位置
		RowCount       int              `json:"rowCount"`                 // sheet 行数
		ColumnCount    int              `json:"columnCount"`              // sheet 列数
		FrozenColCount int              `json:"frozenColCount"`           // sheet 冻结列数
		FrozenRowCount int              `json:"frozenRowCount"`           // sheet 冻结行数
		Merges         []Merges         `json:"merges,omitempty"`         // 该sheet中合并单元格的范围
		ProtectedRange []ProtectedRange `json:"protectedRange,omitempty"` // 该sheet中保护范围
		BlockInfo      BlockInfo        `json:"blockInfo,omitempty"`      // 若含有该字段，则此工作表不为表格
	}
)

type ExcelInfoResp struct {
	Code int        `json:"code"`
	Msg  string     `json:"msg"`
	Data *ExcelInfo `json:"data"`
}

func (l *LarkU) GetExcelInfo(excelToken string, extFields string, userIdType string) (info *ExcelInfo, err error) {
	httpCode, respBody, err := l.LarkGet("/open-apis/sheets/v2/spreadsheets/"+excelToken+"/metainfo", url.Values{
		"extFields":    []string{extFields},
		"user_id_type": []string{userIdType},
	})
	if err != nil {
		err = errors.Errorf("http error: %+v", err)
		return
	}
	if httpCode != http.StatusOK {
		err = errors.Errorf("http error: code= %d | %s", httpCode, string(respBody))
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
		err = errors.Errorf("http error: code= %d | %s", httpCode, string(respBody))
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
		err = errors.Errorf("http error: code= %d | %s", httpCode, string(respBody))
	}
	return
}

type SetCellStyleReq struct {
	AppendStyle AppendStyle `json:"appendStyle,omitempty"`
	ExcelToken  string
}
type Font struct {
	Bold     bool   `json:"bold,omitempty"`
	Italic   bool   `json:"italic,omitempty"`
	FontSize string `json:"fontSize,omitempty"`
	Clean    bool   `json:"clean,omitempty"`
}
type Style struct {
	Font           Font   `json:"font,omitempty"`
	TextDecoration int    `json:"textDecoration,omitempty"`
	Formatter      string `json:"formatter,omitempty"`
	HAlign         int    `json:"hAlign,omitempty"`
	VAlign         int    `json:"vAlign,omitempty"`
	ForeColor      string `json:"foreColor,omitempty"`
	BackColor      string `json:"backColor,omitempty"`
	BorderType     string `json:"borderType,omitempty"`
	BorderColor    string `json:"borderColor,omitempty"`
	Clean          bool   `json:"clean,omitempty"`
}
type AppendStyle struct {
	Range string `json:"range"`
	Style Style  `json:"style"`
}

type SetCellStyleResp struct {
	Code int    `json:"code"`
	Msg  string `json:"msg"`
	Data struct {
		SpreadsheetToken string `json:"spreadsheetToken"`
		UpdatedRange     string `json:"updatedRange"`
		UpdatedRows      int    `json:"updatedRows"`
		UpdatedColumns   int    `json:"updatedColumns"`
		UpdatedCells     int    `json:"updatedCells"`
		Revision         int    `json:"revision"`
	} `json:"data"`
}

func (l *LarkU) SetCellStyle(req *SetCellStyleReq) (err error) {
	httpCode, respBody, err := l.LarkPut("/open-apis/sheets/v2/spreadsheets/"+req.ExcelToken+"/style", map[string]interface{}{
		"appendStyle": req.AppendStyle,
	})
	if err != nil {
		err = errors.Errorf("http error: %+v", err)
		return
	}
	if httpCode != http.StatusOK {
		err = errors.Errorf("http error: code= %d | %s", httpCode, string(respBody))
	}
	res := new(SetCellStyleResp)
	_ = json.Unmarshal(respBody, &res)
	if res.Code != 0 {
		err = errors.Errorf("remote service error: code = %d | %s", res.Code, res.Msg)
	}
	return
}

func (l *LarkU) SetRowOrColumnStyle() (err error) {
	return
}

func (l *LarkU) InsertValueToCell() {

}
