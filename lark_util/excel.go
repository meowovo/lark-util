package lark_util

import (
	"encoding/json"
	"net/http"
	"net/url"

	"github.com/pkg/errors"
)

/** -------------------------------------------------表格-------------------------------------------------------------------- **/

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
		Properties       *ExcelProperties `json:"properties"`
		Sheets           []*Sheets        `json:"sheets"`
		SpreadsheetToken string           `json:"spreadsheetToken"`
	}
	ExcelProperties struct {
		Title       string `json:"title"`       // 表格的标题
		OwnerUserId int    `json:"ownerUserId"` // 所有者的 id,取决于user_id_type的值,仅user_id_type不为空是返回该值
		SheetCount  int    `json:"sheetCount"`  // sheet 数
		Revision    int    `json:"revision"`    // 该sheet的版本
	}
	Merges struct {
		ColumnCount      int `json:"columnCount"`      // 合并单元格范围的列数量
		RowCount         int `json:"rowCount"`         // 合并单元格范围的行数量
		StartColumnIndex int `json:"startColumnIndex"` // 合并单元格范围的开始列下标,index 从 0 开始
		StartRowIndex    int `json:"startRowIndex"`    // 合并单元格范围的开始行下标,index 从 0 开始
	}
	Dimension struct {
		EndIndex       int    `json:"endIndex"`       // 保护行列的结束位置,位置从1开始
		MajorDimension string `json:"majorDimension"` // 若为ROWS,则为保护行;为COLUMNS,则为保护列
		SheetID        string `json:"sheetId"`        // 保护范围所在工作表 ID
		StartIndex     int    `json:"startIndex"`     // 保护行列的起始位置,位置从1开始
	}
	ProtectedRange struct {
		Dimension Dimension `json:"dimension"` // 保护行列的信息,如果为保护工作表,则该字段为空
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
		BlockInfo      BlockInfo        `json:"blockInfo,omitempty"`      // 若含有该字段,则此工作表不为表格
	}
)

type ExcelInfoResp struct {
	Code int        `json:"code"`
	Msg  string     `json:"msg"`
	Data *ExcelInfo `json:"data"`
}

// GetExcelInfo 获取表格的元数据
func (l *LarkU) GetExcelInfo(excelToken string, extFields string, userIdType string) (info *ExcelInfo, err error) {
	var values = url.Values{}
	if extFields != "" {
		values.Set("extFields", extFields)
	}
	if userIdType != "" {
		values.Set("user_id_type", userIdType)
	}
	httpCode, respBody, err := l.LarkGet("/open-apis/sheets/v2/spreadsheets/"+excelToken+"/metainfo", values)
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

type UpdateExcelReq struct {
	ExcelToken string
	Properties struct {
		Title string `json:"title"` // 标题,最大长度100个字符
	} `json:"properties"`
}

// UpdateExcel 更新表格属性,暂时只有更新标题
func (l *LarkU) UpdateExcel(req *UpdateExcelReq) (err error) {
	httpCode, respBody, err := l.LarkPut("/open-apis/sheets/v2/spreadsheets/"+req.ExcelToken+"/properties", map[string]interface{}{
		"properties": req.Properties,
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

/** -------------------------------------------------表格-------------------------------------------------------------------- **/

/** -------------------------------------------------sheet-------------------------------------------------------------------- **/

type (
	HandleSheetReq struct {
		ExcelToken string
		Requests   []*HandleSheetRequest `json:"requests,omitempty"`
	}
	AddSheet struct {
		Properties HandleSheetProperties `json:"properties,omitempty"`
	}
	HandleSheetSource struct {
		SheetID string `json:"sheetId,omitempty"`
	}
	HandleSheetDestination struct {
		Title string `json:"title,omitempty"`
	}
	CopySheet struct {
		Source      *HandleSheetSource      `json:"source,omitempty"`
		Destination *HandleSheetDestination `json:"destination,omitempty"` // 目标工作表名称 不填为old_title(副本_0)
	}
	DeleteSheet struct {
		SheetID string `json:"sheetId,omitempty"`
	}
	HandleSheetProtect struct {
		Lock     string   `json:"lock"`               // LOCK 、UNLOCK 上锁/解锁
		LockInfo string   `json:"lockInfo,omitempty"` // 锁定信息
		UserIDs  []string `json:"userIDs,omitempty"`  // 除了本人与所有者外,添加其他的可编辑人员,user_id_type不为空时使用该字段
	}
	HandleSheetProperties struct {
		SheetID        string              `json:"sheetId,omitempty"`        // read-only ,作为表格唯一识别参数
		Title          string              `json:"title,omitempty"`          // 新增/更改工作表标题
		Index          string              `json:"index,omitempty"`          // 新增工作表的位置,不填默认往前增加工作表 | 移动工作表的位置
		Hidden         string              `json:"hidden,omitempty"`         // 隐藏表格,默认false
		FrozenColCount string              `json:"frozenColCount,omitempty"` // 冻结行数,小于等于工作表的最大行数,0表示取消冻结行
		FrozenRowCount string              `json:"frozenRowCount,omitempty"` // 该 sheet 的冻结列数,小于等于工作表的最大列数,0表示取消冻结列
		Protect        *HandleSheetProtect `json:"protect,omitempty"`        // 锁定表格
	}
	UpdateSheet struct {
		Properties *HandleSheetProperties `json:"properties,omitempty"`
	}
	HandleSheetRequest struct {
		AddSheet    *AddSheet    `json:"addSheet,omitempty"`
		CopySheet   *CopySheet   `json:"copySheet,omitempty"`
		DeleteSheet *DeleteSheet `json:"deleteSheet,omitempty"`
		UpdateSheet *UpdateSheet `json:"updateSheet,omitempty"`
	}
)

// HandleSheet 操作工作表,包括增删复制
func (l *LarkU) HandleSheet(req *HandleSheetReq) (err error) {
	httpCode, respBody, err := l.LarkPost("/open-apis/sheets/v2/spreadsheets/"+req.ExcelToken+"/sheets/sheets_batch_update", map[string]interface{}{
		"requests": req.Requests,
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

/** -------------------------------------------------sheet-------------------------------------------------------------------- **/

/** -------------------------------------------------行列--------------------------------------------------------------------- **/

type (
	AddDimensionReq struct {
		Dimension  *DimensionAdd `json:"dimension"`
		ExcelToken string
	}
	DimensionAdd struct {
		SheetID        string `json:"sheetId"`
		MajorDimension string `json:"majorDimension,omitempty"` // 默认ROWS 可选ROWS、COLUMNS
		Length         int    `json:"length"`                   // 要增加的行/列数,0<length<5000
	}
)

const (
	MajorDimensionRows = "ROWS"
	MajorDimensionCols = "COLUMNS"
)

// AddDimension 增加行列
func (l *LarkU) AddDimension(req *AddDimensionReq) (err error) {
	httpCode, respBody, err := l.LarkPost("/open-apis/sheets/v2/spreadsheets/"+req.ExcelToken+"/dimension_range", map[string]interface{}{
		"dimension": req.Dimension,
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

type (
	InsertDimensionReq struct {
		ExcelToken   string
		Dimension    *InsertDimension `json:"dimension"`
		InheritStyle string           `json:"inheritStyle,omitempty"` // BEFORE 或 AFTER，不填为不继承 style
	}
	InsertDimension struct {
		SheetID        string `json:"sheetId"`
		MajorDimension string `json:"majorDimension,omitempty"`
		StartIndex     int    `json:"startIndex"`
		EndIndex       int    `json:"endIndex"`
	}
)

const (
	InheritStyleBefore = "BEFORE"
	InheritStyleAfter  = "AFTER"
)

// InsertDimension 插入行列 用于根据 spreadsheetToken 和维度信息 插入空行/列。
// 如 startIndex=3,endIndex=7,则从第 4 行开始开始插入行列,一直到第 7 行,共插入 4 行;单次操作不超过5000行或列
func (l *LarkU) InsertDimension(req *InsertDimensionReq) (err error) {
	httpCode, respBody, err := l.LarkPost("/open-apis/sheets/v2/spreadsheets/"+req.ExcelToken+"/dimension_range", map[string]interface{}{
		"dimension": req.Dimension,
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

type (
	UpdateDimensionReq struct {
		ExcelToken          string
		Dimension           *UpdateDimension           `json:"dimension"`
		DimensionProperties *UpdateDimensionProperties `json:"dimensionProperties"`
	}
	UpdateDimension struct {
		SheetID        string `json:"sheetId"`
		MajorDimension string `json:"majorDimension,omitempty"`
		StartIndex     int    `json:"startIndex"`
		EndIndex       int    `json:"endIndex"`
	}
	UpdateDimensionProperties struct {
		Visible   bool `json:"visible,omitempty"`   // true为显示 false为隐藏行列
		FixedSize int  `json:"fixedSize,omitempty"` // 行/列的大小
	}
)

// UpdateDimension 更新行列
func (l *LarkU) UpdateDimension(req *UpdateDimensionReq) (err error) {
	httpCode, respBody, err := l.LarkPut("/open-apis/sheets/v2/spreadsheets/"+req.ExcelToken+"/dimension_range", map[string]interface{}{
		"dimension": req.Dimension,
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

type (
	MoveDimensionReq struct {
		ExcelToken       string
		SheetId          string
		Source           *MoveDimensionSource `json:"source,omitempty"`
		DestinationIndex int                  `json:"destination_index,omitempty"`
	}
	MoveDimensionSource struct {
		MajorDimension string `json:"major_dimension,omitempty"` // 操作行还是列，取值：ROWS、COLUMNS
		StartIndex     int    `json:"start_index,omitempty"`
		EndIndex       int    `json:"end_index,omitempty"`
	}
)

// MoveDimension 移动行列
func (l *LarkU) MoveDimension(req *MoveDimensionReq) (err error) {
	httpCode, respBody, err := l.LarkPost("/open-apis/sheets/v3/spreadsheets/"+req.ExcelToken+"/sheets/"+req.SheetId+"/move_dimension", map[string]interface{}{
		"source":            req.Source,
		"destination_index": req.DestinationIndex,
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

type (
	DelDimensionReq struct {
		ExcelToken string
		Dimension  *DelDimensionDimension `json:"dimension"`
	}
	DelDimensionDimension struct {
		SheetID        string `json:"sheetId"`
		MajorDimension string `json:"majorDimension,omitempty"`
		StartIndex     int    `json:"startIndex"`
		EndIndex       int    `json:"endIndex"`
	}
)

// DeleteDimension 删除行列
func (l *LarkU) DelDimension(req *DelDimensionReq) (err error) {
	httpCode, respBody, err := l.LarkDelete("/open-apis/sheets/v2/spreadsheets/"+req.ExcelToken+"/dimension_range", map[string]interface{}{
		"dimension": req.Dimension,
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

/** -------------------------------------------------行列--------------------------------------------------------------------- **/

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
