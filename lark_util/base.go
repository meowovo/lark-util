package lark_util

import (
	"bytes"
	"encoding/json"
	"io"
	"io/ioutil"
	"net/http"
	"net/url"
	"time"
)

type LarkU struct {
	appId     string
	appSecret string
	client    *http.Client
	larkHost  string
	larkToken string
}

type LarkMeta struct {
	AppId     string
	AppSecret string
}

func NewLarkU(meta *LarkMeta) *LarkU {
	if meta == nil {
		panic("lark meta is nil")
	}
	l := &LarkU{
		appId:     meta.AppId,
		appSecret: meta.AppSecret,
		client: &http.Client{
			Timeout: time.Second * 5,
			Transport: &http.Transport{
				MaxIdleConns:        100,
				MaxIdleConnsPerHost: 100,
				IdleConnTimeout:     60 * time.Second,
				ReadBufferSize:      16 * 1024,
				WriteBufferSize:     8 * 1024,
				DisableCompression:  true,
			},
		},
		larkHost: "open.feishu.cn",
	}
	l.setLarkToken()
	go func() {
		for range time.NewTicker(time.Minute * 5).C {
			l.setLarkToken()
		}
	}()
	return l
}

func (l *LarkU) setLarkToken() {
	b, _ := json.Marshal(map[string]string{
		"app_id":     l.appId,
		"app_secret": l.appSecret,
	})
	resp, err := l.client.Post("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", "application/json; charset=utf-8", bytes.NewReader(b))
	if err != nil {
		panic(err)
	}
	defer resp.Body.Close()
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		panic(err)
	}
	type tmp struct {
		Code              int32  `json:"code,omitempty"`
		TenantAccessToken string `json:"tenant_access_token,omitempty"`
	}
	t := new(tmp)
	_ = json.Unmarshal(body, t)
	if t.TenantAccessToken != "" {
		l.larkToken = t.TenantAccessToken
	}
}

func (l *LarkU) LarkPost(path string, param map[string]interface{}) (httpCode int, responseBody []byte, err error) {
	b, _ := json.Marshal(param)
	var bufferBody *bytes.Buffer
	if b == nil || len(b) <= 0 {
		bufferBody = bytes.NewBuffer([]byte("{}"))
	} else {
		bufferBody = bytes.NewBuffer(b)
	}
	req, err := http.NewRequest(http.MethodPost, "https://"+l.larkHost+path, bufferBody)
	if err != nil {
		return
	}
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Authorization", "Bearer "+l.larkToken)
	resp, err := l.client.Do(req)
	if err != nil {
		httpCode = http.StatusInternalServerError
		return
	}
	responseBody, err = ioutil.ReadAll(resp.Body)
	httpCode = resp.StatusCode
	return
}

func (l *LarkU) LarkGet(path string, form url.Values) (httpCode int, responseBody []byte, err error) {
	urlStr := "https://" + l.larkHost + path
	if len(form) > 0 {
		urlStr += "?" + form.Encode()
	}
	req, _ := http.NewRequest(http.MethodGet, urlStr, nil)
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Authorization", "Bearer "+l.larkToken)
	resp, err := l.client.Do(req)
	if err != nil {
		httpCode = http.StatusInternalServerError
		return
	}
	responseBody, err = ioutil.ReadAll(resp.Body)
	httpCode = resp.StatusCode
	return
}

func (l *LarkU) LarkPut(path string, param map[string]interface{}) (httpCode int, responseBody []byte, err error) {
	b, _ := json.Marshal(param)
	var bufferBody *bytes.Buffer
	if b == nil || len(b) <= 0 {
		bufferBody = bytes.NewBuffer([]byte("{}"))
	} else {
		bufferBody = bytes.NewBuffer(b)
	}
	req, err := http.NewRequest(http.MethodPut, "https://"+l.larkHost+path, bufferBody)
	if err != nil {
		return
	}
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Authorization", "Bearer "+l.larkToken)
	resp, err := l.client.Do(req)
	if err != nil {
		httpCode = http.StatusInternalServerError
		return
	}
	responseBody, err = ioutil.ReadAll(resp.Body)
	httpCode = resp.StatusCode
	return
}
