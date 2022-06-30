package lark_util

import (
	"encoding/json"
	"fmt"
	"net/http"

	"github.com/pkg/errors"
)

func (l *LarkU) GetUserId(email string) (userId string, err error) {
	if email == "" {
		panic("email is empty")
	}

	httpCode, respBody, err := l.LarkPost("/open-apis/contact/v3/users/batch_get_id", map[string]interface{}{
		"emails": []string{email},
	})
	if err != nil {
		err = errors.Errorf("http error: %+v", err)
		return
	}
	if httpCode != http.StatusOK {
		err = errors.Errorf("http error: code= %d | %+v", httpCode, respBody)
		return
	}
	type GetUserIdResp struct {
		Code int32  `json:"code,omitempty"`
		Msg  string `json:"msg"`
		Data struct {
			UserList []struct {
				UserId string `json:"user_id,omitempty"`
				Email  string `json:"email,omitempty"`
			} `json:"user_list,omitempty"`
		}
	}
	m := new(GetUserIdResp)
	_ = json.Unmarshal(respBody, &m)
	if m.Code != 0 {
		err = errors.Errorf("remote service error: code = %d | %s", m.Code, m.Msg)
		return
	}
	userId = fmt.Sprint(m.Data.UserList[0].UserId)
	return
}
