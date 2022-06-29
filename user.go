package lark_util

import (
	"encoding/json"
	"fmt"
)

var (
	userId string
	email  string
)

func getUserId() string {
	if userId != "" {
		return userId
	}

	if email == "" {
		panic("email is empty")
	}

	httpCode, respBody, err := larkPost("/open-apis/contact/v3/users/batch_get_id", map[string]interface{}{
		"emails": []string{email},
	})
	if err != nil {
		return ""
	}
	type GetUserIdResp struct {
		Code int32 `json:"code,omitempty"`
		Data struct {
			UserList []struct {
				UserId string `json:"user_id,omitempty"`
				Email  string `json:"email,omitempty"`
			} `json:"user_list,omitempty"`
		}
	}
	m := new(GetUserIdResp)
	_ = json.Unmarshal(respBody, &m)
	userId = fmt.Sprint(m.Data.UserList[0].UserId)
	return userId
}
