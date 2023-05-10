package utils

import "strings"

// GetFormulaString 获得在；个字符就换行的文本
func GetFormulaString(origin string) string {
	result := ""
	for strings.Index(origin, "；") > 0 {
		result += origin[0:strings.Index(origin, "；")+3] + "\n"
		origin = origin[strings.Index(origin, "；")+3:]
	}

	if result == "" {
		return origin
	} else {
		return result
	}

}

// RemoveStringChineseSpace 去除中文字符串中的空格
func RemoveStringChineseSpace(a string) string {
	return strings.Replace(a, " ", "", -1)
}
