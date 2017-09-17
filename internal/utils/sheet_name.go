package utils

import (
	"reflect"
)

func ParseSheetName(i interface{}) string {
	switch s := i.(type) {
	case string:
		return s
	case interface {
		GetXLSXSheetName() string
	}:
		return s.GetXLSXSheetName()
	default:
		val := reflect.ValueOf(i)
		typ := reflect.Indirect(val).Type()
		return typ.Name()
	}
}
