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
		switch typ.Kind() {
		case reflect.Slice, reflect.Ptr:
			typ = typ.Elem()
			return ParseSheetName(reflect.New(typ).Elem().Interface())
		default:
			return typ.Name()
		}

	}
}
