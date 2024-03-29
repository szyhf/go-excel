package excel

import (
	"fmt"
	"reflect"
)

func (conn *connect) parseSheetName(i interface{}) string {
	// log.Printf("%T", i)
	switch s := i.(type) {
	case string:
		return s
	case int, int8, int32, int64, uint, uint8, uint16, uint32, uint64:
		if name, ok := conn.worksheetIDToNameMap[fmt.Sprintf("%d", s)]; ok {
			return name
		}
		return ""
	case interface {
		GetXLSXSheetName() string
	}:
		return s.GetXLSXSheetName()
	default:
		val := reflect.ValueOf(i)
		typ := reflect.Indirect(val).Type()
		switch typ.Kind() {
		case reflect.Slice, reflect.Array:
			// 数组、切片的元素类型
			typ = typ.Elem()
			if typ.Kind() == reflect.Ptr {
				typ = typ.Elem()
			}
			return conn.parseSheetName(reflect.New(typ).Interface())
		case reflect.Ptr:
			return conn.parseSheetName(reflect.New(typ).Elem().Interface())
		default:
			return typ.Name()
		}

	}
}
