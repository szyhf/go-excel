package internal

import (
	"encoding/xml"
	"errors"
	"fmt"
	"reflect"
	"strings"

	convert "github.com/szyhf/go-convert"
	"github.com/szyhf/go-excel/internal/twenty_six"
)

type Row struct {
	// map[A1]0
	// srcMap map[int]string
	// map[0]A1
	dstMap map[string]int
}

func newRowAsMap(rd *Read) (r *Row, err error) {
	defer func() {
		if rc := recover(); rc != nil {
			err = fmt.Errorf("%s", rc)
		}
	}()
	r = &Row{
		// srcMap: make(map[int]string),
		dstMap: make(map[string]int),
	}
	tempCell := &xlsxC{}
	for t, err := rd.decoder.Token(); err == nil; t, err = rd.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			if token.Name.Local == "c" {
				tempCell.R = ""
				tempCell.T = ""
				for _, a := range token.Attr {
					switch a.Name.Local {
					case "r":
						tempCell.R = a.Value
					case "t":
						tempCell.T = a.Value
					}
				}
			}
		case xml.EndElement:
			if token.Name.Local == ROW {
				// 结束当前行
				return r, nil
			}
		case xml.CharData:
			trimedColumnName := strings.TrimRight(tempCell.R, ALL_NUMBER)
			columnIndex := twentySix.ToDecimalism(trimedColumnName)
			var str string
			if tempCell.T == S {
				// get string from shared
				str = rd.connecter.getSharedString(convert.MustInt(string(token)))
			} else {
				str = string(token)
			}
			// r.srcMap[columnIndex] = str
			r.dstMap[str] = columnIndex
		}
	}

	return nil, errors.New("No row")
}

func newRowBySchame(rd *Read, title *Row, s *Schema, v reflect.Value) (err error) {
	defer func() {
		if rc := recover(); rc != nil {
			err = fmt.Errorf("%s", rc)
		}
	}()

	tempCell := &xlsxC{}
	fieldsMap := title.MapToFields(s)
	for t, err := rd.decoder.Token(); err == nil; t, err = rd.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			if token.Name.Local == "c" {
				tempCell.R = ""
				tempCell.T = ""
				for _, a := range token.Attr {
					switch a.Name.Local {
					case "r":
						tempCell.R = a.Value
					case "t":
						tempCell.T = a.Value
					}
				}
			}
		case xml.EndElement:
			if token.Name.Local == ROW {
				// 结束当前行
				return nil
			}
		case xml.CharData:
			trimedColumnName := strings.TrimRight(tempCell.R, ALL_NUMBER)
			columnIndex := twentySix.ToDecimalism(trimedColumnName)
			var valStr string
			if tempCell.T == S {
				// get string from shared
				valStr = rd.connecter.getSharedString(convert.MustInt(string(token)))
			} else {
				valStr = string(token)
			}

			fields := fieldsMap[columnIndex]
			for _, fieldCnf := range fields {
				fieldValue := v.Field(fieldCnf.FieldIndex)
				if fieldValue.Kind() == reflect.Ptr {
					fieldValue = fieldValue.Elem()
				}
				Scan(valStr, fieldValue.Addr().Interface())
			}

			// r.srcMap[columnIndex] = str
			// r.dstMap[str] = columnIndex
		}
	}

	return errors.New("No row")
}

// return: map[ColumnIndex][]*Field
func (this *Row) MapToFields(s *Schema) (rowToFiled map[int][]*Field) {
	m := make(map[int][]*Field)
	for _, field := range s.Fields {
		var cloIndex int
		if field.ColumnName == "" {
			// Use ColumnIndex as default
			cloIndex = field.ColumnIndex
		} else {
			// Use ColumnName to find index
			if i, ok := this.dstMap[field.ColumnName]; ok {
				cloIndex = i
			}
		}
		if fAry, ok := m[cloIndex]; !ok {
			m[cloIndex] = []*Field{field}
		} else {
			m[cloIndex] = append(fAry, field)
		}
	}
	return m
}
