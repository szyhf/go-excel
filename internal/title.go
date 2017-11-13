package internal

import (
	"encoding/xml"
	"fmt"
	"reflect"
	"strings"

	convert "github.com/szyhf/go-convert"
	"github.com/szyhf/go-excel/internal/twenty_six"
)

type TitleRow struct {
	// map[0]A1
	dstMap map[string]int

	typeFieldMap map[reflect.Type]map[int][]*FieldConfig
}

func newRowAsMap(rd *Read) (r *TitleRow, err error) {
	defer func() {
		if rc := recover(); rc != nil {
			err = fmt.Errorf("%s", rc)
		}
	}()
	r = &TitleRow{
		dstMap: make(map[string]int),
	}
	tempCell := &xlsxC{}
	for t, err := rd.decoder.Token(); err == nil; t, err = rd.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			if token.Name.Local == C {
				tempCell.R = ""
				tempCell.T = ""
				for _, a := range token.Attr {
					switch a.Name.Local {
					case R:
						tempCell.R = a.Value
					case T:
						tempCell.T = a.Value
					}
				}
			}
		case xml.EndElement:
			if token.Name.Local == ROW {
				// 结束当前行
				r.typeFieldMap = make(map[reflect.Type]map[int][]*FieldConfig)
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

	return nil, ErrNoRow
}

// return: a copy of map[ColumnIndex][]*FieldConfig
func (this *TitleRow) MapToFields(s *Schema) (rowToFiled map[int][]*FieldConfig) {
	fieldsMap, ok := this.typeFieldMap[s.Type]
	if !ok {
		fieldsMap = make(map[int][]*FieldConfig)
		for _, field := range s.Fields {
			var cloIndex int
			// Use ColumnName to find index
			if i, ok := this.dstMap[field.ColumnName]; ok {
				cloIndex = i
			} else {
				// Use 26-number-system to find
				// cloIndex = twentySix.ToDecimalism(field.ColumnName)
				panic(fmt.Sprintf("go-excel: column name = \"%s\" is not exist.", field.ColumnName))
			}

			if fAry, ok := fieldsMap[cloIndex]; !ok {
				fieldsMap[cloIndex] = []*FieldConfig{field}
			} else {
				fieldsMap[cloIndex] = append(fAry, field)
			}
		}
		this.typeFieldMap[s.Type] = fieldsMap
	}
	copyMap := make(map[int][]*FieldConfig)
	for k, v := range fieldsMap {
		copyMap[k] = v
	}
	return copyMap
}
