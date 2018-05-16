package excel

import (
	"encoding/xml"
	"fmt"
	"reflect"
	"strings"

	convert "github.com/szyhf/go-convert"

	"github.com/szyhf/go-excel/internal/twenty_six"
)

type titleRow struct {
	// map[0]A1
	dstMap map[string]int

	typeFieldMap map[reflect.Type]map[int][]*fieldConfig
}

func newRowAsMap(rd *read) (r *titleRow, err error) {
	defer func() {
		if rc := recover(); rc != nil {
			err = fmt.Errorf("%s", rc)
		}
	}()
	r = &titleRow{
		dstMap: make(map[string]int),
	}
	tempCell := &xlsxC{}
	for t, err := rd.decoder.Token(); err == nil; t, err = rd.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			if token.Name.Local == _C {
				tempCell.R = ""
				tempCell.T = ""
				for _, a := range token.Attr {
					switch a.Name.Local {
					case _R:
						tempCell.R = a.Value
					case _T:
						tempCell.T = a.Value
					}
				}
			}
		case xml.EndElement:
			if token.Name.Local == _RowPrefix {
				// 结束当前行
				r.typeFieldMap = make(map[reflect.Type]map[int][]*fieldConfig)
				return r, nil
			}
		case xml.CharData:
			trimedColumnName := strings.TrimRight(tempCell.R, _AllNumber)
			columnIndex := twentysix.ToDecimalism(trimedColumnName)
			var str string
			if tempCell.T == _S {
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

// return: a copy of map[ColumnIndex][]*fieldConfig
func (tr *titleRow) MapToFields(s *schema) (rowToFiled map[int][]*fieldConfig, err error) {
	fieldsMap, ok := tr.typeFieldMap[s.Type]
	if !ok {
		fieldsMap = make(map[int][]*fieldConfig)
		for _, field := range s.Fields {
			var cloIndex int
			// Use ColumnName to find index
			if i, ok := tr.dstMap[field.ColumnName]; ok {
				cloIndex = i
			} else if field.IsRequired {
				// Use 26-number-system to find
				// cloIndex = twentysix.ToDecimalism(field.ColumnName)
				return nil, fmt.Errorf("go-excel: column name = \"%s\" is not exist", field.ColumnName)
			} else {
				// continue if is not required.
				continue
			}

			if fAry, ok := fieldsMap[cloIndex]; !ok {
				fieldsMap[cloIndex] = []*fieldConfig{field}
			} else {
				fieldsMap[cloIndex] = append(fAry, field)
			}
		}
		tr.typeFieldMap[s.Type] = fieldsMap
	}
	copyMap := make(map[int][]*fieldConfig)
	for k, v := range fieldsMap {
		copyMap[k] = v
	}
	return copyMap, nil
}
