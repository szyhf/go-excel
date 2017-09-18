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
	// map[0]A1
	dstMap map[string]int

	filedMap map[reflect.Type]map[int][]*Field
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
				r.filedMap = make(map[reflect.Type]map[int][]*Field)
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

// return: map[ColumnIndex][]*Field
func (this *Row) MapToFields(s *Schema) (rowToFiled map[int][]*Field) {
	m, ok := this.filedMap[s.Type]
	if !ok {
		m = make(map[int][]*Field)
		for _, field := range s.Fields {
			var cloIndex int
			// Use ColumnName to find index
			if i, ok := this.dstMap[field.ColumnName]; ok {
				cloIndex = i
			} else {
				// Use 26-number-system to find
				cloIndex = twentySix.ToDecimalism(field.ColumnName)
			}

			if fAry, ok := m[cloIndex]; !ok {
				m[cloIndex] = []*Field{field}
			} else {
				m[cloIndex] = append(fAry, field)
			}
		}
		this.filedMap[s.Type] = m
	}
	return m
}
