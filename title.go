package excel

import (
	"encoding/xml"
	"fmt"
	"reflect"
	"strings"

	convert "github.com/szyhf/go-convert"
	twentysix "github.com/szyhf/go-excel/internal/twenty_six"
)

type titleRow struct {
	// map[A1]0
	dstMap map[string]int
	// map[0]A1
	srcMap map[int]string

	// sorted titles
	titles []string

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
		srcMap: make(map[int]string),
		titles: make([]string, 0),
	}
	tempCell := &xlsxC{}
	for t, err := rd.decoder.Token(); err == nil; t, err = rd.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case _RowPrefix:
				// start of row
			case _C:
				// start of row.cell
				tempCell.R = ""
				tempCell.T = ""
				tempCell.V = ""
				tempCell.columnIndex = -1
				for _, a := range token.Attr {
					switch a.Name.Local {
					case _R:
						tempCell.R = a.Value
					case _T:
						tempCell.T = a.Value
					}
				}
			case _V:
				// start of row.cell.value
			}
		case xml.EndElement:
			switch token.Name.Local {
			case _RowPrefix:
				// end of row
				r.typeFieldMap = make(map[reflect.Type]map[int][]*fieldConfig)
				return r, nil
			case _C:
				// end of row.cell
				for i := len(r.titles); i < tempCell.columnIndex; i++ {
					// fill the skipped empty cell with blank
					const blankText = ""
					r.dstMap[blankText] = i
					r.srcMap[i] = blankText
					r.titles = append(r.titles, blankText)
				}
				r.dstMap[tempCell.V] = tempCell.columnIndex
				r.srcMap[tempCell.columnIndex] = tempCell.V
				r.titles = append(r.titles, tempCell.V)
			case _V:
				// end of row.cell.value
			}
		case xml.CharData:
			trimedColumnName := strings.TrimRight(tempCell.R, _AllNumber)
			tempCell.columnIndex = twentysix.ToDecimalism(trimedColumnName)

			if tempCell.T == _S {
				// get string from shared
				tempCell.V = rd.connecter.getSharedString(convert.MustInt(string(token)))
			} else {
				tempCell.V = string(token)
			}
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
