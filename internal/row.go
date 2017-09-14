package internal

import (
	"encoding/xml"
	"errors"
	"fmt"
	"strings"

	convert "github.com/szyhf/go-convert"
	"github.com/szyhf/go-excel/internal/twenty_six"
)

type Row struct {
	// map[A1]1
	srcMap map[int]string
}

func newRow(rd *Read) (*Row, error) {
	row := &xlsxRow{}
	err := rd.decoder.Decode(row)

	if err != nil {
		fmt.Println(err)
		return nil, err
	}

	r := &Row{
		srcMap: make(map[int]string),
	}
	// for _, cell := range row.C {
	// 	trimedColumnName := strings.TrimRight(cell.R, ALL_NUMBER)
	// 	if cell.T == S {
	// 		// get string from shared
	// 		r.srcMap[trimedColumnName] = rd.connecter.getSharedString(convert.MustInt(cell.V))
	// 	} else {
	// 		r.srcMap[trimedColumnName] = cell.V
	// 	}

	// }
	return r, nil
}

func newRowAsMap(rd *Read) (r *Row, err error) {
	defer func() {
		if rc := recover(); rc != nil {
			err = fmt.Errorf("%s", rc)
		}
	}()
	r = &Row{
		srcMap: make(map[int]string),
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
			if tempCell.T == S {
				// get string from shared
				r.srcMap[columnIndex] = rd.connecter.getSharedString(convert.MustInt(string(token)))
			} else {
				r.srcMap[columnIndex] = string(token)
			}
		}
	}

	return nil, errors.New("No row")
}

func (this *Row) ValueAtColumn(i int) string {
	// expect empty as ""
	v, _ := this.srcMap[i]
	return v
}
