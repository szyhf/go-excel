package excel

import (
	"encoding/xml"
	"fmt"
	"io"
	"reflect"
	"strings"

	convert "github.com/szyhf/go-convert"
	twentysix "github.com/szyhf/go-excel/internal/twenty_six"
)

// read is default implement of reader
type read struct {
	connecter          *connect
	decoder            *xml.Decoder
	decoderReadCloseer io.ReadCloser
	title              *titleRow
	schameMap          map[reflect.Type]*schema
}

// Move the cursor to next row's start.
func (rd *read) Next() bool {
	for t, err := rd.decoder.Token(); err == nil; t, err = rd.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case _RowPrefix:
				return true
			}
		}
	}
	return false
}

// Read current row into an object by its pointer
// return: the last row might be a row with not data,
//
//	in rd case will return io.EOF
func (rd *read) Read(i interface{}) error {
	t := reflect.TypeOf(i)
	switch t.Kind() {
	case reflect.Ptr:
		t = t.Elem()
		switch t.Kind() {
		case reflect.Slice,
			reflect.Array:
			v := reflect.ValueOf(i)
			return rd.readToSlice(t, v)
		case reflect.Struct:
			v := reflect.ValueOf(i)
			return rd.readToStruct(t, v)
		case reflect.Map:
			if t.Key().Kind() != reflect.String {
				return fmt.Errorf("%T should be pointer to map of string key", i)
			}
			v := reflect.ValueOf(i)
			return rd.readToMap(t, v)
		default:
			return fmt.Errorf("%T should be pointer to struct or map[string]string", i)
		}
	default:
		return fmt.Errorf("%T should be pointer to struct or map[string]string", i)
	}
}

func (rd *read) Close() error {
	if rd.decoder != nil {
		rd.decoder = nil
	}
	if rd.decoderReadCloseer != nil {
		rd.decoderReadCloseer.Close()
		rd.decoderReadCloseer = nil
	}
	rd.connecter = nil
	rd.title = nil
	rd.schameMap = nil
	return nil
}

// Read all rows
func (rd *read) ReadAll(container interface{}) error {
	val := reflect.ValueOf(container)
	typ := reflect.Indirect(val).Type()

	if val.Kind() != reflect.Ptr {
		return ErrInvalidConatiner
	}
	if typ.Kind() != reflect.Slice {
		return ErrInvalidConatiner
	}

	elemTyp := typ.Elem()
	if elemTyp.Kind() == reflect.Ptr {
		elemTyp = elemTyp.Elem()
	}
	var err error
	switch elemTyp.Kind() {
	case reflect.Struct:
		elemSchema := newSchema(elemTyp)
		slcVal := val.Elem()
		for rd.Next() {
			elmVal := sliceNextElem(slcVal)
			for err = ErrEmptyRow; err == ErrEmptyRow; {
				err = rd.readToValue(elemSchema, elmVal)
			}
			if err != nil {
				// remove the last row.
				slcVal.SetLen(slcVal.Len() - 1)
				if err != io.EOF {
					// EOF is normal.
					return err
				}
			}
		}
	case reflect.Map:
		slcVal := val.Elem()
		for rd.Next() {
			elmVal := sliceNextElem(slcVal)
			for err = ErrEmptyRow; err == ErrEmptyRow; {
				err = rd.readToMap(elemTyp, elmVal)
			}
			if err != nil {
				// remove the last row.
				slcVal.SetLen(slcVal.Len() - 1)
				if err != io.EOF {
					// EOF is normal.
					return err
				}
			}
		}
	case reflect.Slice,
		reflect.Array:
		slcVal := val.Elem()
		for rd.Next() {
			elmVal := sliceNextElem(slcVal)
			for err = ErrEmptyRow; err == ErrEmptyRow; {
				err = rd.readToSlice(elemTyp, elmVal)
			}
			if err != nil {
				// remove the last row.
				slcVal.SetLen(slcVal.Len() - 1)
				if err != io.EOF {
					// EOF is normal.
					return err
				}
			}
		}
	}

	return nil
}

func (rd *read) GetTitles() []string {
	// prevent unexpect edit
	titles := make([]string, len(rd.title.titles))
	copy(titles, rd.title.titles)
	return titles
}

func (rd *read) readToStruct(t reflect.Type, v reflect.Value) error {
	if len(rd.title.dstMap) != len(rd.title.titles) {
		return ErrDuplicatedTitles
	}

	s := rd.getSchame(t)
	if v.IsNil() {
		v.Set(reflect.New(t))
	}
	v = v.Elem()

	var err error
	for err = ErrEmptyRow; err == ErrEmptyRow; {
		err = rd.readToValue(s, v)
	}
	return err
}

// v should be value of map[string]string
func (rd *read) readToMap(t reflect.Type, v reflect.Value) error {
	if len(rd.title.dstMap) != len(rd.title.titles) {
		return ErrDuplicatedTitles
	}

	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}
	v.Set(reflect.MakeMapWithSize(t, len(rd.title.dstMap)))

	var err error
	for err = ErrEmptyRow; err == ErrEmptyRow; {
		err = rd.readToMapValue(v)
	}

	if v.Len() < len(rd.title.dstMap) {
		for _, keyValue := range v.MapKeys() {
			title := keyValue.String()
			if _, ok := rd.title.dstMap[title]; !ok {
				// fill default value to column not read.
				val := reflect.New(v.Type().Elem())
				v.SetMapIndex(reflect.ValueOf(title), val.Elem())
			}
		}
	}
	return err
}

func (rd *read) readToSlice(t reflect.Type, v reflect.Value) (err error) {
	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}
	// if use rd.title.dstMap as len to slice, the titles with duplcated key will be ignored.
	// in slice case, ignore the duplicated key and read all cells is ok.
	v.Set(reflect.MakeSlice(t, len(rd.title.titles), len(rd.title.titles)))

	for err = ErrEmptyRow; err == ErrEmptyRow; {
		err = rd.readToSliceValue(v)
	}
	return
}

func (rd *read) readToValue(s *schema, v reflect.Value) (err error) {
	if len(rd.title.dstMap) != len(rd.title.titles) {
		return ErrDuplicatedTitles
	}

	tempCell := &xlsxC{}
	fieldsMap, err := rd.title.MapToFields(s)
	if err != nil {
		return err
	}
	scaned := false
	defer func() {
		if !scaned && err == nil {
			err = ErrEmptyRow
		}
	}()

	isV := false
	for t, e := rd.decoder.Token(); e == nil; t, e = rd.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case _C:
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
			case _V:
				isV = true
			}
		case xml.EndElement:
			// 避免上个节点的状态未正确处理以后延伸到下次
			isV = false
			if token.Name.Local == _RowPrefix {
				// fill default value to column not read.
				for _, notFilledFields := range fieldsMap {
					for _, fieldCnf := range notFilledFields {
						fieldValue := v.Field(fieldCnf.FieldIndex)
						// log.Printf("Fill %s = %v with default: %s", v.Type().Field(fieldCnf.FieldIndex).Name, fieldValue.Interface(), fieldCnf.DefaultValue)
						err = fieldCnf.ScanDefault(fieldValue)
						if err != nil {
							return err
						}
					}
				}
				// 结束当前行
				return err
			}
		case xml.CharData:
			if !isV {
				// log.Println(string(token))
				break
			}
			trimedColumnName := strings.TrimRight(tempCell.R, _AllNumber)
			columnIndex := twentysix.ToDecimalism(trimedColumnName)
			fields, ok := fieldsMap[columnIndex]
			if !ok {
				// Not an error, just ignore rd column.
				break
			}

			var valStr string
			if tempCell.T == _S {
				// get string from shared
				valStr = rd.connecter.getSharedString(convert.MustInt(string(token)))
			} else {
				valStr = string(token)
			}
			// println("Key:", trimedColumnName, "Val:", valStr)
			scaned = true
			isV = false
			for _, fieldCnf := range fields {
				fieldValue := v.Field(fieldCnf.FieldIndex)
				err = fieldCnf.scan(valStr, fieldValue)
				if err != nil && len(valStr) > 0 {
					return err
				}
			}
			if err == nil {
				delete(fieldsMap, columnIndex)
			}
		}
	}

	if err != nil {
		return err
	}
	return io.EOF
}

func (rd *read) readToMapValue(v reflect.Value) (err error) {
	if len(rd.title.dstMap) != len(rd.title.titles) {
		return ErrDuplicatedTitles
	}

	tempCell := &xlsxC{}
	scaned := false
	defer func() {
		if !scaned && err == nil {
			err = ErrEmptyRow
		}
	}()
	isV := false
	for t, e := rd.decoder.Token(); e == nil; t, e = rd.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case _C:
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
			case _V:
				isV = true
			}
		case xml.EndElement:
			// 避免上个节点的状态未正确处理以后延伸到下次
			isV = false
			if token.Name.Local == _RowPrefix {
				// end of current row
				return err
			}
		case xml.CharData:
			if !isV {
				break
			}
			trimedColumnName := strings.TrimRight(tempCell.R, _AllNumber)
			var valStr string
			if tempCell.T == _S {
				// get string from shared
				valStr = rd.connecter.getSharedString(convert.MustInt(string(token)))
			} else {
				valStr = string(token)
			}
			val := reflect.New(v.Type().Elem())
			_ = scan(valStr, val.Interface())
			columnIndex := twentysix.ToDecimalism(trimedColumnName)
			title := rd.title.srcMap[columnIndex]
			v.SetMapIndex(reflect.ValueOf(title), val.Elem())
			// log.Println("Key:", trimedColumnName, "Val:", valStr)
			scaned = true
			isV = false
		}
	}

	if err != nil {
		return err
	}
	return io.EOF
}

func (rd *read) readToSliceValue(v reflect.Value) (err error) {
	tempCell := &xlsxC{}
	scaned := false
	defer func() {
		if !scaned && err == nil {
			err = ErrEmptyRow
		}
	}()
	isV := false
	for t, e := rd.decoder.Token(); e == nil; t, e = rd.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case _C:
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
			case _V:
				isV = true
			}
		case xml.EndElement:
			// 避免上个节点的状态未正确处理以后延伸到下次
			isV = false
			if token.Name.Local == _RowPrefix {
				// end of current row
				return err
			}
		case xml.CharData:
			if !isV {
				break
			}
			trimedColumnName := strings.TrimRight(tempCell.R, _AllNumber)
			var valStr string
			if tempCell.T == _S {
				// get string from shared
				valStr = rd.connecter.getSharedString(convert.MustInt(string(token)))
			} else {
				valStr = string(token)
			}

			columnIndex := twentysix.ToDecimalism(trimedColumnName)
			if columnIndex < v.Len() {
				val := v.Index(columnIndex)
				if val.Type().Kind() == reflect.Ptr {
					val.Set(reflect.New(val.Type().Elem()))
					_ = scan(valStr, val.Interface())
				} else if val.CanAddr() {
					_ = scan(valStr, val.Addr().Interface())
				} else {
					return fmt.Errorf("unexpect type of %T, is not ptr and can't addr", v.Interface())
				}

				// } else {
				// log.Printf("columnIndex(%d) < v.Len(%d)", columnIndex, v.Len())
			}
			// log.Println("Key:", trimedColumnName, "Val:", valStr)
			scaned = true
			isV = false
		}
	}

	if err != nil {
		return err
	}
	return io.EOF
}

func (rd *read) getSchame(t reflect.Type) *schema {
	s, ok := rd.schameMap[t]
	if !ok {
		s = newSchema(t)
		rd.schameMap[t] = s
	}
	return s
}

func newReader(cn *connect, workSheetFileReader io.ReadCloser, titleRowIndex, skip int) (Reader, error) {
	rd, err := newBaseReaderByWorkSheetFile(cn, workSheetFileReader)
	if err != nil {
		return nil, err
	}
	// consider title row
	var i = 0
	// <= because Next() have to put the pointer to the Index row.
	for ; i <= titleRowIndex; i++ {
		if !rd.Next() {
			return rd, nil
		}
	}
	rd.title, err = newRowAsMap(rd)

	// consider skip
	// Next() will called before Read() so just skip cursor to the row before first data row.
	// log.Println("Start for skip")
	for i = 0; i < skip; i++ {
		if !rd.Next() {
			return rd, nil
		}
	}
	rd.schameMap = make(map[reflect.Type]*schema)
	return rd, err
}

// Make a base reader to sheet
func newBaseReaderByWorkSheetFile(cn *connect, rc io.ReadCloser) (*read, error) {
	decoder := xml.NewDecoder(rc)
	// step into root [xml.StartElement] token
	func(decoder *xml.Decoder) {
		for t, err := decoder.Token(); err == nil; t, err = decoder.Token() {
			// [xml.ProcInst]
			// [xml.CharData]
			// [xml.StartElement]
			switch t.(type) {
			case xml.StartElement:
				return
			}
		}
	}(decoder)

	err := func(decoder *xml.Decoder) error {
		// use func block to break to 'for' range
		for t, err := decoder.Token(); err == nil; t, err = decoder.Token() {
			// log.Printf("%+v\n\n", t)
			switch token := t.(type) {
			case xml.StartElement:
				switch token.Name.Local {
				case _SheetData:
					return nil
				default:
					if err := decoder.Skip(); err != nil {
						return err
					}
				}
			}
		}
		return nil
	}(decoder)

	if err != nil {
		return nil, err
	}

	rd := &read{
		connecter:          cn,
		decoder:            decoder,
		decoderReadCloseer: rc,
	}

	return rd, nil
}
