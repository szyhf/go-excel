package excel

import (
	"encoding/xml"
	"fmt"
	"io"
	"reflect"
	"strings"

	"github.com/szyhf/go-excel/internal/twenty_six"

	convert "github.com/szyhf/go-convert"
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
//         in rd case will return io.EOF
func (rd *read) Read(i interface{}) error {
	t := reflect.TypeOf(i)
	switch t.Kind() {
	case reflect.Ptr:
		t = t.Elem()
		switch t.Kind() {
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
	}

	return nil
}

func (rd *read) readToStruct(t reflect.Type, v reflect.Value) error {
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

func (rd *read) readToValue(s *schema, v reflect.Value) (err error) {
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
	for t, e := rd.decoder.Token(); e == nil; t, e = rd.decoder.Token() {
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
	tempCell := &xlsxC{}
	scaned := false
	defer func() {
		if !scaned && err == nil {
			err = ErrEmptyRow
		}
	}()
	for t, e := rd.decoder.Token(); e == nil; t, e = rd.decoder.Token() {
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
				// end of current row
				return err
			}
		case xml.CharData:
			trimedColumnName := strings.TrimRight(tempCell.R, _AllNumber)
			var valStr string
			if tempCell.T == _S {
				// get string from shared
				valStr = rd.connecter.getSharedString(convert.MustInt(string(token)))
			} else {
				valStr = string(token)
			}
			val := reflect.New(v.Type().Elem())
			err := scan(valStr, val.Interface())
			if err != nil {
				// skip
			}
			columnIndex := twentysix.ToDecimalism(trimedColumnName)
			title := rd.title.srcMap[columnIndex]
			v.SetMapIndex(reflect.ValueOf(title), val.Elem())
			// log.Println("Key:", trimedColumnName, "Val:", valStr)
			scaned = true
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

	func(decoder *xml.Decoder) {
		// use func block to break to 'for' range
		for t, err := decoder.Token(); err == nil; t, err = decoder.Token() {
			// log.Printf("%+v\n\n", t)
			switch token := t.(type) {
			case xml.StartElement:
				switch token.Name.Local {
				case _SheetData:
					return
				default:
					decoder.Skip()
				}
			}
		}
	}(decoder)

	rd := &read{
		connecter:          cn,
		decoder:            decoder,
		decoderReadCloseer: rc,
	}

	return rd, nil
}
