package internal

import (
	"encoding"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"reflect"
	"strings"

	convert "github.com/szyhf/go-convert"

	"github.com/szyhf/go-excel/internal/twenty_six"
)

var typeOfTextUnmarshaler = reflect.TypeOf(encoding.TextUnmarshaler(nil))

type Read struct {
	connecter *Connect
	decoder   *xml.Decoder
	title     *Row
}

func (this *Read) Next() bool {
	for t, err := this.decoder.Token(); err == nil; t, err = this.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case ROW:
				return true
			}
		}
	}
	return false
}

// Read current row into an object by its pointer
func (this *Read) Read(i interface{}) error {
	t := reflect.TypeOf(i)
	switch t.Kind() {
	case reflect.Slice, reflect.Chan, reflect.Array, reflect.Map, reflect.Ptr:
		t = t.Elem()
	default:
		return fmt.Errorf("%T should be pointer.", i)
	}

	s := newSchema(t)

	v := reflect.ValueOf(i)
	if v.IsNil() {
		v.Set(reflect.New(t))
	}
	v = v.Elem()

	return this.readToValue(s, v)
}

func (this *Read) Close() error {
	if this.decoder != nil {
		this.decoder = nil
	}
	return nil
}

// Read all rows
func (this *Read) ReadAll(container interface{}) error {
	val := reflect.ValueOf(container)
	typ := reflect.Indirect(val).Type()

	if val.Kind() != reflect.Ptr {
		return ErrInvalidConatiner
	}
	if typ.Kind() != reflect.Slice {
		return ErrInvalidConatiner
	}

	elemTyp := typ.Elem()
	elemSchema := newSchema(elemTyp)

	for this.Next() {
		elmVal := SliceNextElem(val.Elem())
		err := this.readToValue(elemSchema, elmVal)
		if err != nil {
			return err
		}
	}
	return nil
}

func (this *Read) readToValue(s *Schema, v reflect.Value) (err error) {
	defer func() {
		if rc := recover(); rc != nil {
			err = fmt.Errorf("%s", rc)
		}
	}()

	tempCell := &xlsxC{}
	fieldsMap := this.title.MapToFields(s)
	for t, err := this.decoder.Token(); err == nil; t, err = this.decoder.Token() {
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
				valStr = this.connecter.getSharedString(convert.MustInt(string(token)))
			} else {
				valStr = string(token)
			}

			fields := fieldsMap[columnIndex]
			for _, fieldCnf := range fields {
				fieldValue := v.Field(fieldCnf.FieldIndex)
				switch fieldValue.Kind() {
				case reflect.Slice, reflect.Array:
					if len(fieldCnf.Split) != 0 {
						// use split
						elems := strings.Split(valStr, fieldCnf.Split)
						fieldValue.Set(reflect.MakeSlice(fieldValue.Type(), 0, len(elems)))
						err = ScanSlice(elems, fieldValue.Addr())
					}
				case reflect.Ptr:
					newValue := fieldValue
					if newValue.IsNil() {
						for newValue.Kind() == reflect.Ptr {
							newValue.Set(reflect.New(newValue.Type().Elem()))
							newValue = newValue.Elem()
						}
					}
					err = Scan(valStr, newValue.Addr().Interface())
				default:
					err = Scan(valStr, fieldValue.Addr().Interface())
				}
				if err != nil {
					println(fieldCnf.ColumnName)
					return err
				}
			}
		}
	}

	return errors.New("No row")
}

func newReader(cn *Connect, workSheetFileReader io.Reader, titleRowIndex, skip int) (Reader, error) {
	rd, err := newBaseReaderByWorkSheetFile(cn, workSheetFileReader)
	if err != nil {
		return nil, err
	}
	// consider title row
	var i = 0
	for ; i < titleRowIndex; i++ {
		if !rd.Next() {
			return rd, nil
		}
	}
	rd.title, err = newRowAsMap(rd)

	// consider skip
	for i = 0; i < skip; i++ {
		if !rd.Next() {
			return rd, nil
		}
	}
	return rd, err
}

// Make a base reader to sheet
func newBaseReaderByWorkSheetFile(cn *Connect, rc io.Reader) (*Read, error) {
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
			// fmt.Printf("%+v\n\n", t)
			switch token := t.(type) {
			case xml.StartElement:
				switch token.Name.Local {
				case SHEET_DATA:
					return
				default:
					decoder.Skip()
				}
			}
		}
	}(decoder)

	rd := &Read{
		connecter: cn,
		decoder:   decoder,
	}

	return rd, nil
}
