package internal

import (
	"encoding/xml"
	"fmt"
	"io"
	"reflect"

	"github.com/szyhf/go-convert"
	"github.com/szyhf/go-excel"
)

type Read struct {
	connecter *Connect
	config    *Config
	decoder   *xml.Decoder

	// The index of the Func Next()
	nextIndex int

	// map[$fieldName]$fileIndex
	fields map[string]int
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
	if t.Kind() != reflect.Ptr {
		return fmt.Errorf("%T should be pointer.", i)
	}
	t = t.Elem()
	v := reflect.ValueOf(i)
	if v.IsNil() {
		v.Set(reflect.New(t))
	}
	v = v.Elem()

	row, err := newRowAsMap(this)
	if err != nil {
		return err
	}
	for index := 0; index < t.NumField(); index++ {
		// fieldType := t.Field(index)
		fieldValue := v.Field(index)
		if fieldValue.Kind() == reflect.Ptr {
			fieldValue = fieldValue.Elem()
		}
		varStr := row.ValueAtColumn(index)
		Scan(varStr, fieldValue.Addr().Interface())
	}
	return nil
}

func (this *Read) Close() error {
	if this.decoder != nil {
		this.decoder = nil
	}
	return nil
}

// Read all rows
func (this *Read) ReadAll(container interface{}) error {
	return nil
}

func newReader(cn *Connect, workSheetFileReader io.Reader) (rdd excel.Reader, erro error) {
	// 先不考虑标题行
	return newBaseReaderByWorkSheetFile(cn, workSheetFileReader)
}

// Make a base reader to sheet
func newBaseReaderByWorkSheetFile(cn *Connect, rc io.Reader) (excel.Reader, error) {
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
		// TODO
		config:  &Config{},
		decoder: decoder,
	}

	return rd, nil
}

// 根据标题行推算有多少列
// 读取tag的配置，生成各个列的读取配置
func readRowStart(token xml.StartElement, column *Column) {
	if len(token.Attr) <= 0 {
		return
	}

}

// 确定当前row token在第几行
func getRowIndexOfToken(rowToken xml.StartElement) int {
	if len(rowToken.Attr) <= 0 {
		panic("unexpect sheet.xml structure.")
	}
	for _, attr := range rowToken.Attr {
		if attr.Name.Local == "r" {
			return convert.MustInt(attr.Value)
		}
	}
	panic("unexpect sheet.xml structure.")
}
