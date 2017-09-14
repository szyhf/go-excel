package internal

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"

	"github.com/szyhf/go-convert"
	"github.com/szyhf/go-excel"
)

type Read struct {
	connecter *Connect
	config    *Config
	// currentRowToken xml.StartElement
	readCloser io.ReadCloser
	decoder    *xml.Decoder

	// The index of the Func Next()
	nextIndex int

	// map[$fieldName]$fileIndex
	fields map[string]int
}

func newReader(cn *Connect, sheetName string) (rdd excel.Reader, erro error) {
	workSheetFile, ok := cn.worksheetNameFileMap[sheetName]
	if !ok {
		return nil, excel.ErrWorksheetsNotExist
	}
	// 先不考虑标题行
	return newBaseReaderByWorkSheetFile(cn, workSheetFile)
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

// Read current row into a object
func (this *Read) Read(v interface{}) error {
	xlRow, err := newRowAsMap(this)
	fmt.Printf("%+v,,,,,,,,%+v\n", xlRow, err)
	// fmt.Printf("%+v\n", this.currentRow.srcMap)
	return nil
}

func (this *Read) Close() error {
	if this.readCloser != nil {
		err := this.readCloser.Close()
		if err != nil {
			return err
		}
	}
	if this.decoder != nil {
		this.decoder = nil
	}
	return nil
}

// Read all rows
func (this *Read) ReadAll(container interface{}) error {
	return nil
}

// Make a base reader to sheet
func newBaseReaderByWorkSheetFile(cn *Connect, workSheetFile *zip.File) (excel.Reader, error) {
	rc, err := workSheetFile.Open()
	if err != nil {
		return nil, err
	}

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
		config:     &Config{},
		readCloser: rc,
		decoder:    decoder,
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
