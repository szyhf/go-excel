package internal

import (
	"archive/zip"

	"github.com/szyhf/go-excel"
)

const (
	// 所有的值汇总
	SHARED_STRING = "xl/sharedStrings.xml"
	// 找个各个sheet的名字的地方
	WORK_BOOK = "xl/workbook.xml"
	// 各个工作表的数据
	WORK_SHEETS = "xl/worksheets/sheet*.xml"
)

type Connect struct {
	FilePath string

	// xl/sharedStrings.xml
	sharedStringsFile *zip.File
	// xl/workbook.xml
	workBookFile *zip.File
	// "xl/worksheets/sheet*.xml" map[sheet*]*zip.File
	sheetFileMap map[string]*zip.File
	zipReader    *zip.ReadCloser
}

func NewConnect() excel.Connecter {
	return &Connect{}
}

// Open a excel file
func (this *Connect) Open(filePath string) error {
	var err error
	this.zipReader, err = zip.OpenReader(filePath)
	if err != nil {
		return err
	}
	return nil
}

// Close file reader
func (this *Connect) Close() error {
	err := this.zipReader.Close()
	if err != nil {
		return err
	}
	this.zipReader = nil
	return nil
}

// Generate an new reader of a sheet
func (this *Connect) NewReader(sheet string) (excel.Reader, error) {
	if this.zipReader == nil {
		return nil, ErrConnectNotOpened
	}
	return nil, nil
}

func (this *Connect) MustReader(sheet string) excel.Reader {
	rd, err := this.NewReader(sheet)
	if err != nil {
		panic(err)
	}
	return rd
}

func (this *Connect) init() {
	// Find file of "workbook.xml", "sharedStrings.xml" and files in worksheets
}
