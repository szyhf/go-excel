package internal

import (
	"archive/zip"
	"errors"
	"strings"

	"github.com/szyhf/go-excel"
)

type Connect struct {
	FilePath string

	// list of sorted sheet name
	sheets        []string
	sharedStrings []string

	// xl/sharedStrings.xml
	sharedStringsFile *zip.File
	// xl/workbook.xml
	workbookFile *zip.File
	// "xl/worksheets/sheet*.xml" map[sheet*]*zip.File
	worksheetIDFileMap   map[string]*zip.File
	worksheetNameFileMap map[string]*zip.File
	zipReader            *zip.ReadCloser
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
	// prepare for files
	err = this.init()
	if err != nil {
		this.zipReader.Close()
		this.zipReader = nil
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
		return nil, excel.ErrConnectNotOpened
	}
	workSheetFile, ok := this.worksheetNameFileMap[sheet]
	if !ok {
		return nil, excel.ErrWorksheetsNotExist
	}
	rc, err := workSheetFile.Open()
	if err != nil {
		return nil, err
	}
	reader, err := newReader(this, rc)
	rc.Close()
	return reader, err
}

func (this *Connect) MustReader(sheet string) excel.Reader {
	rd, err := this.NewReader(sheet)
	if err != nil {
		panic(err)
	}
	return rd
}

func (this *Connect) getSharedString(id int) string {
	return this.sharedStrings[id]
}

func (this *Connect) init() error {
	// Find file of "workbook.xml", "sharedStrings.xml" and files in worksheets
	this.worksheetIDFileMap = make(map[string]*zip.File)
	for _, f := range this.zipReader.File {
		switch f.Name {
		case SHARED_STRING:
			this.sharedStringsFile = f
		case WORK_BOOK:
			this.workbookFile = f
		default:
			if strings.HasPrefix(f.Name, WORK_SHEETS_PREFIX) {
				// Trim left of prefix
				// Trim right of ".xml" as len = 4
				worksheetIDName := f.Name[len(WORK_SHEETS_PREFIX) : len(f.Name)-4]
				// println("WorksheetName:", worksheetName)
				this.worksheetIDFileMap[worksheetIDName] = f
			}
		}
	}
	if this.workbookFile == nil {
		return excel.ErrWorkbookNotExist
	}
	if this.sharedStringsFile == nil {
		return excel.ErrSharedStringsNotExist
	}
	if this.worksheetIDFileMap == nil || len(this.worksheetIDFileMap) == 0 {
		return excel.ErrWorkbookNotExist
	}
	var err error
	// prepare workbook
	err = this.readWorkbook()
	if err != nil {
		return errors.New("read workbook failed:" + err.Error())
	}
	// prepare sharedstring
	rc, err := this.sharedStringsFile.Open()
	if err != nil {
		return err
	}
	this.sharedStrings = readSharedStringsXML(rc)
	rc.Close()
	return nil
}

func (this *Connect) readWorkbook() error {
	// Find name of sheets
	rc, err := this.workbookFile.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	wb, err := readWorkbookXML(rc)
	if err != nil {
		return err
	}
	this.sheets = make([]string, 0, len(wb.Sheets.Sheet))
	this.worksheetNameFileMap = make(map[string]*zip.File)
	for _, sheet := range wb.Sheets.Sheet {
		this.sheets = append(this.sheets, sheet.Name)
		// record the sheet name to *zip.File
		file := this.worksheetIDFileMap[sheet.SheetID]
		this.worksheetNameFileMap[sheet.Name] = file
	}

	return nil
}
