package internal

import (
	"archive/zip"
	"errors"
	"fmt"
	"strings"

	"github.com/szyhf/go-excel/internal/utils"
)

type Connect struct {
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
	worksheetNameList    []string
	zipReader            *zip.ReadCloser
}

func NewConnect() *Connect {
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

	this.sheets = this.sheets[:0]
	this.sharedStrings = this.sharedStrings[:0]
	this.sharedStringsFile = nil
	this.workbookFile = nil

	this.worksheetIDFileMap = nil
	this.worksheetNameFileMap = nil

	return nil
}

// Generate an new reader of a sheet
// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
//        if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
//        otherwise, will use sheetNamer as struct and reflect for it's name.
func (this *Connect) NewReader(sheetNamer interface{}) (Reader, error) {
	return this.NewReaderByConfig(&Config{Sheet: sheetNamer})
}

func (this *Connect) MustReader(sheetNamer interface{}) Reader {
	rd, err := this.NewReader(sheetNamer)
	if err != nil {
		panic(err)
	}
	return rd
}

func (this *Connect) NewReaderByConfig(config *Config) (Reader, error) {
	if this.zipReader == nil {
		return nil, ErrConnectNotOpened
	}
	sheet := utils.ParseSheetName(config.Sheet)
	sheet = config.Prefix + sheet + config.Suffix
	workSheetFile, ok := this.worksheetNameFileMap[sheet]
	if !ok {
		return nil, fmt.Errorf("can not find worksheet named = %s", sheet)
	}
	rc, err := workSheetFile.Open()
	if err != nil {
		return nil, err
	}
	reader, err := newReader(this, rc, config.TitleRowIndex, config.Skip)
	return reader, err
}

func (this *Connect) MustReaderByConfig(config *Config) Reader {
	rd, err := this.NewReaderByConfig(config)
	if err != nil {
		panic(err)
	}
	return rd
}

func (this *Connect) GetSheetNames() []string {
	if len(this.worksheetNameFileMap) != len(this.worksheetNameList) {
		this.worksheetNameList = make([]string, 0, len(this.worksheetNameFileMap))
		for name := range this.worksheetNameFileMap {
			this.worksheetNameList = append(this.worksheetNameList, name)
		}
	}
	dst := make([]string, len(this.worksheetNameList))
	copy(dst, this.worksheetNameList)
	return dst
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
		return ErrWorkbookNotExist
	}
	if this.sharedStringsFile == nil {
		return ErrSharedStringsNotExist
	}
	if this.worksheetIDFileMap == nil || len(this.worksheetIDFileMap) == 0 {
		return ErrWorkbookNotExist
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

	wb, err := readWorkbookXML(rc)
	if err != nil {
		rc.Close()
		return err
	}
	if this.sheets == nil {
		this.sheets = make([]string, 0, len(wb.Sheets.Sheet))
	}
	this.worksheetNameFileMap = make(map[string]*zip.File)
	for _, sheet := range wb.Sheets.Sheet {
		this.sheets = append(this.sheets, sheet.Name)
		// record the sheet name to *zip.File
		sheetID := strings.TrimLeft(sheet.ID, "rId")
		file, ok := this.worksheetIDFileMap[sheetID]
		if !ok {
			return fmt.Errorf("Sheet.ID = %s not exist.", sheetID)
		}
		this.worksheetNameFileMap[sheet.Name] = file
	}
	rc.Close()
	return nil
}
