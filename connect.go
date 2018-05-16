package excel

import (
	"archive/zip"
	"errors"
	"fmt"
	"strings"
)

// connect is default implement of connector.
type connect struct {
	// list of sorted sheet name
	sheets            []string
	sharedStringPaths []string

	// xl/sharedStringPaths.xml
	sharedStringPathsFile *zip.File
	// xl/workbook.xml
	workbookFile *zip.File
	// "xl/worksheets/sheet*.xml" map[sheet*]*zip.File
	worksheetIDFileMap   map[string]*zip.File
	worksheetNameFileMap map[string]*zip.File
	worksheetNameList    []string
	zipReader            *zip.ReadCloser
}

// Open a excel file
func (conn *connect) Open(filePath string) error {
	var err error
	conn.zipReader, err = zip.OpenReader(filePath)
	if err != nil {
		return err
	}
	// prepare for files
	err = conn.init()
	if err != nil {
		conn.zipReader.Close()
		conn.zipReader = nil
		return err
	}
	return nil
}

// Close file reader
func (conn *connect) Close() error {
	err := conn.zipReader.Close()
	if err != nil {
		return err
	}
	conn.zipReader = nil

	conn.sheets = conn.sheets[:0]
	conn.sharedStringPaths = conn.sharedStringPaths[:0]
	conn.sharedStringPathsFile = nil
	conn.workbookFile = nil

	conn.worksheetIDFileMap = nil
	conn.worksheetNameFileMap = nil

	return nil
}

// NewReader generate an new reader of a sheet
// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
//        if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
//        otherwise, will use sheetNamer as struct and reflect for it's name.
func (conn *connect) NewReader(sheetNamer interface{}) (Reader, error) {
	return conn.NewReaderByConfig(&Config{Sheet: sheetNamer})
}

// MustReader will panic instead of return error
func (conn *connect) MustReader(sheetNamer interface{}) Reader {
	rd, err := conn.NewReader(sheetNamer)
	if err != nil {
		panic(err)
	}
	return rd
}

// NewReaderByConfig make a new reader by config
func (conn *connect) NewReaderByConfig(config *Config) (Reader, error) {
	if conn.zipReader == nil {
		return nil, ErrConnectNotOpened
	}
	sheet := parseSheetName(config.Sheet)
	sheet = config.Prefix + sheet + config.Suffix
	workSheetFile, ok := conn.worksheetNameFileMap[sheet]
	if !ok {
		return nil, fmt.Errorf("can not find worksheet named = %s", sheet)
	}
	rc, err := workSheetFile.Open()
	if err != nil {
		return nil, err
	}
	reader, err := newReader(conn, rc, config.TitleRowIndex, config.Skip)
	return reader, err
}

// MustReaderByConfig panic insead of return error
func (conn *connect) MustReaderByConfig(config *Config) Reader {
	rd, err := conn.NewReaderByConfig(config)
	if err != nil {
		panic(err)
	}
	return rd
}

// GetSheetNames return the sheet names.
func (conn *connect) GetSheetNames() []string {
	if len(conn.worksheetNameFileMap) != len(conn.worksheetNameList) {
		conn.worksheetNameList = make([]string, 0, len(conn.worksheetNameFileMap))
		for name := range conn.worksheetNameFileMap {
			conn.worksheetNameList = append(conn.worksheetNameList, name)
		}
	}
	dst := make([]string, len(conn.worksheetNameList))
	copy(dst, conn.worksheetNameList)
	return dst
}

func (conn *connect) getSharedString(id int) string {
	return conn.sharedStringPaths[id]
}

func (conn *connect) init() error {
	// Find file of "workbook.xml", "sharedString.xml" and files in worksheets
	conn.worksheetIDFileMap = make(map[string]*zip.File)
	for _, f := range conn.zipReader.File {
		switch f.Name {
		case _SharedStringPath:
			conn.sharedStringPathsFile = f
		case _WorkBookPath:
			conn.workbookFile = f
		default:
			if strings.HasPrefix(f.Name, _WorkSheetsPrefix) {
				// Trim left of prefix
				// Trim right of ".xml" as len = 4
				worksheetIDName := f.Name[len(_WorkSheetsPrefix) : len(f.Name)-4]
				// println("WorksheetName:", worksheetName)
				conn.worksheetIDFileMap[worksheetIDName] = f
			}
		}
	}
	if conn.workbookFile == nil {
		return ErrWorkbookNotExist
	}
	if conn.sharedStringPathsFile == nil {
		return ErrSharedStringsNotExist
	}
	if conn.worksheetIDFileMap == nil || len(conn.worksheetIDFileMap) == 0 {
		return ErrWorkbookNotExist
	}
	var err error
	// prepare workbook
	err = conn.readWorkbook()
	if err != nil {
		return errors.New("read workbook failed:" + err.Error())
	}
	// prepare sharedstring
	rc, err := conn.sharedStringPathsFile.Open()
	if err != nil {
		return err
	}
	conn.sharedStringPaths = readSharedStringsXML(rc)
	rc.Close()
	return nil
}

func (conn *connect) readWorkbook() error {
	// Find name of sheets
	rc, err := conn.workbookFile.Open()
	if err != nil {
		return err
	}

	wb, err := readWorkbookXML(rc)
	if err != nil {
		rc.Close()
		return err
	}
	if conn.sheets == nil {
		conn.sheets = make([]string, 0, len(wb.Sheets.Sheet))
	}
	conn.worksheetNameFileMap = make(map[string]*zip.File)
	for _, sheet := range wb.Sheets.Sheet {
		conn.sheets = append(conn.sheets, sheet.Name)
		// record the sheet name to *zip.File
		sheetID := strings.TrimLeft(sheet.ID, _RID)
		file, ok := conn.worksheetIDFileMap[sheetID]
		if !ok {
			return fmt.Errorf("Sheet.ID = %s not exist", sheetID)
		}
		conn.worksheetNameFileMap[sheet.Name] = file
	}
	rc.Close()
	return nil
}
