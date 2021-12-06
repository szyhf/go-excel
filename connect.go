package excel

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"io"
	"strings"
)

// connect is default implement of connector.
type connect struct {
	// list of sorted sheet name
	sheets            []string
	sharedStringPaths []string

	// xl/sharedStringPaths.xml
	sharedStringPathsFile *zip.File
	// xl/_rels/workbook.xml.rels
	workbookRels *zip.File
	// map["rId*"]"xl/path/to/target";
	// just support rels for workbook now.
	workbookRelsIDMap map[string]string
	// xl/workbook.xml
	workbookFile *zip.File
	// map["sheet_id"]"sheet_name"
	worksheetIDToNameMap map[string]string
	// "xl/worksheets/sheet*.xml"
	// map["xl/path/to/sheet*.xml"]*zip.File
	worksheetFileMap map[string]*zip.File
	// map["sheet_name"]*zip.File
	worksheetNameFileMap map[string]*zip.File
	worksheetNameList    []string

	// 实际的读取接口
	zipReader *zip.Reader
	// 仅读取文件时有效
	zipReaderCloser io.ReadCloser
}

// Open a excel file
func (conn *connect) Open(filePath string) error {
	var err error
	zipReaderCloser, err := zip.OpenReader(filePath)
	if err != nil {
		return err
	}
	conn.zipReader = &zipReaderCloser.Reader
	// prepare for files
	err = conn.init()
	if err != nil {
		if conn.zipReaderCloser != nil {
			conn.zipReaderCloser.Close()
			conn.zipReaderCloser = nil
		}
		conn.zipReader = nil
		return err
	}
	return nil
}

// OpenReader read a binary of xlsx file.
func (conn *connect) OpenBinary(xlsxData []byte) error {
	rd := bytes.NewReader(xlsxData)
	var err error
	conn.zipReader, err = zip.NewReader(rd, int64(rd.Len()))
	if err != nil {
		return err
	}
	// prepare for files
	err = conn.init()
	if err != nil {
		// 没有zipReader，不用Close
		return err
	}
	return nil
}

// Close file reader
func (conn *connect) Close() error {
	if conn.zipReaderCloser != nil {
		err := conn.zipReaderCloser.Close()
		if err != nil {
			return err
		}
		conn.zipReaderCloser = nil
	}
	conn.zipReader = nil

	conn.sheets = conn.sheets[:0]
	conn.sharedStringPaths = conn.sharedStringPaths[:0]
	conn.sharedStringPathsFile = nil
	conn.workbookFile = nil

	conn.worksheetFileMap = nil
	conn.worksheetNameFileMap = nil

	return nil
}

// NewReader generate an new reader of a sheet
// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
//             if sheetNamer is int, will i'th sheet in the workbook, be careful the hidden sheet is counted. i ∈ [1,+inf]
//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
//             otherwise, will use sheetNamer as struct and reflect for it's name.
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
	sheet := conn.parseSheetName(config.Sheet)
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

func (conn *connect) init() (err error) {
	// Find file of "workbook.xml", "sharedString.xml" and files in worksheets
	conn.worksheetFileMap = make(map[string]*zip.File)
	for _, f := range conn.zipReader.File {
		switch f.Name {
		case _SharedStringPath:
			conn.sharedStringPathsFile = f
		case _WorkBookPath:
			conn.workbookFile = f
		case _WorkBookRels:
			conn.workbookRels = f
		default:
			if strings.HasPrefix(f.Name, _WorkSheetsPrefix) {
				// log.Println("WorksheetName:", f.Name)
				conn.worksheetFileMap[f.Name] = f
			}
		}
	}

	if conn.workbookRels == nil {
		return ErrWorkbookRelsNotExist
	}
	if conn.workbookFile == nil {
		return ErrWorkbookNotExist
	}
	if conn.sharedStringPathsFile == nil {
		// sharedStringPathsFile may not exist.
		// return ErrSharedStringsNotExist
	}
	if conn.worksheetFileMap == nil || len(conn.worksheetFileMap) == 0 {
		return ErrWorkbookNotExist
	}
	// prepare workbook rels
	err = conn.readWorkbookRels()
	if err != nil {
		return errors.New("read workbook rels failed:" + err.Error())
	}
	// prepare workbook
	err = conn.readWorkbook()
	if err != nil {
		return errors.New("read workbook failed:" + err.Error())
	}
	// prepare sharedstring
	err = conn.readSharedString()
	if err != nil {
		return errors.New("read shared string failed:" + err.Error())
	}
	return nil
}

func (conn *connect) readWorkbookRels() error {
	rc, err := conn.workbookRels.Open()
	if err != nil {
		return err
	}
	wbRels, err := readWorkbookRelsXML(rc)
	if err != nil {
		rc.Close()
		return err
	}
	conn.workbookRelsIDMap = make(map[string]string, len(wbRels.Relationships))
	for _, rel := range wbRels.Relationships {
		if rel.Type == _RelTypeWorkSheet {
			// Is a rels for worksheet
			conn.workbookRelsIDMap[rel.ID] = _XL + rel.Target
		}
	}
	// log.Println(conn.workbookRelsIDMap)

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
	conn.worksheetNameFileMap = make(map[string]*zip.File, len(wb.Sheets.Sheet))
	conn.worksheetIDToNameMap = make(map[string]string, len(wb.Sheets.Sheet))
	for _, sheet := range wb.Sheets.Sheet {
		conn.sheets = append(conn.sheets, sheet.Name)
		// record the sheet name to *zip.File
		sheetTargetPath, ok := conn.workbookRelsIDMap[sheet.RID]
		if !ok {
			return fmt.Errorf("Sheet.RID = %s not exist", sheet.RID)
		}
		// log.Println(sheetTargetPath)
		file, ok := conn.worksheetFileMap[sheetTargetPath]
		if !ok {
			return fmt.Errorf("Sheet %s not exist", sheetTargetPath)
		}
		// log.Println(sheet.Name)
		conn.worksheetNameFileMap[sheet.Name] = file
		conn.worksheetIDToNameMap[sheet.SheetID] = sheet.Name
	}
	rc.Close()
	return nil
}

func (conn *connect) readSharedString() error {
	// sharedStringPathsFile may not exist
	if conn.sharedStringPathsFile == nil {
		conn.sharedStringPaths = make([]string, 0)
		return nil
	}
	rc, err := conn.sharedStringPathsFile.Open()
	if err != nil {
		return err
	}
	conn.sharedStringPaths = readSharedStringsXML(rc)
	rc.Close()
	return nil
}
