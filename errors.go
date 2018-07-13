package excel

import "errors"

var (
	// ErrConnectNotOpened means can not open connect to excel.
	ErrConnectNotOpened = errors.New("connect should opened")
	// ErrWorkbookRelsNotExist means can not found the workbook rels of excel.
	ErrWorkbookRelsNotExist = errors.New("parse xlsx file failed: xl/_rels/workbook.xml.rels not exist")
	// ErrWorkbookNotExist means can not found the workbook of excel.
	ErrWorkbookNotExist = errors.New("parse xlsx file failed: xl/workbook.xml not exist")
	// ErrSharedStringsNotExist means can not found the shared of excel.
	ErrSharedStringsNotExist = errors.New("parse xlsx file failed: xl/sharedStringPaths.xml not exist")
	// ErrInvalidConatiner means can not using the container.
	ErrInvalidConatiner = errors.New("container should be ptr to slice")
	// ErrEmptyRow means the row is empty.
	ErrEmptyRow = errors.New("empty row")
	// ErrNoRow means there is no row.
	ErrNoRow = errors.New("no row")
	// ErrScanNil means scan nil.
	ErrScanNil = errors.New("scan(nil)")
)
