package excel

import "errors"

var (
	ErrConnectNotOpened      = errors.New("connect should opened.")
	ErrWorkbookNotExist      = errors.New("xl/workbook.xml not exist.")
	ErrWorksheetsNotExist    = errors.New("xl/worksheets/*.xml not exist.")
	ErrSharedStringsNotExist = errors.New("xl/sharedStrings.xml not exist.")
)
