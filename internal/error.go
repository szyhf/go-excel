package internal

import (
	"errors"
)

var (
	ErrConnectNotOpened      = errors.New("connect should opened.")
	ErrWorkbookNotExist      = errors.New("parse xlsx file failed: xl/workbook.xml not exist.")
	ErrSharedStringsNotExist = errors.New("parse xlsx file failed: xl/sharedStrings.xml not exist.")
	ErrInvalidConatiner      = errors.New("container should be ptr to slice.")
	ErrEmptyRow              = errors.New("Empty row.")
)
