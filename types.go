package excel

import (
	"github.com/szyhf/go-excel/internal"
)

// External of excel reader
type Reader = internal.Reader

// External of excel config
type Config = internal.Config

// An Connecter of excel file
type Connecter interface {
	// Open a file of excel
	Open(filePath string) error
	// Close file reader
	Close() error

	// Get all sheets name
	GetSheetNames() []string

	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	NewReader(sheetNamer interface{}) (Reader, error)
	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	MustReader(sheetNamer interface{}) Reader

	NewReaderByConfig(config *Config) (Reader, error)
	MustReaderByConfig(config *Config) Reader
}
