package internal

type Connecter interface {
	// Open a file of excel
	Open(filePath string) error
	// Close file reader
	Close() error
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

type Reader interface {
	// Read current row into a object
	Read(v interface{}) error
	// Read all rows
	// container: container should be ptr to slice or array.
	ReadAll(container interface{}) error
	// Read next rows
	Next() bool
	// Close the reader
	Close() error
}
