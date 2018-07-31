package excel

// Config of connecter
type Config struct {
	// sheet: if sheet is string, will use sheet as sheet name.
	//        if sheet is int, will i'th sheet in the workbook, be careful the hidden sheet is counted. i ∈ [1,+inf]
	//        if sheet is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//        otherwise, will use sheet as struct and reflect for it's name.
	// 		  if sheet is a slice, the type of element will be used to infer like before.
	Sheet interface{}
	// Use the index row as title, every row before title-row will be ignore, default is 0.
	TitleRowIndex int
	// Skip n row after title, default is 0 (not skip), empty row is not counted.
	Skip int
	// Auto prefix to sheet name.
	Prefix string
	// Auto suffix to sheet name.
	Suffix string
}

// Reader to read excel
type Reader interface {
	// Get all titles sorted
	GetTitles() []string
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

// An Connecter of excel file
type Connecter interface {
	// Open a file of excel
	Open(filePath string) error
	// Open a binary of excel
	OpenBinary(xlsxData []byte) error

	// Close file reader
	Close() error

	// Get all sheets name
	GetSheetNames() []string

	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is int, will i'th sheet in the workbook, be careful the hidden sheet is counted. i ∈ [1,+inf]
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	// 	           if sheetNamer is a slice, the type of element will be used to infer like before.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	NewReader(sheetNamer interface{}) (Reader, error)
	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is int, will i'th sheet in the workbook, be careful the hidden sheet is counted. i ∈ [1,+inf]
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	MustReader(sheetNamer interface{}) Reader

	NewReaderByConfig(config *Config) (Reader, error)
	MustReaderByConfig(config *Config) Reader
}
