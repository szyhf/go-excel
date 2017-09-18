package internal

type Connecter interface {
	// Open a file of excel
	Open(filePath string) error
	// Close file reader
	Close() error
	// Generate an new reader of a sheet
	NewReader(sheetNamer interface{}) (Reader, error)
	// Panic if can not generate a new reader
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
