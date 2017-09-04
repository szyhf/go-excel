package excel

type Connecter interface {
	// Open a file of excel
	Open(filePath string) error
	// Close file reader
	Close() error
	// Generate an new reader of a sheet
	NewReader(sheet string) (Reader, error)
	// Panic if can not generate a new reader
	MustReader(sheet string) Reader
}

type Reader interface {
	// Read current row
	Read(v interface{}) error
	// Read as rows[start:stop]
	ReadBetween(start, stop int, container interface{}) error
	// Read all rows
	ReadAll(container interface{}) error
	// Read next rows
	Next() bool

	// GetBool(index int) bool
	// GetString(index int) string

}
