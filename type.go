package excel

type Connecter interface {
	Open(file string)
	NewReader(sheet string) Reader
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
