package internal

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
