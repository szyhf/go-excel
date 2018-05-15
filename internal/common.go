package internal

// Config of connecter
type Config struct {
	// sheet: if sheet is string, will use sheet as sheet name.
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
