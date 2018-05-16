package excel

// xlsxC directly maps the c element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxC struct {
	R string `xml:"r,attr"`           // Cell ID, e.g. A1
	T string `xml:"t,attr,omitempty"` // Type.
	V string `xml:"v,omitempty"`      // Value
}
