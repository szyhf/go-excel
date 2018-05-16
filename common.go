package excel

// NewConnecter make a new connecter to connect to a exist xlsx file.
func NewConnecter() Connecter {
	return &connect{}
}

// UnmarshalXLSX unmarshal a sheet of XLSX file into a slice container.
// The sheet name will be inferred from element of container
// If container implement the function of GetXLSXSheetName()string, the return string will used.
// Oterwise will use the reflect struct name.
func UnmarshalXLSX(filePath string, container interface{}) error {
	conn := NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		return err
	}

	rd, err := conn.NewReader(container)
	if err != nil {
		conn.Close()
		return err
	}

	err = rd.ReadAll(container)
	if err != nil {
		conn.Close()
		rd.Close()
		return err
	}
	conn.Close()
	rd.Close()
	return nil
}
