package internal

import (
	"encoding/xml"
	"io"
)

// Ref: tealeg/xlsx/xmlWorkbook.go
// I just keep the part I want.

func readWorkbookXML(rc io.ReadCloser) (*xlsxWorkbook, error) {
	var err error
	workbook := new(xlsxWorkbook)
	decoder := xml.NewDecoder(rc)
	err = decoder.Decode(workbook)
	if err != nil {
		return nil, err
	}
	return workbook, nil
}

// xlsxWorkbook directly maps the workbook element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxWorkbook struct {
	Sheets xlsxSheets `xml:"sheets"`
}

// xlsxSheets directly maps the sheets element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxSheets struct {
	Sheet []xlsxSheet `xml:"sheet"`
}

// xlsxSheet directly maps the sheet element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxSheet struct {
	Name    string `xml:"name,attr,omitempty"`
	SheetID string `xml:"sheetId,attr,omitempty"`
}
