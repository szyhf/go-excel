package excel

import (
	"encoding/xml"
	"io"
)

// Ref: tealeg/xlsx/xmlWorkbook.go
// I just keep the part I want.

func readWorkbookXML(rd io.Reader) (*xlsxWorkbook, error) {
	var err error
	workbook := new(xlsxWorkbook)
	decoder := xml.NewDecoder(rd)
	err = decoder.Decode(workbook)
	if err != nil {
		return nil, err
	}
	return workbook, nil
}

func readWorkbookRelsXML(rd io.Reader) (*xlsxWorkbookRels, error) {
	var err error
	workbookRels := new(xlsxWorkbookRels)
	decoder := xml.NewDecoder(rd)
	err = decoder.Decode(workbookRels)
	if err != nil {
		return nil, err
	}
	return workbookRels, nil
}

// xmlxWorkbookRels contains xmlxWorkbookRelations
// which maps sheet id and sheet XML
type xlsxWorkbookRels struct {
	XMLName       xml.Name               `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xlsxWorkbookRelation `xml:"Relationship"`
}

// xmlxWorkbookRelation maps sheet id and xl/worksheets/sheet%d.xml
type xlsxWorkbookRelation struct {
	ID     string `xml:"Id,attr"`
	Target string `xml:",attr"`
	Type   string `xml:",attr"`
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
	RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}
