package internal

import (
	"encoding/xml"
	"io"

	"github.com/szyhf/go-convert"
)

func readSharedStringsXML(rc io.ReadCloser) []string {
	decoder := xml.NewDecoder(rc)

	tStart, rStart := false, false
	var slc []string
	siIndex := 0
	for t, err := decoder.Token(); err == nil; t, err = decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case "t":
				tStart = true
			case "r":
				rStart = true
			case "sst":
				for _, attr := range token.Attr {
					if attr.Name.Local == "uniqueCount" {
						slc = make([]string, convert.MustInt(attr.Value))
					}
				}
			default:
				decoder.Skip()
			}
		case xml.EndElement:
			switch token.Name.Local {
			case "si":
				siIndex++
			case "t":
				tStart = false
			case "r":
				rStart = false
			}
		case xml.CharData:
			if tStart {
				if rStart {
					str := slc[siIndex]
					slc[siIndex] = str + string(token)
				} else {
					slc[siIndex] = string(token)
				}
			}
		}
	}
	return slc
}
