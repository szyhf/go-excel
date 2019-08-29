package excel

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
			case _SI:
				// don't enter default ...
			case _T:
				tStart = true
			case _R:
				rStart = true
			case _SST:
				count := 0
				unqCount := 0
				for _, attr := range token.Attr {
					switch attr.Name.Local {
					case _Count:
						count = convert.MustInt(attr.Value)
					case _UniqueCount:
						unqCount = convert.MustInt(attr.Value)
					}
				}
				if unqCount != 0 {
					slc = make([]string, unqCount)
				} else {
					slc = make([]string, count)
				}
			default:
				_ = decoder.Skip()
			}
		case xml.EndElement:
			switch token.Name.Local {
			case _SI:
				siIndex++
			case _T:
				tStart = false
			case _R:
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
