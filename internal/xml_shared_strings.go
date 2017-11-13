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
			case SI:
				// don't enter default ...
			case T:
				tStart = true
			case R:
				rStart = true
			case SST:
				count := 0
				unqCount := 0
				for _, attr := range token.Attr {
					switch attr.Name.Local {
					case COUNT:
						count = convert.MustInt(attr.Value)
					case UNIQUE_COUNT:
						unqCount = convert.MustInt(attr.Value)
					}
				}
				if unqCount != 0 {
					slc = make([]string, unqCount)
				} else {
					slc = make([]string, count)
				}
			default:
				decoder.Skip()
			}
		case xml.EndElement:
			switch token.Name.Local {
			case SI:
				siIndex++
			case T:
				tStart = false
			case R:
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
