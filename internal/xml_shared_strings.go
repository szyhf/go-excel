package internal

import (
	"encoding/xml"
	"io"
)

var readSharedStringsBuff = make([]string, 0)

func readSharedStringsXML(rc io.ReadCloser) []string {
	decoder := xml.NewDecoder(rc)

	siStart, tStart := false, false
	slc := readSharedStringsBuff[:0]
	for t, err := decoder.Token(); err == nil; t, err = decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case "si":
				siStart = true
			case "t":
				tStart = true
			}
		case xml.EndElement:
			switch token.Name.Local {
			case "si":
				siStart = false
			case "t":
				tStart = false
			}
		case xml.CharData:
			if siStart {
				slc = append(slc, string(token))
			}
			if tStart {
				lastStr := slc[len(slc)-1]
				slc[len(slc)-1] = lastStr + string(token)
			}
		}
	}
	res := make([]string, len(slc))
	copy(res, slc)
	return res
}
