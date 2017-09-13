package internal

import (
	"encoding/xml"
	"io"
)

var readSharedStringsBuff = make([]string, 0)

func readSharedStringsXML(rc io.ReadCloser) []string {
	decoder := xml.NewDecoder(rc)
	tStart := false
	slc := readSharedStringsBuff[:0]
	for t, err := decoder.Token(); err == nil; t, err = decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			if token.Name.Local == "t" {
				tStart = true
				// println(token.Name.Local)
			}
		case xml.EndElement:
			if token.Name.Local == "t" {
				tStart = false
				// println(token.Name.Local)
			}
		case xml.CharData:
			if tStart {
				slc = append(slc, string(token))
				// println(string(token))
			}
		}
	}
	res := make([]string, len(slc))
	copy(res, slc)
	return res
}
