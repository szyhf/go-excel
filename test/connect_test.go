package test

import (
	"fmt"
	"testing"

	"github.com/json-iterator/go"

	convert "github.com/szyhf/go-convert"
	"github.com/szyhf/go-excel/internal"
)

func TestOpen(t *testing.T) {
	// file
	conn := internal.NewConnect()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
	}
	err = conn.Close()
	if err != nil {
		t.Error(err)
	}
}

type Simple struct {
	ID    int      `xlsx:"column(ID)"`
	Name  string   `xlsx:"column(Slice)"`
	Slice *string  `xlsx:"column(C)"`
	Hello []string `xlsx:"column(Hello);split(|)"`
	Temp  *Temp    `xlsx:"Temp"`
}

type Temp struct {
	Foo string
}

func (this *Temp) UnmarshalBinary(d []byte) error {
	return jsoniter.Unmarshal(d, this)
}

func TestNewReader(t *testing.T) {
	// file
	conn := internal.NewConnect()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
	}
	rd, err := conn.NewReader(sheetName)
	if err != nil {
		t.Error(err)
	}
	conn.Close()
	for rd.Next() {
		var s Simple
		err := rd.Read(&s)
		if err != nil {
			t.Error(err)
		} else {
			fmt.Printf("%+v\n", convert.MustJsonPrettyString(s))
		}
	}
}
