package test

import (
	"encoding/json"
	"fmt"
	"testing"

	"github.com/szyhf/go-convert"
	excel "github.com/szyhf/go-excel"
)

func TestOpen(t *testing.T) {
	// file
	conn := excel.NewConnecter()
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
	return json.Unmarshal(d, this)
}

func TestNewReader(t *testing.T) {
	// file
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
	}
	_, err = conn.NewReader(sheetName)
	if err != nil {
		t.Error(err)
	}

	defer conn.Close()
}

func TestRead(t *testing.T) {
	// file
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	defer conn.Close()
	if err != nil {
		t.Error(err)
	}
	rd, err := conn.NewReader(sheetName)
	defer rd.Close()
	if err != nil {
		t.Error(err)
	}
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

func TestReadAll(t *testing.T) {
	// file
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
	}
	rd, err := conn.NewReader(sheetName)
	if err != nil {
		t.Error(err)
	}
	defer conn.Close()

	var slc []Simple
	err = rd.ReadAll(&slc)
	if err != nil {
		t.Error(err)
	}
	fmt.Printf("%s\n", convert.MustJsonPrettyString(slc))
}
