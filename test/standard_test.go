package test

import (
	"encoding/json"
	"reflect"
	"testing"

	convert "github.com/szyhf/go-convert"
	excel "github.com/szyhf/go-excel"
)

func strPtr(s string) *string {
	return &s
}

var expectStandardList = []Standard{
	Standard{
		ID:      1,
		Name:    "Andy",
		NamePtr: strPtr("Andy"),
		Age:     1,
		Slice:   []int{1, 2},
		Temp: &Temp{
			Foo: "Andy",
		},
	},
	Standard{
		ID:      2,
		Name:    "Leo",
		NamePtr: strPtr("Leo"),
		Age:     2,
		Slice:   []int{2, 3, 4},
		Temp: &Temp{
			Foo: "Leo",
		},
	},
	Standard{
		ID:      3,
		Name:    "Ben",
		NamePtr: strPtr("Ben"),
		Age:     3,
		Slice:   []int{3, 4, 5, 6},
		Temp: &Temp{
			Foo: "Ben",
		},
	},
	Standard{
		ID:      4,
		Name:    "Ming",
		NamePtr: strPtr("Ming"),
		Age:     4,
		Slice:   []int{1},
		Temp: &Temp{
			Foo: "Ming",
		},
	},
}

// defined a struct
type Standard struct {
	// use field name as default column name
	ID int
	// column means to map the column name
	Name string `xlsx:"column(NameOf)"`
	// you can map a column into more than one field
	NamePtr *string `xlsx:"column(NameOf)"`
	// omit `column` if only want to map to column name, it's equal to `column(AgeOf)`
	Age int `xlsx:"AgeOf"`
	// split means to split the string into slice by the `|`
	Slice []int `xlsx:"split(|)"`
	Temp  *Temp `xlsx:"column(UnmarshalString)"`
	// use '-' to ignore.
	WantIgnored string `xlsx:"-"`
}

func (this Standard) GetXLSXSheetName() string {
	return "Some sheet name if need"
}

type Temp struct {
	Foo string
}

func (this *Temp) UnmarshalBinary(d []byte) error {
	return json.Unmarshal(d, this)
}

func TestReadStandardSimple(t *testing.T) {
	var stdList []Standard
	err := excel.UnmarshalXLSX(filePath, &stdList)
	if err != nil {
		t.Error(err)
		return
	}
	if !reflect.DeepEqual(stdList, expectStandardList) {
		t.Errorf("unexprect std list: %s", convert.MustJsonPrettyString(stdList))
	}
}

func TestReadStandard(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	rd, err := conn.NewReader(stdSheetName)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	idx := 0
	for rd.Next() {
		var s Standard
		rd.Read(&s)
		expectStd := expectStandardList[idx]
		if !reflect.DeepEqual(s, expectStd) {
			t.Errorf("unexpect std at %d = \n%s", idx, convert.MustJsonPrettyString(expectStd))
		}
		idx++
	}
}

func TestReadStandardAll(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	var stdList []Standard
	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	rd, err := conn.NewReader(stdList)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	err = rd.ReadAll(&stdList)
	if err != nil {
		t.Error(err)
		return
	}
	if !reflect.DeepEqual(expectStandardList, stdList) {
		t.Errorf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}
}
