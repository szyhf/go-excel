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

type Temp struct {
	Foo string
}

func (this *Temp) UnmarshalBinary(d []byte) error {
	return json.Unmarshal(d, this)
}
func TestReadStandard(t *testing.T) {
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
