package excel_test

import (
	"reflect"
	"testing"

	convert "github.com/szyhf/go-convert"
	excel "github.com/szyhf/go-excel"
)

var expectAdvanceList = []Advance{
	{
		ID:      1,
		Name:    "Andy",
		NamePtr: strPtr("Andy"),
		Age:     1,
		Slice:   []int{1, 2},
		Temp: &Temp{
			Foo: "Andy",
		},
	},
	{
		ID:      2,
		Name:    "Leo",
		NamePtr: strPtr("Leo"),
		Age:     2,
		Slice:   []int{2, 3, 4},
		Temp: &Temp{
			Foo: "Leo",
		},
	},
	{
		ID:      3,
		Name:    "",
		NamePtr: strPtr("Ben"),
		Age:     180, //  using default
		Slice:   []int{3, 4, 5, 6},
		Temp: &Temp{
			Foo: "Ben",
		},
	},
	{
		ID:      4,
		Name:    "Ming",
		NamePtr: strPtr("Ming"),
		Age:     4,
		Slice:   []int{1},
		Temp: &Temp{
			Foo: "Default",
		},
	},
}

type Advance struct {
	// use field name as default column name
	ID int
	// column means to map the column name, and skip cell that value equal to "Ben"
	Name string `xlsx:"column(NameOf);nil(Ben);req();"`
	// you can map a column into more than one field
	NamePtr *string `xlsx:"column(NameOf);req();"`
	// omit `column` if only want to map to column name, it's equal to `column(AgeOf)`
	// use 180 as default if cell is empty.
	Age int `xlsx:"column(AgeOf);default(180);req();"`
	// split means to split the string into slice by the `|`
	Slice []int `xlsx:"split(|);req();"`
	// use default also can marshal to struct
	Temp *Temp `xlsx:"column(UnmarshalString);default({\"Foo\":\"Default\"});req();"`
	// use '-' to ignore.
	WantIgnored string `xlsx:"-"`
	// By default, required tag req is not set
	NotRequired string
}

func TestRead(t *testing.T) {
	// file
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
	}
	defer conn.Close()

	rd, err := conn.NewReaderByConfig(&excel.Config{
		// Sheet name as string or sheet model as object or as slice of objecg.
		Sheet: advSheetName,
		// Use the index row as title, every row before title-row will be ignore, default is 0.
		TitleRowIndex: 1,
		// Skip n row after title, default is 0 (not skip), empty row is not counted.
		Skip: 1,
		// Auto prefix to sheet name.
		Prefix: "",
		// Auto suffix to sheet name.
		Suffix: advSheetSuffix,
	})
	if err != nil {
		t.Error(err)
	}
	defer rd.Close()

	idx := 0
	for rd.Next() {
		var a Advance
		err := rd.Read(&a)
		if err != nil {
			t.Error(err)
			return
		}
		expect := expectAdvanceList[idx]
		if !reflect.DeepEqual(expect, a) {
			t.Errorf("unexpect advance at %d = \n%s", idx, convert.MustJsonPrettyString(a))
		}

		idx++
	}
}

func TestReadAll(t *testing.T) {
	// see the Advancd.suffix sheet in simple.xlsx
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
	}
	rd, err := conn.NewReaderByConfig(&excel.Config{
		// Sheet name as string or sheet model as object or a slice of object.
		Sheet: advSheetName,
		// Use the index row as title, every row before title-row will be ignore, default is 0.
		TitleRowIndex: 1,
		// Skip n row after title, default is 0 (not skip), empty row is not counted.
		Skip: 1,
		// Auto prefix to sheet name.
		Prefix: "",
		// Auto suffix to sheet name.
		Suffix: advSheetSuffix,
	})
	if err != nil {
		t.Error(err)
	}
	defer conn.Close()

	var slc []Advance
	err = rd.ReadAll(&slc)
	if err != nil {
		t.Error(err)
		return
	}
	if !reflect.DeepEqual(slc, expectAdvanceList) {
		t.Errorf("unexpect advance list: \n%s", convert.MustJsonPrettyString(slc))
	}

}
