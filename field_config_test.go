package excel_test

import (
	"fmt"
	"reflect"
	"testing"

	convert "github.com/szyhf/go-convert"
	excel "github.com/szyhf/go-excel"
)

type StandardFieldConfig Standard

func (StandardFieldConfig) GetXLSXSheetName() string {
	return "Standard"
}

func (StandardFieldConfig) GetXLSXFieldConfigs() map[string]excel.FieldConfig {
	return map[string]excel.FieldConfig{
		"Name": {
			ColumnName: "NameOf",
		},
		"NameOf": {
			ColumnName: "NameOf",
		},
		"Age": {
			ColumnName: "AgeOf",
		},
		"Slice": {
			Split: "|",
		},
		"Temp": {
			ColumnName: "UnmarshalString",
		},
		"WantIgnored": {
			Ignore: true,
		},
	}
}

var expectStandardFieldConfigList = []StandardFieldConfig{
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
		Name:    "Ben",
		NamePtr: strPtr("Ben"),
		Age:     3,
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
			Foo: "Ming",
		},
	},
}

func TestReadStandardFieldConfigSimple(t *testing.T) {
	var stdList []StandardFieldConfig
	err := excel.UnmarshalXLSX(filePath, &stdList)
	if err != nil {
		t.Error(err)
		return
	}
	if !reflect.DeepEqual(stdList, expectStandardFieldConfigList) {
		t.Errorf("unexprect std list: %s", convert.MustJsonPrettyString(stdList))
	}
}

func TestReadStandardFieldConfig(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is int, will use sheet as sheet name.
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
		var s StandardFieldConfig
		if err := rd.Read(&s); err != nil {
			fmt.Println(err)
			return
		}
		expectStd := expectStandardFieldConfigList[idx]
		if !reflect.DeepEqual(s, expectStd) {
			t.Errorf("unexpect std at %d = \n%s", idx, convert.MustJsonPrettyString(expectStd))
		}
		idx++
	}
}

func TestReadStandardFieldConfigIndex(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is int, will use sheet as sheet name.
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	rd, err := conn.NewReader(2)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	idx := 0
	for rd.Next() {
		var s StandardFieldConfig
		if err := rd.Read(&s); err != nil {
			fmt.Println(err)
			return
		}
		expectStd := expectStandardFieldConfigList[idx]
		if !reflect.DeepEqual(s, expectStd) {
			t.Errorf("unexpect std at %d = \n%s", idx, convert.MustJsonPrettyString(expectStd))
		}
		idx++
	}
}

func TestReadStandardFieldConfigAll(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	var stdList []StandardFieldConfig
	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is int, will i'th sheet in the workbook, be careful the hidden sheet is counted. i ∈ [1,+inf]
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
	if !reflect.DeepEqual(expectStandardFieldConfigList, stdList) {
		t.Errorf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}
}
