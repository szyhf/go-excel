package excel_test

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"reflect"
	"testing"

	convert "github.com/szyhf/go-convert"
	excel "github.com/szyhf/go-excel"
)

func init() {
	log.SetFlags(log.Llongfile)
}

func strPtr(s string) *string {
	return &s
}

var expectStandardList = []Standard{
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

var expectStandardPtrList = []*Standard{
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

var expectStandardMapList = []map[string]string{
	map[string]string{
		"ID":              "1",
		"NameOf":          "Andy",
		"AgeOf":           "1",
		"Slice":           "1|2",
		"UnmarshalString": "{\"Foo\":\"Andy\"}",
	},
	map[string]string{
		"ID":              "2",
		"NameOf":          "Leo",
		"AgeOf":           "2",
		"Slice":           "2|3|4",
		"UnmarshalString": "{\"Foo\":\"Leo\"}",
	},
	map[string]string{
		"ID":              "3",
		"NameOf":          "Ben",
		"AgeOf":           "3",
		"Slice":           "3|4|5|6",
		"UnmarshalString": "{\"Foo\":\"Ben\"}",
	},
	map[string]string{
		"ID":              "4",
		"NameOf":          "Ming",
		"AgeOf":           "4",
		"Slice":           "1",
		"UnmarshalString": "{\"Foo\":\"Ming\"}",
	},
}

var expectStandardSliceList = [][]string{
	[]string{
		"1",
		"Andy",
		"1",
		"1|2",
		"{\"Foo\":\"Andy\"}",
	},
	[]string{
		"2",
		"Leo",
		"2",
		"2|3|4",
		"{\"Foo\":\"Leo\"}",
	},
	[]string{
		"3",
		"Ben",
		"3",
		"3|4|5|6",
		"{\"Foo\":\"Ben\"}",
	},
	[]string{
		"4",
		"Ming",
		"4",
		"1",
		"{\"Foo\":\"Ming\"}",
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

// func (this Standard) GetXLSXSheetName() string {
// 	return "Some sheet name if need"
// }

type Temp struct {
	Foo string
}

func (tmp *Temp) UnmarshalBinary(d []byte) error {
	return json.Unmarshal(d, tmp)
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
		var s Standard
		if err := rd.Read(&s); err != nil {
			fmt.Println(err)
			return
		}
		expectStd := expectStandardList[idx]
		if !reflect.DeepEqual(s, expectStd) {
			t.Errorf("unexpect std at %d = \n%s", idx, convert.MustJsonPrettyString(expectStd))
		}
		idx++
	}
}

func TestReadStandardIndex(t *testing.T) {
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
		var s Standard
		if err := rd.Read(&s); err != nil {
			fmt.Println(err)
			return
		}
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
	if !reflect.DeepEqual(expectStandardList, stdList) {
		t.Errorf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}
}

func TestReadStandardPtrSimple(t *testing.T) {
	var stdList []*Standard
	err := excel.UnmarshalXLSX(filePath, &stdList)
	if err != nil {
		t.Error(err)
		return
	}
	if !reflect.DeepEqual(stdList, expectStandardPtrList) {
		t.Errorf("unexprect std list: %s", convert.MustJsonPrettyString(stdList))
	}
}

func TestReadStandardPtrAll(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	var stdList []*Standard
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
	if !reflect.DeepEqual(expectStandardPtrList, stdList) {
		t.Errorf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}
}

func TestReadBinaryStandardPtrAll(t *testing.T) {
	xlsxData, err := ioutil.ReadFile(filePath)
	if err != nil {
		t.Error(err)
		return
	}

	conn := excel.NewConnecter()
	err = conn.OpenBinary(xlsxData)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	var stdList []*Standard
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
	if !reflect.DeepEqual(expectStandardPtrList, stdList) {
		t.Errorf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}
}

func TestReadStandardMap(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is int, will i'th sheet in the workbook, be careful the hidden sheet is counted. i ∈ [1,+inf]
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
		var m map[string]string
		if err := rd.Read(&m); err != nil {
			fmt.Println(err)
			return
		}

		expectStdMap := expectStandardMapList[idx]
		if !reflect.DeepEqual(m, expectStdMap) {
			t.Errorf("unexpect std at %d = \n%s", idx, convert.MustJsonPrettyString(expectStdMap))
		}
		idx++
	}
}

func TestReadStandardSliceMap(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is int, will i'th sheet in the workbook, be careful the hidden sheet is counted. i ∈ [1,+inf]
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	rd, err := conn.NewReader(stdSheetName)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	var stdMapList []map[string]string
	err = rd.ReadAll(&stdMapList)
	if err != nil {
		t.Error(err)
		return
	}

	if !reflect.DeepEqual(expectStandardMapList, stdMapList) {
		t.Errorf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdMapList))
	}
}

func TestReadStandardSlice(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is int, will i'th sheet in the workbook, be careful the hidden sheet is counted. i ∈ [1,+inf]
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
		var l []string
		if err := rd.Read(&l); err != nil {
			fmt.Println(err)
			return
		}

		expectStdList := expectStandardSliceList[idx]
		if !reflect.DeepEqual(l, expectStdList) {
			t.Errorf("unexpect std at %d %s = \n%s", idx, convert.MustJsonPrettyString(l), convert.MustJsonPrettyString(expectStdList))
		}
		idx++
	}
}

func TestReadStandardSliceList(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is int, will i'th sheet in the workbook, be careful the hidden sheet is counted. i ∈ [1,+inf]
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	rd, err := conn.NewReader(stdSheetName)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	var stdList [][]string
	err = rd.ReadAll(&stdList)
	if err != nil {
		t.Error(err)
		return
	}

	if !reflect.DeepEqual(expectStandardSliceList, stdList) {
		t.Errorf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}
}
