package excel_test

import (
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

type DuplicatedTitle struct {
	DuplicatedTitle string
}

var expectDuplicatedTitleSliceList = [][]string{
	[]string{
		"Value1",
		"EmptyTitleValue1",
		"Value2",
	},
	[]string{
		"Value3",
		"EmptyTitleValue2",
		"Value2",
	},
}

// func (this DuplicatedTitle) GetXLSXSheetName() string {
// 	return "Some sheet name if need"
// }

func TestReadDuplicatedTitles(t *testing.T) {
	var stdList []DuplicatedTitle
	// expect ErrDuplicatedTitles
	err := excel.UnmarshalXLSX(filePath, &stdList)
	if err != excel.ErrDuplicatedTitles {
		t.Errorf("expect ErrDuplicatedTitles but got: %+v", err)
		return
	}
}

func TestReadDuplicatedTitle(t *testing.T) {
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
	rd, err := conn.NewReader(dupSheetName)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	for rd.Next() {
		var s DuplicatedTitle
		if err := rd.Read(&s); err != excel.ErrDuplicatedTitles {
			t.Errorf("expect ErrDuplicatedTitles but got: %+v", err)
			return
		}
	}
}

func TestReadDuplicatedTitleIndex(t *testing.T) {
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
	rd, err := conn.NewReader(3)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	for rd.Next() {
		var s DuplicatedTitle
		if err := rd.Read(&s); err != excel.ErrDuplicatedTitles {
			t.Errorf("expect ErrDuplicatedTitles but got: %+v", err)
			return
		}
	}
}

func TestReadDuplicatedTitleAll(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	var stdList []DuplicatedTitle
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
	if err != excel.ErrDuplicatedTitles {
		t.Errorf("expect ErrDuplicatedTitles but got: %+v", err)
		return
	}
}

func TestReadDuplicatedTitlePtrSimple(t *testing.T) {
	var stdList []*DuplicatedTitle
	err := excel.UnmarshalXLSX(filePath, &stdList)
	if err != excel.ErrDuplicatedTitles {
		t.Errorf("expect ErrDuplicatedTitles but got: %+v", err)
		return
	}
}

func TestReadDuplicatedTitlePtrAll(t *testing.T) {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
		return
	}
	defer conn.Close()

	var stdList []*DuplicatedTitle
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
	if err != excel.ErrDuplicatedTitles {
		t.Errorf("expect ErrDuplicatedTitles but got: %+v", err)
		return
	}
}

func TestReadBinaryDuplicatedTitlePtrAll(t *testing.T) {
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

	var stdList []*DuplicatedTitle
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
	if err != excel.ErrDuplicatedTitles {
		t.Errorf("expect ErrDuplicatedTitles but got: %+v", err)
		return
	}
}

func TestReadDuplicatedTitleMap(t *testing.T) {
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
	rd, err := conn.NewReader(dupSheetName)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	for rd.Next() {
		var m map[string]string
		if err := rd.Read(&m); err != excel.ErrDuplicatedTitles {
			t.Errorf("expect ErrDuplicatedTitles but got: %+v", err)
			return
		}
	}
}

func TestReadDuplicatedTitleSliceMap(t *testing.T) {
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
	rd, err := conn.NewReader(dupSheetName)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	var stdMapList []map[string]string
	err = rd.ReadAll(&stdMapList)
	if err != excel.ErrDuplicatedTitles {
		t.Errorf("expect ErrDuplicatedTitles but got: %+v", err)
		return
	}
}

func TestReadDuplicatedTitleSlice(t *testing.T) {
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
	rd, err := conn.NewReader(dupSheetName)
	if err != nil {
		t.Error(err)
		return
	}
	defer rd.Close()

	idx := 0
	for rd.Next() {
		// in duplicated title case, slice is allowed
		var l []string
		if err := rd.Read(&l); err != nil {
			fmt.Println(err)
			return
		}

		expectStdList := expectDuplicatedTitleSliceList[idx]
		if !reflect.DeepEqual(l, expectStdList) {
			t.Errorf("unexpect std at %d %s = \n%s", idx, convert.MustJsonPrettyString(l), convert.MustJsonPrettyString(expectStdList))
		}
		idx++
	}
}

func TestReadDuplicatedTitleSliceList(t *testing.T) {
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
	rd, err := conn.NewReader(dupSheetName)
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

	if !reflect.DeepEqual(expectDuplicatedTitleSliceList, stdList) {
		t.Errorf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}
}
