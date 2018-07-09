package excel_test

import (
	"fmt"
	"io/ioutil"
	"reflect"

	convert "github.com/szyhf/go-convert"
	excel "github.com/szyhf/go-excel"
)

func ExampleUnmarshalXLSX_struct() {
	var stdList []Standard
	err := excel.UnmarshalXLSX(filePath, &stdList)
	if err != nil {
		fmt.Println(err.Error())
		return
	}
	if !reflect.DeepEqual(stdList, expectStandardList) {
		fmt.Printf("unexprect std list: %s", convert.MustJsonPrettyString(stdList))
	}

	fmt.Printf(convert.MustJsonString(stdList))

	// output:
	// [{"ID":1,"Name":"Andy","NamePtr":"Andy","Age":1,"Slice":[1,2],"Temp":{"Foo":"Andy"},"WantIgnored":""},{"ID":2,"Name":"Leo","NamePtr":"Leo","Age":2,"Slice":[2,3,4],"Temp":{"Foo":"Leo"},"WantIgnored":""},{"ID":3,"Name":"Ben","NamePtr":"Ben","Age":3,"Slice":[3,4,5,6],"Temp":{"Foo":"Ben"},"WantIgnored":""},{"ID":4,"Name":"Ming","NamePtr":"Ming","Age":4,"Slice":[1],"Temp":{"Foo":"Ming"},"WantIgnored":""}]
}

func ExampleReader_readStruct() {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		fmt.Println(err)
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
		fmt.Println(err)
		return
	}
	defer rd.Close()

	idx := 0
	for rd.Next() {
		var s Standard
		rd.Read(&s)
		expectStd := expectStandardList[idx]
		if !reflect.DeepEqual(s, expectStd) {
			fmt.Printf("unexpect std at %d = \n%s", idx, convert.MustJsonPrettyString(expectStd))
		}
		fmt.Printf("%d => %s\n", idx, convert.MustJsonString(s))
		idx++
	}

	// output:
	//0 => {"ID":1,"Name":"Andy","NamePtr":"Andy","Age":1,"Slice":[1,2],"Temp":{"Foo":"Andy"},"WantIgnored":""}
	// 1 => {"ID":2,"Name":"Leo","NamePtr":"Leo","Age":2,"Slice":[2,3,4],"Temp":{"Foo":"Leo"},"WantIgnored":""}
	// 2 => {"ID":3,"Name":"Ben","NamePtr":"Ben","Age":3,"Slice":[3,4,5,6],"Temp":{"Foo":"Ben"},"WantIgnored":""}
	// 3 => {"ID":4,"Name":"Ming","NamePtr":"Ming","Age":4,"Slice":[1],"Temp":{"Foo":"Ming"},"WantIgnored":""}
}

func ExampleReader_readAllSliceStruct() {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		fmt.Println(err)
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
		fmt.Println(err)
		return
	}
	defer rd.Close()

	err = rd.ReadAll(&stdList)
	if err != nil {
		fmt.Println(err)
		return
	}
	if !reflect.DeepEqual(expectStandardList, stdList) {
		fmt.Printf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}

	fmt.Println(convert.MustJsonString(stdList))

	// output:
	// [{"ID":1,"Name":"Andy","NamePtr":"Andy","Age":1,"Slice":[1,2],"Temp":{"Foo":"Andy"},"WantIgnored":""},{"ID":2,"Name":"Leo","NamePtr":"Leo","Age":2,"Slice":[2,3,4],"Temp":{"Foo":"Leo"},"WantIgnored":""},{"ID":3,"Name":"Ben","NamePtr":"Ben","Age":3,"Slice":[3,4,5,6],"Temp":{"Foo":"Ben"},"WantIgnored":""},{"ID":4,"Name":"Ming","NamePtr":"Ming","Age":4,"Slice":[1],"Temp":{"Foo":"Ming"},"WantIgnored":""}]
}

func ExampleUnmarshalXLSX_ptr() {
	var stdList []*Standard
	err := excel.UnmarshalXLSX(filePath, &stdList)
	if err != nil {
		fmt.Println(err)
		return
	}
	if !reflect.DeepEqual(stdList, expectStandardPtrList) {
		fmt.Printf("unexprect std list: %s", convert.MustJsonPrettyString(stdList))
	}
	fmt.Println(convert.MustJsonString(stdList))

	// output:
	// [{"ID":1,"Name":"Andy","NamePtr":"Andy","Age":1,"Slice":[1,2],"Temp":{"Foo":"Andy"},"WantIgnored":""},{"ID":2,"Name":"Leo","NamePtr":"Leo","Age":2,"Slice":[2,3,4],"Temp":{"Foo":"Leo"},"WantIgnored":""},{"ID":3,"Name":"Ben","NamePtr":"Ben","Age":3,"Slice":[3,4,5,6],"Temp":{"Foo":"Ben"},"WantIgnored":""},{"ID":4,"Name":"Ming","NamePtr":"Ming","Age":4,"Slice":[1],"Temp":{"Foo":"Ming"},"WantIgnored":""}]
}

func ExampleReader_readAllSlicePtr() {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer conn.Close()

	var stdList []*Standard
	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	rd, err := conn.NewReader(stdList)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer rd.Close()

	err = rd.ReadAll(&stdList)
	if err != nil {
		fmt.Println(err)
		return
	}
	if !reflect.DeepEqual(expectStandardPtrList, stdList) {
		fmt.Printf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}

	fmt.Println(convert.MustJsonString(stdList))

	// output:
	// [{"ID":1,"Name":"Andy","NamePtr":"Andy","Age":1,"Slice":[1,2],"Temp":{"Foo":"Andy"},"WantIgnored":""},{"ID":2,"Name":"Leo","NamePtr":"Leo","Age":2,"Slice":[2,3,4],"Temp":{"Foo":"Leo"},"WantIgnored":""},{"ID":3,"Name":"Ben","NamePtr":"Ben","Age":3,"Slice":[3,4,5,6],"Temp":{"Foo":"Ben"},"WantIgnored":""},{"ID":4,"Name":"Ming","NamePtr":"Ming","Age":4,"Slice":[1],"Temp":{"Foo":"Ming"},"WantIgnored":""}]
}

func ExampleReader_readBinaryAllSlicePtr() {
	xlsxData, err := ioutil.ReadFile(filePath)
	if err != nil {
		fmt.Println(err)
		return
	}

	conn := excel.NewConnecter()
	err = conn.OpenBinary(xlsxData)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer conn.Close()

	var stdList []*Standard
	// Generate an new reader of a sheet
	// sheetNamer: if sheetNamer is string, will use sheet as sheet name.
	//             if sheetNamer is a object implements `GetXLSXSheetName()string`, the return value will be used.
	//             otherwise, will use sheetNamer as struct and reflect for it's name.
	// 			   if sheetNamer is a slice, the type of element will be used to infer like before.
	rd, err := conn.NewReader(stdList)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer rd.Close()

	err = rd.ReadAll(&stdList)
	if err != nil {
		fmt.Println(err)
		return
	}
	if !reflect.DeepEqual(expectStandardPtrList, stdList) {
		fmt.Printf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdList))
	}

	fmt.Println(convert.MustJsonString(stdList))

	// output:
	// [{"ID":1,"Name":"Andy","NamePtr":"Andy","Age":1,"Slice":[1,2],"Temp":{"Foo":"Andy"},"WantIgnored":""},{"ID":2,"Name":"Leo","NamePtr":"Leo","Age":2,"Slice":[2,3,4],"Temp":{"Foo":"Leo"},"WantIgnored":""},{"ID":3,"Name":"Ben","NamePtr":"Ben","Age":3,"Slice":[3,4,5,6],"Temp":{"Foo":"Ben"},"WantIgnored":""},{"ID":4,"Name":"Ming","NamePtr":"Ming","Age":4,"Slice":[1],"Temp":{"Foo":"Ming"},"WantIgnored":""}]
}

func ExampleReader_readMap() {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		fmt.Println(err)
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
		fmt.Println(err)
		return
	}
	defer rd.Close()

	idx := 0
	for rd.Next() {
		var m map[string]string
		rd.Read(&m)

		expectStdMap := expectStandardMapList[idx]
		if !reflect.DeepEqual(m, expectStdMap) {
			fmt.Printf("unexpect std at %d = \n%s", idx, convert.MustJsonPrettyString(expectStdMap))
		}

		fmt.Printf("%d => %s\n", idx, convert.MustJsonString(m))
		idx++
	}

	// output:
	// 0 => {"A":"1","B":"Andy","C":"1","D":"1|2","E":"{\"Foo\":\"Andy\"}"}
	// 1 => {"A":"2","B":"Leo","C":"2","D":"2|3|4","E":"{\"Foo\":\"Leo\"}"}
	// 2 => {"A":"3","B":"Ben","C":"3","D":"3|4|5|6","E":"{\"Foo\":\"Ben\"}"}
	// 3 => {"A":"4","B":"Ming","C":"4","D":"1","E":"{\"Foo\":\"Ming\"}"}
}

func ExampleReader_readAllSliceMap() {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		fmt.Println(err)
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
		fmt.Println(err)
		return
	}
	defer rd.Close()

	var stdMapList []map[string]string
	err = rd.ReadAll(&stdMapList)
	if err != nil {
		fmt.Println(err)
		return
	}

	if !reflect.DeepEqual(expectStandardMapList, stdMapList) {
		fmt.Printf("unexpect stdlist: \n%s", convert.MustJsonPrettyString(stdMapList))
	}

	fmt.Println(convert.MustJsonString(stdMapList))

	// output:
	// [{"A":"1","B":"Andy","C":"1","D":"1|2","E":"{\"Foo\":\"Andy\"}"},{"A":"2","B":"Leo","C":"2","D":"2|3|4","E":"{\"Foo\":\"Leo\"}"},{"A":"3","B":"Ben","C":"3","D":"3|4|5|6","E":"{\"Foo\":\"Ben\"}"},{"A":"4","B":"Ming","C":"4","D":"1","E":"{\"Foo\":\"Ming\"}"}]
}

func ExampleReader_readAllSliceMapOtherValueType() {
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		fmt.Println(err)
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
		fmt.Println(err)
		return
	}
	defer rd.Close()

	// will fill with default value which cell can not unmarshal to int
	// int is just example, can be other type
	var stdMapList []map[string]int
	err = rd.ReadAll(&stdMapList)
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println(convert.MustJsonString(stdMapList))

	// output:
	// [{"A":1,"B":0,"C":1,"D":0,"E":0},{"A":2,"B":0,"C":2,"D":0,"E":0},{"A":3,"B":0,"C":3,"D":0,"E":0},{"A":4,"B":0,"C":4,"D":1,"E":0}]
}
