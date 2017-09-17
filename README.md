# Intro | 简介

Expect to create a ORM-Like library to read or write relate-db-like excel easily.

---



## RoadMap | 开发计划

+ Read xlsx file and got the expect xml. √
+ Prepare the shared string xml. √
+ Get the correct sheetX.xml. √
+ Read a row of a sheet. √
+ Read a cell of a row, fix the empty cell. √
+ Fill string cell with value of shared string xml. √
+ Can set the column name row, default is the first row. √
+ Read a row to a struct by column name. √

## Example | 用例

Assume you have a xlsx file has a sheet named "Simple" and looks like below:

|ID|NameOf|age|Slice|Temp|
|-|-|-|-|-|
|1|Andy|15|1-2-3|{"Foo":"Bar"}|
|2|Leo||2||
||||||
|9|Ben|14|3||
|10|Ming|10|9-2-3||

---

See the `simple.xlsx` in `testdata`

---

So define a struct like this:

```go
type Simple struct {
	ID    int    // No tag if not needed.
	Name  string `xlsx:column(NameOf)`
	Age   int    `xlsx:default(0)`
	Slice []int  `xlsx:"split(-)"`
	Temp  Temp   // Will auto use the string in cell to Unmarshal to `Temp`
}

type Temp struct{
	Foo string
	Bar int
}
// Implement the `encoding.BinaryMarshaler`
func (this*Temp)UnmarshalBinary(d []data)error{
	return json.Unmarshal(d)
}

```

Then read the xlsx file will like this:

```go
conn := internal.NewConnect()
err := conn.Open("path/to/file.xlsx")
defer conn.Close()
if err != nil {
	panic(err)
}

rd, err := conn.NewReader("Simple")
if err != nil {
	panic(err)
}

for rd.Next() {
	var s Simple
	err := rd.Read(&s)
	if err != nil {
		panic(err)
	} else {
		fmt.Printf("%+v\n", s)
	}
}
```

Or just read to an slice:

```go
rd, err := conn.NewReader("Simple")
if err != nil {
	panic(err)
}
var slc []Simple
err = rd.ReadAll(&slc)
if err != nil {
	panic(err)
}
```

Use config for more optione:

```go
config := &excel.Config{
	// Sheet name as string or sheet model as object.
	Sheet: "Simple"
	// Use the index row as title, every row before title-row will be ignore, default is 0.
	TitleRowIndex: 0
	// Skip n row after title, default is 0 (not skip).
	Skip: 1
}
rd, err := conn.NewReaderByConfig(config)
if err != nil {
	panic(err)
}
// skip the usage.
rd.Next()
```

For more usage to read the `test` directory.

## Tag | 标签使用

### column

Map to field name in title row, by default will use the field name.

### default

Set default value when no value is filled in excel cell, by default is 0 or "".

### split

Split a string and convert them to a slice, it won't work if not set.

## Thinking | 随想

在复杂的系统中（例如游戏）

有时候为了便于非专业人员设置一些配置

会使用Excel作为一种轻量级的关系数据库或者配置文件

毕竟对于很多非开发人员来说

配个Excel要比写json或者yaml什么简单得多

这种场景下

读取特定格式（符合关系数据库特点的表格）的数据会比各种花式写入Excel的功能更重要

毕竟从编辑上来说微软提供的Excel本身功能就非常强大了

而现在我找到的Excel库的功能都过于强大了

用起来有点浪费

于是写了这个简化库

这个库的工作参考了[tealeg/xlsx](github.com/tealeg/xlsx)的部分实现和读取逻辑。

感谢[tealeg](github.com/tealeg)