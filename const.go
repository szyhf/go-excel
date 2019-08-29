package excel

const (
	_XL = "xl/"
	// 所有的值汇总
	_SharedStringPath = "xl/sharedStrings.xml"
	// Resource(例如Sheet)和rId的关系定义
	_WorkBookRels = "xl/_rels/workbook.xml.rels"
	// 找个各个sheet的名字的地方
	_WorkBookPath = "xl/workbook.xml"
	// 各个工作表的数据
	_WorkSheetsPrefix = "xl/worksheets/sheet"
	// worksheet表里的数据字段起始
	_SheetData = "sheetData"
	// worksheet表里的行字段起始
	_RowPrefix = "row"

	_AllNumber = "0123456789"

	// xml
	_S           = "s"
	_SI          = "si"
	_T           = "t"
	_R           = "r"
	_SST         = "sst"
	_Count       = "count"
	_UniqueCount = "uniqueCount"
	_C           = "c"
	_V           = "v"
	// _RID         = "rId"

	// workbook.xml.rels表中描述worksheet类型的类型枚举
	_RelTypeWorkSheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)
