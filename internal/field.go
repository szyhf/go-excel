package internal

import (
	"reflect"
	"strings"

	"github.com/szyhf/go-convert"
)

const tagIdentify = "xlsx"
const tagSplit = ";"

const columnTag = "column"
const splitTag = "split"
const defaultTag = "default"
const indexTag = "index"
const ignoreTag = "-"

type Field struct {
	FieldIndex   int
	ColumnIndex  int
	ColumnName   string
	DefaultValue string
	Split        string
}

type Schema struct {
	// map[FieldIndex]*Field
	Fields []*Field
}

func newSchema(t reflect.Type) *Schema {

	s := &Schema{
		Fields: make([]*Field, 0, t.NumField()),
	}
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		if value, ok := field.Tag.Lookup(tagIdentify); ok {
			if value != ignoreTag {
				fieldCnf := praseTagValue(value)
				fieldCnf.FieldIndex = i
				s.Fields = append(s.Fields, fieldCnf)
			}
		}
	}
	return s
}

func praseTagValue(v string) *Field {
	c := &Field{}
	params := strings.Split(v, tagSplit)

	for _, param := range params {
		if param == "" {
			continue
		}
		cnfKey, cnfVal := getTagParam(param)
		fillField(c, cnfKey, cnfVal)
	}
	// with more params
	return c
}

func getTagParam(v string) (key, value string) {
	// expect v = `field_name` or `column(fieldName)` or `default(0)` and so on ...
	start := strings.Index(v, "(")
	end := strings.Index(v, ")")
	if start > 0 && end == len(v)-1 {
		return v[:start], v[start+1 : end]
	} else {
		// fmt.Printf("Use column as default?[%s]\n", v)
		return columnTag, v
	}
}

func fillField(c *Field, k, v string) {
	switch k {
	case columnTag:
		c.ColumnName = v
	case defaultTag:
		c.DefaultValue = v
	case splitTag:
		c.Split = v
	case indexTag:
		c.ColumnIndex = convert.MustInt(v)
	}
}
