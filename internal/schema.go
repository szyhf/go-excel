package internal

import (
	"reflect"
	"strings"
)

const (
	tagIdentify = "xlsx"
	tagSplit    = ";"

	columnTag  = "column"
	splitTag   = "split"
	defaultTag = "default"
	nilTag     = "nil"
	ignoreTag  = "-"
	reqTag     = "req"
)

type FieldConfig struct {
	FieldIndex int
	// use ptr in order to know if configed.
	ColumnName   string
	DefaultValue string
	Split        string
	// if cell.value == NilValue, will skip this scan
	NilValue string
	// panic if reuqired this column but not set
	IsRequired bool
}

func (this *FieldConfig) Scan(valStr string, fieldValue reflect.Value) error {
	if this.NilValue == valStr {
		// fmt.Printf("Got nil,skip")
		return nil
	}
	var err error
	switch fieldValue.Kind() {
	case reflect.Slice, reflect.Array:
		if len(this.Split) != 0 && len(valStr) > 0 {
			// use split
			elems := strings.Split(valStr, this.Split)
			fieldValue.Set(reflect.MakeSlice(fieldValue.Type(), 0, len(elems)))
			err = ScanSlice(elems, fieldValue.Addr())
		}
	case reflect.Ptr:
		newValue := fieldValue
		if newValue.IsNil() {
			for newValue.Kind() == reflect.Ptr {
				newValue.Set(reflect.New(newValue.Type().Elem()))
				newValue = newValue.Elem()
			}
		}
		err = Scan(valStr, newValue.Addr().Interface())
	default:
		err = Scan(valStr, fieldValue.Addr().Interface())
	}
	return err
}

func (this *FieldConfig) ScanDefault(fieldValue reflect.Value) error {
	err := this.Scan(this.DefaultValue, fieldValue)
	if err != nil && len(this.DefaultValue) > 0 {
		return err
	}
	return nil
}

type Schema struct {
	Type reflect.Type
	// map[FieldIndex]*Field
	Fields []*FieldConfig
}

func newSchema(t reflect.Type) *Schema {
	s := &Schema{
		Fields: make([]*FieldConfig, 0, t.NumField()),
	}
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		if value, ok := field.Tag.Lookup(tagIdentify); ok {
			if value != ignoreTag {
				fieldCnf := praseTagValue(value)
				fieldCnf.FieldIndex = i
				if fieldCnf.ColumnName == "" {
					fieldCnf.ColumnName = field.Name
				}
				s.Fields = append(s.Fields, fieldCnf)
			}
		} else {
			// use default config
			fieldCnf := &FieldConfig{
				FieldIndex: i,
				ColumnName: field.Name,
			}
			s.Fields = append(s.Fields, fieldCnf)
		}
	}
	s.Type = t
	return s
}

func praseTagValue(v string) *FieldConfig {
	c := &FieldConfig{}
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

func fillField(c *FieldConfig, k, v string) {
	switch k {
	case columnTag:
		c.ColumnName = v
	case defaultTag:
		c.DefaultValue = v
	case splitTag:
		c.Split = v
	case nilTag:
		c.NilValue = v
	case reqTag:
		c.IsRequired = true
	}
}
