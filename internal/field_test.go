package internal

import (
	"fmt"
	"reflect"
	"testing"

	convert "github.com/szyhf/go-convert"
)

type SimpleWithTag struct {
	// Int    int    `xlsx:"int_field"`
	IntAry []int `xlsx:"column(int_ary_field);split(|);default(1);"`
	// String string `xlsx:"index(2);"`
	// Uint   uint   `xlsx:"-"`

	// Float64 float64 `excel:"float64_field"`
	// UintAry []uint
}

func TestPraseTag(t *testing.T) {
	s := newSchema(reflect.TypeOf(&SimpleWithTag{}))
	fmt.Printf("%s", convert.MustJsonPrettyString(s))
}

func TestGetTagParam(t *testing.T) {
	expect := map[string][]string{
		"column(hello)":  []string{"column", "hello"},
		"hello":          []string{"column", "hello"},
		"default(hello)": []string{"default", "hello"},
		"split(|)":       []string{"split", "|"},
	}
	for k, v := range expect {
		k1, v1 := getTagParam(k)
		println(k1, v1)
		if v[0] != k1 && v[1] != v1 {
			t.Fail()
		}
	}
}
