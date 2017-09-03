package test

import (
	"testing"

	excel "github.com/szyhf/go-excel-orm"
)

type Simple struct {
	Int int
	// Uint    uint
	String string
	// Float64 float64
}

type SimpleWithTag struct {
	Int    int    `excel:"int_field"`
	IntAry []int  `excel:"int_ary_field" excelsplit:"|"`
	String string `excel:"string_field"`
	// Uint    uint    `excel:"uint_field"`

	// Float64 float64 `excel:"float64_field"`
	// UintAry []uint
}

func TestMarshal(t *testing.T) {
	filePath := "./testdata/simple.xlsx"
	sheetName := "simple"
	conn := excel.NewConnecter()
	conn.Open(filePath)
	rd := conn.NewReader(sheetName)
	var simple Simple
	rd.Read(&simple)
}
