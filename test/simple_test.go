package test

import (
	"testing"
)

type Simple struct {
	Int int
	// Uint    uint
	String string
	// Float64 float64
}

type SimpleWithTag struct {
	Int    int    `excel:"int_field"`
	IntAry []int  `excel:"column(int_ary_field);split(|);default(0);"`
	String string `excel:"index(2);"`
	// Uint    uint    `excel:"uint_field"`

	// Float64 float64 `excel:"float64_field"`
	// UintAry []uint
}

func TestMarshal(t *testing.T) {

	// conn := excel.NewConnecter()
	// conn.Open(filePath)
	// rd := conn.NewReader(sheetName)
	// var simple Simple
	// rd.Read(&simple)
}
