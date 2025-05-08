package twentysix

import (
	"testing"
)

func TestNumOfChar(t *testing.T) {
	if numOfChar('A') != 0 {
		t.Fail()
	}
	if numOfChar('Z') != 25 {
		t.Fail()
	}
}

func TestCharOfNum(t *testing.T) {
	if charOfNum(0) != 'A' {
		t.Fail()
	}
	if charOfNum(25) != 'Z' {
		t.Fail()
	}
}

func TestToDecimalism(t *testing.T) {
	s := "A"
	if ToDecimals(s) != 0 {
		println(s, "=", ToDecimals(s))
		t.Fail()
	}
	s = "Z"
	if ToDecimals(s) != 25 {
		println(s, "=", ToDecimals(s))
		t.Fail()
	}
	s = "ZZ"
	if ToDecimals(s) != 701 {
		println(s, "=", ToDecimals(s))
		t.Fail()
	}
	s = "AAA"
	if ToDecimals(s) != 702 {
		println(s, "=", ToDecimals(s))
		t.Fail()
	}
	s = "AAB"
	if ToDecimals(s) != 703 {
		println(s, "=", ToDecimals(s))
		t.Fail()
	}
	s = "AAZ"
	if ToDecimals(s) != 727 {
		println(s, "=", ToDecimals(s))
		t.Fail()
	}
	s = "GWW"
	if ToDecimals(s) != 5352 {
		println(s, "=", ToDecimals(s))
		t.Fail()
	}
	s = "QIJ"
	if ToDecimals(s) != 11735 {
		println(s, "=", ToDecimals(s))
		t.Fail()
	}
	s = "XFD"
	if ToDecimals(s) != 16383 {
		println(s, "=", ToDecimals(s))
		t.Fail()
	}
}

func TestToColumnName(t *testing.T) {
	tests := []struct {
		input int
		want  string
	}{
		{-1, ""},       // 负数测试
		{0, "A"},       // 边界值测试
		{25, "Z"},      // 边界值测试
		{26, "AA"},     // 进位测试
		{27, "AB"},     // 常规测试
		{51, "AZ"},     // 常规测试
		{701, "ZZ"},    // 边界值测试
		{702, "AAA"},   // 进位测试
		{16383, "XFD"}, // Excel最大列测试
	}

	for _, tt := range tests {
		got := ToColumnName(tt.input)
		if got != tt.want {
			t.Errorf("ToColumnName(%d) = %v, want %v", tt.input, got, tt.want)
		}
	}
}

func TestToCellName(t *testing.T) {
	tests := []struct {
		x    int
		y    int
		want string
	}{
		{0, 0, "A1"},          // 起始位置测试
		{0, 1, "A2"},          // Y轴测试
		{1, 0, "B1"},          // X轴测试
		{25, 25, "Z26"},       // 常规测试
		{26, 1, "AA2"},        // X轴进位测试
		{701, 1000, "ZZ1001"}, // 大数值测试
	}

	for _, tt := range tests {
		got := ToCellName(tt.x, tt.y)
		if got != tt.want {
			t.Errorf("ToCellName(%d, %d) = %v, want %v", tt.x, tt.y, got, tt.want)
		}
	}
}
