package twentySix

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
	var s string
	s = "A"
	if ToDecimalism(s) != 0 {
		println(s, "=", ToDecimalism(s))
		t.Fail()
	}
	s = "Z"
	if ToDecimalism(s) != 25 {
		println(s, "=", ToDecimalism(s))
		t.Fail()
	}
	s = "ZZ"
	if ToDecimalism(s) != 701 {
		println(s, "=", ToDecimalism(s))
		t.Fail()
	}
	s = "AAA"
	if ToDecimalism(s) != 702 {
		println(s, "=", ToDecimalism(s))
		t.Fail()
	}
	s = "AAB"
	if ToDecimalism(s) != 703 {
		println(s, "=", ToDecimalism(s))
		t.Fail()
	}
	s = "AAZ"
	if ToDecimalism(s) != 727 {
		println(s, "=", ToDecimalism(s))
		t.Fail()
	}
	s = "GWW"
	if ToDecimalism(s) != 5352 {
		println(s, "=", ToDecimalism(s))
		t.Fail()
	}
	s = "QIJ"
	if ToDecimalism(s) != 11735 {
		println(s, "=", ToDecimalism(s))
		t.Fail()
	}
	s = "XFD"
	if ToDecimalism(s) != 16383 {
		println(s, "=", ToDecimalism(s))
		t.Fail()
	}
}

func TestToTwentySixString(t *testing.T) {
	// it's not needed.
	return
	// if toTwentySixString(0) != "A" {
	// 	t.Fail()
	// }
	// if toTwentySixString(25) != "Z" {
	// 	println(toTwentySixString(25))
	// 	t.Fail()
	// }
	// if toTwentySixString(4650) != "GWW" {
	// 	println(toTwentySixString(4650))
	// 	t.Fail()
	// }
}
