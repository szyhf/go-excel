package twentySix

import (
	"testing"
)

func TestNumOfChar(t *testing.T) {
	if NumOfChar('A') != 0 {
		t.Fail()
	}
	if NumOfChar('Z') != 25 {
		t.Fail()
	}
}

func TestCharOfNum(t *testing.T) {
	if CharOfNum(0) != 'A' {
		t.Fail()
	}
	if CharOfNum(25) != 'Z' {
		t.Fail()
	}
}

func TestToDecimalism(t *testing.T) {
	if ToDecimalism("A") != 0 {
		println("A=", ToDecimalism("A"))
		t.Fail()
	}
	if ToDecimalism("Z") != 25 {
		println("Z=", ToDecimalism("Z"))
		t.Fail()
	}
	if ToDecimalism("GWW") != 4650 {
		println("GWW=", ToDecimalism("GWW"))
		t.Fail()
	}
}

func TestToTwentySixString(t *testing.T) {
	if ToTwentySixString(0) != "A" {
		t.Fail()
	}
	if ToTwentySixString(25) != "Z" {
		println(ToTwentySixString(25))
		t.Fail()
	}
	if ToTwentySixString(4650) != "GWW" {
		println(ToTwentySixString(4650))
		t.Fail()
	}
}
