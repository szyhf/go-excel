package twentysix

import "strconv"

// a 26-number-system to decoder/encoder column of excel
// It's a very special system, don't use it as default system!!!

func numOfChar(c rune) int {
	a := c - rune('A')
	return int(a)
}

func charOfNum[Integer IntegerType](n Integer) rune {
	return rune('A') + rune(n)
}

// ToDecimals convert string to int
func ToDecimals(s string) int {
	res := 0
	ary := []rune(s)

	for i, j := len(ary)-2, 26; i >= 0; i, j = i-1, j*26 {
		c := ary[i]
		// log.Printf("res(%d)=res(%d)+(Num(%d)+1)*%d, c=%s, i=%d, j=%d\n", res+(NumOfChar(c)+1)*j, res, NumOfChar(c), j, string(c), i, j)
		res = res + (numOfChar(c)+1)*j
	}

	return res + numOfChar(ary[len(ary)-1])
}

// 把第n列转成excel用的字母列名
// 注意，这里的n是从0开始的
// 0 -> A
// 1 -> B
// 25 -> Z
// 26 -> AA
// 27 -> AB
// 51 -> AZ
// 701 -> ZZ
// 702 -> AAA
func ToColumnName[Integer IntegerType](n Integer) string {
	if n < 0 {
		return ""
	}

	if n < 26 {
		return string(charOfNum(n))
	}

	var result string
	for n >= 0 {
		rem := n % 26
		n = n/26 - 1
		result = string(charOfNum(rem)) + result
	}

	return result
}

type IntegerType interface {
	int | int8 | int16 | int32 | int64
}

// ToCellName 把坐标转成excel用的单元格名
// 注意，这里的x和y都是从0开始的
// 0,0 -> A1
// 0,1 -> A2
// 1,0 -> B1
// 25,25 -> ZZ26
func ToCellName[Integer IntegerType](x, y Integer) string {
	return ToColumnName(x) + strconv.FormatInt(int64(y+1), 10)
}
