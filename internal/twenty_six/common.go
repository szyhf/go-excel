package twentysix

// a 26-number-system to decoder/encoder column of excel
// It's a very special system, don't use it as default system!!!

func numOfChar(c rune) int {
	a := c - rune('A')
	return int(a)
}

func charOfNum(n int) rune {
	return rune('A') + rune(n)
}

// ToDecimalism convert string to int
func ToDecimalism(s string) int {
	res := 0
	ary := []rune(s)

	for i, j := len(ary)-2, 26; i >= 0; i, j = i-1, j*26 {
		c := ary[i]
		// log.Printf("res(%d)=res(%d)+(Num(%d)+1)*%d, c=%s, i=%d, j=%d\n", res+(NumOfChar(c)+1)*j, res, NumOfChar(c), j, string(c), i, j)
		res = res + (numOfChar(c)+1)*j
	}

	return res + numOfChar(ary[len(ary)-1])
}

func toTwentySixString(n int) string {
	panic("it's not needed.")
	// s := ""
	// if n == 0 {
	// 	return "A"
	// }
	// for n > 0 {
	// 	m := n % 26
	// 	if m == 0 {
	// 		m = 26
	// 	}
	// 	s = string(charOfNum(m)) + s
	// 	n = (n - m) / 26
	// }
	// return s
}

func pow(x, y int) int {
	res := 1
	for i := 1; i < y; i++ {
		res = res * x
	}
	return res
}
