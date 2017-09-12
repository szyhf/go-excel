package twentySix

// a 26-number-system to decoder/encoder column of excel

func NumOfChar(c rune) int {
	a := c - rune('A')
	return int(a)
}

func CharOfNum(n int) rune {
	return rune('A') + rune(n)
}

func ToDecimalism(s string) int {
	res := 0
	ary := []rune(s)

	// for i := 0; i < len(ary); i++ {
	// 	c := ary[i]
	// 	res = res + NumOfChar(c)*pow(26, len(ary)-i)
	// }
	for i, j := len(ary)-1, 1; i >= 0; i, j = i-1, j*26 {
		c := ary[i]
		res = res + NumOfChar(c)*j
	}

	return res
}

func ToTwentySixString(n int) string {
	s := ""
	if n == 0 {
		return "A"
	}
	for n > 0 {
		m := n % 26
		if m == 0 {
			m = 26
		}
		s = string(CharOfNum(m)) + s
		n = (n - m) / 26
	}
	return s
}

func pow(x, y int) int {
	res := 1
	for i := 1; i < y; i++ {
		res = res * x
	}
	return res
}
