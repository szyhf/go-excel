package excel

import (
	"strconv"
	"testing"
)

func Test_read_RawRow(t *testing.T) {
	var (
		columns = []string{"a", "b", "c"}
		conn    = NewConnecter()
		row     = 2
	)
	conn.Open("./testdata/simple1.xlsx")

	var rd, _ = conn.NewReader(conn.GetSheetNames()[0])

	for rd.Next() {
		val, err := rd.RawRow()
		if err != nil {
			t.Fatal(err)
		}
		rowStr := strconv.Itoa(row)
		for col, val := range val {
			if val != columns[col]+rowStr {
				t.Fatal(col, row, val)
			}
		}
		row++
	}
}
