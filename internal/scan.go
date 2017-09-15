package internal

import (
	"encoding"
	"errors"
	"fmt"
	"reflect"

	"github.com/szyhf/go-convert"
)

// ref: gopkg.in/redis.v5

func Scan(s string, ptr interface{}) error {
	var err error
	switch p := ptr.(type) {
	case nil:
		return errors.New("Scan(nil).")
	case *string:
		*p = s
	case *[]byte:
		*p = []byte(s)
	case *[]rune:
		*p = []rune(s)
	case *int:
		*p, err = convert.ToInt(s)
	case *int8:
		*p, err = convert.ToInt8(s)
	case *int16:
		*p, err = convert.ToInt16(s)
	case *int32:
		*p, err = convert.ToInt32(s)
	case *int64:
		*p, err = convert.ToInt64(s)
	case *uint:
		*p, err = convert.ToUint(s)
	case *uint8:
		*p, err = convert.ToUint8(s)
	case *uint16:
		*p, err = convert.ToUint16(s)
	case *uint32:
		*p, err = convert.ToUint32(s)
	case *uint64:
		*p, err = convert.ToUint64(s)
	case *float32:
		*p, err = convert.ToFloat32(s)
	case *float64:
		*p, err = convert.ToFloat64(s)
	case *bool:
		*p, err = convert.ToBool(s)
	case encoding.BinaryUnmarshaler:
		if err = p.UnmarshalBinary([]byte(s)); err != nil {
			err = fmt.Errorf("can't unmarshar by encoding.BinaryUnmarshaler: %s", err)
		}
	default:
		err = fmt.Errorf(
			"can't unmarshal %T (consider implementing encoding.TextUnmarshaler)", ptr)
	}
	return err
}

func ScanSlice(data []string, sliceValue reflect.Value) error {
	if !sliceValue.IsValid() {
		return fmt.Errorf("ScanSlice(nil)")
	}
	if sliceValue.Kind() != reflect.Ptr {
		return fmt.Errorf("ScanSlice(non-pointer %s)", sliceValue.Kind())
	}
	sliceValue = sliceValue.Elem()
	if sliceValue.Kind() != reflect.Slice {
		return fmt.Errorf("ScanSlice(non-slice %s)", sliceValue.Kind())
	}

	for i, s := range data {
		elem := SliceNextElem(sliceValue)
		if err := Scan(s, elem.Addr().Interface()); err != nil {
			return fmt.Errorf("ScanSlice(index=%d value=%q) failed: %s", i, s, err)
		}
	}

	return nil
}

func SliceNextElem(v reflect.Value) reflect.Value {
	if v.Len() < v.Cap() {
		v.Set(v.Slice(0, v.Len()+1))
		return v.Index(v.Len() - 1)
	}

	elemType := v.Type().Elem()

	if elemType.Kind() == reflect.Ptr {
		elem := reflect.New(elemType.Elem())
		v.Set(reflect.Append(v, elem))
		return elem.Elem()
	}

	v.Set(reflect.Append(v, reflect.Zero(elemType)))
	return v.Index(v.Len() - 1)
}
