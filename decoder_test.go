package excelstruct

import (
	"fmt"
	"reflect"
	"strconv"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

var (
	_ ValueUnmarshaler = (*stringType)(nil)
	_ ValueMarshaler   = (*stringType)(nil)
)

type stringType string

func (s *stringType) MarshalXLSXValue() ([]string, error) {
	return []string{string(*s) + " world"}, nil
}

func (s *stringType) UnmarshalXLSXValue(value []string) error {
	*s = stringType(value[0]) + " world"
	return nil
}

type baseType struct {
	Int     int       `excel:"int"`
	Int8    int8      `excel:"int8"`
	Int16   int16     `excel:"int16"`
	Int32   int32     `excel:"int32"`
	Int64   int64     `excel:"int64"`
	Uint    uint      `excel:"uint"`
	Uint8   uint8     `excel:"uint8"`
	Uint16  uint16    `excel:"uint16"`
	Uint32  uint32    `excel:"uint32"`
	Uint64  uint64    `excel:"uint64"`
	Float32 float32   `excel:"float32"`
	Float64 float64   `excel:"float64"`
	String  string    `excel:"string"`
	Bool    bool      `excel:"bool"`
	Date    time.Time `excel:"date"`

	PInt     *int       `excel:"pint"`
	PInt8    *int8      `excel:"pint8"`
	PInt16   *int16     `excel:"pint16"`
	PInt32   *int32     `excel:"pint32"`
	PInt64   *int64     `excel:"pint64"`
	PUint    *uint      `excel:"puint"`
	PUint8   *uint8     `excel:"puint8"`
	PUint16  *uint16    `excel:"puint16"`
	PUint32  *uint32    `excel:"puint32"`
	PUint64  *uint64    `excel:"puint64"`
	PFloat32 *float32   `excel:"pfloat32"`
	PFloat64 *float64   `excel:"pfloat64"`
	PString  *string    `excel:"pstring"`
	PBool    *bool      `excel:"pbool"`
	PDate    *time.Time `excel:"pdate"`
}

type unmarshalerType struct {
	String  stringType  `excel:"string"`
	PString *stringType `excel:"pstring"`
}

type sliceType struct {
	Strings  []string  `excel:"strings"`
	Bytes    []byte    `excel:"bytes"`
	PStrings []*string `excel:"pstrings"`
}

type arrayType struct {
	Strings  [3]string  `excel:"strings"`
	PStrings [3]*string `excel:"pstrings"`
}

type inlineType struct {
	B baseType  `excel:",inline"`
	C *struct{} `excel:",inline"`
}

func TestUnmarshal(t *testing.T) {
	t.Parallel()

	t.Run("simple", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/type.xlsx"})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewDecoder[baseType](f, DecoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		var got []baseType
		require.NoError(t, sheet.All(&got))

		want := []baseType{{
			Int:     1,
			Int8:    8,
			Int16:   16,
			Int32:   32,
			Int64:   64,
			Uint:    2,
			Uint8:   28,
			Uint16:  216,
			Uint32:  232,
			Uint64:  264,
			Float32: 32.0,
			Float64: 64.0,
			String:  "hello",
			Bool:    true,
			Date:    time.Date(2023, 12, 7, 0, 0, 0, 0, time.UTC),

			PInt:     ptrV(1),
			PInt8:    ptrV(int8(8)),
			PInt16:   ptrV(int16(16)),
			PInt32:   ptrV(int32(32)),
			PInt64:   ptrV(int64(64)),
			PUint:    ptrV(uint(2)),
			PUint8:   ptrV(uint8(28)),
			PUint16:  ptrV(uint16(216)),
			PUint32:  ptrV(uint32(232)),
			PUint64:  ptrV(uint64(264)),
			PFloat32: ptrV(float32(32.0)),
			PFloat64: ptrV(64.0),
			PString:  ptrV("hello"),
			PBool:    ptrV(true),
			PDate:    ptrV(time.Date(2023, 12, 7, 0, 0, 0, 0, time.UTC)),
		}}
		assert.Equal(t, want, got)
	})

	t.Run("slice", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/array.xlsx"})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewDecoder[sliceType](f, DecoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		var got []sliceType
		require.NoError(t, sheet.All(&got))
		assert.Equal(t, []sliceType{{
			Strings:  []string{"hello", "world"},
			PStrings: []*string{ptrV("phello"), ptrV("pworld")},
		}}, got)
	})

	t.Run("array", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/array.xlsx"})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewDecoder[arrayType](f, DecoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		var got []arrayType
		require.NoError(t, sheet.All(&got))
		assert.Equal(t, []arrayType{{
			Strings:  [3]string{"hello", "world", ""},
			PStrings: [3]*string{ptrV("phello"), ptrV("pworld"), nil},
		}}, got)
	})

	t.Run("unmarshaler", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/type.xlsx"})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewDecoder[unmarshalerType](f, DecoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		var got []unmarshalerType
		require.NoError(t, sheet.All(&got))
		assert.Equal(t, []unmarshalerType{{
			String:  "hello world",
			PString: ptrV(stringType("hello world")),
		}}, got)
	})

	t.Run("inline", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/type.xlsx"})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewDecoder[inlineType](f, DecoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		var got []inlineType
		require.NoError(t, sheet.All(&got))
		assert.Equal(t, []inlineType{{
			B: baseType{
				Int:     1,
				Int8:    8,
				Int16:   16,
				Int32:   32,
				Int64:   64,
				Uint:    2,
				Uint8:   28,
				Uint16:  216,
				Uint32:  232,
				Uint64:  264,
				Float32: 32.0,
				Float64: 64.0,
				String:  "hello",
				Bool:    true,
				Date:    time.Date(2023, 12, 7, 0, 0, 0, 0, time.UTC),

				PInt:     ptrV(1),
				PInt8:    ptrV(int8(8)),
				PInt16:   ptrV(int16(16)),
				PInt32:   ptrV(int32(32)),
				PInt64:   ptrV(int64(64)),
				PUint:    ptrV(uint(2)),
				PUint8:   ptrV(uint8(28)),
				PUint16:  ptrV(uint16(216)),
				PUint32:  ptrV(uint32(232)),
				PUint64:  ptrV(uint64(264)),
				PFloat32: ptrV(float32(32.0)),
				PFloat64: ptrV(64.0),
				PString:  ptrV("hello"),
				PBool:    ptrV(true),
				PDate:    ptrV(time.Date(2023, 12, 7, 0, 0, 0, 0, time.UTC)),
			},
			C: nil,
		}}, got)
	})

	t.Run("nil", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/type_nil.xlsx"})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewDecoder[baseType](f, DecoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		var got []baseType
		require.NoError(t, sheet.All(&got))

		want := []baseType{{Int: 1}}
		assert.Equal(t, want, got)
	})
}

func TestInvalidUnmarshalError(t *testing.T) {
	t.Parallel()

	f, err := OpenFile(OpenFileOptions{FilePath: "testdata/type.xlsx"})
	require.NoError(t, err)

	sheet, err := NewDecoder[baseType](f, DecoderOptions{})
	require.NoError(t, err)
	defer sheet.Close()

	t.Run("nil", func(t *testing.T) {
		t.Parallel()

		err := sheet.Decode(nil)
		got := new(InvalidUnmarshalError)
		require.ErrorAs(t, err, &got)

		want := &InvalidUnmarshalError{Type: reflect.TypeOf((*baseType)(nil))}
		assert.Equal(t, want, got)
	})
}

func TestUnmarshalError(t *testing.T) {
	t.Parallel()

	t.Run("type", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/type_error.xlsx"})
		require.NoError(t, err)

		sheet, err := NewDecoder[baseType](f, DecoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		var got []baseType
		for sheet.Next() {
			var v baseType
			err := sheet.Decode(&v)
			if err == nil {
				got = append(got, v)
				continue
			}

			got := new(UnmarshalError)
			require.ErrorAs(t, err, &got)

			typeErr := got.AsTypeError()
			require.Equal(t, 28, len(typeErr))

			want := UnmarshalTypeError{
				Value: "a",
				Type:  reflect.TypeOf(int(0)),
				Field: "int",
				Err:   nil,
			}
			assert.Equal(t, want.Value, typeErr[0].Value)
			assert.Equal(t, want.Type, typeErr[0].Type)
			assert.Equal(t, want.Field, typeErr[0].Field)
		}

		want := []baseType{{
			Int:     1,
			Int8:    8,
			Int16:   16,
			Int32:   32,
			Int64:   64,
			Uint:    2,
			Uint8:   28,
			Uint16:  216,
			Uint32:  232,
			Uint64:  264,
			Float32: 32.0,
			Float64: 64.0,
			String:  "hello",
			Bool:    true,
			Date:    time.Date(2023, 12, 7, 0, 0, 0, 0, time.UTC),

			PInt:     ptrV(1),
			PInt8:    ptrV(int8(8)),
			PInt16:   ptrV(int16(16)),
			PInt32:   ptrV(int32(32)),
			PInt64:   ptrV(int64(64)),
			PUint:    ptrV(uint(2)),
			PUint8:   ptrV(uint8(28)),
			PUint16:  ptrV(uint16(216)),
			PUint32:  ptrV(uint32(232)),
			PUint64:  ptrV(uint64(264)),
			PFloat32: ptrV(float32(32.0)),
			PFloat64: ptrV(64.0),
			PString:  ptrV("hello"),
			PBool:    ptrV(true),
			PDate:    ptrV(time.Date(2023, 12, 7, 0, 0, 0, 0, time.UTC)),
		}}
		assert.Equal(t, want, got)
	})

	t.Run("convert", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/type_error.xlsx"})
		require.NoError(t, err)

		sheet, err := NewDecoder[baseType](f, DecoderOptions{
			StringConv: func(_ string, v string) (string, error) {
				if len(v) <= 2 {
					return "", fmt.Errorf("string error")
				}
				return v, nil
			},
			BoolConv: func(_ string, v string) (bool, error) {
				return strconv.ParseBool(v)
			},
			TimeConv: func(v string) (time.Time, error) {
				return defaultTimeConv(v)
			},
		})

		require.NoError(t, err)
		defer sheet.Close()

		var got []baseType
		for sheet.Next() {
			var v baseType
			err := sheet.Decode(&v)
			if err == nil {
				got = append(got, v)
				continue
			}

			got := new(UnmarshalError)
			require.ErrorAs(t, err, &got)

			convertErr := got.AsConvertValueError()
			require.Equal(t, 4, len(convertErr))

			want := ConvertValueError{
				Value: "m",
				Field: "string",
				Err:   nil,
			}
			assert.Equal(t, want.Value, convertErr[0].Value)
			assert.Equal(t, want.Field, convertErr[0].Field)
		}

		want := []baseType{{
			Int:     1,
			Int8:    8,
			Int16:   16,
			Int32:   32,
			Int64:   64,
			Uint:    2,
			Uint8:   28,
			Uint16:  216,
			Uint32:  232,
			Uint64:  264,
			Float32: 32.0,
			Float64: 64.0,
			String:  "hello",
			Bool:    true,
			Date:    time.Date(2023, 12, 7, 0, 0, 0, 0, time.UTC),

			PInt:     ptrV(1),
			PInt8:    ptrV(int8(8)),
			PInt16:   ptrV(int16(16)),
			PInt32:   ptrV(int32(32)),
			PInt64:   ptrV(int64(64)),
			PUint:    ptrV(uint(2)),
			PUint8:   ptrV(uint8(28)),
			PUint16:  ptrV(uint16(216)),
			PUint32:  ptrV(uint32(232)),
			PUint64:  ptrV(uint64(264)),
			PFloat32: ptrV(float32(32.0)),
			PFloat64: ptrV(64.0),
			PString:  ptrV("hello"),
			PBool:    ptrV(true),
			PDate:    ptrV(time.Date(2023, 12, 7, 0, 0, 0, 0, time.UTC)),
		}}
		assert.Equal(t, want, got)
	})
}

func TestDecoder_Count(t *testing.T) {
	t.Parallel()

	t.Run("offset=0", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/rows3.xlsx"})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewDecoder[baseType](f, DecoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		want := 3
		assert.Equal(t, want, sheet.Count())
	})

	t.Run("offset=1", func(t *testing.T) {
		t.Parallel()

		f, err := OpenFile(OpenFileOptions{FilePath: "testdata/offset1rows3.xlsx"})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewDecoder[baseType](f, DecoderOptions{TitleRowIndex: 2})
		require.NoError(t, err)
		defer sheet.Close()

		want := 3
		assert.Equal(t, want, sheet.Count())
	})

}

func ptrV[T any](v T) *T {
	return &v
}
