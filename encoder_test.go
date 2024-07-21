package excelstruct

import (
	"fmt"
	"reflect"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

type marshalerType struct {
	String  stringType  `excel:"string"`
	PString *stringType `excel:"pstring"`
}

func TestMarshal(t *testing.T) {
	t.Parallel()

	t.Run("simple", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[baseType](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		assert.NoError(t, sheet.Encode(&baseType{
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
		}))

		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)
		assert.Equal(t, [][]string{
			{"int", "1"},
			{"int8", "8"},
			{"int16", "16"},
			{"int32", "32"},
			{"int64", "64"},
			{"uint", "2"},
			{"uint8", "28"},
			{"uint16", "216"},
			{"uint32", "232"},
			{"uint64", "264"},
			{"float32", "32"},
			{"float64", "64"},
			{"string", "hello"},
			{"bool", "TRUE"},
			{"date", "12/7/23 00:00"},
			{"pint", "1"},
			{"pint8", "8"},
			{"pint16", "16"},
			{"pint32", "32"},
			{"pint64", "64"},
			{"puint", "2"},
			{"puint8", "28"},
			{"puint16", "216"},
			{"puint32", "232"},
			{"puint64", "264"},
			{"pfloat32", "32"},
			{"pfloat64", "64"},
			{"pstring", "hello"},
			{"pbool", "TRUE"},
			{"pdate", "12/7/23 00:00"},
		}, got)
	})

	t.Run("slice", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[sliceType](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		assert.NoError(t, sheet.Encode(&sliceType{
			Strings:  []string{"hello", "world", ""},
			Bytes:    []byte("hello"),
			PStrings: []*string{ptrV("hello"), ptrV("world"), nil},
		}))

		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)
		assert.Equal(t, [][]string{
			{"strings", "hello"},
			{"strings", "world"},
			{"strings", ""},
			{"bytes", "hello"},
			{"pstrings", "hello"},
			{"pstrings", "world"},
			{"pstrings"},
		}, got)
	})

	t.Run("slice_nil", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[sliceType](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		assert.NoError(t, sheet.Encode(&sliceType{}))
		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)
		assert.Equal(t, [][]string{
			{"strings"},
			{"bytes"},
			{"pstrings"},
		}, got)
	})

	t.Run("array", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[arrayType](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		assert.NoError(t, sheet.Encode(&arrayType{
			Strings:  [3]string{"hello", "world", ""},
			PStrings: [3]*string{ptrV("hello"), ptrV("world"), nil}},
		))

		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)
		assert.Equal(t, [][]string{
			{"strings", "hello"},
			{"strings", "world"},
			{"strings", ""},
			{"pstrings", "hello"},
			{"pstrings", "world"},
			{"pstrings"},
		}, got)
	})

	t.Run("map", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[map[string]any](f, EncoderOptions{TitleName: []string{"string"}})
		require.NoError(t, err)
		defer sheet.Close()

		assert.NoError(t, sheet.Encode(&map[string]any{"string": "hello"}))
		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		want := [][]string{{"string", "hello"}}
		assert.Equal(t, want, got)
	})

	t.Run("marshaler", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[marshalerType](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		assert.NoError(t, sheet.Encode(&marshalerType{
			String:  "hello",
			PString: ptrV(stringType("hello")),
		}))

		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)
		assert.Equal(t, [][]string{
			{"string", "hello world"},
			{"pstring", "hello world"},
		}, got)
	})

	t.Run("nil", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[baseType](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		assert.NoError(t, sheet.Encode(&baseType{}))
		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)
		assert.Equal(t, [][]string{
			{"int", "0"},
			{"int8", "0"},
			{"int16", "0"},
			{"int32", "0"},
			{"int64", "0"},
			{"uint", "0"},
			{"uint8", "0"},
			{"uint16", "0"},
			{"uint32", "0"},
			{"uint64", "0"},
			{"float32", "0"},
			{"float64", "0"},
			{"string", ""},
			{"bool", "FALSE"},
			{"date", "0001-01-01T00:00:00Z"},
			{"pint"},
			{"pint8"},
			{"pint16"},
			{"pint32"},
			{"pint64"},
			{"puint"},
			{"puint8"},
			{"puint16"},
			{"puint32"},
			{"puint64"},
			{"pfloat32"},
			{"pfloat64"},
			{"pstring"},
			{"pbool"},
			{"pdate"},
		}, got)
	})

	t.Run("omitempty", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type v struct {
			Int       int            `excel:"int,omitempty"`
			Int8      int8           `excel:"int8,omitempty"`
			Int16     int16          `excel:"int16,omitempty"`
			Int32     int32          `excel:"int32,omitempty"`
			Int64     int64          `excel:"int64,omitempty"`
			Uint      uint           `excel:"uint,omitempty"`
			Uint8     uint8          `excel:"uint8,omitempty"`
			Uint16    uint16         `excel:"uint16,omitempty"`
			Uint32    uint32         `excel:"uint32,omitempty"`
			Uint64    uint64         `excel:"uint64,omitempty"`
			Float32   float32        `excel:"float32,omitempty"`
			Float64   float64        `excel:"float64,omitempty"`
			String    string         `excel:"string,omitempty"`
			Slice     []string       `excel:"slice,omitempty"`
			Array     [0]string      `excel:"array,omitempty"`
			Map       map[string]int `excel:"map,omitempty"`
			Bool      bool           `excel:"bool,omitempty"`
			Date      time.Time      `excel:"date,omitempty"`
			PInt      *int           `excel:"pint,omitempty"`
			PInt8     *int8          `excel:"pint8,omitempty"`
			PInt16    *int16         `excel:"pint16,omitempty"`
			PInt32    *int32         `excel:"pint32,omitempty"`
			PInt64    *int64         `excel:"pint64,omitempty"`
			PUint     *uint          `excel:"puint,omitempty"`
			PUint8    *uint8         `excel:"puint8,omitempty"`
			PUint16   *uint16        `excel:"puint16,omitempty"`
			PUint32   *uint32        `excel:"puint32,omitempty"`
			PUint64   *uint64        `excel:"puint64,omitempty"`
			PFloat32  *float32       `excel:"pfloat32,omitempty"`
			PFloat64  *float64       `excel:"pfloat64,omitempty"`
			PString   *string        `excel:"pstring,omitempty"`
			PBool     *bool          `excel:"pbool,omitempty"`
			PDate     *time.Time     `excel:"pdate,omitempty"`
			Interface ValueMarshaler `excel:"marshaller,omitempty"`
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()
		assert.NoError(t, sheet.Encode(&v{}))

		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)
		assert.Equal(t, [][]string{
			{"int"},
			{"int8"},
			{"int16"},
			{"int32"},
			{"int64"},
			{"uint"},
			{"uint8"},
			{"uint16"},
			{"uint32"},
			{"uint64"},
			{"float32"},
			{"float64"},
			{"string"},
			{"slice"},
			{"array"},
			{"map"},
			{"bool"},
			{"date"},
			{"pint"},
			{"pint8"},
			{"pint16"},
			{"pint32"},
			{"pint64"},
			{"puint"},
			{"puint8"},
			{"puint16"},
			{"puint32"},
			{"puint64"},
			{"pfloat32"},
			{"pfloat64"},
			{"pstring"},
			{"pbool"},
			{"pdate"},
			{"marshaller"},
		}, got)
	})
}

func TestMarshal_ValueConv(t *testing.T) {
	t.Parallel()

	t.Run("string", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type v struct {
			V string `excel:"v"`
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{
			StringConv: func(title string, v string) (string, error) {
				return v + " world", nil
			},
		})
		require.NoError(t, err)
		defer sheet.Close()

		require.NoError(t, sheet.Encode(&v{V: "hello"}))
		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		want := [][]string{
			{"v", "hello world"},
		}
		assert.Equal(t, want, got)
	})

	t.Run("bool", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type v struct {
			V string `excel:"v"`
			B bool   `excel:"b"`
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{
			BoolConv: func(title string, v bool) (string, error) {
				if v {
					return "yes", nil
				}
				return "no", nil
			},
		})
		require.NoError(t, err)
		defer sheet.Close()

		require.NoError(t, sheet.Encode(&v{V: "v", B: true}))
		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		want := [][]string{
			{"v", "v"},
			{"b", "yes"},
		}
		assert.Equal(t, want, got)
	})
}

func TestEncoder_Orientation(t *testing.T) {
	t.Parallel()

	type v struct {
		V []int `excel:"v"`
	}

	t.Run("row", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[v](f, EncoderOptions{Orientation: OrientationRow})
		require.NoError(t, err)

		require.NoError(t, sheet.Encode(&v{V: []int{1, 2, 3}}))
		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		want := [][]string{
			{"v", "1"},
			{"v", "2"},
			{"v", "3"},
		}
		assert.Equal(t, want, got)
	})

	t.Run("column", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[v](f, EncoderOptions{Orientation: OrientationColumn})
		require.NoError(t, err)
		defer sheet.Close()

		require.NoError(t, sheet.Encode(&v{V: []int{1, 2, 3}}))
		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		want := [][]string{
			{"v", "1", "2", "3"},
		}
		assert.Equal(t, want, got)
	})

}

func TestEncoder_DataValidation(t *testing.T) {
	t.Parallel()

	type v struct {
		V []int `excelstruct:"v"`
	}

	t.Run("default", func(t *testing.T) {
		t.Parallel()

		dv := excelize.NewDataValidation(true)
		dv.SetSqref("A1:B2")
		require.NoError(t, dv.SetRange(10, 20, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween))
		dv.SetError(excelize.DataValidationErrorStyleStop, "error title", "error body")

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[v](f, EncoderOptions{DataValidation: func(title string) (*excelize.DataValidation, error) {
			dv := *dv
			return &dv, nil
		}})
		require.NoError(t, err)

		require.NoError(t, sheet.Encode(&v{V: []int{9}}))
		require.NoError(t, sheet.Close())

		got, err := f.File.GetDataValidations(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		dv.Sqref = "A2:A2"
		assert.Equal(t, []*excelize.DataValidation{dv}, got)
	})

	t.Run("over=1", func(t *testing.T) {
		t.Parallel()

		dv := excelize.NewDataValidation(true)
		dv.SetSqref("A1:B2")
		require.NoError(t, dv.SetRange(10, 20, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween))
		dv.SetError(excelize.DataValidationErrorStyleStop, "error title", "error body")

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[v](f, EncoderOptions{
			DataValidation: func(title string) (*excelize.DataValidation, error) {
				dv := *dv
				return &dv, nil
			},
			ValidationOverRow: 1,
		})
		require.NoError(t, err)

		require.NoError(t, sheet.Encode(&v{V: []int{9}}))
		require.NoError(t, sheet.Close())

		got, err := f.File.GetDataValidations(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		dv.Sqref = "A2:A3"
		assert.Equal(t, []*excelize.DataValidation{dv}, got)
	})

	t.Run("slice", func(t *testing.T) {
		t.Parallel()

		dv := excelize.NewDataValidation(true)
		dv.SetSqref("A1:B2")
		require.NoError(t, dv.SetRange(10, 20, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween))
		dv.SetError(excelize.DataValidationErrorStyleStop, "error title", "error body")

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewEncoder[v](f, EncoderOptions{
			DataValidation: func(title string) (*excelize.DataValidation, error) {
				dv := *dv
				return &dv, nil
			},
		})
		require.NoError(t, err)

		require.NoError(t, sheet.Encode(&v{V: []int{9, 10}}))
		require.NoError(t, sheet.Close())

		got, err := f.File.GetDataValidations(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		dv.Sqref = "A2:B2"
		assert.Equal(t, []*excelize.DataValidation{dv}, got)
	})
}

func TestEncoder_Style(t *testing.T) {
	t.Parallel()

	type v struct {
		V int `excel:"v"`
	}

	f, err := WriteFile(WriteFileOptions{})
	require.NoError(t, err)
	defer f.Close()

	sheet, err := NewEncoder[v](f, EncoderOptions{Style: NameStyle{
		"v": &excelize.Style{Border: []excelize.Border{}},
	}})
	require.NoError(t, err)

	require.NoError(t, sheet.Encode(&v{V: 9}))
	require.NoError(t, sheet.Close())
}

func TestEncoder_Width(t *testing.T) {
	t.Parallel()

	f, err := WriteFile(WriteFileOptions{
		FilePath: "testdata/width.xlsx",
	})
	require.NoError(t, err)
	defer f.Close()

	type v struct {
		V int `excelstruct:"v"`
	}

	sheet, err := NewEncoder[v](f, EncoderOptions{
		TitleScaleAutoWidth: DefaultScaleAutoWidth,
		TitleMaxWidth:       func(title string) float64 { return -1 },
	})
	require.NoError(t, err)
	defer sheet.Close()

	require.NoError(t, sheet.Encode(&v{V: 9}))
	require.NoError(t, sheet.Close())

	got, err := f.File.GetColWidth(sheet.enc.title.config.sheetName, "A")
	require.NoError(t, err)

	want := DefaultScaleAutoWidth(len("9"))
	assert.Equal(t, want, got)
}

func TestMarshal_Panic(t *testing.T) {
	t.Parallel()

	t.Run("recover", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type v struct {
			V string `excel:"v"`
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{
			StringConv: func(title string, v string) (string, error) {
				return "", fmt.Errorf("msg")
			},
		})
		require.NoError(t, err)
		defer sheet.Close()

		want := "msg"
		got := sheet.Encode(&v{V: ""})
		assert.Equal(t, want, got.Error())
	})

	t.Run("not recover", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type v struct {
			V string `excel:"v"`
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{
			StringConv: func(title string, v string) (string, error) {
				panic("msg")
			},
		})
		require.NoError(t, err)
		defer sheet.Close()

		require.PanicsWithValue(t, "msg", func() {
			_ = sheet.Encode(&v{V: ""})
		})
	})
}

func TestMarshal_StructInline(t *testing.T) {
	t.Parallel()

	t.Run("ok", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type b struct {
			A string
		}

		type v struct {
			B b `excel:",inline"`
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()
		require.NoError(t, sheet.Encode(&v{B: b{A: "hello"}}))

		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		want := [][]string{{"A", "hello"}}
		assert.Equal(t, want, got)
	})

	t.Run("not export", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type b struct {
			A string
		}

		type v struct {
			b b `excel:",inline"`
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()
		require.NoError(t, sheet.Encode(&v{b: b{A: "hello"}}))

		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		want := [][]string{}
		assert.Equal(t, want, got)
	})

	t.Run("not inline", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type b struct {
			A string
		}

		type v struct {
			B b
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()
		require.NoError(t, sheet.Encode(&v{B: b{A: "hello"}}))

		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)

		want := [][]string{}
		assert.Equal(t, want, got)
	})
}

func TestMarshal_UnsupportedType(t *testing.T) {
	t.Parallel()

	t.Run("func", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type v struct {
			V func() `excel:"v"`
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		err = sheet.Encode(&v{})
		got := &UnsupportedTypeError{}
		assert.ErrorAs(t, err, &got)

		want := &UnsupportedTypeError{Type: reflect.TypeOf(func() {})}
		assert.Equal(t, want, got)
	})

	t.Run("ignore slice", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type v struct {
			V []struct{} `excel:"v"`
		}

		sheet, err := NewEncoder[v](f, EncoderOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		err = sheet.Encode(&v{V: []struct{}{{}}})
		require.NoError(t, err)
	})
}
