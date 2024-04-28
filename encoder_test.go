package excelstruct

import (
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

func TestWriteFile_Encode(t *testing.T) {
	t.Parallel()

	t.Run("simple", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewWWorkSpace[baseType](f, WWorkSpaceOptions{})
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
			{"pdate", "12/7/23 00:00"},
		}, got)
	})

	t.Run("slice", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewWWorkSpace[sliceType](f, WWorkSpaceOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		assert.NoError(t, sheet.Encode(&sliceType{
			Strings:  []string{"hello", "world", ""},
			PStrings: []*string{ptrV("hello"), ptrV("world"), nil},
		}))

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

	t.Run("slice_nil", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewWWorkSpace[sliceType](f, WWorkSpaceOptions{})
		require.NoError(t, err)
		defer sheet.Close()

		assert.NoError(t, sheet.Encode(&sliceType{}))
		got, err := f.File.GetCols(sheet.enc.title.config.sheetName)
		require.NoError(t, err)
		assert.Equal(t, [][]string{
			{"strings"},
			{"pstrings"},
		}, got)
	})

	t.Run("array", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewWWorkSpace[arrayType](f, WWorkSpaceOptions{})
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

		sheet, err := NewWWorkSpace[map[string]any](f, WWorkSpaceOptions{TitleName: []string{"string"}})
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

		sheet, err := NewWWorkSpace[marshalerType](f, WWorkSpaceOptions{})
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

		sheet, err := NewWWorkSpace[baseType](f, WWorkSpaceOptions{})
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
			{"pdate"},
		}, got)
	})
}

func TestWriteFile_ValueConv(t *testing.T) {
	t.Parallel()

	t.Run("bool", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		type v struct {
			V string `excel:"v"`
			B bool   `excel:"b"`
		}

		sheet, err := NewWWorkSpace[v](f, WWorkSpaceOptions{
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

func TestWriteFile_Orientation(t *testing.T) {
	t.Parallel()

	type v struct {
		V []int `excel:"v"`
	}

	t.Run("row", func(t *testing.T) {
		t.Parallel()

		f, err := WriteFile(WriteFileOptions{})
		require.NoError(t, err)
		defer f.Close()

		sheet, err := NewWWorkSpace[v](f, WWorkSpaceOptions{Orientation: OrientationRow})
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

		sheet, err := NewWWorkSpace[v](f, WWorkSpaceOptions{Orientation: OrientationColumn})
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

func TestWriteFile_DataValidation(t *testing.T) {
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

		sheet, err := NewWWorkSpace[v](f, WWorkSpaceOptions{DataValidation: func(title string) (*excelize.DataValidation, error) {
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

		sheet, err := NewWWorkSpace[v](f, WWorkSpaceOptions{
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

		sheet, err := NewWWorkSpace[v](f, WWorkSpaceOptions{
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

func TestWriteFile_Style(t *testing.T) {
	t.Parallel()

	type v struct {
		V int `excelstruct:"v"`
	}

	f, err := WriteFile(WriteFileOptions{})
	require.NoError(t, err)
	defer f.Close()

	sheet, err := NewWWorkSpace[v](f, WWorkSpaceOptions{Style: NameStyle{
		"v": &excelize.Style{Border: []excelize.Border{}},
	}})
	require.NoError(t, err)

	require.NoError(t, sheet.Encode(&v{V: 9}))
	require.NoError(t, sheet.Close())
}

func TestWriteFile_Width(t *testing.T) {
	t.Parallel()

	f, err := WriteFile(WriteFileOptions{
		FilePath: "testdata/width.xlsx",
	})
	require.NoError(t, err)
	defer f.Close()

	type v struct {
		V int `excelstruct:"v"`
	}

	sheet, err := NewWWorkSpace[v](f, WWorkSpaceOptions{
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
