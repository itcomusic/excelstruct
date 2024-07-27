# Excel Go binding struct

[![build-img]][build-url]
[![pkg-img]][pkg-url]
[![coverage-img]][coverage-url]

**excelstruct** is a comprehensive Go package that simplifies working with Excel files by allowing you to easily encode and decode structs.
Built on top of the powerful [excelize](https://github.com/qax-os/excelize) library, it offers a solution for dealing with excel data in a structured and type-safe manner.

## Features
- **Type Safety**: Work directly with Go structs, leveraging Go's type system
- **Custom Type Support**: Easily handle custom types with marshaler and unmarshaler interfaces
- **Slice and Array Support**: Encode and decode slices and arrays seamlessly
- **Flexible Type Conversion**: Built-in type conversion options eliminate the need for custom types in many cases
- **Style Support**: Apply Excel border styles and auto alignment

## Installation

Go version 1.21+

```bash
go get github.com/itcomusic/excelstruct
```

## Usage
```go
package main

import (
	"fmt"

	"github.com/itcomusic/excelstruct"
)

type ReadExcel struct {
	Int         int              `excel:"int"`
	String      string           `excel:"string"`
	Slice       []string         `excel:"slice"`
	Unmarshaler valueUnmarshaler `excel:"unmarshaler"`
}

type valueUnmarshaler struct {
	value string
}

func (v *valueUnmarshaler) UnmarshalXLSXValue(value string) error {
	v.value = value
	return nil
}

func main() {
	f, _ := excelstruct.OpenFile(excelstruct.OpenFileOptions{FilePath: "read.xlsx"})
	defer f.Close()

	sheet, _ := excelstruct.NewDecoder[ReadExcel](f, excelstruct.DecoderOptions{})
	defer sheet.Close()

	var got []ReadExcel
	_ = sheet.All(&got)
	fmt.Println(got)
}
```

```go
package main

import (
	"github.com/itcomusic/excelstruct"
)

type WriteExcel struct {
	Int       int             `excel:"int"`
	String    string          `excel:"string"`
	Slice     []string        `excel:"slice"`
	Marshaler *valueMarshaler `excel:"marshaler"`
}

type valueMarshaler struct {
	value string
}

func (v *valueMarshaler) MarshalXLSXValue() ([]string, error) {
	return []string{v.value}, nil
}

func main() {
	f, _ := excelstruct.WriteFile(excelstruct.WriteFileOptions{FilePath: "write.xlsx"})
	defer f.Close()

	sheet, _ := excelstruct.NewEncoder[WriteExcel](f, excelstruct.EncoderOptions{})
	defer sheet.Close()

	_ = sheet.Encode(&WriteExcel{
		Int:       1,
		String:    "string",
		Slice:     []string{"value1", "value2"},
		Marshaler: &valueMarshaler{value: "marshaler"},
	})
}
```

#### Write column oriented helping for sqref and data validation
```go
package main

import (
	"fmt"

	"github.com/itcomusic/excelstruct"
	"github.com/xuri/excelize/v2"
)

const columnName = "name"

type WriteExcel struct {
	ID   int    `excel:"id"`
	Name string `excel:"name"`
}

func main() {
	f, _ := excelstruct.WriteFile(excelstruct.WriteFileOptions{FilePath: "dv.xlsx"})
	defer f.Close()

	list, _ := excelstruct.NewEncoder[map[string][]string](f, excelstruct.EncoderOptions{
		SheetName:   "list",
		TitleName:   []string{columnName},
		Orientation: excelstruct.OrientationColumn,
	})
	defer list.Close()
	_ = list.Encode(&map[string][]string{columnName: {"Gopher", "Rob Pike"}})

	sheet, _ := excelstruct.NewEncoder[WriteExcel](f, excelstruct.EncoderOptions{
		DataValidation: dataValidation(list),
	})
	defer sheet.Close()

	_ = sheet.Encode(&WriteExcel{
		ID:   1,
		Name: "gopher",
	})
}

func dataValidation(sheet *excelstruct.Encoder[map[string][]string]) func(title string) (*excelize.DataValidation, error) {
	return func(title string) (*excelize.DataValidation, error) {
		switch title {
		case columnName:
			dv := excelize.NewDataValidation(true)
			dv.ShowInputMessage = true
			dv.SetError(excelize.DataValidationErrorStyleStop, "Must select a value from the list", "Value not found")

			sqref, err := sheet.SqrefRow(title)
			if err != nil {
				return nil, fmt.Errorf("sqref row: %w", err)
			}
			dv.SetSqrefDropList(sqref)
			return dv, nil
		}
		return nil, nil
	}
}
```

#### Write custom styles to excel cells
This includes setting borders, alignment, and other cell formatting options to enhance the appearance of your Excel sheets.
```go
package main

import (
	"time"

	"github.com/xuri/excelize/v2"

	"github.com/itcomusic/excelstruct"
)

type WriteExcel struct {
	Int    int       `excel:"i"`
	String string    `excel:"s"`
	Bool   bool      `excel:"b"`
	Time   time.Time `excel:"t"`
}

var border = []excelize.Border{
	{
		Type:  "left",
		Color: "000000",
		Style: 1,
	},
	{
		Type:  "top",
		Color: "000000",
		Style: 1,
	},
	{
		Type:  "bottom",
		Color: "000000",
		Style: 1,
	},
	{
		Type:  "right",
		Color: "000000",
		Style: 1,
	},
}

func main() {
	f, _ := excelstruct.WriteFile(excelstruct.WriteFileOptions{FilePath: "write.xlsx"})
	defer f.Close()

	sheet, _ := excelstruct.NewEncoder[WriteExcel](f, excelstruct.EncoderOptions{
		TitleScaleAutoWidth: excelstruct.DefaultScaleAutoWidth,
		CellStyle:           &excelize.Style{Border: border},
		Style: excelstruct.NameStyle{
			"center": excelize.Style{
				// dotting
				Border: []excelize.Border{
					{
						Type:  "left",
						Color: "000000",
						Style: 6,
					},
					{
						Type:  "top",
						Color: "000000",
						Style: 6,
					},
					{
						Type:  "bottom",
						Color: "000000",
						Style: 6,
					},
					{
						Type:  "right",
						Color: "000000",
						Style: 6,
					},
				},
				Alignment: &excelize.Alignment{
					Horizontal:      "center",
					Indent:          1,
					JustifyLastLine: true,
					ReadingOrder:    0,
					RelativeIndent:  1,
					ShrinkToFit:     false,
					Vertical:        "center",
				},
			},
		},
		TitleStyle: map[string]string{
			"s": "center",
			"t": "center",
		},
	})
	defer sheet.Close()

	_ = sheet.Encode(&WriteExcel{
		Int:    1,
		String: "string",
		Bool:   true,
		Time:   time.Now(),
	})
}
```

## License

[MIT License](LICENSE)

[build-img]: https://github.com/itcomusic/excelstruct/workflows/build/badge.svg

[build-url]: https://github.com/itcomusic/excelstruct/actions

[pkg-img]: https://pkg.go.dev/badge/github.com/itcomusic/excelstruct.svg

[pkg-url]: https://pkg.go.dev/github.com/itcomusic/excelstruct

[coverage-img]: https://codecov.io/gh/itcomusic/excelstruct/branch/main/graph/badge.svg

[coverage-url]: https://codecov.io/gh/itcomusic/excelstruct