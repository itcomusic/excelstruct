# Excel Go binding struct

[![build-img]][build-url]
[![pkg-img]][pkg-url]
[![coverage-img]][coverage-url]

**excelstruct** is a comprehensive Go package that simplifies working with Excel files by allowing you to easily encode and decode structs [example](https://github.com/itcomusic/excelstruct/tree/main/_example).
Built on top of the powerful [excelize](https://github.com/qax-os/excelize) library, it offers a solution for dealing with excel data in a structured and type-safe manner.

## Features
- **Type Safety**: Work directly with Go structs, leveraging Go's type system
- **Custom Type Support**: Easily handle custom types with marshaler and unmarshaler interfaces
- **Slice and Array Support**: Encode and decode slices and arrays seamlessly
- **Flexible Type Conversion**: Built-in type conversion options eliminate the need for custom types in many cases
- **Style Support**: Apply Excel styles
- **Data validation**: Add data validation and column oriented helping for sqref

## Installation

Go version 1.22+

```bash
go get github.com/itcomusic/excelstruct
```

## Usage

```go
type ReadExcel struct {
    Int         int              `excel:"i"`
    String      string           `excel:"s"`
    Slice       []string         `excel:"a"`
    Unmarshaler valueUnmarshaler `excel:"m"`
}

type valueUnmarshaler struct {
    value string
}

func (v *valueUnmarshaler) UnmarshalXLSXValue(value []string) error {
    v.value = value[0]
    return nil
}

func main() {
    f, _ := excelstruct.OpenFile(excelstruct.OpenFileOptions{FilePath: "read.xlsx"})
    defer f.Close()

    sheet, _ := excelstruct.NewDecoder[ReadExcel](f, excelstruct.DecoderOptions{})
    defer sheet.Close()

    var got []ReadExcel
    _ = sheet.All(&got)
    fmt.Printf("%+v, got) // [{Int:1 String:str Slice:[value1 value2] Unmarshaler:{value:unmarshaler}}]
}

```

#### Write with custom styles to excel
Setting borders, alignment, and other cell formatting options to enhance the appearance of your Excel sheets.
```go
type WriteExcel struct {
    Int       int             `excel:"i"`
    String    string          `excel:"s"`
    Slice     []string        `excel:"a"`
    Marshaler *valueMarshaler `excel:"m"`
}

type valueMarshaler struct {
    value string
}

func (v *valueMarshaler) MarshalXLSXValue() ([]string, error) {
    return []string{v.value}, nil
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
    CellStyle: &excelize.Style{Border: border},
        Style: excelstruct.NameStyle{
        "center": excelize.Style{
            Border: []excelize.Border{ // double line
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
        },
    })
    defer sheet.Close()
	
    _ = sheet.Encode(&WriteExcel{
        Int:       1,
        String:    "string",
        Slice:     []string{"value1", "value2"},
        Marshaler: &valueMarshaler{value: "marshaler"},
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