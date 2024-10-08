package main

import (
	"github.com/itcomusic/excelstruct"
)

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
