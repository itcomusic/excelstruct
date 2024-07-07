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

	sheet, _ := excelstruct.NewWWorkSpace[WriteExcel](f, excelstruct.WWorkSpaceOptions{})
	defer sheet.Close()

	_ = sheet.Encode(&WriteExcel{
		Int:       1,
		String:    "string",
		Slice:     []string{"value1", "value2"},
		Marshaler: &valueMarshaler{value: "marshaler"},
	})
}
