package main

import (
	"fmt"

	"github.com/itcomusic/excelstruct"
)

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
	fmt.Printf("%+v", got)
}
