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

	sheet, _ := excelstruct.NewRWorkSpace[ReadExcel](f, excelstruct.RWorkSpaceOptions{})
	defer sheet.Close()

	var got []ReadExcel
	_ = sheet.All(&got)
	fmt.Println(got)
}
