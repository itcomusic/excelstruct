package main

import (
	"github.com/itcomusic/excelstruct"
)

type WriteExcel struct {
	ID   int    `excel:"id"`
	Name string `excel:"name"`
}

func main() {
	f, _ := excelstruct.WriteFile(excelstruct.WriteFileOptions{FilePath: "write.xlsx"})
	defer f.Close()

	sheet, _ := excelstruct.NewWWorkSpace[WriteExcel](f, excelstruct.WWorkSpaceOptions{})
	defer sheet.Close()

	_ = sheet.Encode(&WriteExcel{
		ID:   1,
		Name: "Gopher",
	})
}
