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
