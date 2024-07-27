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
