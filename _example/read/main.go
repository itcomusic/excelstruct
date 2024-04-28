package main

import (
	"fmt"

	"github.com/itcomusic/excelstruct"
)

type ReadExcel struct {
	ID   int    `excel:"ID"`
	Name string `excel:"Name"`
}

func main() {
	if models, err := excelstruct.ReadFile[*ReadExcel]("/to/path.xlsx"); err != nil {
		fmt.Println("read excel err:" + err.Error())
	} else {
		fmt.Printf("read excel num: %d\n", len(models))
	}
}
