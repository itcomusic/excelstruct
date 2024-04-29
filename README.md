# Excel Go binding struct

[![build-img]][build-url]
[![pkg-img]][pkg-url]
[![coverage-img]][coverage-url]

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
	ID   int    `excel:"id"`
	Name string `excel:"name"`
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
```

```go
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
```

#### Write column oriented helping for sqref and data validation

#### Auto converts data for encoding/decoding

#### Style: write border

## License

[MIT License](LICENSE)

[build-img]: https://github.com/itcomusic/excelstruct/workflows/build/badge.svg

[build-url]: https://github.com/itcomusic/excelstruct/actions

[pkg-img]: https://pkg.go.dev/badge/github.com/itcomusic/excelstruct.svg

[pkg-url]: https://pkg.go.dev/github.com/itcomusic/excelstruct

[coverage-img]: https://codecov.io/gh/itcomusic/excelstruct/branch/main/graph/badge.svg

[coverage-url]: https://codecov.io/gh/itcomusic/excelstruct