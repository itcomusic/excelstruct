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

## License

[MIT License](LICENSE)

[build-img]: https://github.com/itcomusic/excelstruct/workflows/build/badge.svg

[build-url]: https://github.com/itcomusic/excelstruct/actions

[pkg-img]: https://pkg.go.dev/badge/github.com/itcomusic/excelstruct.svg

[pkg-url]: https://pkg.go.dev/github.com/itcomusic/excelstruct

[coverage-img]: https://codecov.io/gh/itcomusic/excelstruct/branch/main/graph/badge.svg

[coverage-url]: https://codecov.io/gh/itcomusic/excelstruct