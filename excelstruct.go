// Package excelstruct provides a reading and writing XLSX files using structs.
package excelstruct

import (
	"fmt"
	"os"
	"reflect"

	"github.com/xuri/excelize/v2"
)

// ValueMarshaler is the interface implemented by an object that can
// marshal value itself into a string form.
type ValueMarshaler interface {
	MarshalXLSXValue() ([]string, error)
}

// ValueUnmarshaler is the interface implemented by an object that can
// unmarshal value a string representation of itself.
type ValueUnmarshaler interface {
	UnmarshalXLSXValue(value []string) error
}

type Read struct {
	*excelize.File
}

// OpenFile opens a xlsx file.
func OpenFile(o OpenFileOptions) (*Read, error) {
	var opts []excelize.Options
	if o.Excel != nil {
		opts = append(opts, *o.Excel)
	}

	file, err := excelize.OpenFile(o.FilePath, opts...)
	if err != nil {
		return nil, fmt.Errorf("excelstruct: file: %w", err)
	}

	return &Read{File: file}, nil
}

// Close closes the file.
func (r *Read) Close() {
	defer r.File.Close()
}

// Decoder is a workspace for reading data from a file.
type Decoder[T any] struct {
	*excelize.File
	sheetName string
	cursor    *excelize.Rows
	rtype     reflect.Type
	dec       *decodeState
}

// NewDecoder creates a decoder with the specified titles and struct.
func NewDecoder[T any](r *Read, opts DecoderOptions) (*Decoder[T], error) {
	opts.initDefault()

	// read title
	cursor, err := r.Rows(opts.SheetName)
	if err != nil {
		return nil, fmt.Errorf("excel: rows: %w", err)
	}

	title, err := newTitleFromFile(titleConfig{
		tag:      opts.StructTag,
		rowIndex: opts.TitleRowIndex,
		conv:     opts.TitleConv,
	}, cursor)
	if err != nil {
		cursor.Close()
		return nil, fmt.Errorf("excelstruct: title: %w", err)
	}

	decodeState := &decodeState{
		opts: decOpts{
			tag:        opts.StructTag,
			stringConv: opts.StringConv,
			boolConv:   opts.BoolConv,
			timeConv:   opts.TimeConv,
		},
		title: title,
		row:   opts.TitleRowIndex + 1, // position the first row of data
	}

	return &Decoder[T]{
		File:      r.File,
		sheetName: opts.SheetName,
		cursor:    cursor,
		rtype:     reflect.TypeFor[T](),
		dec:       decodeState,
	}, nil
}

// Next moves the cursor to the next row.
func (c *Decoder[T]) Next() bool {
	return c.cursor.Next()
}

// Decode decodes the row to the struct.
func (c *Decoder[T]) Decode(res *T) error {
	column, err := c.cursor.Columns()
	if err != nil {
		return fmt.Errorf("excelstruct: get columns: %w", err)
	}

	if err := c.dec.unmarshal(column, res); err != nil {
		return err
	}
	return nil
}

// Count returns the number of rows.
func (c *Decoder[T]) Count() int {
	cursor, err := c.File.Rows(c.sheetName)
	if err != nil {
		return 0
	}
	defer cursor.Close()

	i := 1
	// <= because Next() have to put the pointer to the index row
	for ; i <= c.dec.title.config.rowIndex; i++ {
		if !cursor.Next() {
			return 0
		}
	}

	count := 0
	for cursor.Next() {
		count++
	}
	return count
}

// All decodes all rows to the struct.
func (c *Decoder[T]) All(res *[]T) error {
	for c.cursor.Next() {
		v := reflect.New(c.rtype).Elem().Interface().(T)
		if err := c.Decode(&v); err != nil {
			return err
		}
		*res = append(*res, v)
	}
	return nil
}

// Close closes the cursor.
func (c *Decoder[T]) Close() {
	defer c.cursor.Close()
}

type Write struct {
	*excelize.File
	config writeConfig
}

// WriteFile writes a xlsx file.
func WriteFile(opts WriteFileOptions) (*Write, error) {
	opts.initDefault()

	file, err := openExcelFile(opts.FilePath, opts.Excel)
	if err != nil {
		return nil, fmt.Errorf("excelstruct: open file %q: %w", opts.FilePath, err)
	}

	return &Write{
		File: file,
		config: writeConfig{
			filePath:  opts.FilePath,
			structTag: opts.StructTag,
		},
	}, nil
}

// Close closes the file.
func (w *Write) Close() error {
	defer w.File.Close()
	return w.File.SaveAs(w.config.filePath)
}

// openExcelFile opens or creates a xlsx file.
func openExcelFile(path string, o *excelize.Options) (*excelize.File, error) {
	var opts []excelize.Options
	if o != nil {
		opts = append(opts, *o)
	}

	if _, err := os.Stat(path); err != nil {
		if !os.IsNotExist(err) {
			return nil, fmt.Errorf("stat file: %w", err)
		}
		return excelize.NewFile(opts...), nil
	}
	return excelize.OpenFile(path, opts...)
}

// Encoder is a writing data to a file.
type Encoder[T any] struct {
	*excelize.File
	enc       *encodeState
	cellStyle *excelize.Style
	close     bool
	err       error
	style     NameStyle
}

// NewEncoder creates encoder the specified titles and struct.
func NewEncoder[T any](w *Write, opts EncoderOptions) (*Encoder[T], error) {
	opts.initDefault()

	if _, err := w.File.NewSheet(opts.SheetName); err != nil {
		return nil, fmt.Errorf("excelstruct: new sheet %q: %w", opts.SheetName, err)
	}

	title, err := newTitleFromStruct[T](titleConfig{
		tag:               w.config.structTag,
		rowIndex:          opts.TitleRowIndex,
		sheetName:         opts.SheetName,
		name:              opts.TitleName,
		conv:              opts.TitleConv,
		maxWidth:          opts.TitleMaxWidth,
		scaleAutoWidth:    opts.TitleScaleAutoWidth,
		dataValidation:    opts.DataValidation,
		validationOverRow: opts.ValidationOverRow,
		orient:            opts.Orientation,
		numFmt:            opts.CellNumFmt,
		titleNumFmt:       opts.TitleNumFmt,
		titleStyle:        opts.TitleStyle,
	}, w.File)
	if err != nil {
		return nil, fmt.Errorf("excelstruct: init title: %w", err)
	}

	// write title
	if err := title.write(); err != nil {
		return nil, fmt.Errorf("excelstruct: write title: %w", err)
	}

	return &Encoder[T]{
		File: w.File,
		enc: &encodeState{
			encOpts: encOpts{
				stringConv: opts.StringConv,
				boolConv:   opts.BoolConv,
			},
			typeOpts: typeOpts{
				structTag: w.config.structTag,
			},
			orient:                opts.Orientation,
			disallowUnknownFields: opts.DisallowUnknownFields,
			title:                 title,
			file:                  w.File,
			row:                   opts.TitleRowIndex + 1, // position the first row of data
		},
		cellStyle: opts.CellStyle,
		style:     opts.Style,
	}, nil
}

// Encode writes v to file.
func (e *Encoder[T]) Encode(v *T) error {
	return e.enc.marshal(v)
}

// All writes all values to file.
func (e *Encoder[T]) All(v []T) error {
	for i := range v {
		if err := e.Encode(&v[i]); err != nil {
			return fmt.Errorf("excelstruct: marshal: %w", err)
		}
	}
	return nil
}

// SqrefRow returns the range of the row by title.
func (e *Encoder[T]) SqrefRow(title string) (string, error) {
	return e.enc.title.sqrefByRow(title)
}

// Close applies the style and data validation.
func (e *Encoder[T]) Close() (err error) {
	if e.close {
		return e.err
	}

	defer func() {
		e.close = true
		if err != nil {
			e.err = err
		}
	}()

	if err := e.enc.title.writeDataValidation(); err != nil {
		return fmt.Errorf("write data validation: %w", err)
	}

	if err := e.enc.title.writeWidth(); err != nil {
		return fmt.Errorf("title width: %w", err)
	}

	if err := e.applyCellStyle(); err != nil {
		return fmt.Errorf("global style: %w", err)
	}
	return nil
}

func (e *Encoder[T]) applyCellStyle() error {
	// aligns by max row data
	nextRow := e.enc.title.config.rowIndex
	for _, n := range e.enc.title.name {
		for _, v := range n.RowData {
			if v > nextRow {
				nextRow = v
			}
		}
	}

	for _, t := range e.enc.title.name {
		numFmtID := e.enc.title.numFmt[t.Name]
		var style excelize.Style
		if e.cellStyle != nil {
			style = *e.cellStyle
		}
		style.NumFmt = numFmtID
		if err := e.initStyle(t.Name, &style); err != nil {
			return fmt.Errorf("init style: %w", err)
		}

		styleID, err := e.File.NewStyle(&style)
		if err != nil {
			return fmt.Errorf("new style: %w", err)
		}

		// max column a title
		var maxColumn int
		for _, c := range t.Column {
			if c > maxColumn {
				maxColumn = c
			}
		}

		h, err := excelize.CoordinatesToCellName(t.Column[0], e.enc.title.config.rowIndex)
		if err != nil {
			return fmt.Errorf("coordinates to cell name: %w", err)
		}

		v, err := excelize.CoordinatesToCellName(maxColumn, nextRow)
		if err != nil {
			return fmt.Errorf("coordinates to cell name: %w", err)
		}

		if err := e.SetCellStyle(e.enc.title.config.sheetName, h, v, styleID); err != nil {
			return fmt.Errorf("set cell style: %w", err)
		}
	}
	return nil
}

// initStyle initializes the style by title.
func (e *Encoder[T]) initStyle(title string, style *excelize.Style) error {
	styleName, ok := e.enc.title.config.titleStyle[title]
	if !ok {
		return nil
	}

	s, ok := e.style[styleName]
	if !ok {
		return fmt.Errorf("style %q not found", styleName)
	}

	if s.NumFmt == 0 {
		s.NumFmt = style.NumFmt
	}
	*style = s
	return nil
}

type writeConfig struct {
	filePath  string
	structTag string
}
