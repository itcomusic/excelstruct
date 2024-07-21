package excelstruct

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

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
