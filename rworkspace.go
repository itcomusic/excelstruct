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

// RWorkSpace is a workspace for reading data from a file.
type RWorkSpace[T any] struct {
	*excelize.File
	sheetName string
	cursor    *excelize.Rows
	rtype     reflect.Type
	dec       *decodeState
}

// NewRWorkSpace creates a work space with the specified titles and struct.
func NewRWorkSpace[T any](r *Read, opts RWorkSpaceOptions) (*RWorkSpace[T], error) {
	opts.initDefault()

	// read title
	cursor, err := r.Rows(opts.SheetName)
	if err != nil {
		return nil, fmt.Errorf("excel: rows: %w", err)
	}

	title, err := newTitleFromExcel(titleConfig{
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

	return &RWorkSpace[T]{
		File:      r.File,
		sheetName: opts.SheetName,
		cursor:    cursor,
		rtype:     reflect.TypeOf(new(T)).Elem(),
		dec:       decodeState,
	}, nil
}

// Next moves the cursor to the next row.
func (c *RWorkSpace[T]) Next() bool {
	return c.cursor.Next()
}

// Decode decodes the row to the struct.
func (c *RWorkSpace[T]) Decode(res *T) error {
	column, err := c.cursor.Columns()
	if err != nil {
		return fmt.Errorf("excelstruct: get columns: %w", err)
	}

	if err := c.dec.unmarshal(column, res); err != nil {
		return err
	}
	return nil
}

// All decodes all rows to the struct.
func (c *RWorkSpace[T]) All(res *[]T) error {
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
func (c *RWorkSpace[T]) Close() {
	defer c.cursor.Close()
}
