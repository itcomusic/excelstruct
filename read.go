package excelstruct

import (
	"fmt"

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
