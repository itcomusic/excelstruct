package excelstruct

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

type Read struct {
	*excelize.File
	style styleID
}

// OpenFile opens a xlsx file.
func OpenFile(opts OpenFileOptions) (*Read, error) {
	var opt []excelize.Options
	if opts.Opts != nil {
		opt = append(opt, *opts.Opts)
	}

	file, err := excelize.OpenFile(opts.FilePath, opt...)
	if err != nil {
		return nil, fmt.Errorf("excelstruct: file: %w", err)
	}

	// init style
	style, err := initStyle(file, opts.Style)
	if err != nil {
		return nil, fmt.Errorf("excelstruct: init style: %w", err)
	}

	return &Read{
		File:  file,
		style: style,
	}, nil
}

func (r *Read) Close() {
	defer r.File.Close()
}
