package excelstruct

import (
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

type writeConfig struct {
	filePath  string
	structTag string
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
