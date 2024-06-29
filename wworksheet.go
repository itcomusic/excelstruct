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

// WWorkSpace is a workspace for writing data to a file.
type WWorkSpace[T any] struct {
	*excelize.File
	enc        *encodeState
	cellBorder []excelize.Border
	close      bool
	err        error
}

// NewWWorkSpace creates a work space with the specified titles and struct.
func NewWWorkSpace[T any](w *Write, opts WWorkSpaceOptions) (*WWorkSpace[T], error) {
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
	}, w.File)
	if err != nil {
		return nil, fmt.Errorf("excelstruct: init title: %w", err)
	}

	// write title
	if err := title.write(); err != nil {
		return nil, fmt.Errorf("excelstruct: write title: %w", err)
	}

	return &WWorkSpace[T]{
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
		cellBorder: opts.CellBorder,
	}, nil
}

// Encode writes v to file.
func (w *WWorkSpace[T]) Encode(v *T) error {
	return w.enc.marshal(v)
}

// All writes all values to file.
func (w *WWorkSpace[T]) All(v []T) error {
	for i := range v {
		if err := w.Encode(&v[i]); err != nil {
			return fmt.Errorf("excelstruct: marshal: %w", err)
		}
	}
	return nil
}

// SqrefRow returns the range of the row by title.
func (w *WWorkSpace[T]) SqrefRow(title string) (string, error) {
	return w.enc.title.sqrefByRow(title)
}

// Close applies the style and data validation.
func (w *WWorkSpace[T]) Close() (err error) {
	if w.close {
		return w.err
	}

	defer func() {
		w.close = true
		if err != nil {
			w.err = err
		}
	}()

	if err := w.enc.title.writeDataValidation(); err != nil {
		return fmt.Errorf("write data validation: %w", err)
	}

	if err := w.enc.title.writeWidth(); err != nil {
		return fmt.Errorf("title width: %w", err)
	}

	if err := w.applyCellStyle(); err != nil {
		return fmt.Errorf("global style: %w", err)
	}
	return nil
}

func (w *WWorkSpace[T]) applyCellStyle() error {
	// aligns by max row data
	nextRow := w.enc.title.config.rowIndex
	for _, n := range w.enc.title.name {
		for _, v := range n.RowData {
			if v > nextRow {
				nextRow = v
			}
		}
	}

	for _, t := range w.enc.title.name {
		numFmtID := w.enc.title.numFmt[t.Name]
		styleID, err := w.File.NewStyle(&excelize.Style{
			Border: w.cellBorder,
			NumFmt: numFmtID,
		})
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

		h, err := excelize.CoordinatesToCellName(t.Column[0], w.enc.title.config.rowIndex)
		if err != nil {
			return fmt.Errorf("coordinates to cell name: %w", err)
		}

		v, err := excelize.CoordinatesToCellName(maxColumn, nextRow)
		if err != nil {
			return fmt.Errorf("coordinates to cell name: %w", err)
		}

		if err := w.SetCellStyle(w.enc.title.config.sheetName, h, v, styleID); err != nil {
			return fmt.Errorf("set cell style: %w", err)
		}
	}
	return nil
}
