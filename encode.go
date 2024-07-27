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
