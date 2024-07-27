package excelstruct

import (
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const DefaultSheetName = "Sheet1"

type Orientation string

const (
	OrientationRow    Orientation = "row"
	OrientationColumn Orientation = "col"
)

const (
	defaultTimeFormat = "01-02-06"
)

var defaultCellType = map[excelize.CellType]int{
	excelize.CellTypeDate: 14, // "dd.mm.yyyy"
}

// TitleConv is the function to convert title name.
type TitleConv func(title string) string

// TitleMaxWidth is the function to set in cell max width. (0 - disable, -1 - no limit).
type TitleMaxWidth func(title string) float64

// DataValidation is the function to set in cell data validation.
type DataValidation func(title string) (*excelize.DataValidation, error)

// NameStyle is the naming style to quick search style.
type NameStyle map[string]excelize.Style

// ScaleAutoWidth is a scale function for auto width.
type ScaleAutoWidth func(len int) float64

// WriteStringConv is the function to convert string value.
type WriteStringConv func(title string, v string) (string, error)

// WriteBoolConv is the function to convert bool to string.
type WriteBoolConv func(title string, v bool) (string, error)

// ReadStringConv is the function to convert value to string.
type ReadStringConv func(title string, v string) (string, error)

// ReadBoolConv is the function to convert value to bool.
type ReadBoolConv func(title string, v string) (bool, error)

// ReadTimeConv is the function to convert value to time.Time.
type ReadTimeConv func(v string) (time.Time, error)

// DefaultScaleAutoWidth is the scale function for default font size.
//
// PR: https://github.com/qax-os/excelize/pull/1386
var DefaultScaleAutoWidth = ScaleAutoWidth(func(len int) float64 {
	return float64(len) + 2.0
})

// infinityTitleMaxWith is the scale function for infinity font size.
var infinityTitleMaxWith = TitleMaxWidth(func(_ string) float64 { return -1 })

func defaultTitleConv(title string) string {
	return title
}

func defaultTimeConv(v string) (time.Time, error) {
	et, err := time.Parse(defaultTimeFormat, v)
	if err != nil {
		return time.Time{}, err
	}
	return et, nil
}

// OpenFileOptions is the options for open file.
type OpenFileOptions struct {
	FilePath string
	Style    NameStyle
	Excel    *excelize.Options
}

// WriteFileOptions is the options for write file.
type WriteFileOptions struct {
	FilePath  string
	StructTag string
	Excel     *excelize.Options
}

func (o *WriteFileOptions) initDefault() {
	if o.StructTag == "" {
		o.StructTag = defaultTag
	}
}

// EncoderOptions is the options for write workspace.
type EncoderOptions struct {
	SheetName             string
	TitleRowIndex         int
	TitleName             []string
	DisallowUnknownFields bool
	TitleConv             TitleConv
	TitleMaxWidth         TitleMaxWidth
	TitleScaleAutoWidth   ScaleAutoWidth
	DataValidation        DataValidation
	ValidationOverRow     int
	StringConv            WriteStringConv
	BoolConv              WriteBoolConv
	Orientation           Orientation

	CellNumFmt  map[excelize.CellType]int
	TitleNumFmt map[string]int

	CellStyle  *excelize.Style
	Style      NameStyle
	TitleStyle map[string]string // map[title]style
}

func (o *EncoderOptions) initDefault() {
	if o.SheetName == "" {
		o.SheetName = DefaultSheetName
	}

	if o.TitleRowIndex < 1 {
		o.TitleRowIndex = 1
	}

	if o.TitleConv == nil {
		o.TitleConv = defaultTitleConv
	}

	if o.Orientation == "" {
		o.Orientation = OrientationRow
	}

	// merge with a default cell type
	if len(o.CellNumFmt) == 0 {
		o.CellNumFmt = make(map[excelize.CellType]int, len(defaultCellType))
	}

	for k, v := range defaultCellType {
		if _, ok := o.CellNumFmt[k]; !ok {
			o.CellNumFmt[k] = v
		}
	}

	if o.TitleNumFmt == nil {
		o.TitleNumFmt = make(map[string]int)
	}

	if o.TitleScaleAutoWidth != nil {
		o.TitleMaxWidth = infinityTitleMaxWith
	}

	// case-insensitive name
	nameStyle := make(NameStyle, len(o.Style))
	for k, v := range o.Style {
		nameStyle[strings.ToLower(k)] = v
	}
	o.Style = nameStyle

	titleStyle := make(map[string]string, len(o.TitleStyle))
	for k, v := range o.TitleStyle {
		titleStyle[k] = strings.ToLower(v)
	}
	o.TitleStyle = titleStyle

	// remove duplicate title name
	unique := make(map[string]struct{}, len(o.TitleName))
	for i := 0; i < len(o.TitleName); i++ {
		if _, ok := unique[o.TitleName[i]]; ok {
			o.TitleName = append(o.TitleName[:i], o.TitleName[i+1:]...)
			i--
			continue
		}
		unique[o.TitleName[i]] = struct{}{}
	}
}

// DecoderOptions is the options for read workspace.
type DecoderOptions struct {
	SheetName     string
	TitleRowIndex int
	TitleConv     TitleConv
	StringConv    ReadStringConv
	BoolConv      ReadBoolConv
	TimeConv      ReadTimeConv
	StructTag     string
}

func (o *DecoderOptions) initDefault() {
	if o.SheetName == "" {
		o.SheetName = DefaultSheetName
	}

	if o.TitleRowIndex < 1 {
		o.TitleRowIndex = 1
	}

	if o.TitleConv == nil {
		o.TitleConv = defaultTitleConv
	}

	if o.StructTag == "" {
		o.StructTag = defaultTag
	}

	if o.TimeConv == nil {
		o.TimeConv = defaultTimeConv
	}
}
