package excelstruct

import (
	"fmt"
	"reflect"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

var defaultTileConfig = titleConfig{
	tag:       defaultTag,
	sheetName: DefaultSheetName,
	conv:      defaultTitleConv,
	rowIndex:  1,
}

type titleConfig struct {
	tag               string
	rowIndex          int
	sheetName         string
	name              []string
	conv              TitleConv
	maxWidth          TitleMaxWidth
	scaleAutoWidth    ScaleAutoWidth
	dataValidation    DataValidation
	validationOverRow int
	orient            Orientation
	numFmt            map[excelize.CellType]int
	titleNumFmt       map[string]int
	titleStyle        map[string]string
}

// A title is the title of Excel.
type title struct {
	file   *excelize.File
	name   []titleName
	idx    map[string]int
	numFmt map[string]int
	config titleConfig
}

type titleName struct {
	Name   string
	Column []int
	Width  map[int]float64 // map[indexColumn]width, using index-column because column may be shifted

	// RowData is a row where data not empty
	RowData map[int]int // map[indexColumn]row
}

// newTitleFromFile returns the title after reading the Excel.
func newTitleFromFile(config titleConfig, rows *excelize.Rows) (*title, error) {
	i := 1
	// <= because Next() have to put the pointer to the index row
	for ; i <= config.rowIndex; i++ {
		if !rows.Next() {
			return nil, fmt.Errorf("title row index out of range")
		}
	}

	column, err := rows.Columns()
	if err != nil {
		return nil, fmt.Errorf("get title columns: %w", err)
	}

	title := &title{
		name:   make([]titleName, 0, len(column)),
		idx:    make(map[string]int, len(column)),
		config: config,
	}

	for i, name := range column {
		name = title.config.conv(name)
		idx, ok := title.idx[name]
		if ok {
			title.name[idx].Column = append(title.name[idx].Column, i+1)
			continue
		}

		title.name = append(title.name, titleName{
			Name:   name,
			Column: []int{i + 1},
		})
		title.idx[name] = len(title.name) - 1
	}
	return title, nil
}

// newTitleFromStruct returns the title after reading the struct.
func newTitleFromStruct[T any](config titleConfig, file *excelize.File) (*title, error) {
	var (
		inputName []string
		numFmt    map[string]int
	)

	v := reflect.ValueOf(new(T)).Elem()
	switch v.Kind() {
	case reflect.Struct:
		ff := cachedTypeFields(v.Type(), typeOpts{structTag: config.tag}).list
		inputName = make([]string, 0, len(ff))
		nameIndex := make(map[string]int, len(ff))
		for i, v := range ff {
			inputName = append(inputName, v.name)
			nameIndex[v.name] = i
		}

		// struct must have a fields in specific title
		if len(config.name) > 0 {
			exist := make(map[string]struct{}, len(inputName))
			for _, v := range inputName {
				exist[v] = struct{}{}
			}

			for _, v := range config.name {
				if _, ok := exist[v]; !ok {
					return nil, fmt.Errorf("title %q not found in struct tag %q", v, config.tag)
				}
			}
			inputName = config.name
		}

		// numFmt
		numFmt = make(map[string]int, len(inputName))
		for _, name := range inputName {
			// support only time.Time
			switch ff[nameIndex[name]].typ.Kind() {
			case reflect.Struct:
				if ff[nameIndex[name]].typ != timeType {
					continue
				}

				nf, ok := config.numFmt[excelize.CellTypeDate]
				if !ok {
					return nil, fmt.Errorf("numFmt date not found")
				}
				numFmt[name] = nf

				if nf, ok := config.titleNumFmt[name]; ok {
					numFmt[name] = nf
				}
			}
		}

	case reflect.Map:
		if len(config.name) == 0 {
			return nil, fmt.Errorf("type %q not support without title name", v.Type().String())
		}
		inputName = config.name
		numFmt = make(map[string]int, len(inputName))

	default:
		return nil, fmt.Errorf("type %q not support", v.Type().String())
	}

	name := make([]titleName, 0, len(inputName))
	idx := make(map[string]int, len(inputName))
	for i, f := range inputName {
		name = append(name, titleName{
			Name:    f,
			Column:  []int{i + 1},
			Width:   map[int]float64{},
			RowData: map[int]int{0: config.rowIndex}, // pointer to the header
		})
		idx[f] = i
	}

	return &title{
		file:   file,
		name:   name,
		idx:    idx,
		config: config,
		numFmt: numFmt,
	}, nil
}

// sqrefByRow returns the first range data by the row.
func (t *title) sqrefByRow(name string) (string, error) {
	idx, ok := t.idx[name]
	if !ok {
		return "", fmt.Errorf("title %q not found", name)
	}

	title := t.name[idx]
	col := title.Column[0]

	start, err := excelize.CoordinatesToCellName(col, t.config.rowIndex+1, true)
	if err != nil {
		return "", fmt.Errorf("title %q coordinates to start name: %w", name, err)
	}

	end, err := excelize.CoordinatesToCellName(col, title.RowData[0], true)
	if err != nil {
		return "", fmt.Errorf("title %q coordinates to end name: %w", name, err)
	}
	return fmt.Sprintf("%s!%s:%s", t.config.sheetName, start, end), nil
}

// columnIndex returns the column index of the title.
func (t *title) columnIndex(name string) ([]int, bool) {
	if idx, ok := t.idx[name]; ok {
		return t.name[idx].Column, true
	}
	return nil, false
}

// resizeColumnIndex adds the column index of the title.
func (t *title) resizeColumnIndex(name string, lenSlice int) ([]int, error) {
	idx, ok := t.idx[name]
	if !ok {
		return nil, fmt.Errorf("title %q not found", name)
	}

	col := t.name[idx].Column
	if len(col) >= lenSlice {
		return col, nil
	}

	// add column index
	endCol := col[len(col)-1]
	diff := lenSlice - len(col)
	for i := 1; i <= diff; i++ {
		col = append(col, col[len(col)-1]+1)
	}

	if err := t.insertCols(endCol, diff); err != nil {
		return nil, fmt.Errorf("title %q insert column: %w", name, err)
	}

	t.name[idx].Column = col
	if err := t.writeTitle(name); err != nil {
		return nil, fmt.Errorf("title %q write: %w", name, err)
	}

	// shift another column index for right-hand side
	for i := idx + 1; i < len(t.name); i++ {
		for j := range t.name[i].Column {
			t.name[i].Column[j] += diff
		}
	}
	return col, nil
}

// rowData returns the row data of the title.
func (t *title) rowData(name string) (int, bool) {
	if idx, ok := t.idx[name]; ok {
		return t.name[idx].RowData[0], true
	}
	return 0, false
}

// incRowData increments the row data of the title.
func (t *title) incRowData(name string, col int) {
	if idx, ok := t.idx[name]; ok {
		for i, v := range t.name[idx].Column {
			if v != col {
				continue
			}

			t.name[idx].RowData[i]++
			return
		}
	}
}

// setWidth sets the width of column.
func (t *title) setWidth(name string, col int, cell string) error {
	if t.config.scaleAutoWidth == nil {
		return nil
	}

	// cell width
	maxWidth := t.config.maxWidth(name)
	if maxWidth == 0 {
		return nil
	}

	value, err := t.file.GetCellValue(t.config.sheetName, cell)
	if err != nil {
		return fmt.Errorf("cell %q get value: %w", cell, err)
	}

	cellWidth := t.config.scaleAutoWidth(utf8.RuneCountInString(value))
	if maxWidth != -1 && cellWidth > maxWidth {
		cellWidth = maxWidth
	}

	// compare prev
	titleIdx := t.idx[name]
	for idx, v := range t.name[titleIdx].Column {
		if v != col {
			continue
		}

		prev := t.name[titleIdx].Width[idx]
		if prev < cellWidth {
			t.name[titleIdx].Width[idx] = cellWidth
		}
		break
	}
	return nil
}

// insertCols inserts the columns.
func (t *title) insertCols(idx, count int) error {
	col, _ := excelize.ColumnNumberToName(idx)
	if err := t.file.InsertCols(t.config.sheetName, col, count); err != nil {
		return err
	}
	return nil
}

// writeTitle writes the title.
func (t *title) writeTitle(name string) error {
	idx, ok := t.idx[name]
	if !ok {
		return fmt.Errorf("title %q not found", name)
	}

	for _, c := range t.name[idx].Column {
		cell, _ := excelize.CoordinatesToCellName(c, t.config.rowIndex)
		if err := t.file.SetCellStr(t.config.sheetName, cell, t.config.conv(name)); err != nil {
			return fmt.Errorf("cell %q write: %w", cell, err)
		}
	}
	return nil
}

func (t *title) maxRowData() int {
	maxRow := t.config.rowIndex
	for _, n := range t.name {
		for _, v := range n.RowData {
			maxRow = max(maxRow, v)
		}
	}
	return maxRow
}

// writeDataValidation writes the data validation.
func (t *title) writeDataValidation() error {
	if t.config.dataValidation == nil {
		return nil
	}

	maxRow := t.maxRowData()
	if t.config.validationOverRow > 0 {
		maxRow += t.config.validationOverRow
	}

	// no needs a data validation because there is no written data
	if maxRow == t.config.rowIndex {
		return nil
	}

	for _, n := range t.name {
		dv, err := t.config.dataValidation(n.Name)
		if err != nil {
			return fmt.Errorf("title %q data validation func has error: %w", n.Name, err)
		}

		if dv == nil {
			continue
		}

		startCol, endCol := n.Column[0], n.Column[len(n.Column)-1]
		from, _ := excelize.CoordinatesToCellName(startCol, t.config.rowIndex+1) // row: +1 for header
		to, _ := excelize.CoordinatesToCellName(endCol, maxRow)
		cell := from + ":" + to
		dv.Sqref = cell

		if err := t.file.AddDataValidation(t.config.sheetName, dv); err != nil {
			return fmt.Errorf("title %q %q data validation: %w", n.Name, cell, err)
		}
	}
	return nil
}

// writeWidth writes the width of column.
func (t *title) writeWidth() error {
	if t.config.scaleAutoWidth == nil {
		return nil
	}

	for _, n := range t.name {
		if t.config.maxWidth(n.Name) == 0 {
			continue
		}

		for idx, c := range n.Column {
			col, _ := excelize.ColumnNumberToName(c)
			if err := t.file.SetColWidth(t.config.sheetName, col, col, n.Width[idx]); err != nil {
				return fmt.Errorf("title %q column %q width: %w", n.Name, col, err)
			}
		}
	}
	return nil
}

// write writes the titles.
func (t *title) write() error {
	for _, n := range t.name {
		if err := t.writeTitle(n.Name); err != nil {
			return fmt.Errorf("title %q write: %w", n.Name, err)
		}
	}
	return nil
}
