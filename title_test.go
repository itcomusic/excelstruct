package excelstruct

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func TestNewTitleFromExcel(t *testing.T) {
	t.Parallel()

	t.Run("default", func(t *testing.T) {
		t.Parallel()

		file, err := excelize.OpenFile("testdata/title.xlsx")
		require.NoError(t, err)

		cursor, err := file.Rows(DefaultSheetName)
		require.NoError(t, err)
		defer cursor.Close()

		got, err := newTitleFromExcel(titleConfig{rowIndex: 1, conv: defaultTitleConv}, cursor)
		require.NoError(t, err)

		want := &title{
			name: []titleName{
				{
					Name:   "A",
					Column: []int{1},
				},
				{
					Name: "B",
					Column: []int{
						2,
						3,
					},
				},
				{
					Name: "C",
					Column: []int{
						4,
					},
				},
			},
			idx: map[string]int{
				"A": 0,
				"B": 1,
				"C": 2,
			},
		}
		require.Equal(t, want.name, got.name)
		require.Equal(t, want.idx, got.idx)
	})

	t.Run("row index", func(t *testing.T) {
		t.Parallel()

		file, err := excelize.OpenFile("testdata/title.xlsx")
		require.NoError(t, err)

		cursor, err := file.Rows(DefaultSheetName)
		require.NoError(t, err)
		defer cursor.Close()

		got, err := newTitleFromExcel(titleConfig{rowIndex: 2, conv: defaultTitleConv}, cursor)
		require.NoError(t, err)

		want := &title{
			name: []titleName{
				{
					Name:   "C",
					Column: []int{1, 3},
				},
				{
					Name: "D",
					Column: []int{
						2,
						4,
					},
				},
			},
			idx: map[string]int{
				"C": 0,
				"D": 1,
			},
		}
		require.Equal(t, want.name, got.name)
		require.Equal(t, want.idx, got.idx)
	})

	t.Run("conv", func(t *testing.T) {
		t.Parallel()

		file, err := excelize.OpenFile("testdata/title.xlsx")
		require.NoError(t, err)

		cursor, err := file.Rows(DefaultSheetName)
		require.NoError(t, err)
		defer cursor.Close()

		titleAlias := map[string]string{
			"A": "a",
			"B": "b",
			"C": "c",
		}
		got, err := newTitleFromExcel(titleConfig{rowIndex: 1, conv: func(title string) string {
			return titleAlias[title]
		}}, cursor)
		require.NoError(t, err)

		want := &title{
			name: []titleName{
				{
					Name:   "a",
					Column: []int{1},
				},
				{
					Name: "b",
					Column: []int{
						2,
						3,
					},
				},
				{
					Name: "c",
					Column: []int{
						4,
					},
				},
			},
			idx: map[string]int{
				"a": 0,
				"b": 1,
				"c": 2,
			},
		}
		require.Equal(t, want.name, got.name)
		require.Equal(t, want.idx, got.idx)
	})
}

func TestNewTitleFromStruct(t *testing.T) {
	t.Parallel()

	t.Run("default", func(t *testing.T) {
		t.Parallel()

		type in struct {
			A string `excel:"A"`
			B []int  `excel:"B"`
		}

		got, err := newTitleFromStruct[in](titleConfig{tag: defaultTag}, nil)
		require.NoError(t, err)

		want := &title{
			name: []titleName{
				{
					Name:    "A",
					Column:  []int{1},
					Width:   map[int]float64{},
					RowData: map[int]int{0: 0},
				},
				{
					Name:    "B",
					Column:  []int{2},
					Width:   map[int]float64{},
					RowData: map[int]int{0: 0},
				},
			},
			idx: map[string]int{
				"A": 0,
				"B": 1,
			},
		}
		require.Equal(t, want.name, got.name)
		require.Equal(t, want.idx, got.idx)
	})

	t.Run("inline", func(t *testing.T) {
		t.Parallel()

		type in struct {
			A string `excel:"A"`
			B *struct {
				B int `excel:"B"`
			} `excel:",inline"`
			C struct {
				C time.Time `excel:"C"`
			} `excel:",inline"`
		}

		got, err := newTitleFromStruct[in](titleConfig{numFmt: defaultCellType, tag: defaultTag}, nil)
		require.NoError(t, err)

		want := &title{
			name: []titleName{
				{
					Name:    "A",
					Column:  []int{1},
					Width:   map[int]float64{},
					RowData: map[int]int{0: 0},
				},
				{
					Name:    "B",
					Column:  []int{2},
					Width:   map[int]float64{},
					RowData: map[int]int{0: 0},
				},
				{
					Name:    "C",
					Column:  []int{3},
					Width:   map[int]float64{},
					RowData: map[int]int{0: 0},
				},
			},
			idx: map[string]int{
				"A": 0,
				"B": 1,
				"C": 2,
			},
		}
		require.Equal(t, want.name, got.name)
	})
}

func TestTitle_Write(t *testing.T) {
	t.Parallel()

	tit := &title{
		file: excelize.NewFile(),
		name: []titleName{
			{
				Name: "A",
				Column: []int{
					1,
				},
			},
			{
				Name: "B",
				Column: []int{
					2,
				},
			},
		},
		idx:    map[string]int{"A": 0, "B": 1},
		config: defaultTileConfig,
	}
	require.NoError(t, tit.write())
	got, err := tit.file.GetCols(tit.config.sheetName)
	require.NoError(t, err)
	assert.Equal(t, [][]string{{"A"}, {"B"}}, got)
}

func TestTitle_ResizeColumnIndex(t *testing.T) {
	t.Parallel()

	tit := &title{
		file: excelize.NewFile(),
		name: []titleName{
			{
				Name: "A",
				Column: []int{
					1,
				},
			},
			{
				Name: "B",
				Column: []int{
					2,
				},
			},
			{
				Name: "C",
				Column: []int{
					3,
				},
			},
		},
		idx:    map[string]int{"A": 0, "B": 1, "C": 2},
		config: defaultTileConfig,
	}
	require.NoError(t, tit.write())
	got, err := tit.resizeColumnIndex("B", 3)
	require.NoError(t, err)

	assert.Equal(t, []int{2, 3, 4}, got)
	want := &title{
		name: []titleName{
			{
				Name:   "A",
				Column: []int{1},
			},
			{
				Name: "B",
				Column: []int{
					2,
					3,
					4,
				},
			},
			{
				Name: "C",
				Column: []int{
					5,
				},
			},
		},
		idx: map[string]int{
			"A": 0,
			"B": 1,
			"C": 2,
		},
	}
	assert.Equal(t, want.name, tit.name)
	assert.Equal(t, want.idx, tit.idx)

	gotCols, err := tit.file.GetCols(tit.config.sheetName)
	require.NoError(t, err)
	assert.Equal(t, [][]string{{"A"}, {"B"}, {"B"}, {"B"}, {"C"}}, gotCols)
}
