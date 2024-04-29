package excelstruct

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestSimpleLetterEqualFold(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name string
		s    []byte
		t    []byte
		want bool
	}{
		{
			name: "equal bytes same case",
			s:    []byte("Hello"),
			t:    []byte("Hello"),
			want: true,
		},
		{
			name: "equal bytes different case",
			s:    []byte("Hello"),
			t:    []byte("HELLO"),
			want: true,
		},
		{
			name: "different bytes",
			s:    []byte("Hello"),
			t:    []byte("World"),
			want: false,
		},
	}

	for _, tt := range tests {
		tt := tt
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()

			assert.Equal(t, tt.want, simpleLetterEqualFold(tt.s, tt.t))
		})
	}
}
