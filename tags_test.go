package excelstruct

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestParseTag(t *testing.T) {
	name, opts := parseTag("field,foobar,foo")
	assert.Equal(t, "field", name)

	for _, tt := range []struct {
		opt  string
		want bool
	}{
		{"foobar", true},
		{"foo", true},
		{"bar", false},
	} {
		assert.Equal(t, tt.want, opts.Contains(tt.opt))
	}
}
