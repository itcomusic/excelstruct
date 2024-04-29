// Package excelstruct provides a reading and writing XLSX files using structs.
package excelstruct

// ValueMarshaler is the interface implemented by an object that can
// marshal value itself into a string form.
type ValueMarshaler interface {
	MarshalXLSXValue() ([]string, error)
}

// ValueUnmarshaler is the interface implemented by an object that can
// unmarshal value a string representation of itself.
type ValueUnmarshaler interface {
	UnmarshalXLSXValue(value []string) error
}
