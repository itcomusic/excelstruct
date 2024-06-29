package excelstruct

import (
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"strings"
)

// An InvalidUnmarshalError describes an invalid argument passed to Unmarshal.
// (The argument to Unmarshal must be a non-nil pointer.)
type InvalidUnmarshalError struct {
	Type reflect.Type
}

func (e *InvalidUnmarshalError) Error() string {
	if e.Type == nil {
		return "excelstruct: Unmarshal(nil)"
	}

	if e.Type.Kind() != reflect.Pointer {
		return "excelstruct: Unmarshal(non-pointer " + e.Type.String() + ")"
	}
	return "excelstruct: Unmarshal(nil " + e.Type.String() + ")"
}

// An UnmarshalTypeError describes a EXCEL value that was
// not appropriate for a value of a specific Go type.
type UnmarshalTypeError struct {
	Value string       // description of string value - "bool", "array", "number -5"
	Type  reflect.Type // type of Go value it could not be assigned to
	Field string       // the full path from root node to the field
	Err   error        // the error returns convert function string to type
}

func (e *UnmarshalTypeError) Error() string {
	return fmt.Sprintf("excelstruct: cannot unmarshal %q into Go struct field %q of type %s: %v", e.Value, e.Field, e.Type, e.Err)
}

// An ConvertValueError describes a value that was cannot convert to a specific user value.
type ConvertValueError struct {
	Value string
	Field string
	Err   error
}

func (e *ConvertValueError) Error() string {
	return fmt.Sprintf("excelstruct: cannot convert value %q into Go struct field %q: %v", e.Value, e.Field, e.Err)
}

// An UnmarshalError describes an error that was occurred during unmarshal.
type UnmarshalError struct {
	Row int
	Err []error
}

func (e *UnmarshalError) saveError(err error) {
	e.Err = append(e.Err, err)
}

// AsTypeError returns the all UnmarshalTypeError in UnmarshalError.
func (e *UnmarshalError) AsTypeError() []UnmarshalTypeError {
	var res []UnmarshalTypeError
	for _, v := range e.Err {
		if err := new(UnmarshalTypeError); errors.As(v, &err) {
			res = append(res, *err)
		}
	}
	return res
}

// AsConvertValueError returns the all ConvertValueError in UnmarshalError.
func (e *UnmarshalError) AsConvertValueError() []ConvertValueError {
	var res []ConvertValueError
	for _, v := range e.Err {
		if err := new(ConvertValueError); errors.As(v, &err) {
			res = append(res, *err)
		}
	}
	return res
}

// Error returns the all error in UnmarshalError.
func (e *UnmarshalError) Error() string {
	causes := make([]string, 0, 2)
	for _, v := range e.Err {
		causes = append(causes, v.Error())
	}

	message := "excelstruct: unmarshal error: "
	if len(causes) == 0 {
		return message + "no causes"
	}

	return message + strings.Join(causes, ", ")
}

type decOpts struct {
	tag        string
	stringConv ReadStringConv
	boolConv   ReadBoolConv
	timeConv   ReadTimeConv
}

type decodeState struct {
	opts     decOpts
	title    *title
	field    string
	row      int
	colIndex int
	col      []int
}

func (d *decodeState) unmarshal(item []string, v any) error {
	rv := reflect.ValueOf(v)
	if rv.Kind() != reflect.Pointer || rv.IsNil() {
		return &InvalidUnmarshalError{reflect.TypeOf(v)}
	}

	if err := d.value(item, rv); err != nil {
		return err
	}

	d.row++
	return nil
}

func (d *decodeState) value(item []string, v reflect.Value) error {
	if len(item) == 0 {
		return nil
	}

	u, pv := indirect(v)
	if u != nil {
		return u.UnmarshalXLSXValue(item)
	}
	v = pv

	switch v.Kind() {
	case reflect.Slice, reflect.Array:
		return d.array(item, v)

	case reflect.Struct:
		return d.object(item, v)

	case
		reflect.Int,
		reflect.Int8,
		reflect.Int16,
		reflect.Int32,
		reflect.Int64,
		reflect.Uint,
		reflect.Uint8,
		reflect.Uint16,
		reflect.Uint32,
		reflect.Uint64,
		reflect.Float32,
		reflect.Float64,
		reflect.Bool,
		reflect.String:
		return d.literalStore(item[0], v)

	default:
		return &UnmarshalTypeError{Value: item[0], Type: v.Type(), Field: d.field}
	}
}

// literalStore decodes a literal stored in item into v.
func (d *decodeState) literalStore(item string, v reflect.Value) error {
	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		n, err := strconv.ParseInt(item, 10, 64)
		if err != nil {
			return &UnmarshalTypeError{
				Value: item,
				Type:  v.Type(),
				Field: d.field,
				Err:   err,
			}
		}
		v.SetInt(n)

	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		n, err := strconv.ParseUint(item, 10, 64)
		if err != nil {
			return &UnmarshalTypeError{
				Value: item,
				Type:  v.Type(),
				Field: d.field,
				Err:   err,
			}
		}
		v.SetUint(n)

	case reflect.Float32, reflect.Float64:
		n, err := strconv.ParseFloat(item, 64)
		if err != nil {
			return &UnmarshalTypeError{
				Value: item,
				Type:  v.Type(),
				Field: d.field,
				Err:   err,
			}
		}
		v.SetFloat(n)

	case reflect.Bool:
		if d.opts.boolConv != nil {
			b, err := d.opts.boolConv(d.field, item)
			if err != nil {
				return &ConvertValueError{
					Value: item,
					Field: d.field,
					Err:   err,
				}
			}

			v.SetBool(b)
			return nil
		}

		b, err := strconv.ParseBool(item)
		if err != nil {
			return &UnmarshalTypeError{
				Value: item,
				Type:  v.Type(),
				Field: d.field,
				Err:   err,
			}
		}
		v.SetBool(b)

	case reflect.String:
		if d.opts.stringConv != nil {
			s, err := d.opts.stringConv(d.field, item)
			if err != nil {
				return &ConvertValueError{
					Value: item,
					Field: d.field,
					Err:   err,
				}
			}

			v.SetString(s)
			return nil
		}
		v.SetString(item)

	default:
		return fmt.Errorf("unsupported type %q", v.Type().String())
	}
	return nil
}

func (d *decodeState) time(item string, v reflect.Value) error {
	et, err := d.opts.timeConv(item)
	if err != nil {
		return &UnmarshalTypeError{
			Value: item,
			Type:  v.Type(),
			Field: d.field,
			Err:   err,
		}
	}

	v.Set(reflect.ValueOf(et))
	return nil
}

// array consumes an array from d.data decoding into v.
func (d *decodeState) array(item []string, v reflect.Value) error {
	if v.Kind() == reflect.Slice {
		s := reflect.MakeSlice(v.Type(), len(item), len(item))
		v.Set(s)
	}

	for i := range item {
		d.colIndex = i
		if err := d.value([]string{item[i]}, v.Index(i)); err != nil {
			return fmt.Errorf("decode array: %w", err)
		}
	}
	return nil
}

// object consumes an object from d.data[d.off-1:], decoding into v.
func (d *decodeState) object(data []string, v reflect.Value) error {
	t := v.Type()
	if t == timeType {
		return d.time(data[0], v)
	}

	unmarshalError := &UnmarshalError{Row: d.row}
	fields := cachedTypeFields(t, typeOpts{structTag: d.opts.tag})
	for i := range fields.list {
		col, ok := d.title.columnIndex(fields.list[i].name)
		if !ok {
			continue
		}

		item := make([]string, 0, len(col))
		for _, v := range col {
			// empty value doesn't include item column, so skip index out of range
			if v-1 >= len(data) {
				continue
			}

			d := strings.TrimSpace(data[v-1])
			if isEmptyString(d) {
				continue
			}
			item = append(item, d)
		}

		subv := v
		f := &fields.list[i]
		for _, i := range f.index {
			if subv.Kind() == reflect.Pointer {
				if subv.IsNil() {
					// If a struct embeds a pointer to an unexported type,
					// it is not possible to set a newly allocated value
					// since the field is unexported.
					//
					// See https://golang.org/issue/21357
					if !subv.CanSet() {
						unmarshalError.saveError(fmt.Errorf("exelstruct: cannot set embedded pointer to unexported struct: %v", subv.Type().Elem()))
						// Invalidate subv to ensure d.value(subv) skips over
						// the value without assigning it to subv.
						subv = reflect.Value{}
						break
					}
					subv.Set(reflect.New(subv.Type().Elem()))
				}
				subv = subv.Elem()
			}
			subv = subv.Field(i)
		}

		d.field = f.name
		d.col = col
		if err := d.value(item, subv); err != nil {
			unmarshalError.saveError(err)
		}
	}

	if len(unmarshalError.Err) > 0 {
		return unmarshalError
	}
	return nil
}

// indirect walks down v allocating pointers as needed, until it gets to a non-pointer.
// If it encounters an Unmarshaler, indirect stops and returns that.
func indirect(v reflect.Value) (ValueUnmarshaler, reflect.Value) {
	// Issue #24153 indicates that it is generally not a guaranteed property
	// that you may round-trip a reflect.Value by calling Value.Addr().Elem()
	// and expect the value to still be settable for values derived from
	// unexported embedded struct fields.
	//
	// The logic below effectively does this when it first addresses the value
	// (to satisfy possible pointer methods) and continues to dereference
	// subsequent pointers as necessary.
	//
	// After the first round-trip, we set v back to the original value to
	// preserve the original RW flags contained in reflect.Value.
	v0 := v
	haveAddr := false

	// If v is a named type and is addressable,
	// start with its address, so that if the type has pointer methods,
	// we find them.
	if v.Kind() != reflect.Pointer && v.Type().Name() != "" && v.CanAddr() {
		haveAddr = true
		v = v.Addr()
	}

	for {
		// Load value from interface, but only if the result will be usefully addressable.
		if v.Kind() == reflect.Interface && !v.IsNil() {
			e := v.Elem()
			if e.Kind() == reflect.Pointer && !e.IsNil() {
				haveAddr = false
				v = e
				continue
			}
		}

		if v.Kind() != reflect.Pointer {
			break
		}

		// Prevent infinite loop if v is an interface pointing to its own address:
		//     var v interface{}
		//     v = &v
		if v.Elem().Kind() == reflect.Interface && v.Elem().Elem() == v {
			v = v.Elem()
			break
		}

		if v.IsNil() {
			v.Set(reflect.New(v.Type().Elem()))
		}

		if v.Type().NumMethod() > 0 && v.CanInterface() {
			if u, ok := v.Interface().(ValueUnmarshaler); ok {
				return u, reflect.Value{}
			}
		}

		if haveAddr {
			v = v0 // restore original value after round-trip Value.Addr().Elem()
			haveAddr = false
		} else {
			v = v.Elem()
		}
	}
	return nil, v
}

func isEmptyString(v string) bool {
	return v == ""
}
