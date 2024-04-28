package excelstruct

import (
	"fmt"
	"reflect"
	"slices"
	"sort"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

type encoderFunc func(e *encodeState, v reflect.Value, opts encOpts)

var (
	encoderCache sync.Map // map[reflect.Type]encoderFunc
	fieldCache   sync.Map // map[reflect.Type]structFields
)

var (
	timeType           = reflect.TypeOf((*time.Time)(nil)).Elem()
	valueMarshalerType = reflect.TypeOf((*ValueMarshaler)(nil)).Elem()
)

// An UnsupportedTypeError is returned by Marshal when attempting
// to encode an unsupported value type.
type UnsupportedTypeError struct {
	Type reflect.Type
}

func (e *UnsupportedTypeError) Error() string {
	return "excelstruct: unsupported type: " + e.Type.String()
}

// A MarshalerError represents an error from calling a MarshalXLSXValue.
type MarshalerError struct {
	Type       reflect.Type
	Err        error
	sourceFunc string
}

func (e *MarshalerError) Error() string {
	srcFunc := e.sourceFunc
	if srcFunc == "" {
		srcFunc = "MarshalXLSXValue"
	}

	return "excelstruct: error calling " + srcFunc + " for type " + e.Type.String() + ": " + e.Err.Error()
}

// Unwrap returns the underlying error.
func (e *MarshalerError) Unwrap() error { return e.Err }

type encOpts struct {
	stringConv WriteStringConv
	boolConv   WriteBoolConv
}

type typeOpts struct {
	structTag string
	// TODO: deny encode many slices of struct
}

type encodeState struct {
	encOpts               encOpts
	typeOpts              typeOpts
	orient                Orientation
	disallowUnknownFields bool
	title                 *title
	file                  *excelize.File
	field                 string
	row                   int
	col                   int
}

// excelError is an error wrapper type for internal use only.
// Panics with errors are wrapped in excelError so that the top-level recover
// can distinguish intentional panics from this package.
type excelError struct{ error }

func (e *encodeState) marshal(v any) (err error) {
	defer func() {
		if r := recover(); r != nil {
			if je, ok := r.(excelError); ok {
				err = je.error
			} else {
				panic(r)
			}
			return
		}

		if e.orient == OrientationRow {
			e.row++
		}
	}()

	e.reflectValue(reflect.ValueOf(v), e.encOpts)
	return nil
}

func isEmptyValue(v reflect.Value) bool {
	switch v.Kind() {
	case reflect.Array, reflect.Map, reflect.Slice, reflect.String:
		return v.Len() == 0
	case reflect.Bool:
		return !v.Bool()
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return v.Int() == 0
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		return v.Uint() == 0
	case reflect.Float32, reflect.Float64:
		return v.Float() == 0
	case reflect.Interface, reflect.Pointer:
		return v.IsNil()
	}
	return false
}

func (e *encodeState) setField(field string) bool {
	e.field = field
	col, ok := e.title.columnIndex(field)
	if !ok {
		if !e.disallowUnknownFields {
			return false
		}
		e.error(fmt.Errorf("excelstruct: title %q not found", field))
	}
	e.col = col[0]
	return true
}

func (e *encodeState) rangeColumnIndex(v reflect.Value, encElem encoderFunc, opts encOpts) {
	switch e.orient {
	case OrientationRow:
		col, err := e.title.resizeColumnIndex(e.field, v.Len())
		if err != nil {
			e.error(fmt.Errorf("excelstruct: title %q resize column: %w", e.field, err))
		}

		for i := 0; i < v.Len(); i++ {
			e.col = col[i]
			encElem(e, v.Index(i), opts)
		}

	case OrientationColumn:
		col, ok := e.title.columnIndex(e.field)
		if !ok {
			e.error(fmt.Errorf("excelstruct: title %q not found", e.field))
		}
		e.col = col[0]

		for i := 0; i < v.Len(); i++ {
			row, ok := e.title.rowData(e.field)
			if !ok {
				e.error(fmt.Errorf("excelstruct: title %q not found", e.field))
			}

			e.row = row + 1 // +1 shift on next row
			encElem(e, v.Index(i), opts)
		}
	}
}

func (e *encodeState) reflectValue(v reflect.Value, opts encOpts) {
	e.valueEncoder(v)(e, v, opts)
}

// writeValue writes value to cell.
func (e *encodeState) writeValue(value any) {
	cell := e.cell()
	if err := e.file.SetCellValue(e.title.config.sheetName, cell, value); err != nil {
		e.error(fmt.Errorf("excelstruct: field %q set cell value: %w", e.field, err))
	}

	e.title.incRowData(e.field, e.col)
	if err := e.title.setWidth(e.field, e.col, cell); err != nil {
		e.error(fmt.Errorf("excelstruct: field %q set width: %w", e.field, err))
	}
}

// cell returns the cell name.
func (e *encodeState) cell() string {
	cell, _ := excelize.CoordinatesToCellName(e.col, e.row)
	return cell
}

func (e *encodeState) error(err error) {
	panic(excelError{err})
}

func (e *encodeState) valueEncoder(v reflect.Value) encoderFunc {
	if !v.IsValid() {
		return invalidValueEncoder
	}
	return typeEncoder(v.Type(), e.typeOpts)
}

func typeEncoder(t reflect.Type, opts typeOpts) encoderFunc {
	if fi, ok := encoderCache.Load(t); ok {
		return fi.(encoderFunc)
	}

	// To deal with recursive types, populate the map with an
	// indirect func before we build it. This type waits on the
	// real func (f) to be ready and then calls it. This indirect
	// func is only used for recursive types.
	var (
		wg sync.WaitGroup
		f  encoderFunc
	)

	wg.Add(1)
	fi, loaded := encoderCache.LoadOrStore(t, encoderFunc(func(e *encodeState, v reflect.Value, opt encOpts) {
		wg.Wait()
		f(e, v, opt)
	}))
	if loaded {
		return fi.(encoderFunc)
	}

	// compute the real encoder and replace the indirect func with it.
	f = newTypeEncoder(t, opts, true)
	wg.Done()
	encoderCache.Store(t, f)
	return f
}

func newTypeEncoder(t reflect.Type, opts typeOpts, allowAddr bool) encoderFunc {
	if t.Kind() != reflect.Pointer && allowAddr && reflect.PointerTo(t).Implements(valueMarshalerType) {
		return newCondAddrEncoder(addrValueMarshalerEncoder, newTypeEncoder(t, opts, false))
	}

	if t.Implements(valueMarshalerType) {
		return valueMarshalerEncoder
	}

	switch t.Kind() {
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
		return defaultEncoder
	case reflect.Interface:
		return interfaceEncoder
	case reflect.Struct:
		if t == timeType {
			return timeEncoder
		}
		return newStructEncoder(t, opts)
	case reflect.Slice:
		return newSliceEncoder(t, opts)
	case reflect.Array:
		return newArrayEncoder(t, opts)
	case reflect.Map:
		return newMapEncoder(t, opts)
	case reflect.Pointer:
		return newPtrEncoder(t, opts)
	default:
		return unsupportedTypeEncoder
	}
}

func unsupportedTypeEncoder(e *encodeState, v reflect.Value, _ encOpts) {
	e.error(&UnsupportedTypeError{v.Type()})
}

func invalidValueEncoder(_ *encodeState, _ reflect.Value, _ encOpts) {}

func defaultEncoder(e *encodeState, v reflect.Value, opts encOpts) {
	switch v.Kind() {
	case reflect.String:
		if opts.stringConv != nil {
			nv, err := opts.stringConv(e.field, v.String())
			if err != nil {
				e.error(err)
				return
			}

			e.writeValue(nv)
			return
		}

	case reflect.Bool:
		if opts.boolConv != nil {
			nv, err := opts.boolConv(e.field, v.Bool())
			if err != nil {
				e.error(err)
				return
			}

			e.writeValue(nv)
			return
		}
	}
	e.writeValue(v.Interface())
}

func timeEncoder(e *encodeState, v reflect.Value, _ encOpts) {
	e.writeValue(v.Interface())
}

func interfaceEncoder(e *encodeState, v reflect.Value, opts encOpts) {
	if v.IsNil() {
		return
	}
	e.reflectValue(v.Elem(), opts)
}

type structEncoder struct {
	fields structFields
}

type structFields struct {
	list      []field
	nameIndex map[string]int
}

func (se structEncoder) encode(e *encodeState, v reflect.Value, opts encOpts) {
FieldLoop:
	for i := range se.fields.list {
		f := &se.fields.list[i]

		// find the nested struct field by following f.index.
		fv := v
		for _, i := range f.index {
			if fv.Kind() == reflect.Pointer {
				if fv.IsNil() {
					continue FieldLoop
				}
				fv = fv.Elem()
			}
			fv = fv.Field(i)
		}

		if f.omitEmpty && isEmptyValue(fv) {
			continue
		}

		if !e.setField(f.name) {
			continue
		}
		f.encoder(e, fv, opts)
	}
}

func newStructEncoder(t reflect.Type, opts typeOpts) encoderFunc {
	se := structEncoder{fields: cachedTypeFields(t, opts)}
	return se.encode
}

func encodeByteSlice(e *encodeState, v reflect.Value, opts encOpts) {
	if v.IsNil() {
		return
	}

	e.writeValue(v.Interface())
}

// sliceEncoder just wraps an arrayEncoder, checking to make sure the value isn't nil.
type sliceEncoder struct {
	arrayEnc encoderFunc
}

func (se sliceEncoder) encode(e *encodeState, v reflect.Value, opts encOpts) {
	if v.IsNil() {
		return
	}

	se.arrayEnc(e, v, opts)
}

func newSliceEncoder(t reflect.Type, opts typeOpts) encoderFunc {
	// byte slices get special treatment; arrays don't.
	if t.Elem().Kind() == reflect.Uint8 {
		p := reflect.PointerTo(t.Elem())
		if !p.Implements(valueMarshalerType) {
			return encodeByteSlice
		}
	}

	enc := sliceEncoder{newArrayEncoder(t, opts)}
	return enc.encode
}

type arrayEncoder struct {
	elemEnc encoderFunc
}

func (ae arrayEncoder) encode(e *encodeState, v reflect.Value, opts encOpts) {
	e.rangeColumnIndex(v, ae.elemEnc, opts)
}

func newArrayEncoder(t reflect.Type, opts typeOpts) encoderFunc {
	enc := arrayEncoder{typeEncoder(t.Elem(), opts)}
	return enc.encode
}

type mapEncoder struct {
	elemEnc encoderFunc
}

func (me mapEncoder) encode(e *encodeState, v reflect.Value, opts encOpts) {
	if v.IsNil() {
		return
	}

	// key is string
	keys := make([]string, 0, len(v.MapKeys()))
	for _, v := range v.MapKeys() {
		if v.Kind() != reflect.String {
			e.error(fmt.Errorf("excelstruct: map key must be string"))
		}
		keys = append(keys, v.String())
	}
	slices.Sort(keys)

	for _, k := range keys {
		if !e.setField(k) {
			continue
		}
		me.elemEnc(e, v.MapIndex(reflect.ValueOf(k)), opts)
	}
}

func newMapEncoder(t reflect.Type, opts typeOpts) encoderFunc {
	enc := mapEncoder{typeEncoder(t.Elem(), opts)}
	return enc.encode
}

type ptrEncoder struct {
	elemEnc encoderFunc
}

func (pe ptrEncoder) encode(e *encodeState, v reflect.Value, opts encOpts) {
	if v.IsNil() {
		return
	}

	pe.elemEnc(e, v.Elem(), opts)
}

func newPtrEncoder(t reflect.Type, opts typeOpts) encoderFunc {
	enc := ptrEncoder{typeEncoder(t.Elem(), opts)}
	return enc.encode
}

func valueMarshalerEncoder(e *encodeState, v reflect.Value, opts encOpts) {
	if v.Kind() == reflect.Pointer && v.IsNil() {
		return
	}

	m, ok := v.Interface().(ValueMarshaler)
	if !ok {
		return
	}

	s, err := m.MarshalXLSXValue()
	if err != nil {
		e.error(&MarshalerError{v.Type(), err, ""})
	}

	e.rangeColumnIndex(reflect.ValueOf(s), defaultEncoder, opts)
}

func addrValueMarshalerEncoder(e *encodeState, v reflect.Value, opts encOpts) {
	va := v.Addr()
	if va.IsNil() {
		return
	}

	m := va.Interface().(ValueMarshaler)
	s, err := m.MarshalXLSXValue()
	if err != nil {
		e.error(err)
	}

	e.rangeColumnIndex(reflect.ValueOf(s), defaultEncoder, opts)
}

type condAddrEncoder struct {
	canAddrEnc, elseEnc encoderFunc
}

func (ce condAddrEncoder) encode(e *encodeState, v reflect.Value, opts encOpts) {
	if v.CanAddr() {
		ce.canAddrEnc(e, v, opts)
	} else {
		ce.elseEnc(e, v, opts)
	}
}

// newCondAddrEncoder returns an encoder that checks whether its value
// CanAddr and delegates to canAddrEnc if so, else to elseEnc.
func newCondAddrEncoder(canAddrEnc, elseEnc encoderFunc) encoderFunc {
	enc := condAddrEncoder{canAddrEnc: canAddrEnc, elseEnc: elseEnc}
	return enc.encode
}

// A field represents a single field found in a struct.
type field struct {
	name      string
	nameBytes []byte                 // []byte(name)
	equalFold func(s, t []byte) bool // bytes.EqualFold or equivalent

	tag       bool
	index     []int
	typ       reflect.Type
	omitEmpty bool

	encoder encoderFunc
}

// byIndex sorts field by index sequence.
type byIndex []field

func (x byIndex) Len() int { return len(x) }

func (x byIndex) Swap(i, j int) { x[i], x[j] = x[j], x[i] }

func (x byIndex) Less(i, j int) bool {
	for k, xik := range x[i].index {
		if k >= len(x[j].index) {
			return false
		}
		if xik != x[j].index[k] {
			return xik < x[j].index[k]
		}
	}
	return len(x[i].index) < len(x[j].index)
}

// typeFields returns a list of fields that should recognize for the given type.
// The algorithm is breadth-first search over the set of structs to include - the top struct
// and then any reachable anonymous structs.
func typeFields(t reflect.Type, opts typeOpts) structFields {
	// Anonymous fields to explore at the current level and the next.
	current := []field{}
	next := []field{{typ: t}}

	// Count of queued names for current level and the next.
	var count, nextCount map[reflect.Type]int

	// Types already visited at an earlier level.
	visited := map[reflect.Type]bool{}

	// Fields found.
	var fields []field

	for len(next) > 0 {
		current, next = next, current[:0]
		count, nextCount = nextCount, map[reflect.Type]int{}

		for _, f := range current {
			if visited[f.typ] {
				continue
			}
			visited[f.typ] = true

			// Scan f.typ for fields to include.
			for i := 0; i < f.typ.NumField(); i++ {
				sf := f.typ.Field(i)
				tag := sf.Tag.Get(opts.structTag)
				if tag == ignoreField {
					continue
				}

				name, opts := parseTag(tag)
				if !isValidTag(name) {
					name = ""
				}

				anonymous := sf.Anonymous || opts.Contains(optInline)
				if anonymous {
					t := sf.Type
					if t.Kind() == reflect.Pointer {
						t = t.Elem()
					}
					if !sf.IsExported() && t.Kind() != reflect.Struct {
						// Ignore embedded fields of unexported non-struct types.
						continue
					}
					// Do not ignore embedded fields of unexported struct types
					// since they may have exported fields.
				} else if !sf.IsExported() {
					// Ignore unexported non-embedded fields.
					continue
				}

				index := make([]int, len(f.index)+1)
				copy(index, f.index)
				index[len(f.index)] = i

				ft := sf.Type
				if ft.Name() == "" && ft.Kind() == reflect.Pointer {
					// Follow pointer.
					ft = ft.Elem()
				}

				// Record found field and index sequence.
				if name != "" || !anonymous || ft.Kind() != reflect.Struct {
					tagged := name != ""
					if name == "" {
						name = sf.Name
					}

					field := field{
						name:      name,
						tag:       tagged,
						index:     index,
						typ:       ft,
						omitEmpty: opts.Contains(optOmitempty),
					}
					field.nameBytes = []byte(field.name)
					field.equalFold = simpleLetterEqualFold

					fields = append(fields, field)
					if count[f.typ] > 1 {
						// If there were multiple instances, add a second,
						// so that the annihilation code will see a duplicate.
						// It only cares about the distinction between 1 or 2,
						// so don't bother generating any more copies.
						fields = append(fields, fields[len(fields)-1])
					}
					continue
				}

				// Record new anonymous struct to explore in next round.
				nextCount[ft]++
				if nextCount[ft] == 1 {
					next = append(next, field{name: ft.Name(), index: index, typ: ft})
				}
			}
		}
	}

	sort.Slice(fields, func(i, j int) bool {
		x := fields
		// sort field by name, breaking ties with depth, then
		// breaking ties with "name came from excelstruct tag", then
		// breaking ties with an index sequence.
		if x[i].name != x[j].name {
			return x[i].name < x[j].name
		}
		if len(x[i].index) != len(x[j].index) {
			return len(x[i].index) < len(x[j].index)
		}
		if x[i].tag != x[j].tag {
			return x[i].tag
		}
		return byIndex(x).Less(i, j)
	})

	// Delete all fields that are hidden by the Go rules for embedded fields,
	// except that fields with EXELX tags are promoted.

	// The fields are sorted in primary order of name, secondary order
	// of field index length. Loop over names; for each name, delete
	// hidden fields by choosing the one dominant field that survives.
	out := fields[:0]
	for advance, i := 0, 0; i < len(fields); i += advance {
		// One iteration per name.
		// Find the sequence of fields with the name of this first field.
		fi := fields[i]
		name := fi.name
		for advance = 1; i+advance < len(fields); advance++ {
			fj := fields[i+advance]
			if fj.name != name {
				break
			}
		}
		if advance == 1 { // Only one field with this name
			out = append(out, fi)
			continue
		}
		dominant, ok := dominantField(fields[i : i+advance])
		if ok {
			out = append(out, dominant)
		}
	}

	fields = out
	sort.Sort(byIndex(fields))

	for i := range fields {
		f := &fields[i]
		f.encoder = typeEncoder(typeByIndex(t, f.index), opts)
	}
	nameIndex := make(map[string]int, len(fields))
	for i, field := range fields {
		nameIndex[field.name] = i
	}
	return structFields{fields, nameIndex}
}

// dominantField looks through the fields, all of which are known to
// have the same name, to find the single field that dominates the
// others using Go's embedding rules, modified by the presence of
// EXELX tags. If there are multiple top-level fields, the boolean
// will be false: This condition is an error in Go and we skip all
// the fields.
func dominantField(fields []field) (field, bool) {
	// The fields are sorted in increasing index-length order, then by presence of tag.
	// That means that the first field is the dominant one. We need only check
	// for error cases: two fields at top level, either both tagged or neither tagged.
	if len(fields) > 1 && len(fields[0].index) == len(fields[1].index) && fields[0].tag == fields[1].tag {
		return field{}, false
	}
	return fields[0], true
}

func typeByIndex(t reflect.Type, index []int) reflect.Type {
	for _, i := range index {
		if t.Kind() == reflect.Pointer {
			t = t.Elem()
		}
		t = t.Field(i).Type
	}
	return t
}

// cachedTypeFields is like typeFields but uses a cache to avoid repeated work.
func cachedTypeFields(t reflect.Type, opts typeOpts) structFields {
	if f, ok := fieldCache.Load(t); ok {
		return f.(structFields)
	}
	f, _ := fieldCache.LoadOrStore(t, typeFields(t, opts))
	return f.(structFields)
}
