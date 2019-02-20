package comutil

import (
	"fmt"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/google/uuid"
)

// CreateObject supports local creation of a single component object
// model interface. The class identified by the given class ID will be asked to
// create an instance of the supplied interface ID. If creation fails an error
// will be returned.
//
// It is the caller's responsibility to cast the returned interface to the
// correct type. This is typically done with an unsafe pointer cast.
func CreateObject(clsid uuid.UUID, iid uuid.UUID) (iface *ole.IUnknown, err error) {
	serverInfo := &CoServerInfo{}

	var context uint = ole.CLSCTX_SERVER

	results := make([]MultiQI, 0, 1)
	results = append(results, MultiQI{IID: GUID(iid)})

	err = CreateInstanceEx(clsid, context, serverInfo, results)
	if err != nil {
		return nil, err
	}

	iface = results[0].Interface
	if results[0].HR != ole.S_OK {
		err = ole.NewError(results[0].HR)
	} else if iface == nil {
		err = ErrCreationFailed
	}
	return
}

// CreateRemoteObject supports remote creation of a single component object
// model interface. The class identified by the given class ID will be asked to
// create an instance of the supplied interface ID. If creation fails an error
// will be returned.
//
// If the provided server name is empty, this function will create an instance
// on the local machine. It is then the same as calling CreateObject.
//
// It is the caller's responsibility to cast the returned interface to the
// correct type. This is typically done with an unsafe pointer cast.
func CreateRemoteObject(server string, clsid uuid.UUID, iid uuid.UUID) (iface *ole.IUnknown, err error) {
	var bserver *int16
	if len(server) > 0 {
		bserver = ole.SysAllocStringLen(server)
		if bserver == nil {
			return nil, ole.NewError(ole.E_OUTOFMEMORY)
		}
		defer ole.SysFreeString(bserver)
	}

	serverInfo := &CoServerInfo{
		Name: bserver,
	}

	var context uint
	if server == "" {
		context = ole.CLSCTX_SERVER
	} else {
		context = ole.CLSCTX_REMOTE_SERVER
	}

	results := make([]MultiQI, 0, 1)
	results = append(results, MultiQI{IID: GUID(iid)})

	err = CreateInstanceEx(clsid, context, serverInfo, results)
	if err != nil {
		return nil, err
	}

	iface = results[0].Interface
	if results[0].HR != ole.S_OK {
		err = ole.NewError(results[0].HR)
	} else if iface == nil {
		err = ErrCreationFailed
	}
	return
}

// SafeArrayFromStringSlice creates a SafeArray from the given slice of strings.
//
// See http://www.roblocher.com/whitepapers/oletypes.html
func SafeArrayFromStringSlice(slice []string) *ole.SafeArray {
	array, _ := SafeArrayCreateVector(ole.VT_BSTR, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []string to SAFEARRAY")
	}
	for i, v := range slice {
		element := ole.SysAllocStringLen(v)
		err := SafeArrayPutElement(array, int32(i), unsafe.Pointer(element))
		ole.SysFreeString(element)
		if err != nil {
			panic(err)
		}
	}
	return array
}

// VariantToValue attempts to convert the given variant to a native Go
// representation.
//
// If the value contains an IUnknown or IDispatch interface, it is the
// caller's responsibility to release it.
//
// If the value contains an an array with IUnknown or IDispatch members,
// it is the caller's responsibility to release them.
func VariantToValue(variant *ole.VARIANT) (value interface{}, err error) {
	if array := variant.ToArray(); array != nil {
		return SafeArrayToSlice(array)
	}
	return variant.Value(), nil
}

// SafeArrayToSlice converts the given array to a native Go representation. A
// slice of appropriately typed elements will be returned.
//
// If the array contains IUnknown or IDispatch members, it is the
// caller's responsibility to release them.
func SafeArrayToSlice(array *ole.SafeArrayConversion) (value interface{}, err error) {
	vt, err := array.GetType()
	if err != nil {
		return
	}

	if ole.VT(vt) == ole.VT_VARIANT {
		return SafeArrayToVariantSlice(array)
	}

	return SafeArrayToConcreteSlice(array)
}

// SafeArrayToConcreteSlice converts the given non-variant array to a native Go
// representation. A slice of appropriately typed elements will be returned.
//
// If the array contains variant elements an error will be returned.
//
// Only arrays of integers and bytes are supported. Support for additional
// types may be added in the future.
func SafeArrayToConcreteSlice(array *ole.SafeArrayConversion) (value interface{}, err error) {
	vt, elems, err := arrayDetails(array)
	if err != nil {
		return
	}
	if vt == ole.VT_VARIANT {
		return nil, ErrVariantArray
	}

	switch vt {
	case ole.VT_UI1:
		out := make([]byte, elems)
		for i := int32(0); i < elems; i++ {
			copyArrayElement(array.Array, i, unsafe.Pointer(&out[i]), &err)
		}
		value = out
	case ole.VT_I1:
		out := make([]int8, elems)
		for i := int32(0); i < elems; i++ {
			copyArrayElement(array.Array, i, unsafe.Pointer(&out[i]), &err)
		}
		value = out
	case ole.VT_UI2:
		out := make([]uint16, elems)
		for i := int32(0); i < elems; i++ {
			copyArrayElement(array.Array, i, unsafe.Pointer(&out[i]), &err)
		}
		value = out
	case ole.VT_I2:
		out := make([]int16, elems)
		for i := int32(0); i < elems; i++ {
			copyArrayElement(array.Array, i, unsafe.Pointer(&out[i]), &err)
		}
		value = out
	case ole.VT_UI4:
		out := make([]uint32, elems)
		for i := int32(0); i < elems; i++ {
			copyArrayElement(array.Array, i, unsafe.Pointer(&out[i]), &err)
		}
		value = out
	case ole.VT_I4:
		out := make([]int32, elems)
		for i := int32(0); i < elems; i++ {
			copyArrayElement(array.Array, i, unsafe.Pointer(&out[i]), &err)
		}
		value = out
	case ole.VT_UI8:
		out := make([]uint64, elems)
		for i := int32(0); i < elems; i++ {
			copyArrayElement(array.Array, i, unsafe.Pointer(&out[i]), &err)
		}
		value = out
	case ole.VT_I8:
		out := make([]int64, elems)
		for i := int32(0); i < elems; i++ {
			copyArrayElement(array.Array, i, unsafe.Pointer(&out[i]), &err)
		}
		value = out
	default:
		err = ErrUnsupportedArray
	}
	return
}

// SafeArrayToVariantSlice converts the given variant array to a native Go
// representation. A slice of interface{} members will be returned.
//
// If the array does not contain variant members an error will be returned.
//
// If the array contains IUnknown or IDispatch members, it is the
// caller's responsibility to release them.
func SafeArrayToVariantSlice(array *ole.SafeArrayConversion) (values []interface{}, err error) {
	vt, elems, err := arrayDetails(array)
	if err != nil {
		return
	}
	if vt != ole.VT_VARIANT {
		return nil, ErrNonVariantArray
	}

	for i := int32(0); i < elems; i++ {
		element := &ole.VARIANT{}
		ole.VariantInit(element)
		err = SafeArrayGetElement(array.Array, i, unsafe.Pointer(element))
		if err != nil {
			err = fmt.Errorf("unable to retrieve array element %d: %v", i, err)
			ole.VariantClear(element)
			return
		}
		value, valueErr := VariantToValue(element)
		if valueErr != nil {
			if err == nil {
				err = fmt.Errorf("unable to interpret array element %d: %v", i, valueErr)
			}
		} else {
			values = append(values, value)
		}
		// When returning a pointer to a COM interface, add a reference
		// so that VariantClear doesn't release the interface below.
		switch v := value.(type) {
		case *ole.IDispatch:
			v.AddRef()
		case *ole.IUnknown:
			v.AddRef()
		}
		ole.VariantClear(element)
	}

	return
}

func arrayDetails(array *ole.SafeArrayConversion) (vt ole.VT, elems int32, err error) {
	_vt, err := array.GetType()
	if err != nil {
		return
	}
	vt = ole.VT(_vt)

	dims, _ := SafeArrayGetDim(array.Array) // Error intentionally ignored
	if dims != 1 {
		err = ErrMultiDimArray
		return
	}

	elems, err = array.TotalElements(0)
	return
}

func copyArrayElement(from *ole.SafeArray, index int32, to unsafe.Pointer, err *error) {
	e := SafeArrayGetElement(from, index, to)
	if e != nil && *err == nil {
		*err = fmt.Errorf("unable to retrieve array element %d: %v", index, e)
	}
}
