//go:build !windows
// +build !windows

package comutil

import (
	"errors"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/google/uuid"
)

var (
	ErrUnsupported = errors.New("call is unsupported on this platform")
)

func CreateInstanceEx(clsid uuid.UUID, context uint, serverInfo *CoServerInfo, results []MultiQI) (err error) {
	return ErrUnsupported
}

func IIDFromString(value string) (iid *ole.GUID, err error) {
	return nil, ErrUnsupported
}

func SafeArrayCopy(original *ole.SafeArray) (duplicate *ole.SafeArray, err error) {
	return nil, ErrUnsupported
}

func SafeArrayCreateVector(variantType ole.VT, lowerBound int32, length uint32) (safearray *ole.SafeArray, err error) {
	return nil, ErrUnsupported
}

func SafeArrayGetElement(safearray *ole.SafeArray, index int32, element unsafe.Pointer) (err error) {
	return ErrUnsupported
}

func SafeArrayPutElement(safearray *ole.SafeArray, index int32, element unsafe.Pointer) (err error) {
	return ErrUnsupported
}

func SafeArrayGetDim(safearray *ole.SafeArray) (dimensions uint32, err error) {
	return 0, ErrUnsupported
}

func BuildVarArrayStr(elements ...string) (v *ole.VARIANT, err error) {
	return nil, ErrUnsupported
}
