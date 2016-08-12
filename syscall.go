package comutil

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

var (
	modole32, _    = syscall.LoadDLL("ole32.dll")
	modoleaut32, _ = syscall.LoadDLL("oleaut32.dll")
)

var (
	procCoCreateInstanceEx, _    = modole32.FindProc("CoCreateInstanceEx")
	procIIDFromString, _         = modole32.FindProc("IIDFromString")
	procSafeArrayCopy, _         = modoleaut32.FindProc("SafeArrayCopy")
	procSafeArrayCreateVector, _ = modoleaut32.FindProc("SafeArrayCreateVector")
	procSafeArrayGetElement, _   = modoleaut32.FindProc("SafeArrayGetElement")
	procSafeArrayPutElement, _   = modoleaut32.FindProc("SafeArrayPutElement")
	procSafeArrayGetDim, _       = modoleaut32.FindProc("SafeArrayGetDim")
)

// CreateInstanceEx supports remote creation of multiple interfaces within one
// class.
//
// This is a low-level function. Use of the higher level object creation
// functions like CreateObject and CreateRemoteObject are recommended unless
// specific creation parameters are required.
//
// MSDN: https://msdn.microsoft.com/library/ms680701
func CreateInstanceEx(clsid *ole.GUID, context uint, serverInfo *CoServerInfo, results []MultiQI) (err error) {
	var _p0 *MultiQI
	if len(results) > 0 {
		_p0 = &results[0]
	}
	hr, _, _ := procCoCreateInstanceEx.Call(
		uintptr(unsafe.Pointer(clsid)),
		0,
		uintptr(context),
		uintptr(unsafe.Pointer(serverInfo)),
		uintptr(len(results)),
		uintptr(unsafe.Pointer(_p0)))
	if hr != 0 {
		err = ole.NewError(hr)
	}
	return
}

// IIDFromString takes the given value and attempts to convert it into a valid
// GUID. If it fails it returns an error. It does not provide any additional
// validation, such as checking the Windows registry for its registration.
//
// It is safe to use this function to parse any GUID, not just COM interface
// identifiers.
//
// MSDN: https://msdn.microsoft.com/library/ms687262
// Raymond Chen: https://blogs.msdn.microsoft.com/oldnewthing/20151015-00/?p=91351
func IIDFromString(value string) (iid *ole.GUID, err error) {
	bvalue := ole.SysAllocStringLen(value)
	if bvalue == nil {
		return nil, ole.NewError(ole.E_OUTOFMEMORY)
	}
	defer ole.SysFreeString(bvalue)
	iid = new(ole.GUID)
	hr, _, _ := procIIDFromString.Call(
		uintptr(unsafe.Pointer(bvalue)),
		uintptr(unsafe.Pointer(iid)))
	if hr != 0 {
		err = ole.NewError(hr)
	}
	return
}

// SafeArrayCopy returns a copy of the given SafeArray.
//
// AKA: SafeArrayCopy in Windows API.
func SafeArrayCopy(original *ole.SafeArray) (duplicate *ole.SafeArray, err error) {
	hr, _, _ := procSafeArrayCopy.Call(
		uintptr(unsafe.Pointer(original)),
		uintptr(unsafe.Pointer(&duplicate)))
	if hr != 0 {
		err = ole.NewError(hr)
	}
	return
}

// SafeArrayCreateVector creates SafeArray.
//
// AKA: SafeArrayCreateVector in Windows API.
func SafeArrayCreateVector(variantType ole.VT, lowerBound int32, length uint32) (safearray *ole.SafeArray, err error) {
	sa, _, err := procSafeArrayCreateVector.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length))
	safearray = (*ole.SafeArray)(unsafe.Pointer(uintptr(sa)))
	return
}

// SafeArrayGetElement stores the data element at the specified location in the
// array.
//
// AKA: SafeArrayGetElement in Windows API.
func SafeArrayGetElement(safearray *ole.SafeArray, index int64, element unsafe.Pointer) (err error) {
	hr, _, _ := procSafeArrayGetElement.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(&index)),
		uintptr(element))
	if hr != 0 {
		err = ole.NewError(hr)
	}
	return
}

// SafeArrayPutElement stores the data element at the specified location in the
// array.
//
// AKA: SafeArrayPutElement in Windows API.
func SafeArrayPutElement(safearray *ole.SafeArray, index int64, element unsafe.Pointer) (err error) {
	hr, _, _ := procSafeArrayPutElement.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(&index)),
		uintptr(element))
	if hr != 0 {
		err = ole.NewError(hr)
	}
	return
}

// SafeArrayGetDim returns the number of dimensions in the given safe array.
//
// AKA: SafeArrayGetDim in Windows API.
func SafeArrayGetDim(safearray *ole.SafeArray) (dimensions uint32, err error) {
	d, _, err := procSafeArrayGetDim.Call(uintptr(unsafe.Pointer(safearray)))
	dimensions = uint32(d)
	return
}
