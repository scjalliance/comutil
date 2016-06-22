package comutil

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

var (
	modole32, _ = syscall.LoadDLL("ole32.dll")
)

var (
	procCoCreateInstanceEx, _ = modole32.FindProc("CoCreateInstanceEx")
	procIIDFromString, _      = modole32.FindProc("IIDFromString")
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
