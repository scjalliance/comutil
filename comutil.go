package comutil

import (
	"unsafe"

	"github.com/go-ole/go-ole"
)

// CreateObject supports local creation of a single component object
// model interface. The class identified by the given class ID will be asked to
// create an instance of the supplied interface ID. If creation fails an error
// will be returned.
//
// It is the caller's responsibility to cast the returned interface to the
// correct type. This is typically done with an unsafe pointer cast.
func CreateObject(clsid *ole.GUID, iid *ole.GUID) (iface *ole.IUnknown, err error) {
	serverInfo := &CoServerInfo{}

	var context uint = ole.CLSCTX_SERVER

	results := make([]MultiQI, 0, 1)
	results = append(results, MultiQI{IID: iid})

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
func CreateRemoteObject(server string, clsid *ole.GUID, iid *ole.GUID) (iface *ole.IUnknown, err error) {
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
	results = append(results, MultiQI{IID: iid})

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

// SafeArrayFromStringSlice creats a SafeArray from the given slice of strings.
//
// See http://www.roblocher.com/whitepapers/oletypes.html
func SafeArrayFromStringSlice(slice []string) *ole.SafeArray {
	array, _ := SafeArrayCreateVector(ole.VT_BSTR, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []string to SAFEARRAY")
	}
	// SysAllocStringLen(s)
	for i, v := range slice {
		SafeArrayPutElement(array, int64(i), unsafe.Pointer(ole.SysAllocStringLen(v)))
	}
	return array
}
