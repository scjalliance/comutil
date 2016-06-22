package comutil

import "github.com/go-ole/go-ole"

// CLSCTX MSDN: https://msdn.microsoft.com/en-us/library/ms693716

// CoAuthInfo represents the COAUTHINFO structure expected by the Windows COM
// api. It is a low-level structure that is used behind the scenes by the
// CreateInstance functions.
//
// MSDN: https://msdn.microsoft.com/library/ms688552
type CoAuthInfo struct {
	AuthenticationService uint32
	AuthorizationService  uint32
	ServerPrincipalName   *int16
	AuthenticationLevel   uint32
	ImpersonationLevel    uint32
	AuthIdentityData      *CoAuthIdentity
	Capabilities          uint32
}

// CoAuthInfo represents the COAUTHITDENTIY structure expected by the Windows
// COM api. It is a low-level structure that is used behind the scenes by the
// CreateInstance functions.
//
// MSDN: https://msdn.microsoft.com/library/ms693358
type CoAuthIdentity struct {
	User           *uint16
	UserLength     uint32
	Domain         *uint16
	DomainLength   uint32
	Password       *uint16
	PasswordLength uint32
	Flags          uint32
}

// CoServerInfo represents the COSERVERINFO structure expected by the Windows
// COM api. It is a low-level structure that is used behind the scenes by the
// CreateInstance functions.
//
// MSDN: https://msdn.microsoft.com/library/ms687322
type CoServerInfo struct {
	_        uint32 // reserved
	Name     *int16
	AuthInfo *CoAuthInfo
	_        uint32 // reserved
}

// MultiQI represents the MULTI_QI structure expected by the Windows COM api.
// It is a low-level structure that is used behind the scenes by the
// CreateInstance functions.
//
// MSDN: https://msdn.microsoft.com/library/ms687289
type MultiQI struct {
	IID       *ole.GUID
	Interface *ole.IUnknown
	HR        uintptr
}
