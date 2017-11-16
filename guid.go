package comutil

import (
	ole "github.com/go-ole/go-ole"
	"github.com/google/uuid"
)

// GUID converts the given uuid to a Windows-style guid that can be used by
// api calls. This is often necessary to ensure correct byte order.
func GUID(id uuid.UUID) *ole.GUID {
	return &ole.GUID{
		Data1: uint32(id[0])<<24 | uint32(id[1])<<16 | uint32(id[2])<<8 | uint32(id[3]),
		Data2: uint16(id[4])<<8 | uint16(id[5]),
		Data3: uint16(id[6])<<8 | uint16(id[7]),
		Data4: [8]byte{id[8], id[9], id[10], id[11], id[12], id[13], id[14], id[15]},
	}
}
