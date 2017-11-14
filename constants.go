package comutil

import "errors"

var (
	// ErrCreationFailed is returned when instance creation fails for an
	// unspecified reason. This generally means that NULL was returned from
	// CoCreateInstanceEx but an HRESULT value of S_OK was returned instead of a
	// proper error code. Under normal circumstances with correctly written
	// component object model servers this error should never occur.
	ErrCreationFailed = errors.New("unable to create interface instance")
)
