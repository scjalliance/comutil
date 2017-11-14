package comutil

import "errors"

var (
	// ErrCreationFailed is returned when instance creation fails for an
	// unspecified reason. This generally means that NULL was returned from
	// CoCreateInstanceEx but an HRESULT value of S_OK was returned instead of a
	// proper error code. Under normal circumstances with correctly written
	// component object model servers this error should never occur.
	ErrCreationFailed = errors.New("unable to create interface instance")

	// ErrMultiDimArray is returned when a safe array contains more than one
	// one dimension. Multi-dimensional array parsing is not currently supported.
	ErrMultiDimArray = errors.New("attribute contains a multi-dimensional array of values")

	// ErrUnsupportedArray is returned when members of a safe array members are
	// supported for type conversion.
	ErrUnsupportedArray = errors.New("unsupported safe array member type")

	// ErrVariantArray is returned when members of a safe array members were not
	// expected to variants but are.
	ErrVariantArray = errors.New("attribute contains variant array members")

	// ErrNonVariantArray is returned when members of a safe array members were
	// expected to variants but are not.
	ErrNonVariantArray = errors.New("attribute contains non-variant array members")
)
