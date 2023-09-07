package gord

import "github.com/go-ole/go-ole"

// noinspection GoNameStartsWithPackageName
type (
	HasReleaser interface {
		Releaser() *Releaser
	}

	HasComObject interface {
		ComObject() *ole.IDispatch
	}

	HasGord interface {
		Gord() *Gord
	}

	ComReleaser interface {
		HasReleaser
		HasComObject
	}

	GordObject interface {
		HasGord
		ComReleaser
	}
)
