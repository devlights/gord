package gord

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Replacement struct {
		gordObj GordObject
		comObj  *ole.IDispatch
	}
)

func newReplacement(gordObj GordObject, comObj *ole.IDispatch) *Replacement {
	r := &Replacement{gordObj: gordObj, comObj: comObj}
	r.Releaser().Add(func() error {
		r.ComObject().Release()
		return nil
	})

	return r
}

func NewReplacement(f *Find, comObject *ole.IDispatch) *Replacement {
	return newReplacement(f, comObject)
}

func (r *Replacement) ComObject() *ole.IDispatch {
	return r.comObj
}

func (r *Replacement) Gord() *Gord {
	return r.gordObj.Gord()
}

func (r *Replacement) Releaser() *Releaser {
	return r.Gord().Releaser()
}

func (r *Replacement) Text(s string) error {
	_, err := oleutil.PutProperty(r.ComObject(), "Text", s)
	if err != nil {
		return err
	}

	return nil
}
