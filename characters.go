package gord

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Characters struct {
		gordObj GordObject
		comObj  *ole.IDispatch
	}
)

func newCharacters(gordObj GordObject, comObj *ole.IDispatch) *Characters {
	c := &Characters{gordObj: gordObj, comObj: comObj}
	c.Releaser().Add(func() error {
		c.ComObject().Release()
		return nil
	})

	return c
}

func NewCharacters(document *Document, comObject *ole.IDispatch) *Characters {
	return newCharacters(document, comObject)
}

func (c *Characters) ComObject() *ole.IDispatch {
	return c.comObj
}

func (c *Characters) Gord() *Gord {
	return c.gordObj.Gord()
}

func (c *Characters) Releaser() *Releaser {
	return c.Gord().Releaser()
}

func (c *Characters) Count() (int32, error) {
	result, err := oleutil.GetProperty(c.ComObject(), "Count")
	if err != nil {
		return 0, err
	}

	return result.Value().(int32), nil
}

func (c *Characters) First() (*Range, error) {
	result, err := oleutil.GetProperty(c.ComObject(), "First")
	if err != nil {
		return nil, err
	}

	r := NewRangeFromCharacters(c, result.ToIDispatch())
	return r, nil
}

func (c *Characters) Last() (*Range, error) {
	result, err := oleutil.GetProperty(c.ComObject(), "Last")
	if err != nil {
		return nil, err
	}

	r := NewRangeFromCharacters(c, result.ToIDispatch())
	return r, nil
}
