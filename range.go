package gord

import (
	"github.com/devlights/gord/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Range struct {
		gordObj GordObject
		comObj  *ole.IDispatch
	}
)

func newRange(gordObj GordObject, comObj *ole.IDispatch) *Range {
	r := &Range{gordObj: gordObj, comObj: comObj}
	r.Releaser().Add(func() error {
		r.ComObject().Release()
		return nil
	})

	return r
}

func NewRange(document *Document, comObject *ole.IDispatch) *Range {
	return newRange(document, comObject)
}

func NewRangeFromCharacters(c *Characters, comObject *ole.IDispatch) *Range {
	return newRange(c, comObject)
}

func NewRangeFromRange(ra *Range, comObject *ole.IDispatch) *Range {
	return newRange(ra, comObject)
}

func (r *Range) ComObject() *ole.IDispatch {
	return r.comObj
}

func (r *Range) Gord() *Gord {
	return r.gordObj.Gord()
}

func (r *Range) Releaser() *Releaser {
	return r.Gord().Releaser()
}

func (r *Range) Find() (*Find, error) {
	result, err := oleutil.GetProperty(r.ComObject(), "Find")
	if err != nil {
		return nil, err
	}

	f := NewFind(r, result.ToIDispatch())
	return f, nil
}

func (r *Range) Delete() error {
	_, err := oleutil.CallMethod(r.ComObject(), "Delete")
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) Next() (*Range, error) {
	result, err := oleutil.CallMethod(r.ComObject(), "Next")
	if err != nil {
		return nil, err
	}

	newR := NewRangeFromRange(r, result.ToIDispatch())
	return newR, nil
}

func (r *Range) Previous() (*Range, error) {
	result, err := oleutil.CallMethod(r.ComObject(), "Previous")
	if err != nil {
		return nil, err
	}

	newR := NewRangeFromRange(r, result.ToIDispatch())
	return newR, nil
}

func (r *Range) Collapse(direction constants.WdCollapseDirection) error {
	_, err := oleutil.CallMethod(r.ComObject(), "Collapse", int(direction))
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) InsertFile(fileName string) error {
	_, err := oleutil.CallMethod(r.ComObject(), "InsertFile", fileName)
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) InsertBreak(breakType constants.WdBreakType) error {
	_, err := oleutil.CallMethod(r.ComObject(), "InsertBreak", int(breakType))
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) InsertPageBreak() error {
	return r.InsertBreak(constants.WdBreakTypePageBreak)
}

func (r *Range) Copy() error {
	_, err := oleutil.CallMethod(r.ComObject(), "Copy")
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) PasteAndFormat(recoveryType constants.WdRecoveryType) error {
	_, err := oleutil.CallMethod(r.ComObject(), "PasteAndFormat", int(recoveryType))
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) Start() (int32, error) {
	result, err := oleutil.GetProperty(r.ComObject(), "Start")
	if err != nil {
		return 0, err
	}

	return result.Value().(int32), nil
}

func (r *Range) End() (int32, error) {
	result, err := oleutil.GetProperty(r.ComObject(), "End")
	if err != nil {
		return 0, err
	}

	return result.Value().(int32), nil
}

func (r *Range) Text() (string, error) {
	result, err := oleutil.GetProperty(r.ComObject(), "Text")
	if err != nil {
		return "", err
	}

	return result.Value().(string), nil
}

func (r *Range) SetText(s string) error {
	_, err := oleutil.PutProperty(r.ComObject(), "Text", s)
	if err != nil {
		return err
	}

	return nil
}
