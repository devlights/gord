package gord

import (
	"github.com/devlights/gord/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Find struct {
		gordObj GordObject
		comObj  *ole.IDispatch
	}
)

func newFind(gordObj GordObject, comObj *ole.IDispatch) *Find {
	r := &Find{gordObj: gordObj, comObj: comObj}
	r.Releaser().Add(func() error {
		r.ComObject().Release()
		return nil
	})

	return r
}

func NewFind(r *Range, comObject *ole.IDispatch) *Find {
	return newFind(r, comObject)
}

func (f *Find) ComObject() *ole.IDispatch {
	return f.comObj
}

func (f *Find) Gord() *Gord {
	return f.gordObj.Gord()
}

func (f *Find) Releaser() *Releaser {
	return f.Gord().Releaser()
}

func (f *Find) ClearFormatting() error {
	_, err := oleutil.CallMethod(f.ComObject(), "ClearFormatting")
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) Text(s string) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Text", s)
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) Forward(forward bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Forward", forward)
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) Wrap(t constants.WdFindWrap) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Wrap", int32(t))
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) Replacement() (*Replacement, error) {
	result, err := oleutil.GetProperty(f.ComObject(), "Replacement")
	if err != nil {
		return nil, err
	}

	r := NewReplacement(f, result.ToIDispatch())
	return r, nil
}

func (f *Find) Format(includeFormatting bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "Format", includeFormatting)
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) MatchCase(caseInsensitive bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "MatchCase", caseInsensitive)
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) MatchWholeWord(onlyEntireWords bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "MatchWholeWord", onlyEntireWords)
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) MatchWildcards(on bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "MatchWildcards", on)
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) MatchSoundsLike(on bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "MatchSoundsLike", on)
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) MatchAllWordForms(on bool) error {
	_, err := oleutil.PutProperty(f.ComObject(), "MatchAllWordForms", on)
	if err != nil {
		return err
	}

	return nil
}

func (f *Find) Execute() (bool, error) {
	found, err := oleutil.CallMethod(f.ComObject(), "Execute")
	if err != nil {
		return false, err
	}

	return found.Value().(bool), nil
}

func (f *Find) Found() (bool, error) {
	found, err := oleutil.GetProperty(f.ComObject(), "Found")
	if err != nil {
		return false, err
	}

	return found.Value().(bool), nil
}
