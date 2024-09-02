package gord

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Paragraph struct {
		gordObj GordObject
		comObj  *ole.IDispatch
	}
)

func newParagraph(gordObj GordObject, comObj *ole.IDispatch) *Paragraph {
	p := &Paragraph{gordObj: gordObj, comObj: comObj}
	p.Releaser().Add(func() error {
		p.ComObject().Release()
		return nil
	})

	return p
}

func NewParagraph(p *Paragraphs, comObject *ole.IDispatch) *Paragraph {
	return newParagraph(p, comObject)
}

func (p *Paragraph) ComObject() *ole.IDispatch {
	return p.comObj
}

func (p *Paragraph) Gord() *Gord {
	return p.gordObj.Gord()
}

func (p *Paragraph) Releaser() *Releaser {
	return p.Gord().Releaser()
}

func (p *Paragraph) Range() (*Range, error) {
	result, err := oleutil.GetProperty(p.ComObject(), "Range")
	if err != nil {
		return nil, err
	}

	r := NewRangeFromParagraph(p, result.ToIDispatch())
	return r, nil
}
