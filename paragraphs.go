package gord

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Paragraphs struct {
		gordObj GordObject
		comObj  *ole.IDispatch
	}
)

func newParagraphs(gordObj GordObject, comObj *ole.IDispatch) *Paragraphs {
	p := &Paragraphs{gordObj: gordObj, comObj: comObj}
	p.Releaser().Add(func() error {
		p.ComObject().Release()
		return nil
	})

	return p
}

func NewParagraphs(document *Document, comObject *ole.IDispatch) *Paragraphs {
	return newParagraphs(document, comObject)
}

func NewParagraphsFromRange(r *Range, comObject *ole.IDispatch) *Paragraphs {
	return newParagraphs(r, comObject)
}

func (p *Paragraphs) ComObject() *ole.IDispatch {
	return p.comObj
}

func (p *Paragraphs) Gord() *Gord {
	return p.gordObj.Gord()
}

func (p *Paragraphs) Releaser() *Releaser {
	return p.Gord().Releaser()
}

func (p *Paragraphs) Add(r *Range) (*Paragraph, error) {
	var (
		result *ole.VARIANT
		err    error
		para   *Paragraph
	)

	if r == nil {
		result, err = oleutil.CallMethod(p.ComObject(), "Add")
	} else {
		result, err = oleutil.CallMethod(p.ComObject(), "Add", r.ComObject())
	}

	if err != nil {
		return nil, err
	}

	para = NewParagraph(p, result.ToIDispatch())
	return para, nil
}

func (p *Paragraphs) Item(i int32) (*Paragraph, error) {
	result, err := oleutil.CallMethod(p.ComObject(), "Item", i)
	if err != nil {
		return nil, err
	}

	para := NewParagraph(p, result.ToIDispatch())
	return para, nil
}
