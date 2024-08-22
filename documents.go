package gord

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Documents struct {
		g      *Gord
		comObj *ole.IDispatch
	}
)

func NewDocuments(g *Gord, docs *ole.IDispatch) *Documents {
	d := &Documents{
		g:      g,
		comObj: docs,
	}

	d.Releaser().Add(func() error {
		d.ComObject().Release()
		return nil
	})

	return d
}

func (d *Documents) ComObject() *ole.IDispatch {
	return d.comObj
}

func (d *Documents) Gord() *Gord {
	return d.g
}

func (d *Documents) Releaser() *Releaser {
	return d.Gord().Releaser()
}

func (d *Documents) Add() (*Document, ReleaseFunc, error) {
	dc, err := oleutil.CallMethod(d.ComObject(), "Add")
	if err != nil {
		return nil, nil, err
	}

	doc, releaseFn := NewDocument(d, dc.ToIDispatch())

	return doc, releaseFn, nil
}

func (d *Documents) MustAdd() (*Document, ReleaseFunc) {
	dc, fn, err := d.Add()
	if err != nil {
		panic(err)
	}

	return dc, fn
}

func (d *Documents) Open(filePath string) (*Document, ReleaseFunc, error) {
	dc, err := oleutil.CallMethod(d.ComObject(), "Open", filePath)
	if err != nil {
		return nil, nil, err
	}

	doc, releaseFn := NewDocument(d, dc.ToIDispatch())

	return doc, releaseFn, nil
}

func (d *Documents) MustOpen(filePath string) (*Document, ReleaseFunc) {
	dc, fn, err := d.Open(filePath)
	if err != nil {
		panic(err)
	}

	return dc, fn
}
