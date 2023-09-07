package gord

import (
	"github.com/devlights/gord/constants"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Document struct {
		gordObj GordObject
		comObj  *ole.IDispatch
	}
)

func NewDocument(docs *Documents, doc *ole.IDispatch) (*Document, ReleaseFunc) {
	d := &Document{
		gordObj: docs,
		comObj:  doc,
	}

	d.Releaser().Add(func() error {
		d.ComObject().Release()
		return nil
	})

	r := func() {
		_ = d.SetSaved(true)
		_ = d.Close()
	}

	return d, r
}

func (w *Document) ComObject() *ole.IDispatch {
	return w.comObj
}

func (w *Document) Gord() *Gord {
	return w.gordObj.Gord()
}

func (w *Document) Releaser() *Releaser {
	return w.Gord().Releaser()
}

func (w *Document) Save() error {
	_, err := oleutil.CallMethod(w.ComObject(), "Save")
	return err
}

func (w *Document) SaveAs2(filePath string, format constants.WdSaveFormat) error {
	return w.SaveAsWithFileFormat(filePath, format)
}

func (w *Document) SaveAsWithFileFormat(filePath string, format constants.WdSaveFormat) error {
	_, err := oleutil.CallMethod(w.ComObject(), "SaveAs2", filePath, int(format))
	return err
}

func (w *Document) SetSaved(value bool) error {
	_, err := oleutil.PutProperty(w.ComObject(), "Saved", value)
	return err
}

func (w *Document) Close() error {
	_, err := oleutil.CallMethod(w.ComObject(), "Close", false)
	return err
}

func (w *Document) PrintOut() error {
	_, err := oleutil.CallMethod(w.ComObject(), "PrintOut", nil)
	return err
}

func (w *Document) ExportAsFixedFormat(path string, fmtType constants.WdExportFormat) error {
	_, err := oleutil.CallMethod(w.ComObject(), "ExportAsFixedFormat", path, int(fmtType))
	return err
}
