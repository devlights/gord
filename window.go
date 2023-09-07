package gord

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type (
	Window struct {
		g *Gord
		w *ole.IDispatch
	}
)

func NewWindow(g *Gord, w *ole.IDispatch) *Window {
	win := &Window{
		g: g,
		w: w,
	}

	win.Releaser().Add(func() error {
		win.ComObject().Release()
		return nil
	})

	return win
}

func (w *Window) ComObject() *ole.IDispatch {
	return w.w
}

func (w *Window) Gord() *Gord {
	return w.g
}

func (w *Window) Releaser() *Releaser {
	return w.Gord().Releaser()
}

func (w *Window) SetZoom(zoomRate int) error {
	_, err := oleutil.PutProperty(w.ComObject(), "Zoom", zoomRate)
	return err
}
