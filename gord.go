package gord

import (
	"fmt"
	"log"
	"runtime"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

var (
	_releaser = NewReleaser()
)

type (
	Gord struct {
		word *ole.IDispatch
	}

	ReleaseFunc func()
)

func InitGord() (func(), error) {
	runtime.LockOSThread()

	return func() {
		runtime.UnlockOSThread()
	}, nil
}

func MustInitGord() func() {
	fn, err := InitGord()
	if err != nil {
		panic(err)
	}

	return fn
}

func NewGord() (*Gord, ReleaseFunc, error) {
	g := new(Gord)

	err := g.init()

	g.Releaser().Add(func() error {
		_ = g.quit()
		_ = g.release()

		return nil
	})

	startReleaserFunc := func() {
		e := g.Releaser().Release()
		if e != nil {
			log.Println(e)
		}
	}

	return g, startReleaserFunc, err
}

func MustNewGord() (*Gord, ReleaseFunc) {
	g, fn, err := NewGord()
	if err != nil {
		panic(err)
	}

	return g, fn
}

func (g *Gord) init() error {
	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		return err
	}

	unknown, err := oleutil.CreateObject("Word.Application")
	if err != nil {
		return err
	}

	word, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}

	g.word = word

	return nil
}

func (g *Gord) quit() error {
	_, err := oleutil.CallMethod(g.ComObject(), "Quit")
	if err != nil {
		log.Println(err)
	}

	return nil
}

func (g *Gord) release() error {
	g.word.Release()
	ole.CoUninitialize()

	return nil
}

func (g *Gord) Gord() *Gord {
	return g
}

func (g *Gord) ComObject() *ole.IDispatch {
	return g.word
}

func (g *Gord) Releaser() *Releaser {
	return _releaser
}

func (g *Gord) EnableEvents() (bool, error) {
	r, err := oleutil.GetProperty(g.ComObject(), "EnableEvents")
	if err != nil {
		return false, err
	}

	enabled, ok := r.Value().(bool)
	if !ok {
		return false, fmt.Errorf("can't cast to bool (EnableEvents)")
	}

	return enabled, nil
}

func (g *Gord) SetEnableEvents(value bool) error {
	_, err := oleutil.PutProperty(g.ComObject(), "EnableEvents", value)
	return err
}

func (g *Gord) ScreenUpdating() (bool, error) {
	r, err := oleutil.GetProperty(g.ComObject(), "ScreenUpdating")
	if err != nil {
		return false, err
	}

	enabled, ok := r.Value().(bool)
	if !ok {
		return false, fmt.Errorf("can't cast to bool (ScreenUpdating)")
	}

	return enabled, nil
}

func (g *Gord) SetScreenUpdating(value bool) error {
	_, err := oleutil.PutProperty(g.ComObject(), "ScreenUpdating", value)
	return err
}

func (g *Gord) DisplayAlerts() (bool, error) {
	r, err := oleutil.GetProperty(g.ComObject(), "DisplayAlerts")
	if err != nil {
		return false, err
	}

	enabled, ok := r.Value().(bool)
	if !ok {
		return false, fmt.Errorf("can't cast to bool (DisplayAlerts)")
	}

	return enabled, nil
}

func (g *Gord) SetDisplayAlerts(value bool) error {
	_, err := oleutil.PutProperty(g.ComObject(), "DisplayAlerts", value)
	return err
}

func (g *Gord) Silent(visible bool) error {
	if err := g.SetDisplayAlerts(false); err != nil {
		return err
	}

	if err := g.SetEnableEvents(false); err != nil {
		return err
	}

	if err := g.SetScreenUpdating(false); err != nil {
		return err
	}

	if err := g.SetVisible(visible); err != nil {
		return err
	}

	return nil
}

func (g *Gord) MustSilent(visible bool) {
	err := g.Silent(visible)
	if err != nil {
		panic(err)
	}
}

func (g *Gord) SetVisible(value bool) error {
	_, err := oleutil.PutProperty(g.ComObject(), "Visible", value)
	return err
}

func (g *Gord) MustSetVisible(value bool) {
	if err := g.SetVisible(value); err != nil {
		panic(err)
	}
}

func (g *Gord) Documents() (*Documents, error) {
	wb, err := oleutil.GetProperty(g.ComObject(), "Documents")
	if err != nil {
		return nil, err
	}

	Documents := NewDocuments(g, wb.ToIDispatch())

	return Documents, nil
}

func (g *Gord) MustDocuments() *Documents {
	wb, err := g.Documents()
	if err != nil {
		panic(err)
	}

	return wb
}

func (g *Gord) ActiveWindow() (*Window, error) {
	w, err := oleutil.GetProperty(g.ComObject(), "ActiveWindow")
	if err != nil {
		return nil, err
	}

	window := NewWindow(g, w.ToIDispatch())

	return window, nil
}

func (g *Gord) ActiveDocument() (*Document, ReleaseFunc, error) {
	wbs, err := g.Documents()
	if err != nil {
		return nil, nil, err
	}

	w, err := oleutil.GetProperty(g.ComObject(), "ActiveDocument")
	if err != nil {
		return nil, nil, err
	}

	Document, wbReleaseFn := NewDocument(wbs, w.ToIDispatch())

	return Document, wbReleaseFn, nil
}
