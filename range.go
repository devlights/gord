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

func NewRangeFromParagraph(p *Paragraph, comObject *ole.IDispatch) *Range {
	return newRange(p, comObject)
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

func (r *Range) Paragraphs() (*Paragraphs, error) {
	result, err := oleutil.GetProperty(r.ComObject(), "Paragraphs")
	if err != nil {
		return nil, err
	}

	p := NewParagraphsFromRange(r, result.ToIDispatch())
	return p, nil
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

func (r *Range) InsertBefore(text string) error {
	_, err := oleutil.CallMethod(r.ComObject(), "InsertBefore", text)
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) InsertAfter(text string) error {
	_, err := oleutil.CallMethod(r.ComObject(), "InsertAfter", text)
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) InsertParagraphBefore() error {
	_, err := oleutil.CallMethod(r.ComObject(), "InsertParagraphBefore")
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) InsertParagraphAfter() error {
	_, err := oleutil.CallMethod(r.ComObject(), "InsertParagraphAfter")
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) Cut() error {
	_, err := oleutil.CallMethod(r.ComObject(), "Cut")
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) Copy() error {
	_, err := oleutil.CallMethod(r.ComObject(), "Copy")
	if err != nil {
		return err
	}

	return nil
}

func (r *Range) CopyAsPicture() error {
	_, err := oleutil.CallMethod(r.ComObject(), "CopyAsPicture")
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

func (r *Range) PasteSpecial(dataType constants.WdPasteDataType) error {
	var (
		iconIndex     = int32(0)                              // DisplayAsIconがTrueの場合、この引数はIconFilenameで指定されたプログラム・ファイルで使用したいアイコンに対応する番号になります
		link          = false                                 // Trueを指定すると、クリップボードの内容のソース・ファイルへのリンクが作成されます。 デフォルト値は False です。
		placement     = int32(constants.WdOLEPlacementInLine) // WdOLEPlacement 定数 wdFloatOverText または wdInLine のいずれかを指定する。 デフォルト値は wdInLine です。
		displayAsIcon = false                                 // リンクをアイコンとして表示するにはTrueを指定します。 デフォルト値はFalseです。
	)
	_, err := oleutil.CallMethod(r.ComObject(), "PasteSpecial", iconIndex, link, placement, displayAsIcon, int32(dataType))
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

func (r *Range) Information(kind constants.WdInformation) (any, error) {
	result, err := oleutil.GetProperty(r.ComObject(), "Information", int(kind))
	if err != nil {
		return nil, err
	}

	return result.Value(), nil
}

func (r *Range) PageNumber() (int32, error) {
	result, err := r.Information(constants.WdInformationActiveEndPageNumber)
	if err != nil {
		return -1, err
	}

	return result.(int32), nil
}

func (r *Range) LineNo() (int32, error) {
	result, err := r.Information(constants.WdInformationFirstCharacterLineNumber)
	if err != nil {
		return -1, err
	}

	return result.(int32), nil
}
