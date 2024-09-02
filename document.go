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

func (d *Document) ComObject() *ole.IDispatch {
	return d.comObj
}

func (d *Document) Gord() *Gord {
	return d.gordObj.Gord()
}

func (d *Document) Releaser() *Releaser {
	return d.Gord().Releaser()
}

func (d *Document) Content() (*Range, error) {
	result, err := oleutil.GetProperty(d.ComObject(), "Content")
	if err != nil {
		return nil, err
	}

	r := NewRange(d, result.ToIDispatch())
	return r, nil
}

func (d *Document) AllRange() (*Range, error) {
	result, err := oleutil.CallMethod(d.ComObject(), "Range")
	if err != nil {
		return nil, err
	}

	r := NewRange(d, result.ToIDispatch())
	return r, nil
}

func (d *Document) Range(start, end int32) (*Range, error) {
	result, err := oleutil.CallMethod(d.ComObject(), "Range", start, end)
	if err != nil {
		return nil, err
	}

	r := NewRange(d, result.ToIDispatch())
	return r, nil
}

func (d *Document) Characters() (*Characters, error) {
	result, err := oleutil.GetProperty(d.ComObject(), "Characters")
	if err != nil {
		return nil, err
	}

	c := NewCharacters(d, result.ToIDispatch())
	return c, nil
}

func (d *Document) Save() error {
	_, err := oleutil.CallMethod(d.ComObject(), "Save")
	return err
}

func (d *Document) SaveAs2(filePath string, format constants.WdSaveFormat) error {
	return d.SaveAsWithFileFormat(filePath, format)
}

func (d *Document) SaveAsWithFileFormat(filePath string, format constants.WdSaveFormat) error {
	_, err := oleutil.CallMethod(d.ComObject(), "SaveAs2", filePath, int(format))
	return err
}

func (d *Document) SaveAsWithFileFormatDefault(filePath string) error {
	_, err := oleutil.CallMethod(d.ComObject(), "SaveAs2", filePath, int(constants.WdSaveFormatDocumentDefault))
	return err
}

func (d *Document) SetSaved(value bool) error {
	_, err := oleutil.PutProperty(d.ComObject(), "Saved", value)
	return err
}

func (d *Document) Close() error {
	_, err := oleutil.CallMethod(d.ComObject(), "Close", false)
	return err
}

func (d *Document) PrintOut() error {
	_, err := oleutil.CallMethod(d.ComObject(), "PrintOut", nil)
	return err
}

//goland:noinspection GoSnakeCaseUsage,GoBoolExpressions
func (d *Document) ExportAsFixedFormat(path string, fmtType constants.WdExportFormat) error {
	var (
		outputFileName     = path                                                         // PDF ファイルまたは XPS ファイルのパスとファイル名
		exportFormat       = int(fmtType)                                                 // PDF 形式または XPS 形式 (WdExportFormat)
		openAfterExport    = false                                                        // コンテンツをエクスポートした後で新しいファイルを開くかどうか
		optimizeFor        = int(constants.WdExportOptimizeForPrint)                      // 画面または印刷用に最適化するかどうか
		exportRange        = int(constants.WdExportRangeAllDocument)                      // エクスポートする範囲 (WdExportRange)
		from               = 0                                                            // exportRange パラメーターが wdExportFromTo に設定されている場合は、開始ページ番号を指定
		to                 = 0                                                            // Range パラメーターが wdExportFromTo に設定されている場合は、終了ページ番号を指定
		item               = int(constants.WdExportItemDocumentContent)                   // エクスポート プロセスにテキストのみを含めるか、テキストとマークアップ コードを含めるかを指定 (WdExportItem)
		includeDocProps    = true                                                         // 新たにエクスポートするファイルに文書のプロパティを含めるかどうかを指定
		keepIRM            = true                                                         // ExportFormat が wdExportFormatPDF の場合、このフラグはラベルを PDF にコピーするかどうかを指定
		createBookmarks    = int(constants.WdExportCreateBookmarksCreateHeadingBookmarks) // ブックマークをエクスポートするかどうか、およびエクスポートするブックマークの種類を指定 (WdExportCreateBookmarks)
		docStructureTags   = true                                                         // フローとコンテンツの論理的な構成に関する情報など、スクリーン リーダーのために余分なデータを含めるかどうかを指定
		bitmapMissingFonts = true                                                         // テキストのビットマップを含めるかどうかを指定
		useISO19005_1      = false                                                        // ISO 19005-1 として標準化された PDF サブセットに PDF の使用を制限するかどうかを指定
	)
	_, err := oleutil.CallMethod(
		d.ComObject(),
		"ExportAsFixedFormat",
		outputFileName,
		exportFormat,
		openAfterExport,
		optimizeFor,
		exportRange,
		from,
		to,
		item,
		includeDocProps,
		keepIRM,
		createBookmarks,
		docStructureTags,
		bitmapMissingFonts,
		useISO19005_1)

	return err
}

func (d *Document) GetTotalPages() (int32, error) {
	ra, err := d.AllRange()
	if err != nil {
		return 0, err
	}

	v, err := ra.Information(constants.WdInformationNumberOfPagesInDocument)
	if err != nil {
		return 0, err
	}

	return v.(int32), nil
}

func (d *Document) GoTo(what constants.WdGoToItem, which constants.WdGoToDirection, count int32) (*Range, error) {
	result, err := oleutil.CallMethod(d.ComObject(), "GoTo", int32(what), int32(which), count)
	if err != nil {
		return nil, err
	}

	ra := NewRange(d, result.ToIDispatch())
	return ra, nil
}

func (d *Document) GetPageRange(page int32) (*Range, error) {
	var (
		allRange   *Range
		startRange *Range
		endRange   *Range
		result     *Range

		startPos   int32
		endPos     int32
		totalPages int32

		err error
	)

	startRange, err = d.GoTo(constants.WdGoToItemGoToPage, constants.WdGoToDirectionGoToAbsolute, page)
	if err != nil {
		return nil, err
	}

	startPos, err = startRange.Start()
	if err != nil {
		return nil, err
	}

	totalPages, err = d.GetTotalPages()
	if err != nil {
		return nil, err
	}

	if page < totalPages {
		endRange, err = d.GoTo(constants.WdGoToItemGoToPage, constants.WdGoToDirectionGoToAbsolute, page+1)
		if err != nil {
			return nil, err
		}

		endPos, err = endRange.End()
		if err != nil {
			return nil, err
		}
	} else {
		allRange, err = d.AllRange()
		if err != nil {
			return nil, err
		}

		endPos, err = allRange.End()
		if err != nil {
			return nil, err
		}
		endPos -= 1
	}

	result, err = d.Range(startPos, endPos)
	if err != nil {
		return nil, err
	}

	return result, nil
}
