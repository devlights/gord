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

func (w *Document) AllRange() (*Range, error) {
	result, err := oleutil.CallMethod(w.ComObject(), "Range")
	if err != nil {
		return nil, err
	}

	r := NewRange(w, result.ToIDispatch())
	return r, nil
}

func (w *Document) Range(start, end int32) (*Range, error) {
	result, err := oleutil.CallMethod(w.ComObject(), "Range", start, end)
	if err != nil {
		return nil, err
	}

	r := NewRange(w, result.ToIDispatch())
	return r, nil
}

func (w *Document) Characters() (*Characters, error) {
	result, err := oleutil.GetProperty(w.ComObject(), "Characters")
	if err != nil {
		return nil, err
	}

	c := NewCharacters(w, result.ToIDispatch())
	return c, nil
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

func (w *Document) SaveAsWithFileFormatDefault(filePath string) error {
	_, err := oleutil.CallMethod(w.ComObject(), "SaveAs2", filePath, int(constants.WdSaveFormatDocumentDefault))
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

//goland:noinspection GoSnakeCaseUsage,GoBoolExpressions
func (w *Document) ExportAsFixedFormat(path string, fmtType constants.WdExportFormat) error {
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
		w.ComObject(),
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
