package constants

type (
	// WdExportOptimizeFor は、エクスポートされる文書の解像度と画質を指定します。
	WdExportOptimizeFor int
)

// WdExportOptimizeFor -- エクスポートされる文書の解像度と画質を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdexportoptimizefor
//
//goland:noinspection GoUnusedConst
const (
	WdExportOptimizeForPrint    WdExportOptimizeFor = 0 // 印刷用にエクスポートします。これは品質が高く、ファイル サイズが大きくなります。
	WdExportOptimizeForOnScreen WdExportOptimizeFor = 1 // 画質が粗くファイル サイズが小さい、画面用のエクスポートです。
)
