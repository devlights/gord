package constants

type (
	// WdExportRange は、エクスポートする文書の範囲を指定します。
	WdExportRange int
)

// WdExportRange -- エクスポートする文書の範囲を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdexportrange
//
//goland:noinspection GoUnusedConst
const (
	WdExportRangeAllDocument WdExportRange = 0 // 文書全体をエクスポートします。
	WdExportRangeCurrentPage WdExportRange = 2 // 現在のページをエクスポートします。
	WdExportRangeFromTo      WdExportRange = 3 // 開始位置と終了位置を使用して、指定範囲をエクスポートします。
	WdExportRangeSelection   WdExportRange = 1 // 現在の選択範囲をエクスポートします。
)
