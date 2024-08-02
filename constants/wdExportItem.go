package constants

type (
	// WdExportItem は、変更履歴とコメントを含めて文書をエクスポートするかどうかを指定します。
	WdExportItem int
)

// WdExportItem -- 変更履歴とコメントを含めて文書をエクスポートするかどうかを指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdexportitem
//
//goland:noinspection GoUnusedConst
const (
	WdExportItemDocumentContent    WdExportItem = 0 // 変更履歴とコメントを含めずに文書をエクスポートします。
	WdExportItemDocumentWithMarkup WdExportItem = 7 // 変更履歴とコメントを含めて文書をエクスポートします。
)
