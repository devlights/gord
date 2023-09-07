package constants

type (
	// WdExportFormat は、文書のエクスポートに使用する形式を指定します。
	WdExportFormat int
)

// WdExportFormat -- 文書のエクスポートに使用する形式を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdexportformat
const (
	WdExportFormatPDF WdExportFormat = 17 // 文書を PDF 形式にエクスポートします。
	WdExportFormatXPS WdExportFormat = 18 // 文書を XML Paper Specification (XPS) 形式にエクスポートします。
)
