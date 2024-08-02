package constants

type (
	// WdExportCreateBookmarks は、文書をエクスポートするときに含めるブックマークを指定します。
	WdExportCreateBookmarks int
)

// WdExportCreateBookmarks -- 文書をエクスポートするときに含めるブックマークを指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdexportcreatebookmarks
//
//goland:noinspection GoUnusedConst
const (
	WdExportCreateBookmarksCreateHeadingBookmarks WdExportCreateBookmarks = 1 // Microsoft Word の見出し (メイン文書とテキスト ボックス内の見出しのみを含み、ヘッダー、フッター、文末脚注、脚注、またはコメント内の見出しは含まない) ごとに、エクスポートされた文書でブックマークを作成します。
	WdExportCreateBookmarksCreateNoBookmarks      WdExportCreateBookmarks = 0 // エクスポートされた文書でブックマークを作成しません。
	WdExportCreateBookmarksCreateWordBookmarks    WdExportCreateBookmarks = 2 // Word のブックマーク (ヘッダーとフッター内に含まれるブックマークを除くすべてのブックマークを含む) ごとに、エクスポートされた文書でブックマークを作成します。
)
