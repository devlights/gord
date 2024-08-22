package constants

type (
	// WdBreakType は、改ページの種類を指定します。
	WdBreakType int
)

// WdBreakType -- 改ページの種類を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdbreaktype
//
//goland:noinspection GoUnusedConst
const (
	WdBreakTypeColumnBreak            WdBreakType = 8  // 挿入位置に段区切りを挿入します。
	WdBreakTypeLineBreak              WdBreakType = 6  // 改行します。
	WdBreakTypeLineBreakClearLeft     WdBreakType = 9  // 改行します。
	WdBreakTypeLineBreakClearRight    WdBreakType = 10 // 改行します。
	WdBreakTypePageBreak              WdBreakType = 7  // 挿入位置で改ページします。
	WdBreakTypeSectionBreakContinuous WdBreakType = 3  // 改ページなしで新しいセクションを開始します。
	WdBreakTypeSectionBreakEvenPage   WdBreakType = 4  // セクション区切りを挿入し、次の偶数ページから次のセクションを開始します。 セクション区切りを偶数ページに挿入した場合、次の奇数ページは空白になります。
	WdBreakTypeSectionBreakNextPage   WdBreakType = 2  // 次のページ上にセクション区切りを挿入します。
	WdBreakTypeSectionBreakOddPage    WdBreakType = 5  // セクション区切りを挿入し、次の奇数ページから次のセクションを開始します。 セクション区切りを奇数ページに挿入した場合、次の偶数ページは空白になります。
	WdBreakTypeTextWrappingBreak      WdBreakType = 11 // 現在の行を終了し、画像、表、またはその他の項目の下に続きの文字列を配置します。 続きの文字列の開始位置は、左端または右端に揃えられた表を含まない、次の空白行です。
)
