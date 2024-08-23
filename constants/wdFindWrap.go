package constants

type (
	// WdFindWrap は、検索対象の選択範囲または指定範囲内に検索文字列が見つからなかった場合の、折り返し動作を指定します。
	WdFindWrap int
)

// WdFindWrap -- 検索対象の選択範囲または指定範囲内に検索文字列が見つからなかった場合の、折り返し動作を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdfindwrap
//
//goland:noinspection GoUnusedConst
const (
	WdFindWrapFindAsk      WdFindWrap = 2 // 選択範囲または指定範囲を検索し、文書の残りの部分も検索するかどうかをたずねるメッセージを表示します。
	WdFindWrapFindContinue WdFindWrap = 1 // 検索範囲の先頭または末尾まで検索し、さらに検索を続けます。
	WdFindWrapFindStop     WdFindWrap = 0 // 検索範囲の先頭または末尾まで検索したら、検索を終了します。
)
