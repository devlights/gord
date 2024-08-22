package constants

type (
	// WdCollapseDirection は、指定範囲または選択範囲を折りたたむ方向を指定します。
	WdCollapseDirection int
)

// WdCollapseDirection -- 指定範囲または選択範囲を折りたたむ方向を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdcollapsedirection
//
//goland:noinspection GoUnusedConst
const (
	WdCollapseDirectionStart WdCollapseDirection = 1 // 範囲を始点の方向に折りたたみます。
	WdCollapseDirectionEnd   WdCollapseDirection = 0 // 範囲を終点の方向に折りたたみます。
)
