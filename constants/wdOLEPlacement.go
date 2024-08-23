package constants

type (
	// WdOLEPlacement は、OLE オブジェクトの配置を指定します。
	WdOLEPlacement int
)

// WdOLEPlacement -- OLE オブジェクトの配置を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdoleplacement
//
//goland:noinspection GoUnusedConst
const (
	WdOLEPlacementInLine        WdOLEPlacement = 0 // 行内
	WdOLEPlacementFloatOverText WdOLEPlacement = 1 // 位置を固定しない
)
