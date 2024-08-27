package constants

type (
	// WdGoToDirection は、選択範囲または挿入ポイントの移動先の位置を、オブジェクトまたは移動前の位置を基準にして指定します。
	WdGoToDirection int
)

// WdGoToDirection -- 選択範囲または挿入ポイントの移動先の位置を、オブジェクトまたは移動前の位置を基準にして指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdgotodirection
//
//goland:noinspection GoUnusedConst
const (
	WdGoToDirectionGoToAbsolute WdGoToDirection = 1  // 絶対位置
	WdGoToDirectionGoToFirst    WdGoToDirection = 1  // 指定したオブジェクトの最初のインスタンス
	WdGoToDirectionGoToLast     WdGoToDirection = -1 // 指定したオブジェクトの最後のインスタンス
	WdGoToDirectionGoToNext     WdGoToDirection = 2  // 指定したオブジェクトの次のインスタンス
	WdGoToDirectionGoToPrevious WdGoToDirection = 3  // 指定したオブジェクトの前のインスタンス
	WdGoToDirectionGoToRelative WdGoToDirection = 2  // 現在の位置を基準にした位置
)
