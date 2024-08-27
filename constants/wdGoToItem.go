package constants

type (
	// WdGoToItem は、カーソルまたは直前の選択範囲の移動先となる項目の種類を指定します。
	WdGoToItem int
)

// WdGoToItem -- カーソルまたは直前の選択範囲の移動先となる項目の種類を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdgotoitem
//
//goland:noinspection GoUnusedConst
const (
	WdGoToItemGoToBookmark          WdGoToItem = -1 // ブックマーク
	WdGoToItemGoToComment           WdGoToItem = 6  // コメント
	WdGoToItemGoToEndnote           WdGoToItem = 5  // 文末脚注
	WdGoToItemGoToEquation          WdGoToItem = 10 // 数式
	WdGoToItemGoToField             WdGoToItem = 7  // フィールド
	WdGoToItemGoToFootnote          WdGoToItem = 4  // 脚注
	WdGoToItemGoToGrammaticalError  WdGoToItem = 14 // 文法上のエラー
	WdGoToItemGoToGraphic           WdGoToItem = 8  // グラフィックス
	WdGoToItemGoToHeading           WdGoToItem = 11 // 見出し
	WdGoToItemGoToLine              WdGoToItem = 3  // 行
	WdGoToItemGoToObject            WdGoToItem = 9  // オブジェクト
	WdGoToItemGoToPage              WdGoToItem = 1  // ページ
	WdGoToItemGoToPercent           WdGoToItem = 12 // パーセント
	WdGoToItemGoToProofreadingError WdGoToItem = 15 // 文書校正エラー
	WdGoToItemGoToSection           WdGoToItem = 0  // セクション
	WdGoToItemGoToSpellingError     WdGoToItem = 13 // スペル ミス
	WdGoToItemGoToTable             WdGoToItem = 2  // 表
)
