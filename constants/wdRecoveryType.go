package constants

type (
	// WdRecoveryType は、選択したセルを貼り付けるときに使用する書式を指定します。
	WdRecoveryType int
)

// WdRecoveryType -- 選択したセルを貼り付けるときに使用する書式を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdrecoverytype
//
//goland:noinspection GoUnusedConst
const (
	WdRecoveryTypeChart                                   WdRecoveryType = 14 // Microsoft Office Excel のグラフを埋め込みの OLE オブジェクトとして貼り付けます。
	WdRecoveryTypeChartLinked                             WdRecoveryType = 15 // Excel のグラフを貼り付け、元の Excel ワークシートとのリンクを設定します。
	WdRecoveryTypeChartPicture                            WdRecoveryType = 13 // Excel のグラフを図として貼り付けます。
	WdRecoveryTypeFormatOriginalFormatting                WdRecoveryType = 16 // 貼り付けるオブジェクトの元の書式を保持します。
	WdRecoveryTypeFormatPlainText                         WdRecoveryType = 22 // 書式設定されていないテキスト形式の文字列として貼り付けます。
	WdRecoveryTypeFormatSurroundingFormattingWithEmphasis WdRecoveryType = 20 // 貼り付ける文字列の書式に、周囲の文字列と同じ書式を適用します。
	WdRecoveryTypeListCombineWithExistingList             WdRecoveryType = 24 // 貼り付けるリストと、隣接するリストを結合します。
	WdRecoveryTypeListContinueNumbering                   WdRecoveryType = 7  // 貼り付けるリストに、文書内のリストの続きの番号を振ります。
	WdRecoveryTypeListDontMerge                           WdRecoveryType = 25 // サポートされていません。
	WdRecoveryTypeListRestartNumbering                    WdRecoveryType = 8  // 貼り付けるリストの番号を開始番号から振り直します。
	WdRecoveryTypePasteDefault                            WdRecoveryType = 0  // サポートされていません。
	WdRecoveryTypeSingleCellTable                         WdRecoveryType = 6  // 1 つのセルの表を、別の表として貼り付けます。
	WdRecoveryTypeSingleCellText                          WdRecoveryType = 5  // 1 つのセルを文字列として貼り付けます。
	WdRecoveryTypeTableAppendTable                        WdRecoveryType = 10 // 貼り付けるセルを既存の表に結合します。このとき、貼り付ける行を選択した行の間に挿入します。
	WdRecoveryTypeTableInsertAsRows                       WdRecoveryType = 11 // 貼り付ける表を、貼り付け先の表の 2 つの行の間に行として挿入します。
	WdRecoveryTypeTableOriginalFormatting                 WdRecoveryType = 12 // 末尾に追加する表を、表のスタイルを結合せずに貼り付けます。
	WdRecoveryTypeTableOverwriteCells                     WdRecoveryType = 23 // 表のセルを貼り付け、既存の表のセルを上書きします。
	WdRecoveryTypeUseDestinationStylesRecovery            WdRecoveryType = 19 // 貼り付け先の文書で使用されているスタイルを使用します。
)
