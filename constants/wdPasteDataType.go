package constants

type (
	// WdPasteDataType は、クリップボードの内容を文書に挿入するときの書式を指定します。
	WdPasteDataType int
)

// WdPasteDataType -- クリップボードの内容を文書に挿入するときの書式を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdpastedatatype
//
//goland:noinspection GoUnusedConst
const (
	WdPasteDataTypePasteBitmap                  WdPasteDataType = 4  // ビットマップ
	WdPasteDataTypePasteDeviceIndependentBitmap WdPasteDataType = 5  // デバイスに依存しないビットマップ
	WdPasteDataTypePasteEnhancedMetafile        WdPasteDataType = 9  // 拡張メタファイル
	WdPasteDataTypePasteHTML                    WdPasteDataType = 10 // HTML
	WdPasteDataTypePasteHyperlink               WdPasteDataType = 7  // ハイパーリンク
	WdPasteDataTypePasteMetafilePicture         WdPasteDataType = 3  // メタファイル
	WdPasteDataTypePasteOLEObject               WdPasteDataType = 0  // OLE オブジェクト
	WdPasteDataTypePasteRTF                     WdPasteDataType = 1  // リッチ テキスト形式 (RTF)
	WdPasteDataTypePasteShape                   WdPasteDataType = 8  // 図形
	WdPasteDataTypePasteText                    WdPasteDataType = 2  // 文字列
)
