package constants

type (
	// WdInformation は、指定された選択範囲または指定範囲に関して取得される情報の種類を指定します。
	WdInformation int
)

// WdInformation -- 指定された選択範囲または指定範囲に関して取得される情報の種類を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdinformation
//
//goland:noinspection GoUnusedConst
const (
	WdInformationActiveEndAdjustedPageNumber              WdInformation = 1  // 指定された選択範囲または指定範囲のアクティブな終点が含まれるページの数を返します。 開始ページ番号を設定した場合、またはその他の手動調整を行った場合は、調整されたページ番号が返されます ( WdInformationActiveEndPageNumber とは異なります)。
	WdInformationActiveEndPageNumber                      WdInformation = 3  // 指定された選択範囲または文書の先頭から数えて、範囲のアクティブな終点が含まれるページの数を返します。 ページ番号の手動調整は無視されます ( WdInformationActiveEndAdjustedPageNumber とは異なります)。
	WdInformationActiveEndSectionNumber                   WdInformation = 2  // 指定された選択範囲または指定範囲の終了位置を含むセクション番号を取得します。
	WdInformationAtEndOfRowMarker                         WdInformation = 31 // 指定された選択範囲または指定範囲が表の中の行区切り記号である場合、値は True です。
	WdInformationCapsLock                                 WdInformation = 21 // Returns True if Caps Lock is in effect.
	WdInformationEndOfRangeColumnNumber                   WdInformation = 17 // 指定された選択範囲または指定範囲の終了位置の列番号を取得します。
	WdInformationEndOfRangeRowNumber                      WdInformation = 14 // 指定された選択範囲または指定範囲の終了位置の行番号を取得します。
	WdInformationFirstCharacterColumnNumber               WdInformation = 9  // 指定された選択範囲または指定範囲の開始位置を取得します。 選択範囲または指定範囲が解除されている場合、範囲の右側の文字番号 (ステータス バーで "桁" の後に表示される文字の列番号と同じ) を取得します。
	WdInformationFirstCharacterLineNumber                 WdInformation = 10 // 指定された選択範囲または指定範囲の開始位置を取得します。 選択範囲または指定範囲が解除されている場合は、範囲の右側の文字番号 (ステータス バーで "行" の後に表示される文字の行番号と同じ) を取得します。
	WdInformationFrameIsSelected                          WdInformation = 11 // 指定された選択範囲または指定範囲がレイアウト枠またはテキスト ボックス全体である場合、値は True です。
	WdInformationHeaderFooterType                         WdInformation = 33 // 指定された選択範囲または指定範囲を含むヘッダーまたはフッターの種類を示す値を取得します。 詳細については、「備考」の表を参照してください。
	WdInformationHorizontalPositionRelativeToPage         WdInformation = 5  // 指定した選択範囲または範囲の水平方向の位置を返します。これは、選択範囲または範囲の左端からページの左端までの距離です (1 ポイント = 20 twips、72 ポイント = 1 インチ)。 選択範囲または範囲が画面領域内にない場合は、-1 を返します。
	WdInformationHorizontalPositionRelativeToTextBoundary WdInformation = 7  // 指定した選択範囲または範囲を囲む最も近いテキスト境界の左端を基準にした水平方向の位置をポイント (1 ポイント = 20 twips、72 ポイント = 1 インチ) で返します。 選択範囲または範囲が画面領域内にない場合は、-1 を返します。
	WdInformationInBibliography                           WdInformation = 42 // 文献目録には、指定された選択範囲または指定範囲の場合は True を返します。
	WdInformationInCitation                               WdInformation = 43 // 指定された選択範囲または指定範囲が引用文献の場合は True を返します。
	WdInformationInClipboard                              WdInformation = 38 // この定数の詳細については、Microsoft Office Macintosh Edition に含まれているランゲージ リファレンスのヘルプを参照してください。
	WdInformationInCommentPane                            WdInformation = 26 // 指定された選択範囲または指定範囲がコメント ウィンドウ枠にある場合、値は True です。
	WdInformationInContentControl                         WdInformation = 46 // 指定された選択範囲または指定範囲がコンテンツ コントロール内にある場合は True を返します。
	WdInformationInCoverPage                              WdInformation = 41 // 送付状には、指定された選択範囲または指定範囲の場合は True を返します。
	WdInformationInEndnote                                WdInformation = 36 // 標準表示モードで、文末脚注または文末脚注ウィンドウ枠で印刷レイアウト表示で、選択範囲または指定範囲がの場合 True を返します。
	WdInformationInFieldCode                              WdInformation = 44 // フィールド コードでは、指定された選択範囲または指定範囲の場合は True を返します。
	WdInformationInFieldResult                            WdInformation = 45 // フィールドの実行結果は、指定された選択範囲または指定範囲の場合は True を返します。
	WdInformationInFootnote                               WdInformation = 35 // 標準表示モードで脚注領域または印刷レイアウト表示で、脚注ウィンドウ枠で、選択範囲または指定範囲がの場合 True を返します。
	WdInformationInFootnoteEndnotePane                    WdInformation = 25 // 選択範囲または指定範囲の脚注または文末脚注のウィンドウで印刷レイアウト表示の脚注または文末脚注領域または標準表示モードでは、 True を返します。 詳細については、 WdInformationInFootnote および WdInformationInEndnote を上記の説明を参照してください。
	WdInformationInHeaderFooter                           WdInformation = 28 // 場合は、選択範囲または指定範囲がヘッダーまたはフッターのウィンドウまたは、ヘッダーまたはフッターを印刷レイアウト表示では、 True を返します。
	WdInformationInMasterDocument                         WdInformation = 34 // 選択範囲または指定範囲がグループ文書 (少なくとも 1 つのサブ文書を含む文書) 内の場合は True を返します。
	WdInformationInWordMail                               WdInformation = 37 // 場合は、選択範囲または指定範囲がヘッダーまたはフッターのウィンドウまたは、ヘッダーまたはフッターを印刷レイアウト表示では、 True を返します。
	WdInformationMaximumNumberOfColumns                   WdInformation = 18 // 選択範囲または指定範囲に含まれる表の列の最大の列数を取得します。
	WdInformationMaximumNumberOfRows                      WdInformation = 15 // 指定された選択範囲または指定範囲の表の最大の行数を取得します。
	WdInformationNumberOfPagesInDocument                  WdInformation = 4  // 選択範囲または指定範囲と関連する文書のページ数を取得します。
	WdInformationNumLock                                  WdInformation = 22 // Returns True if Num Lock is in effect.
	WdInformationOverType                                 WdInformation = 23 // 上書きモードの場合、値は True です。 Overtype プロパティを使用して上書きモードの状態を変更できます。
	WdInformationReferenceOfType                          WdInformation = 32 // 「備考」の表に示すとおり、選択範囲が脚注、文末脚注、またはコメントの参照範囲の中にあるかどうかを示す値を取得します。
	WdInformationRevisionMarking                          WdInformation = 24 // 変更履歴の記録がオンの場合、値は True です。
	WdInformationSelectionMode                            WdInformation = 20 // 次の表に示すように、現在の選択モードを示す値を取得します。
	WdInformationStartOfRangeColumnNumber                 WdInformation = 16 // 選択範囲または指定範囲の先頭を含む表の列番号を取得します。
	WdInformationStartOfRangeRowNumber                    WdInformation = 13 // 選択範囲または指定範囲の先頭を含む表の行番号を取得します。
	WdInformationVerticalPositionRelativeToPage           WdInformation = 6  // 選択範囲または範囲の垂直方向の位置を返します。これは、選択範囲の上端からページの上端までの距離です (1 ポイント = 20 twips、72 ポイント = 1 インチ)。 選択範囲がドキュメント ウィンドウに表示されない場合は、-1 を返します。
	WdInformationVerticalPositionRelativeToTextBoundary   WdInformation = 8  // 選択範囲またはポイント (1 ポイント = 20 twip、72 ポイント = 1 インチ) で、それを囲む隣接する境界線の上端を基準にして範囲の垂直方向の位置を返します。 枠または表のセル内に挿入ポイントの位置を決定するのに便利です。 選択範囲が表示されない場合は、-1 を返します。
	WdInformationWithInTable                              WdInformation = 12 // 選択範囲が表の中にある場合、値は True です。
	WdInformationZoomPercentage                           WdInformation = 19 // 割合 のプロパティが設定されている拡大率の現在の割合を返します。
)
