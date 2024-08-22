package constants

type (
	// WdSaveFormat は、文書を保存するときに使用する形式を指定します。
	WdSaveFormat int
)

// WdSaveFormat -- 文書を保存するときに使用する形式を指定します。
//
// # REFERENCES:
//
//   - https://learn.microsoft.com/ja-jp/office/vba/api/word.wdsaveformat
//
//goland:noinspection GoUnusedConst
const (
	WdSaveFormatDocument                    WdSaveFormat = 0  // Microsoft Office Word 97 - 2003 binary file format.
	WdSaveFormatDOSText                     WdSaveFormat = 4  // Microsoft DOS text format.
	WdSaveFormatDOSTextLineBreaks           WdSaveFormat = 5  // Microsoft DOS text with line breaks preserved.
	WdSaveFormatEncodedText                 WdSaveFormat = 7  // Encoded text format.
	WdSaveFormatFilteredHTML                WdSaveFormat = 10 // Filtered HTML format.
	WdSaveFormatFlatXML                     WdSaveFormat = 19 // Open XML file format saved as a single XML file.
	WdSaveFormatFlatXMLMacroEnabled         WdSaveFormat = 20 // Open XML file format with macros enabled saved as a single XML file.
	WdSaveFormatFlatXMLTemplate             WdSaveFormat = 21 // Open XML template format saved as a XML single file.
	WdSaveFormatFlatXMLTemplateMacroEnabled WdSaveFormat = 22 // Open XML template format with macros enabled saved as a single XML file.
	WdSaveFormatOpenDocumentText            WdSaveFormat = 23 // OpenDocument Text format.
	WdSaveFormatHTML                        WdSaveFormat = 8  // Standard HTML format.
	WdSaveFormatRTF                         WdSaveFormat = 6  // Rich text format (RTF).
	WdSaveFormatStrictOpenXMLDocument       WdSaveFormat = 24 // Strict Open XML document format.
	WdSaveFormatTemplate                    WdSaveFormat = 1  // Word template format.
	WdSaveFormatText                        WdSaveFormat = 2  // Microsoft Windows text format.
	WdSaveFormatTextLineBreaks              WdSaveFormat = 3  // Windows text format with line breaks preserved.
	WdSaveFormatUnicodeText                 WdSaveFormat = 7  // Unicode text format.
	WdSaveFormatWebArchive                  WdSaveFormat = 9  // Web archive format.
	WdSaveFormatXML                         WdSaveFormat = 11 // Extensible Markup Language (XML) format.
	WdSaveFormatDocument97                  WdSaveFormat = 0  // Microsoft Word 97 document format.
	WdSaveFormatDocumentDefault             WdSaveFormat = 16 // Word default document file format. For Word, this is the DOCX format.
	WdSaveFormatPDF                         WdSaveFormat = 17 // PDF format.
	WdSaveFormatTemplate97                  WdSaveFormat = 1  // Word 97 template format.
	WdSaveFormatXMLDocument                 WdSaveFormat = 12 // XML document format.
	WdSaveFormatXMLDocumentMacroEnabled     WdSaveFormat = 13 // XML document format with macros enabled.
	WdSaveFormatXMLTemplate                 WdSaveFormat = 14 // XML template format.
	WdSaveFormatXMLTemplateMacroEnabled     WdSaveFormat = 15 // XML template format with macros enabled.
	WdSaveFormatXPS                         WdSaveFormat = 18 // XPS format.
)
