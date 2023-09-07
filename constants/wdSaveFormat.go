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
	WdFormatDocument                    WdSaveFormat = 0  // Microsoft Office Word 97 - 2003 binary file format.
	WdFormatDOSText                     WdSaveFormat = 4  // Microsoft DOS text format.
	WdFormatDOSTextLineBreaks           WdSaveFormat = 5  // Microsoft DOS text with line breaks preserved.
	WdFormatEncodedText                 WdSaveFormat = 7  // Encoded text format.
	WdFormatFilteredHTML                WdSaveFormat = 10 // Filtered HTML format.
	WdFormatFlatXML                     WdSaveFormat = 19 // Open XML file format saved as a single XML file.
	WdFormatFlatXMLMacroEnabled         WdSaveFormat = 20 // Open XML file format with macros enabled saved as a single XML file.
	WdFormatFlatXMLTemplate             WdSaveFormat = 21 // Open XML template format saved as a XML single file.
	WdFormatFlatXMLTemplateMacroEnabled WdSaveFormat = 22 // Open XML template format with macros enabled saved as a single XML file.
	WdFormatOpenDocumentText            WdSaveFormat = 23 // OpenDocument Text format.
	WdFormatHTML                        WdSaveFormat = 8  // Standard HTML format.
	WdFormatRTF                         WdSaveFormat = 6  // Rich text format (RTF).
	WdFormatStrictOpenXMLDocument       WdSaveFormat = 24 // Strict Open XML document format.
	WdFormatTemplate                    WdSaveFormat = 1  // Word template format.
	WdFormatText                        WdSaveFormat = 2  // Microsoft Windows text format.
	WdFormatTextLineBreaks              WdSaveFormat = 3  // Windows text format with line breaks preserved.
	WdFormatUnicodeText                 WdSaveFormat = 7  // Unicode text format.
	WdFormatWebArchive                  WdSaveFormat = 9  // Web archive format.
	WdFormatXML                         WdSaveFormat = 11 // Extensible Markup Language (XML) format.
	WdFormatDocument97                  WdSaveFormat = 0  // Microsoft Word 97 document format.
	WdFormatDocumentDefault             WdSaveFormat = 16 // Word default document file format. For Word, this is the DOCX format.
	WdFormatPDF                         WdSaveFormat = 17 // PDF format.
	WdFormatTemplate97                  WdSaveFormat = 1  // Word 97 template format.
	WdFormatXMLDocument                 WdSaveFormat = 12 // XML document format.
	WdFormatXMLDocumentMacroEnabled     WdSaveFormat = 13 // XML document format with macros enabled.
	WdFormatXMLTemplate                 WdSaveFormat = 14 // XML template format.
	WdFormatXMLTemplateMacroEnabled     WdSaveFormat = 15 // XML template format with macros enabled.
	WdFormatXPS                         WdSaveFormat = 18 // XPS format.
)
