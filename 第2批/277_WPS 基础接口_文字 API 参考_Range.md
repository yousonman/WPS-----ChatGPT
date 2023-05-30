# Range 对象

## 简介

代表文档中的一个连续区域。每个 Range 对象由一个起始字符位置和一个终止字符位置定义

## 说明

与书签在文档中的使用方法类似， Range 对象在 Visual Basic 过程中用来标识文档的特定部分。但与书签不同的是， Range 对象只在定义该对象的过程运行时才存在。 Range 对象独立于所选内容。也就是说，您可以定义和处理一个范围而无需更改所选内容。还可以在文档中定义多个范围，但每个窗格中只能有一个所选内容。

## 方法

| 序号 | 名称                                      | 说明                                                                           |
|:----:|:------------------------------------------|:-------------------------------------------------------------------------------|
|  1   | [Collapse](#Range.Collapse)               | 将某一区域或所选内容折叠到起始位置或结束位置。折叠之后起始位置和结束位置相同。 |
|  2   | [Copy](#Range.Copy)                       | 将指定范围复制到剪贴板。                                                       |
|  3   | [Cut](#Range.Cut)                         | 将指定对象从文档中移到剪贴板上                                                 |
|  4   | [Delete](#Range.Delete)                   | 删除指定数量的字符或单词                                                       |
|  5   | [Expand](#Range.Expand)                   | 扩展指定区域或选定内容。返回添至该区域或选定内容的字符数。 Long 类型。         |
|  6   | [GoTo](#Range.GoTo)                       | 返回一个 Range 对象，该对象代表指定项（如页、书签或域）的起始位置。            |
|  7   | [InsertAfter](#Range.InsertAfter)         | 在范围的末尾插入指定文本。                                                     |
|  8   | [InsertBefore](#Range.InsertBefore)       | 在指定的范围前插入指定文本。                                                   |
|  9   | [InsertBreak](#Range.InsertBreak)         | 插入分页符、分栏符或分节符。                                                   |
|  10  | [InsertFile](#Range.InsertFile)           | 插入指定文件的全部或一部分。                                                   |
|  11  | [InsertParagraph](#Range.InsertParagraph) | 用新段落替换指定范围。                                                         |
|  12  | [Paste](#Range.Paste)                     | 将“剪贴板”中的内容插入指定范围。                                               |
|  13  | [Range](#Range.Range)                     | 将“剪贴板”中的内容插入指定范围。                                               |
|  14  | [Select](#Range.Select)                   | 选择指定的范围。                                                               |
|  15  | [XML](#Range.XML)                         | 返回一个 String 类型的值，该值代表指定对象中的 XML 文本。                      |

## 属性

| 序号 | 名称                                                          | 说明                                                                                                                                                                    |
|:----:|:--------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Range.Application)                             | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                          |
|  2   | [Bold](#Range.Bold)                                           | 如果选定区域中字体的格式为加粗，则该属性值为 True 。 Long 类型，可读写。                                                                                                |
|  3   | [BoldBi](#Range.BoldBi)                                       | 如果字体或区域为加粗格式，则为 True 。该属性值返回 True 、 False 或 wdUndefined （用于加粗和非加粗混合文本）。可设置为 True 、 False 或 wdToggle 。 Long 类型，可读写。 |
|  4   | [BookmarkID](#Range.BookmarkID)                               | 返回位于指定区域开始位置的书签编号。如果没有相应的书签，则返回 0（零）。 Long 类型，只读。                                                                              |
|  5   | [Bookmarks](#Range.Bookmarks)                                 | 返回一个 Bookmarks 集合。该集合代表某一文档、区域或选定内容中的所有书签。只读。                                                                                         |
|  6   | [Borders](#Range.Borders)                                     | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                                                                                   |
|  7   | [Case](#Range.Case)                                           | 返回或设置一个 WdCharacterCase 常量，该常量代表指定区域中文字的大小写。可读写。                                                                                         |
|  8   | [Cells](#Range.Cells)                                         | 返回一个 Cells 集合，该集合代表在某区域中的表格单元格。只读。                                                                                                           |
|  9   | [Characters](#Range.Characters)                               | 返回一个 Characters 集合，该集合代表区域中的字符。只读。                                                                                                                |
|  10  | [CharacterStyle](#Range.CharacterStyle)                       | 返回一个 Variant 类型的值，该值代表用于设置一个或多个字符格式的样式。只读。                                                                                             |
|  11  | [CharacterWidth](#Range.CharacterWidth)                       | 返回或设置指定区域的字符宽度。 WdCharacterWidth 类型，可读写。                                                                                                          |
|  12  | [Columns](#Range.Columns)                                     | 返回一个 Columns 集合，该集合代表区域中的所有表格列。只读。                                                                                                             |
|  13  | [CombineCharacters](#Range.CombineCharacters)                 | 如果指定区域包含合并字符，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                   |
|  14  | [Comments](#Range.Comments)                                   | 返回一个 Comments 集合，该集合代表指定文档、选定内容或区域中的所有批注。只读。                                                                                          |
|  15  | [Conflicts](#Range.Conflicts)                                 | 返回 Conflicts 集合对象，该对象包含范围中的所有冲突对象。只读。                                                                                                         |
|  16  | [ContentControls](#Range.ContentControls)                     | 返回一个 ContentControls 集合，该集合代表一个区域中包含的内容控件。只读。                                                                                               |
|  17  | [Creator](#Range.Creator)                                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                            |
|  18  | [DisableCharacterSpaceGrid](#Range.DisableCharacterSpaceGrid) | 如果 WPS 忽略相应 Range 对象的每行字符数，则该属性值为 True 。 Boolean 类型，可读写。                                                                                   |
|  19  | [Document](#Range.Document)                                   | 返回与指定区域相关的 Document 对象。只读。                                                                                                                              |
|  20  | [Duplicate](#Range.Duplicate)                                 | 返回一个只读 Range 对象，该对象代表指定区域的所有属性。                                                                                                                 |
|  21  | [Editors](#Range.Editors)                                     | 返回一个 Editors 对象，该对象代表已授权修改文档中选定内容或区域的所有用户。                                                                                             |
|  22  | [EmphasisMark](#Range.EmphasisMark)                           | 返回或设置字符或指定的字符串的着重号。 WdEmphasisMark 类型，可读写。                                                                                                    |
|  23  | [End](#Range.End)                                             | 返回或设置某区域中结束字符的位置。可读/写 Long 类型。                                                                                                                   |
|  24  | [EndnoteOptions](#Range.EndnoteOptions)                       | 返回一个 EndnoteOptions 对象，该对象代表区域中的尾注。                                                                                                                  |
|  25  | [Endnotes](#Range.Endnotes)                                   | 返回一个 Endnotes 集合，该集合代表区域中的所有尾注。只读。                                                                                                              |
|  26  | [EnhMetaFileBits](#Range.EnhMetaFileBits)                     | 返回一个 Variant 类型的值，该值代表文本区域的显示方式的图片代表形式。                                                                                                   |
|  27  | [Fields](#Range.Fields)                                       | 返回一个 Fields 集合，该集合代表区域中的所有域。只读。                                                                                                                  |
|  28  | [Find](#Range.Find)                                           | 返回一个 Find 对象，该对象包含查找操作所需的条件。只读。                                                                                                                |
|  29  | [FitTextWidth](#Range.FitTextWidth)                           | 该属性返回或设置 WPS 在当前选定内容或区域中填入文字的宽度（使用当前的度量单位）。 Single 类型，可读写。                                                                 |
|  30  | [Font](#Range.Font)                                           | 返回或设置 Font 对象，该对象代表指定对象的字符格式。 Font 类型，可读写。                                                                                                |
|  31  | [FootnoteOptions](#Range.FootnoteOptions)                     | 返回 FootnoteOptions 对象，该对象代表选定内容或区域的脚注。                                                                                                             |
|  32  | [Footnotes](#Range.Footnotes)                                 | 返回一个 Footnotes 集合，该集合代表区域中的所有脚注。只读。                                                                                                             |
|  33  | [FormattedText](#Range.FormattedText)                         | 返回或设置一个 Range 对象，该对象包含指定区域或选定内容中进行过格式编排的文字。可读写。                                                                                 |
|  34  | [FormFields](#Range.FormFields)                               | 返回一个 FormFields 集合，该集合代表区域中的所有窗体域。只读。                                                                                                          |
|  35  | [Frames](#Range.Frames)                                       | 返回一个 Frames 集合，该集合代表区域中的所有图文框。只读。                                                                                                              |
|  36  | [GrammarChecked](#Range.GrammarChecked)                       | 如果已经检查了指定范围或文档的语法，则该属性值为 True 。 Boolean 类型，可读写。                                                                                         |
|  37  | [GrammaticalErrors](#Range.GrammaticalErrors)                 | 返回一个 ProofreadingErrors 集合，该集合代表指定文档或区域中有语法检查错误的句子。只读。                                                                                |
|  38  | [HighlightColorIndex](#Range.HighlightColorIndex)             | 返回或设置指定区域的突出显示颜色。 WdColorIndex 类型，可读写。                                                                                                          |
|  39  | [HorizontalInVertical](#Range.HorizontalInVertical)           | 返回或设置位于垂直排列文字中的水平排列文字的格式。 WdHorizontalInVerticalType 类型，可读写。                                                                            |
|  40  | [HTMLDivisions](#Range.HTMLDivisions)                         | 返回一个 HTMLDivisions 对象，该对象代表 Web 文档中的 HTML 划分。                                                                                                        |
|  41  | [Hyperlinks](#Range.Hyperlinks)                               | 返回一个 Hyperlinks 集合，该集合代表指定范围内的所有超链接。只读。                                                                                                      |
|  42  | [ID](#Range.ID)                                               | 返回或设置特定范围的标识名称。可读写 String 类型。                                                                                                                      |
|  43  | [Information](#Range.Information)                             | 返回有关指定范围的信息。只读 Variant 类型。                                                                                                                             |
|  44  | [InlineShapes](#Range.InlineShapes)                           | 返回一个 InlineShape 集合，该集合代表范围中的所有 InlineShapes 对象。只读。                                                                                             |
|  45  | [IsEndOfRowMark](#Range.IsEndOfRowMark)                       | 如果指定范围被折叠且位于表格中的行尾标志处，则该属性值为 True 。只读 Boolean 类型。                                                                                     |
|  46  | [Italic](#Range.Italic)                                       | 如果将字体或范围设置为倾斜格式，则该属性值为 True 。可读写 Long 类型。                                                                                                  |
|  47  | [ItalicBi](#Range.ItalicBi)                                   | 如果将字体或范围设置为倾斜格式，则该属性值为 True 。可读写 Long 类型。                                                                                                  |
|  48  | [Kana](#Range.Kana)                                           | 返回或设置日文文本的指定区域是平假名还是片假名。 WdKana 类型，可读写。                                                                                                  |
|  49  | [LanguageDetected](#Range.LanguageDetected)                   | 返回或设置一个值，该值指定 WPS 是否已经检测过指定文本的语言。可读/写 Boolean 类型。                                                                                     |
|  50  | [LanguageID](#Range.LanguageID)                               | 返回或设置一个 WdLanguageID 常量，该常量代表指定范围的语言。可读写。                                                                                                    |
|  51  | [LanguageIDFarEast](#Range.LanguageIDFarEast)                 | 返回或设置指定对象的东亚语言。可读写 WdLanguageID 类型。                                                                                                                |
|  52  | [LanguageIDOther](#Range.LanguageIDOther)                     | 返回或设置指定范围的语言。可读写 WdLanguageID 类型。                                                                                                                    |
|  53  | [ListFormat](#Range.ListFormat)                               | 返回一个 ListFormat 对象，该对象代表某区域中所有的列表格式特征。只读。                                                                                                  |
|  54  | [ListParagraphs](#Range.ListParagraphs)                       | 返回一个 ListParagraphs 集合，该集合代表范围中的所有编号段落。只读。                                                                                                    |
|  55  | [ListStyle](#Range.ListStyle)                                 | 返回一个 Variant 类型的值，该值代表用于设置项目符号列表或编号列表的样式。只读。                                                                                         |
|  56  | [Locks](#Range.Locks)                                         | 返回 CoAuthLocks 集合对象，该对象代表范围中的所有锁。只读。                                                                                                             |
|  57  | [NextStoryRange](#Range.NextStoryRange)                       | 返回一个 Range 对象，该对象代表下一个 文章 。 Range 类型，只读。                                                                                                        |
|  58  | [NoProofing](#Range.NoProofing)                               | 如果拼写和语法检查程序忽略指定文本，则该属性值为 True 。可读写 Long 类型。                                                                                              |
|  59  | [OMaths](#Range.OMaths)                                       | 返回一个 OMaths 集合，该集合代表指定区域内的 OMath 对象。只读。                                                                                                         |
|  60  | [Orientation](#Range.Orientation)                             | 在启用了“文字方向”功能时返回或设置范围中文字的方向。可读写 WdTextOrientation 类型。                                                                                     |
|  61  | [PageSetup](#Range.PageSetup)                                 | 返回一个 PageSetup 对象，该对象与指定范围相关联。                                                                                                                       |
|  62  | [ParagraphFormat](#Range.ParagraphFormat)                     | 返回或设置一个 ParagraphFormat 对象，该对象代表指定范围的段落设置。可读/写。                                                                                            |
|  63  | [Paragraphs](#Range.Paragraphs)                               | 返回一个 Paragraphs 集合，该集合代表指定范围中的所有段落。只读。                                                                                                        |
|  64  | [ParagraphStyle](#Range.ParagraphStyle)                       | 返回一个 Variant 类型的值，该值代表用于设置段落格式的样式。只读。                                                                                                       |
|  65  | [Parent](#Range.Parent)                                       | 返回一个 Object 类型值，该值代表指定 Range 对象的父对象。                                                                                                               |
|  66  | [ParentContentControl](#Range.ParentContentControl)           | 返回一个 ContentControl 对象，该对象代表指定区域的父内容控件。只读。                                                                                                    |
|  67  | [PreviousBookmarkID](#Range.PreviousBookmarkID)               | 返回最后一个书签的编号，该书签从指定范围的前面或与指定范围相同的位置开始。只读 Long 类型。                                                                              |
|  68  | [ReadabilityStatistics](#Range.ReadabilityStatistics)         | 返回一个 ReadabilityStatistics 集合，该集合代表指定文档或范围的可读性统计信息。只读。                                                                                   |
|  69  | [Revisions](#Range.Revisions)                                 | 返回一个 Revisions 集合，该集合代表范围中的修订。只读。                                                                                                                 |
|  70  | [Rows](#Range.Rows)                                           | 返回一个 Rows 集合，该集合代表范围中的所有表格行。只读。                                                                                                                |
|  71  | [Scripts](#Range.Scripts)                                     | 返回一个 Scripts 集合，该集合代表指定对象中 HTML 脚本的集合。                                                                                                           |
|  72  | [Sections](#Range.Sections)                                   | 返回一个 Sections 集合，该集合代表指定范围中的各节。只读。                                                                                                              |
|  73  | [Sentences](#Range.Sentences)                                 | 返回一个 Sentences 集合，该集合代表范围中的所有句子。只读。                                                                                                             |
|  74  | [Shading](#Range.Shading)                                     | 返回一个 Shading 对象，该对象代表指定对象的底纹格式。                                                                                                                   |
|  75  | [ShapeRange](#Range.ShapeRange)                               | 返回一个 ShapeRange 集合，该集合代表指定范围中的所有 Shape 对象。只读。                                                                                                 |
|  76  | [ShowAll](#Range.ShowAll)                                     | 如果显示所有非打印字符（如隐藏文字、制表符、空格和段落标记），则该属性值为 True 。可读写 Boolean 类型。                                                                 |
|  77  | [SpellingChecked](#Range.SpellingChecked)                     | 如果已对指定的区域或文档完成拼写检查，则该属性值为 True 。如果所有或部分区域或文档尚未进行拼写检查，则该属性值为 False 。 Boolean 类型，可读写。                        |
|  78  | [SpellingErrors](#Range.SpellingErrors)                       | 返回一个 ProofreadingErrors 集合，该集合代表指定范围中标识为拼写错误的单词。只读。                                                                                      |
|  79  | [Start](#Range.Start)                                         | 返回或设置某区域中起始字符的位置。 Long 类型，可读写。                                                                                                                  |
|  80  | [StoryLength](#Range.StoryLength)                             | 返回包含指定区域的 文字部分 中的字符数。 Long 类型，只读。                                                                                                              |
|  81  | [StoryType](#Range.StoryType)                                 | 返回指定范围、所选内容或书签的 文字部分 类型。只读 WdStoryType 类型。                                                                                                   |
|  82  | [Style](#Range.Style)                                         | 返回或设置指定对象的样式。可读写 Variant 类型。                                                                                                                         |
|  83  | [Subdocuments](#Range.Subdocuments)                           | 返回一个 Subdocuments 集合，该集合代表指定范围或文档中的所有子文档。只读。                                                                                              |
|  84  | [SynonymInfo](#Range.SynonymInfo)                             | 返回一个 SynonymInfo 对象，该对象包含同义词库中有关某范围的内容的同义词、反义词或相关单词和表达方式的信息。                                                             |
|  85  | [Tables](#Range.Tables)                                       | 返回一个 Tables 集合，该集合代表指定范围内的所有表格。只读。                                                                                                            |
|  86  | [TableStyle](#Range.TableStyle)                               | 返回一个 Variant 类型的值，该值代表用于设置表格格式的样式。只读。                                                                                                       |
|  87  | [Text](#Range.Text)                                           | 返回或设置指定区域或选定内容中的文本。 String 类型，可读写。                                                                                                            |
|  88  | [TextRetrievalMode](#Range.TextRetrievalMode)                 | 返回一个 TextRetrievalMode 对象，该对象控制从指定 区域 检索文字的方式。可读写。                                                                                         |
|  89  | [TopLevelTables](#Range.TopLevelTables)                       | 返回一个 Tables 集合，该集合代表当前范围最外部嵌套层上的表格。只读。                                                                                                    |
|  90  | [TwoLinesInOne](#Range.TwoLinesInOne)                         | 返回或设置 WPS 是否将两行文本合并为一行，并指定括住文本的字符（如果有）。 WdTwoLinesInOneType 类型，可读写。                                                            |
|  91  | [Underline](#Range.Underline)                                 | 返回或设置应用于范围的下划线的类型。可读写 WdUnderline 类型。                                                                                                           |
|  92  | [Updates](#Range.Updates)                                     | 返回 CoAuthUpdates 集合对象，该对象代表范围中的所有可用更新。只读。                                                                                                     |
|  93  | [WordOpenXML](#Range.WordOpenXML)                             | 返回一个 String 类型的值，该值以 WPS Open XML 格式表示区域中包含的 XML。只读。                                                                                          |
|  94  | [Words](#Range.Words)                                         | 返回一个 Words 集合，该集合代表范围中的所有单词。只读。                                                                                                                 |
|  95  | [XMLNodes](#Range.XMLNodes)                                   | 返回一个 XMLNodes 集合，该集合代表指定区域中的 XML 元素（包括任何只是部分属于该区域的元素）。只读。                                                                     |
|  96  | [XMLParentNode](#Range.XMLParentNode)                         | 返回一个 XMLNode 对象，该对象代表区域的父级 XML 节点。只读。                                                                                                            |

## 成员方法

### Range.Collapse

将某一区域或所选内容折叠到起始位置或结束位置。折叠之后起始位置和结束位置相同。

**语法**

**express.Collapse ()**

\*express是一个代表 Range 对象的变量。

**说明**

如果用 **wdCollapseEnd** 折叠一个代表完整段落的区域，则该区域将定位于段落结束标记之后（即下段开头）。但是，在该区域折叠后，可以用 **MoveEnd** 方法将区域回移一个字符

**示例**

``` JavaScript
/*光标定位第一段段位然后再最后插入内容*/
function test()
{
  let mgRange = Application.ActiveDocument.Paragraphs.Item(1).Range
  mgRange.Collapse(wdCollapseEnd)
  mgRange.Text = "111"
}
```

### Range.Copy

将指定范围复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例复制活动文档的第一段，并将该段落粘贴到文档的末尾。*/

function test()
{
  let doc = Application.ActiveDocument
  doc.Paragraphs.Item(1).Range.Copy()
  let myRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
  myRange.Paste()
}
```

### Range.Cut

将指定对象从文档中移到剪贴板上

**语法**

**express.Cut ()**

\*express是一个代表 Range 对象的变量。

**说明**

将区域的内容移动到剪贴板上，但折叠区域仍保留在文档中。

**示例**

``` JavaScript
/*
本示例剪切活动文档的第一段，并将该段落粘贴到文档的末尾。
*/
function test()
{
    let doc = Application.ActiveDocument
    doc.Paragraphs.Item(1).Range.Cut()
    let myRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    myRange.Paste()
}
```

### Range.Delete

删除指定数量的字符或单词

**语法**

**express.Delete ( *Unit* , *Count* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                 |
|---------|-----------|----------|----------------------------------------------------------------------------------------------------------------------|
| *Unit*  | 可选      | Variant  | 删除折叠区域时所基于的单位。可以是 WdUnits 常量之一。                                                                |
| *Count* | 可选      | Variant  | 要删除的单位数。要删除该区域后的单位，请折叠该区域并使用一个正值。要删除该区域前的单位，请折叠该区域并使用一个负值。 |

**说明**

要删除的单位数。要删除该区域后的单位，请折叠该区域并使用一个正值。要删除该区域前的单位，请折叠该区域并使用一个负值。

**示例**

``` JavaScript
/*
本示例选择并删除活动文档中的内容。
*/
function test()
{
    Application.ActiveDocument.Content.Select()
    Application.Selection.Range.Delete()
}
```

### Range.Expand

扩展指定区域或选定内容。返回添至该区域或选定内容的字符数。 **Long** 类型。

**语法**

**express.Expand ( *Unit* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                                                        |
|--------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Unit* | 可选      | Variant  | 扩展区域时所基于的单位。可以是下列 WdUnits 常量之一：wdCharacter、wdWord、wdSentence、wdParagraph、wdSection、wdStory、wdCell、wdColumn、wdRow 或 wdTable。 |

**示例**

``` JavaScript
/*本示例创建指向活动文档第一个单词的区域，然后将该区域扩展至文档首段。*/
function test() {
  let myRange = Application.ActiveDocument.Words.Item(1)
  myRange.Expand(wdParagraph)
}

/*本示例先将选定内容的首个字符设为大写，再将选定内容扩展至整句。*/
function test() {
  let characters = Application.Selection
  characters.Characters.Item(1).Case = wdTitleSentence
  characters.Expand(wdSentence)
}
```

### Range.GoTo

返回一个 **Range** 对象，该对象代表指定项（如页、书签或域）的起始位置。

**语法**

**express.GoTo ( *What* , *Which* , *Count* , *Name* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                        |
|---------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *What*  | 可选      | Variant  | 范围要移动到的项的类别。可以是 WdGoToItem 常量之一。                                                                                                                                                        |
| *Which* | 可选      | Variant  | 范围要移动到的项。可以是 WdGoToDirection 常量之一。                                                                                                                                                         |
| *Count* | 可选      | Variant  | 文档中的项数。默认值为 1。只有正值有效。若要指定一个位于该范围之前的项，可将 wdGoToPrevious 用作 Which 参数，并指定一个 Count 值。                                                                          |
| *Name*  | 可选      | Variant  | 如果 What 参数为 wdGoToBookmark、wdGoToComment、wdGoToField 或 wdGoToObject，则此参数指定一个名称。只有正值有效。若要指定一个位于该范围之前的项，可将 wdGoToPrevious 用作 Which 参数，并指定一个 Count 值。 |

**说明**

以下示例将范围向上移动两行。

``` JavaScript
Application.ActiveDocument.Range().GoTo(wdGoToLine, wdGoToPrevious, 2)
```

以下示例将所选内容移至下一个 DATE 域。

``` JavaScript
Application.ActiveDocument.Range().GoTo(wdGoToField, null, null, "Date") 
```

以下示例将范围移至文档中的第四行。

``` JavaScript
Application.ActiveDocument.Range().GoTo(wdGoToLine, wdGoToAbsolute, 4)
```

下列示例的功能等效，都将范围移至文档中的第一个标题处。

``` JavaScript
Application.ActiveDocument.Range().GoTo(wdGoToHeading, wdGoToFirst)
Application.ActiveDocument.Range().GoTo(wdGoToHeading, wdGoToAbsolute, 1)
```

将 **GoTo** 方法与 **wdGoToGrammaticalError** 、 **wdGoToProofreadingError** 或 **wdGoToSpellingError** 常量一起使用时，返回的 **Range** 对象中包括所有含语法或拼写错误的文本。

**示例**

``` JavaScript
/*以下示例将插入点移至活动文档中第五个尾注引用标记的前面。*/
function test() {
  if(Application.ActiveDocument.Endnotes.Count >= 5){
      Application.ActiveDocument.Range().GoTo(wdGoToEndnote, wdGoToAbsolute, 5)
  }
}

/*以下示例将 R1 设置为等于活动文档中第一个脚注引用标记。*/
function test() {
  if(Application.ActiveDocument.Footnotes.Count >= 1){
      let R1 = Application.ActiveDocument.Range().GoTo(wdGoToFootnote, wdGoToFirst)
      R1.Expand(wdCharacter)
  }
}
```

### Range.InsertAfter

在范围的末尾插入指定文本。

**语法**

**express.InsertAfter ( *Text* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Text* | 可选      | String   | 要插入的文本。 |

**说明**

应用此方法之后，该范围将扩展，以包含新文本。

使用 Visual Basic **Chr** 函数和 **InsertAfter** 方法，可以插入引号、制表符和不间断连字符等字符。还可以使用下列 Visual Basic 常量： **vbCr** 、 **vbLf** 、 **vbCrLf** 和 **vbTab** 。

如果对引用整个段落的范围使用此方法，则在末段标记之后插入文本（插入的文本将出现在下一段的开头）。要在段尾插入文本，请先确定终点，再从该位置减去 1（因为段落标记是一个字符），如以下示例所示。

``` JavaScript
function test() {
    let doc = Application.ActiveDocument
    let rngRange = doc.Range(doc.Paragraphs.Item(1).Start, doc.Paragraphs.Item(1).End - 1)
    rngRange.InsertAfter(" This is now the last sentence in paragraph one.")
}
```

然而，如果该范围以一个段落标记结尾，而该段落标记正好又是文档的末尾，则 WPS 在末段标记前插入文本，而不是在文档末尾创建一个新段落。

同样，如果该范围是书签， WPS 将插入指定的文本，但不会扩展范围或书签以包含新文本。

**示例**

``` JavaScript
/*以下示例在活动文档的末尾插入文本。Content 属性返回一个 Range 对象。*/
Application.ActiveDocument.Content.InsertAfter("end of document")

/*以下示例将输入框中的文本作为活动文档的第二段插入到文档中。*/
function test() {
  let response = prompt("Type some text")
  let range2 = Application.ActiveDocument.Paragraphs.Item(1).Range
  range2.InsertAfter("1." + "\t" + response)
  range2.InsertParagraphAfter()
}
```

### Range.InsertBefore

在指定的范围前插入指定文本。

**语法**

**express.InsertBefore ( *Text* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Text* | 必选      | String   | 要插入的文本。 |

**说明**

在插入文本之后，该范围将扩展，以包含新文本。如果该范围是书签，则书签也会扩展，以包含新文本。

使用 Visual Basic **Chr** 函数和 **InsertBefore** 方法，可以插入引号、制表符和不间断连字符等字符。还可以使用下列 Visual Basic 常量： **vbCr** 、 **vbLf** 、 **vbCrLf** 和 **vbTab** 。

**示例**

``` JavaScript
/*以下示例在活动文档的开头插入文本“Introduction”，并将其作为一个单独段落。*/
function test() {
  let content = Application.ActiveDocument.Content
  content.InsertParagraphBefore()
  content.InsertBefore("Introduction")
}
```

### Range.InsertBreak

插入分页符、分栏符或分节符。

**语法**

**express.InsertBreak ( *Type* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                      |
|--------|-----------|----------|-------------------------------------------------------------------------------------------|
| *Type* | 必选      | Variant  | 要插入的分隔符类型。可以是 WdBreakType 常量之一。如果省略该参数，则默认值为 wdPageBreak。 |

**说明**

当插入分页符或分栏符时，范围将被插入的分隔符所替换。如果不希望替换该范围，可在使用 **InsertBreak** 方法之前使用 **Collapse** 方法。在插入分节符时，此分节符紧接在 **Range** 之前插入。

根据您所选择或安装的语言支持（例如，美国英语），上述部分常量可能不可用。

**示例**

``` JavaScript
/*以下示例紧接在活动文档第二段之后插入一个分页符。*/
function test() {
  let myRange = Application.ActiveDocument.Paragraphs.Item(2).Range
  myRange.Collapse(wdCollapseEnd)
  myRange.InsertBreak(wdPageBreak)
}
```

### Range.InsertFile

插入指定文件的全部或一部分。

**语法**

**express.InsertFile ( *FileName* , *Range* , *ConfirmConversions* , *Link* , *Attachment* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                                                                        |
|----------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName*           | 必选      | String   | 要插入的文件的路径及文件名。如果没有指定路径，则 WPS 假定文件位于当前文件夹中。                                                             |
| *Range*              | 可选      | Variant  | 如果指定的文件是 WPS 文档，则此参数代表书签。如果该文件为其他类型（例如，ET 工作表），则此参数代表命名区域或单元格区域（例如，R1C1:R3C4）。 |
| *ConfirmConversions* | 可选      | Variant  | 如果该参数值为 True，则在插入非 WPS 文档格式的文件时， WPS 将提示您确认转换。                                                               |
| *Link*               | 可选      | Variant  | 如果该参数值为 True，则使用 INCLUDETEXT 域插入该文件。                                                                                      |
| *Attachment*         | 可选      | Variant  | 如果该参数值为 True，则将该文件以附件形式插入电子邮件中。                                                                                   |

**示例**

``` JavaScript
/* 在当前文档的最后插入文件"MyDoc.doc"的内容 */
function test() {
    let end = Application.ActiveDocument.Range().End;
    Application.Selection.Start = end;
    Application.Selection.End = end;    
    Application.Selection.Range.InsertFile("C:\\MyDoc.doc");
}
```

### Range.InsertParagraph

用新段落替换指定范围。

**语法**

**express.InsertParagraph ()**

\*express是一个代表 Range 对象的变量。

**说明**

使用此方法后，该范围将成为一个新段落。

如果您不希望替换该范围，可在使用此方法之前先使用 **Collapse** 方法。 **InsertParagraphAfter** 方法可在 **Range** 对象后插入一个新段落。

**示例**

``` JavaScript
/*以下示例在活动文档的开头插入一个新段落。*/
function test() {
  let myRange = Application.ActiveDocument.Range(0, 0)
  myRange.InsertParagraph()
  myRange.InsertBefore("Dear Sirs,")
}
```

### Range.Paste

将“剪贴板”中的内容插入指定范围。

**语法**

**express.Paste ()**

\*express是一个代表 Range 对象的变量。

**说明**

如果不希望替换该范围的内容，可在使用此方法之前先使用 **Collapse** 方法。

如果您将此方法与 **Range** 对象一起使用，该范围将扩展以包含“剪贴板”中的内容。

**示例**

``` JavaScript
/*以下示例复制活动文档中第一个表格，并将其粘贴至新文档。*/
function test() {
  if(Application.ActiveDocument.Tables.Count >= 1){
      Application.ActiveDocument.Tables.Item(1).Range.Copy()
      Application.Documents.Add().Content.Paste()
  }
}

/*以下示例复制所选内容并将其粘贴到文档末尾。*/
function test() {
  if(Application.Selection.Type != wdSelectionIP){
      Application.Selection.Copy()
      let Range2 = Application.ActiveDocument.Content
      Range2.Collapse(wdCollapseEnd)
      Range2.Paste()
  }
}
```

### Range.Range

将“剪贴板”中的内容插入指定范围。

**语法**

**express.Range ()**

\*express是一个代表 Range 对象的变量。

**说明**

如果不希望替换该范围的内容，可在使用此方法之前先使用 **Collapse** 方法。

如果您将此方法与 **Range** 对象一起使用，该范围将扩展以包含“剪贴板”中的内容。

**示例**

``` JavaScript
/*
复制活动文档中第一个表格，并将其粘贴至新文档
*/
function test()
{
    let doc = Application.ActiveDocument
    if (doc.Tables.Count)
        doc.Tables.Item(1).Range.Copy()
    Application.Documents.Add().Content.Paste();
}
```

### Range.Select

选择指定的范围。

**语法**

**express.Select ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例选择活动文档中的第一个段落。*/
function test(){
  Application.ActiveDocument.Paragraphs.Item(1).Range.Select()
  Application.Selection.Font.Bold = true
}
```

### Range.XML

返回一个 **String** 类型的值，该值代表指定对象中的 XML 文本。

**语法**

**express.XML ( *DataOnly* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                          |
|------------|-----------|----------|-------------------------------------------------------------------------------|
| *DataOnly* | 可选      | Boolean  | 如果该属性值为 True，则返回不带 WPS XML 标记的 XML 的文本。默认设置为 False。 |

## 成员属性

### Range.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Range 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Range.Bold

如果选定区域中字体的格式为加粗，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Bold**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*
把第一段的内容属性加粗
*/
Application.ActiveDocument.Paragraphs.Item(1).Range.Bold = true
```

### Range.BoldBi

如果字体或区域为加粗格式，则为 **True** 。该属性值返回 **True** 、 **False** 或 **wdUndefined** （用于加粗和非加粗混合文本）。可设置为 **True** 、 **False** 或 **wdToggle** 。 **Long** 类型，可读写。

**语法**

**express.BoldBi**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将从右向左语言的活动文档中第一段的格式设为加粗。*/
ActiveDocument.Paragraphs.Item(1).Range.BoldBi = true
```

### Range.BookmarkID

返回位于指定区域开始位置的书签编号。如果没有相应的书签，则返回 0（零）。 **Long** 类型，只读。

**语法**

**express.BookmarkID**

\*express是一个代表 Range 对象的变量。

**说明**

编号或书签对应于书签在文档中的位置：1 对应于第一个书签，2 对应于第二个，以此类推。

**示例**

``` JavaScript
/*如果文档的开始位置没有设置书签，则本示例添加一个名为“temp”的书签。*/
function test()
{
let myRange = ActiveDocument.Content
myRange.Collapse(wdCollapseStart)
if(myRange.BookmarkID == 0){
    ActiveDocument.Bookmarks.Add("temp", myRange)
}
}
```

### Range.Bookmarks

返回一个 **Bookmarks** 集合。该集合代表某一文档、区域或选定内容中的所有书签。只读。

**语法**

**express.Bookmarks**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.Case

返回或设置一个 **WdCharacterCase** 常量，该常量代表指定区域中文字的大小写。可读写。

**语法**

**express.Case**

\*express是一个代表 Range 对象的变量。

**说明**

由于您选择或安装的语言支持不同（例如美国英语），上述部分常量可能无法使用。

### Range.Cells

返回一个 **Cells** 集合，该集合代表在某区域中的表格单元格。只读。

**语法**

**express.Cells**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例创建一个 3x3 表格并给表格中的每个单元格依次分配一个单元格编号。*/
function test()
{
let newDoc = Documents.Add()
let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
let i = 1
for(let c = 1; c <= myTable.Range.Cells.Count; c++){
    myTable.Range.Cells.Item(c).Range.InsertAfter("Cell " + i)
    i = i + 1
}
}
```

### Range.Characters

返回一个 **Characters** 集合，该集合代表区域中的字符。只读。

**语法**

**express.Characters**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.CharacterStyle

返回一个 **Variant** 类型的值，该值代表用于设置一个或多个字符格式的样式。只读。

**语法**

**express.CharacterStyle**

\*express是一个代表 Range 对象的变量。

### Range.CharacterWidth

返回或设置指定区域的字符宽度。 **WdCharacterWidth** 类型，可读写。

**语法**

**express.CharacterWidth**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将当前选定内容转换为半角字符。*/
Selection.Range.CharacterWidth = wdWidthHalfWidth
```

### Range.Columns

返回一个 **Columns** 集合，该集合代表区域中的所有表格列。只读。

**语法**

**express.Columns**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示活动文档的第一个表格中的列数。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1){
    MsgBox(ActiveDocument.Tables.Item(1).Columns.Count)
}
}
```

``` JavaScript
/*本示例将当前列的宽度设置为 1 英寸。*/
function test()
{
if(Selection.Information(wdWithInTable) == true) {
    Selection.Columns.SetWidth(InchesToPoints(1), wdAdjustProportional)
}
}
```

### Range.CombineCharacters

如果指定区域包含合并字符，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CombineCharacters**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例合并选定区域中的字符。*/
Selection.Range.CombineCharacters = true
```

### Range.Comments

返回一个 **Comments** 集合，该集合代表指定文档、选定内容或区域中的所有批注。只读。

**语法**

**express.Comments**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.Conflicts

返回 Conflicts 集合对象，该对象包含范围中的所有冲突对象。只读。

**语法**

**express.Conflicts**

\*express是一个代表 Range 对象的变量。

**说明**

使用 **Conflicts** 属性可以返回文档的 Conflicts 集合对象。使用 Conflicts( *Index* )（其中 *Index* 为冲突索引号）可以返回一个 Conflict 对象。

| 注释                                                                                             |
|--------------------------------------------------------------------------------------------------|
| 该属性仅可用于支持共同创作的文档。如果尝试在不支持共同创作的文档中访问该属性，将导致运行时错误。 |

**示例**

``` JavaScript
/*下面的代码示例显示活动文档的第一个段落中的冲突数量。*/
MsgBox(ActiveDocument.Paragraphs.Item(1).Range.Conflicts.Count)
```

### Range.ContentControls

返回一个 **ContentControls** 集合，该集合代表一个区域中包含的内容控件。只读。

**语法**

**express.ContentControls**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*下面的示例将一个下拉列表内容控件插入到活动文档中的指定位置。*/
function test()
{
let objRange = ActiveDocument.Range(200, 200)
let objCC = objRange.ContentControls.Add(wdContentControlDropdownList)

//List entries
objCC.DropdownListEntries.Add("Cat")
objCC.DropdownListEntries.Add("Dog")
objCC.DropdownListEntries.Add("Horse")
objCC.DropdownListEntries.Add("Monkey")
objCC.DropdownListEntries.Add("Snake")
objCC.DropdownListEntries.Add("Other")
}
```

### Range.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Range 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Range.DisableCharacterSpaceGrid

如果 WPS 忽略相应 **Range** 对象的每行字符数，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisableCharacterSpaceGrid**

\*express是一个代表 Range 对象的变量。

**说明**

如果仅将部分指定字体或区域的 **DisableCharacterSpaceGrid** 属性设置为 **True** ，则返回 **wdUndefined** 。

### Range.Document

返回与指定区域相关的 **Document** 对象。只读。

**语法**

**express.Document**

\*express是一个代表 Range 对象的变量。

### Range.Duplicate

返回一个只读 **Range** 对象，该对象代表指定区域的所有属性。

**语法**

**express.Duplicate**

\*express是一个代表 Range 对象的变量。

**说明**

通过复制 **Range** 对象，可更改所复制的区域的开始或结尾字符的位置，而不会更改源区域。

### Range.Editors

返回一个 **Editors** 对象，该对象代表已授权修改文档中选定内容或区域的所有用户。

**语法**

**express.Editors**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例给当前用户分配编辑权限以修改活动的选定内容。*/
let objEditor = Selection.Editors.Add(wdEditorCurrent)
```

### Range.EmphasisMark

返回或设置字符或指定的字符串的着重号。 **WdEmphasisMark** 类型，可读写。

**语法**

**express.EmphasisMark**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：在活动文档的第四个单词上面设置着重号：逗号。*/
ActiveDocument.Words.Item(4).EmphasisMark = wdEmphasisMarkOverComma
```

### Range.End

返回或设置某区域中结束字符的位置。可读/写 **Long** 类型。

**语法**

**express.End**

\*express是一个代表 Range 对象的变量。

**说明**

**Range** 对象均包含开始位置和结束位置。结束位置是距文档开头部分最远的点。如果该属性的设置值小于 **Start** 属性值，则 **Start** 属性将设为同一值（即 **Start** 与 **End** 属性值相等）。

该属性返回结束字符相对于文档开头部分的位置。文档主体部分 ( **wdMainTextStory** ) 的起始字符位置为 0（零）。设置该属性可以改变选定内容、区域或者书签的大小。

**示例**

``` JavaScript
/*本示例将 myRange 的结束位置移动一个字符。*/
function test() {
  let myRange = Application.ActiveDocument.Paragraphs.Item(1).Range
  myRange.End = myRange.End - 1
}
```

### Range.EndnoteOptions

返回一个 **EndnoteOptions** 对象，该对象代表区域中的尾注。

**语法**

**express.EndnoteOptions**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*如果活动文档第二节中尾注的起始号码不是 1，则本示例将其设为 1。*/
function SetEndnoteOptionsRange() {
    let Range2 = ActiveDocument.Sections.Item(2).Range.EndnoteOptions
    if(Range2.StartingNumber != 1) {
        Range2.StartingNumber = 1
    }
}
```

### Range.Endnotes

返回一个 **Endnotes** 集合，该集合代表区域中的所有尾注。只读。

**语法**

**express.Endnotes**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.EnhMetaFileBits

返回一个 **Variant** 类型的值，该值代表文本区域的显示方式的图片代表形式。

**语法**

**express.EnhMetaFileBits**

\*express是一个代表 Range 对象的变量。

**说明**

**EnhMetaFileBits** 属性返回一个字节数组，该数组可在 Visual Basic 或 Microsoft C++ 开发环境中通过 Microsoft Windows 32 应用程序编程接口来使用。

### Range.Fields

返回一个 **Fields** 集合，该集合代表区域中的所有域。只读。

**语法**

**express.Fields**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例删除活动文档正文和页脚中的所有域。*/
function test()
{
for(let aField = 1; aField <= ActiveDocument.Fields.Count; aField++) {
    ActiveDocument.Fields.Item(aField).Delete()
}
let myRange = ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).Range
for(let a = 1; a <= myRange.Fields.Count; a++) {
    myRange.Fields.Item(a).Delete()
}
}
```

### Range.Find

返回一个 **Find** 对象，该对象包含查找操作所需的条件。只读。

**语法**

**express.Find**

\*express是一个代表 Range 对象的变量。

### Range.FitTextWidth

该属性返回或设置 WPS 在当前选定内容或区域中填入文字的宽度（使用当前的度量单位）。 **Single** 类型，可读写。

**语法**

**express.FitTextWidth**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将当前选定内容填入 5 厘米宽的空间。*/
Selection.FitTextWidth = CentimetersToPoints(5)
```

### Range.Font

返回或设置 **Font** 对象，该对象代表指定对象的字符格式。 **Font** 类型，可读写。

**语法**

**express.Font**

\*express是一个代表 Range 对象的变量。

**说明**

要设置该属性，需指定一个返回 **Font** 对象的表达式。

**示例**

``` JavaScript
/*本示例将取消活动文档的“标题 1”样式中的加粗格式。*/
Application.ActiveDocument.Styles.Item(wdStyleHeading1).Font.Bold = false

/*本示例在 Arial 和 Times New Roman 之间切换活动文档中第二段的字体。*/
function test() {
  let myRange = Application.ActiveDocument.Paragraphs.Item(2).Range
  if(myRange.Font.Name == "Times New Roman"){
      myRange.Font.Name = "Arial"
  }
  else{
      myRange.Font.Name = "Times New Roman"
  }
}
```

### Range.FootnoteOptions

返回 **FootnoteOptions** 对象，该对象代表选定内容或区域的脚注。

**语法**

**express.FootnoteOptions**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例设置第二节中的编号规则以便在新节开始时重新编号。*/
function SetFootnoteOptionsRange() {
    ActiveDocument.Sections.Item(2).Range.FootnoteOptions
        .NumberingRule = wdRestartSection
}
```

### Range.Footnotes

返回一个 **Footnotes** 集合，该集合代表区域中的所有脚注。只读。

**语法**

**express.Footnotes**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.FormattedText

返回或设置一个 **Range** 对象，该对象包含指定区域或选定内容中进行过格式编排的文字。可读写。

**语法**

**express.FormattedText**

\*express是一个代表 Range 对象的变量。

**说明**

此属性返回 **Range** 对象，以及指定的区域或所选内容中的字符格式和文本。如果在区域或所选内容中有一个段落标记，则 **Range** 对象中包含段落格式。

在设置此属性时，区域中的文本会被格式文本替换。如果不想替换现有的文本，则在使用此属性之前使用 **Collapse** 方法（见第一个示例）。

### Range.FormFields

返回一个 **FormFields** 集合，该集合代表区域中的所有窗体域。只读。

**语法**

**express.FormFields**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例检索第二节第一个窗体域的类型。*/
function test()
{
let myType = ActiveDocument.Sections.Item(2).Range.FormFields.Item(1).Type
switch(myType){
    case wdFieldFormTextInput: thetype = "TextBox"
        break
    case wdFieldFormDropDown : thetype = "DropDown"
        break
    case wdFieldFormCheckBox : thetype = "CheckBox"
        break
}
}
```

### Range.Frames

返回一个 **Frames** 集合，该集合代表区域中的所有图文框。只读。

**语法**

**express.Frames**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例使活动文档第一节中的文字环绕图文框。*/
function test()
{
let rng = ActiveDocument.Sections.Item(1).Range
for(let aFrame = 1; aFrame <= rng.Frames.Count; aFrame++) {
    rng.Frames.Item(aFrame).TextWrap = true
}
}
```

### Range.GrammarChecked

如果已经检查了指定范围或文档的语法，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.GrammarChecked**

\*express是一个代表 Range 对象的变量。

**说明**

如果指定的区域或文档的部分或全部未进行语法检查，则该属性值为 **False** 。要重新检查区域或文档的语法，可将 **GrammarChecked** 属性设置为 **False** 。

### Range.GrammaticalErrors

返回一个 **ProofreadingErrors** 集合，该集合代表指定文档或区域中有语法检查错误的句子。只读。

**语法**

**express.GrammaticalErrors**

\*express是一个代表 Range 对象的变量。

**说明**

在每个句子中可有多个错误。如果文档中没有语法错误，则由 **GrammaticalErrors** 属性返回的 **ProofreadingErrors** 对象的 **Count** 属性的返回值为 0（零）。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例检查活动文档的第三段的语法错误，并显示包含一个或多个错误的每一句。*/
function test()
{
let myErrors = ActiveDocument.Paragraphs.Item(3).Range.GrammaticalErrors
for(let myerr = 1; myerr <= myErrors.Count; myerr++) {
    MsgBox(myErrors.Item(myerr).Text)
}
}
```

### Range.HighlightColorIndex

返回或设置指定区域的突出显示颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.HighlightColorIndex**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例删除选定内容中的突出显示格式。*/
Selection.Range.HighlightColorIndex = wdNoHighlight
```

``` JavaScript
/*本示例用黄色来突出显示活动文档中的每个书签。*/
function test()
{
let myBookmarks = ActiveDocument.Bookmarks
for(let abookmark = 1; abookmark <= myBookmarks.Count; abookmark++) {
    myBookmarks.Item(abookmark).Range.HighlightColorIndex = wdYellow
}
}
```

### Range.HorizontalInVertical

返回或设置位于垂直排列文字中的水平排列文字的格式。 **WdHorizontalInVerticalType** 类型，可读写。

**语法**

**express.HorizontalInVertical**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将垂直排列文字中的局部所选内容设置为水平排列的文字格式，并调整文字以适应垂直排列文字的行宽。*/
Selection.Range.HorizontalInVertical = wdHorizontalInVerticalFitInLine
```

### Range.HTMLDivisions

返回一个 **HTMLDivisions** 对象，该对象代表 Web 文档中的 HTML 划分。

**语法**

**express.HTMLDivisions**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例设置活动文档中三个嵌套划分的格式。本示例假定该活动文档为具有至少三个划分的 HTML 文档。*/
function test()
{
let htmlDiv1 = ActiveDocument.Range().HTMLDivisions.Item(1)
    htmlDiv1.Borders.Item(wdBorderLeft).Color = wdColorRed
    htmlDiv1.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleSingle
    htmlDiv1.Borders.Item(wdBorderRight).Color = wdColorRed
    htmlDiv1.Borders.Item(wdBorderRight).LineStyle = wdLineStyleSingle
    
    let htmlDiv2 = htmlDiv1.HTMLDivisions.Item(1)
    htmlDiv2.LeftIndent = InchesToPoints.Item(1)
    htmlDiv2.RightIndent = InchesToPoints.Item(1)
    htmlDiv2.Borders.Item(wdBorderTop).Color = wdColorBlue
    htmlDiv2.Borders.Item(wdBorderTop).LineStyle = wdLineStyleDouble
    htmlDiv2.Borders.Item(wdBorderBottom).Color = wdColorBlue
    htmlDiv2.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleDouble
        
    let htmlDiv3 = htmlDiv2.HTMLDivisions.Item(1)
    htmlDiv3.LeftIndent = InchesToPoints(1)
    htmlDiv3.RightIndent = InchesToPoints(1)
    htmlDiv3.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleDot
    htmlDiv3.Borders.Item(wdBorderRight).LineStyle = wdLineStyleDot
    htmlDiv3.Borders.Item(wdBorderTop).LineStyle = wdLineStyleDot
    htmlDiv3.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleDot

}
```

### Range.Hyperlinks

返回一个 **Hyperlinks** 集合，该集合代表指定范围内的所有超链接。只读。

**语法**

**express.Hyperlinks**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例显示活动文档中前十个段落中的每个超链接的名称。*/
function test()
{
let objRange = ActiveDocument.Range(
    ActiveDocument.Paragraphs.Item(1).Range.Start,
    ActiveDocument.Paragraphs.Item(10).Range.End)

let hplink = objRange.Hyperlinks   
for(let objLink = 1; objLink <= hplink.Count; objLink++) {
    if((hplink.Item(objLink).Address.toLowerCase().search("microsoft")) != -1) {
        MsgBox(hplink.Item(objLink).Name)
    }
}
}
```

### Range.ID

返回或设置特定范围的标识名称。可读写 **String** 类型。

**语法**

**express.ID**

\*express是一个代表 Range 对象的变量。

### Range.Information

返回有关指定范围的信息。只读 **Variant** 类型。

**语法**

**express.Information**

\*express是一个代表 Range 对象的变量。

**说明**

| **名称** | **必选/可选** | **数据类型**      | **说明**   |
|----------|---------------|-------------------|------------|
| *Type*   | 必选          | **WdInformation** | 消息类型。 |

**示例**

``` JavaScript
/*如果第十个单词位于某个表格中，则以下示例选定该表格。*/
function test() {
  if(Application.ActiveDocument.Words.Item(10).Information(wdWithInTable)) {
      Application.ActiveDocument.Words.Item(10).Tables.Item(1).Select()
  }
}
```

### Range.InlineShapes

返回一个 **InlineShape** 集合，该集合代表范围中的所有 **InlineShapes** 对象。只读。

**语法**

**express.InlineShapes**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例显示活动文档中的形状和内嵌形状的数目。*/
function test() {
  let doc = Application.ActiveDocument
  alert("InlineShape = " + doc.InlineShapes.Count + "\r" + "Shapes = " + doc.Shapes.Count)
}
```

### Range.IsEndOfRowMark

如果指定范围被折叠且位于表格中的行尾标志处，则该属性值为 **True** 。只读 **Boolean** 类型。

**语法**

**express.IsEndOfRowMark**

\*express是一个代表 Range 对象的变量。

**说明**

该属性与下面的表达式等效：

``` JavaScript
ActiveDocument.Range().Information(wdAtEndOfRowMarker)
```

**示例**

``` JavaScript
/*以下示例实现的功能是：折叠所选内容，如果插入点位于行尾（刚好在行结束标记之前），则选定当前行。*/
function test()
{
ActiveDocument.Range().Collapse(wdCollapseEnd)
if(ActiveDocument.Range().IsEndOfRowMark == true) {
    ActiveDocument.Range().Rows.Item(1).Select()
}
}
```

### Range.Italic

如果将字体或范围设置为倾斜格式，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.Italic**

\*express是一个代表 Range 对象的变量。

**说明**

此属性返回 **True** 、 **False** 或 **wdUndefined** （ **True** 和 **False** 混合组成），并且可设置为 **True** 、 **False** 或 **wdToggle** 。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个单词的格式设置为倾斜。*/
ActiveDocument.Words.Item(1).Italic = true
```

### Range.ItalicBi

如果将字体或范围设置为倾斜格式，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.ItalicBi**

\*express是一个代表 Range 对象的变量。

**说明**

该属性返回 **True** 、 **False** 或 **wdUndefined** （对于倾斜和非倾斜的混合文本）。可设置为 **True** 、 **False** 或 **wdToggle** 。

| 注释                                        |
|---------------------------------------------|
| **ItalicBi** 属性应用于从右向左语言的文本。 |

**示例**

``` JavaScript
/*以下示例将从右向左语言的活动文档中的第一段设置为倾斜格式。*/
ActiveDocument.Paragraphs.Item(1).Range.ItalicBi = true
```

### Range.Kana

返回或设置日文文本的指定区域是平假名还是片假名。 **WdKana** 类型，可读写。

**语法**

**express.Kana**

\*express是一个代表 Range 对象的变量。

**说明**

如果指定区域含有平假名和片假名的混合文本，或包含其他非日文文本，本属性返回 **wdUndefined** 。如果将 **Kana** 属性设为 **wdUndefined** ，则会产生一个错误。

**示例**

``` JavaScript
/*本示例显示当前选定内容中所含的日文文本的类型。*/
function test()
{
myKana = Selection.Range.Kana
switch(myKana) {
    case wdKanaHiragana:
        MsgBox("This text is hiragana.")
        break
    case wdKanaKatakana:
        MsgBox("This text is katakana.")
        break
    case wdUndefined:
        MsgBox("This text is a mix of " 
            + "hiragana and katakana.")
        break
}
}
```

### Range.LanguageDetected

返回或设置一个值，该值指定 WPS 是否已经检测过指定文本的语言。可读/写 **Boolean** 类型。

**语法**

**express.LanguageDetected**

\*express是一个代表 Range 对象的变量。

**说明**

检查以前所有语言检测结果的 **LanguageID** 属性。

调用 **DetectLanguage** 方法时， **LanguageDetected** 属性被设置为 **True** 。要重新检测指定文本的语言，必须先将 **LanguageDetected** 属性设置为 **False** 。

**示例**

``` JavaScript
/*本示例检查活动文档以确定用于编写它的语言，然后显示结果。*/
function test()
{
let rng = ActiveDocument.Range
    if(rng.LanguageDetected == true) {
        let x = MsgBox("This document has already " 
            + "been checked. Do you want to check " 
            + "it again?",jsYesNo)
        if(x == jsResultYes) {
            rng.LanguageDetected = false
            rng.DetectLanguage()
        }
    }
    else {
        rng.DetectLanguage()
    }
    if(rng.Range.LanguageID == wdEnglishUS) {
        MsgBox("This is a U.S. English document.")
    }
    else {
        MsgBox("This is not a U.S. English document.")
    }
}
```

### Range.LanguageID

返回或设置一个 **WdLanguageID** 常量，该常量代表指定范围的语言。可读写。

**语法**

**express.LanguageID**

\*express是一个代表 Range 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdLanguageID** 常量可能不可用。

**示例**

``` JavaScript
/*以下示例将活动文档的第二段格式设置为法语格式，然后添加新的自定义词典，此词典用于查阅法语文本。*/
function test()
{
ActiveDocument.Paragraphs.Item(2).Range.LanguageID = wdFrench
let myDictionary = CustomDictionaries.Add("French.dic")
    myDictionary.LanguageSpecific = true
    myDictionary.LanguageID = wdFrench
}
```

### Range.LanguageIDFarEast

返回或设置指定对象的东亚语言。可读写 **WdLanguageID** 类型。

**语法**

**express.LanguageIDFarEast**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的语言设置为朝鲜语。*/
ActiveDocument.Paragraphs.Item(1).Range.LanguageIDFarEast = wdKorean
```

### Range.LanguageIDOther

返回或设置指定范围的语言。可读写 **WdLanguageID** 类型。

**语法**

**express.LanguageIDOther**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例将所选内容的语言设置为法语。*/
Selection.Range.LanguageIDOther = wdFrench
```

### Range.ListFormat

返回一个 **ListFormat** 对象，该对象代表某区域中所有的列表格式特征。只读。

**语法**

**express.ListFormat**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中 3 到 6 段中的一个区域赋给变量 myDoc。然后，根据此区域中是否存在列表格式，来设置默认的多级符号列表格式或清除已有编号格式。*/
function test()
{
let myDoc = ActiveDocument
let myRange = 
    myDoc.Range(myDoc.Paragraphs.Item(3).Range.Start, 
    myDoc.Paragraphs.Item(6).Range.End)
myRange.ListFormat.ApplyOutlineNumberDefault()
}
```

``` JavaScript
/*本示例对选定内容中所有段落应用“项目符号和编号”对话框“编号”选项卡中的第二个列表模板。*/
function test()
{
Selection.Range.ListFormat.ApplyListTemplate(
    ListGalleries.Item(wdNumberGallery).ListTemplates.Item(2))
}
```

### Range.ListParagraphs

返回一个 **ListParagraphs** 集合，该集合代表范围中的所有编号段落。只读。

**语法**

**express.ListParagraphs**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.ListStyle

返回一个 **Variant** 类型的值，该值代表用于设置项目符号列表或编号列表的样式。只读。

**语法**

**express.ListStyle**

\*express是一个代表 Range 对象的变量。

### Range.Locks

返回 CoAuthLocks 集合对象，该对象代表范围中的所有锁。只读。

**语法**

**express.Locks**

\*express是一个代表 Range 对象的变量。

**说明**

使用 **Locks** 属性可返回 CoAuthLocks 集合。

| 注释                                                                                             |
|--------------------------------------------------------------------------------------------------|
| 该属性仅可用于支持共同创作的文档。如果尝试在不支持共同创作的文档中访问该属性，将导致运行时错误。 |

**示例**

``` JavaScript
/*下面的代码示例显示活动文档的第一个段落中的锁数量。*/
MsgBox(ActiveDocument.Paragraphs.Item(1).Range.Locks.Count)
```

### Range.NextStoryRange

返回一个 **Range** 对象，该对象代表下一个 文章 。 **Range** 类型，只读。

**语法**

**express.NextStoryRange**

\*express是一个代表 Range 对象的变量。

**说明**

下表列出了依据文章的类型而返回的区域。

| 文章的类型                                                                                                                                                                   | 用 NextStoryRange 方法返回的项目 |
|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------------------|
| **wdMainTextStory** 、 **wdFootnotesStory** 、 **wdEndnotesStory** 和 **wdCommentsStory**                                                                                    | 总是返回 **Nothing**             |
| **wdTextFrameStory**                                                                                                                                                         | 下一组链接文本框的文章           |
| **wdEvenPagesHeaderStory** 、 **wdPrimaryHeaderStory** 、 **wdEvenPagesFooterStory** 、 **wdPrimaryFooterStory** 、 **wdFirstPageHeaderStory** 、 **wdFirstPageFooterStory** | 下一节中相同类型的文章           |

### Range.NoProofing

如果拼写和语法检查程序忽略指定文本，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.NoProofing**

\*express是一个代表 Range 对象的变量。

**说明**

如果只将某些指定文本的 **NoProofing** 属性设置为 **True** ，则该属性返回 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例实现的功能是：标记当前所选内容，以便拼写和语法检查程序将其忽略。*/
Selection.Range.NoProofing = true
```

### Range.OMaths

返回一个 **OMaths** 集合，该集合代表指定区域内的 **OMath** 对象。只读。

**语法**

**express.OMaths**

\*express是一个代表 Range 对象的变量。

### Range.Orientation

在启用了“文字方向”功能时返回或设置范围中文字的方向。可读写 **WdTextOrientation** 类型。

**语法**

**express.Orientation**

\*express是一个代表 Range 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdTextOrientation** 常量可能不可用。

您可以设置文本框架或者恰好位于文本框架内的范围的方向。有关文本框架和文本框之间的区别的信息，请参阅 **TextFrame** 对象。

### Range.PageSetup

返回一个 **PageSetup** 对象，该对象与指定范围相关联。

**语法**

**express.PageSetup**

\*express是一个代表 Range 对象的变量。

### Range.ParagraphFormat

返回或设置一个 **ParagraphFormat** 对象，该对象代表指定范围的段落设置。可读/写。

**语法**

**express.ParagraphFormat**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例对包含 MyDoc.doc 所有内容的有关范围设置段落格式：2 倍行距，并且在 0.25 英寸的位置设置一个自定义制表位。*/
function test()
{
let myRange = Documents.Item("MyDoc.doc").Content
let myPFormat = myRange.ParagraphFormat
    myPFormat.Space2()
    myPFormat.TabStops.Add(InchesToPoints(.25))
}
```

### Range.Paragraphs

返回一个 **Paragraphs** 集合，该集合代表指定范围中的所有段落。只读。

**语法**

**express.Paragraphs**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例将活动文档中第一节中所有段落的集合的行距设置为单倍行距。*/
function test()
{
ActiveDocument.Sections.Item(1).Range.Paragraphs.LineSpacingRule = 
    wdLineSpaceSingle
}
```

### Range.ParagraphStyle

返回一个 **Variant** 类型的值，该值代表用于设置段落格式的样式。只读。

**语法**

**express.ParagraphStyle**

\*express是一个代表 Range 对象的变量。

### Range.Parent

返回一个 **Object** 类型值，该值代表指定 **Range** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Range 对象的变量。

### Range.ParentContentControl

返回一个 **ContentControl** 对象，该对象代表指定区域的父内容控件。只读。

**语法**

**express.ParentContentControl**

\*express是一个代表 Range 对象的变量。

**说明**

如果指定区域没有父内容控件，则此属性返回 **Nothing** 。

### Range.PreviousBookmarkID

返回最后一个书签的编号，该书签从指定范围的前面或与指定范围相同的位置开始。只读 **Long** 类型。

**语法**

**express.PreviousBookmarkID**

\*express是一个代表 Range 对象的变量。

**说明**

如果没有相应的书签，则该属性返回 0（零）。

**示例**

``` JavaScript
/*以下示例显示第二段前面的书签的名称。*/
function test()
{
let num = ActiveDocument.Paragraphs.Item(2).Range.PreviousBookmarkID
if(num != 0) {
    MsgBox(ActiveDocument.Content.Bookmarks.Item(num).Name)
}
}
```

### Range.ReadabilityStatistics

返回一个 **ReadabilityStatistics** 集合，该集合代表指定文档或范围的可读性统计信息。只读。

**语法**

**express.ReadabilityStatistics**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

``` JavaScript
/*以下示例显示文档 1 的每一种可读性统计信息及其值。*/
function test()
{
let RStatistics = Documents.Item(1).ReadabilityStatistics
for(let rs = 1; rs <= RStatistics.Count; rs++) {
    MsgBox(RStatistics.Item(rs).Name + " - " + rs.Value)
}
}
```

### Range.Revisions

返回一个 **Revisions** 集合，该集合代表范围中的修订。只读。

**语法**

**express.Revisions**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

``` JavaScript
/*以下示例显示活动文档中第一节中修订的数目。*/
MsgBox(ActiveDocument.Sections.Item(1).Range.Revisions.Count)
```

``` JavaScript
/*以下示例接受所选内容的第一段中的所有修订。*/
function test()
{
let myRange = Selection.Paragraphs.Item(1).Range
myRange.Revisions.AcceptAll()
}
```

### Range.Rows

返回一个 **Rows** 集合，该集合代表范围中的所有表格行。只读。

**语法**

**express.Rows**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

### Range.Scripts

返回一个 **Scripts** 集合，该集合代表指定对象中 HTML 脚本的集合。

**语法**

**express.Scripts**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例测试指定范围中的第二个 Script 对象以确定其语言。*/
function test()
{
let lage = Selection.Range.Scripts.Item(2).Language
switch(lage) {
    case msoScriptLanguageASP:
        MsgBox("Active Server Pages")
        break
    case msoScriptLanguageVisualBasic:
        MsgBox("VBScript")
        break
    case msoScriptLanguageJava:
        MsgBox("JavaScript")
        break
    case msoScriptLanguageOther:
        MsgBox("Unknown type of script")
}
}
```

### Range.Sections

返回一个 **Sections** 集合，该集合代表指定范围中的各节。只读。

**语法**

**express.Sections**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

### Range.Sentences

返回一个 **Sentences** 集合，该集合代表范围中的所有句子。只读。

**语法**

**express.Sentences**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

``` JavaScript
/*以下示例显示活动文档中第一段的句子数。*/
function test()
{
MsgBox(ActiveDocument.Paragraphs.Item(1).Range 
    .Sentences.Count + " sentences")
}
```

### Range.Shading

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例对所选内容的第一段应用黄色底纹。*/
function test()
{
let myShading = Selection.Paragraphs.Item(1).Shading
    myShading.Texture = wdTexture12Pt5Percent
    myShading.BackgroundPatternColorIndex = wdYellow
    myShading.ForegroundPatternColorIndex = wdBlack
}
```

``` JavaScript
/*以下示例将活动文档中第一个单词的底纹设置为 10%。*/
ActiveDocument.Words.Item(1).Shading.Texture = wdTexture10Percent
```

### Range.ShapeRange

返回一个 **ShapeRange** 集合，该集合代表指定范围中的所有 **Shape** 对象。只读。

**语法**

**express.ShapeRange**

\*express是一个代表 Range 对象的变量。

**说明**

图形范围可以包含绘图、图形、图片、OLE 对象、ActiveX 控件、文本对象和标注。

**示例**

``` JavaScript
/*以下示例将活动文档中所有图形的填充前景色设置为紫色。*/
Application.ActiveDocument.Content.ShapeRange.Fill.ForeColor.RGB = (255, 0, 255)
```

### Range.ShowAll

如果显示所有非打印字符（如隐藏文字、制表符、空格和段落标记），则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.ShowAll**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动文档中显示所有非打印字符。*/
ActiveDocument.Range.ShowAll = true
```

### Range.SpellingChecked

如果已对指定的区域或文档完成拼写检查，则该属性值为 **True** 。如果所有或部分区域或文档尚未进行拼写检查，则该属性值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.SpellingChecked**

\*express是一个代表 Range 对象的变量。

**说明**

若要重新检查某一区域或文档的拼写，请将 **SpellingChecked** 属性设置为 **False** 。

若要查看该区域或文档中是否含有拼写错误，请使用 **SpellingErrors** 属性。

**示例**

``` JavaScript
/*以下示例判定是否已经检查过活动文档第一节中的拼写。如果没有，则开始拼写检查。*/
function test()
{
let myRange = ActiveDocument.Sections.Item(1).Range
let isChecked = myRange.SpellingChecked
if(isChecked == false) {
    myRange.CheckSpelling()
}
else {
    MsgBox("Spelling has already been checked in the range.")
}
}
```

### Range.SpellingErrors

返回一个 **ProofreadingErrors** 集合，该集合代表指定范围中标识为拼写错误的单词。只读。

**语法**

**express.SpellingErrors**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

``` JavaScript
/*以下示例检查指定范围是否存在拼写错误，并显示找到的每个错误。*/
function test()
{
let myErrors = ActiveDocument.Paragraphs.Item(3).Range.SpellingErrors
if(myErrors.Count == 0) {
    MsgBox("No spelling errors found.")
}
else {
    for(let myErr = 1;myErr <= myErrors.Count;myErr++) {
        Msgbox(myErrors.Item(myErr).Text)
    }
}
}
```

### Range.Start

返回或设置某区域中起始字符的位置。 **Long** 类型，可读写。

**语法**

**express.Start**

\*express是一个代表 Range 对象的变量。

**说明**

**Range** 对象包括起始字符和结束字符位置。起始字符位置是指距文档开头部分最近的字符位置。如果将该属性的值设置为大于 **End** 属性的值，则 **End** 属性值会设置为与 **Start** 属性值相同。

该属性返回起始字符相对于文档开头部分的位置。文本主体部分 ( **wdMainTextStory** ) 的起始字符位置为 0（零）。通过设置该属性可以更改选定内容、区域或书签的大小。

**示例**

``` JavaScript
/*本示例返回活动文档第二段的起始位置和第四段的结束位置。这些字符位置用于创建区域 myRange。*/
function test() {
  let pos = Application.ActiveDocument.Paragraphs.Item(2).Range.Start
  let pos2 = Application.ActiveDocument.Paragraphs.Item(4).Range.End
  let myRange = Application.ActiveDocument.Range(pos,pos2)
}

/*本示例将 myRange 起始字符的位置向右移动一个字符（这会使该区域缩小一个字符）。*/
function test() {
  let myRange = Application.Selection.Range
  myRange.SetRange(myRange.Start + 1, myRange.End)
}
```

### Range.StoryLength

返回包含指定区域的 文字部分 中的字符数。 **Long** 类型，只读。

**语法**

**express.StoryLength**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例确定活动文档的页眉部分是否为空。如果页眉部分不为空，则在消息框中显示页眉的内容。如果为空，则 StoryLength 属性返回 1 作为最后的段落标记。*/
function test()
{
let myRange = ActiveDocument.Sections.Item(1) 
    .Headers.Item(wdHeaderFooterPrimary).Range
if(myRange.StoryLength > 1) {
    MsgBox(myRange.Text)
}
}
```

``` JavaScript
/*如果文档为空，则以下示例在关闭文档时不保存所做的更改。*/
function test()
{
if(ActiveDocument.Content.StoryLength == 1) {
    ActiveDocument.Close(wdDoNotSaveChanges)
}
}
```

### Range.StoryType

返回指定范围、所选内容或书签的 文字部分 类型。只读 **WdStoryType** 类型。

**语法**

**express.StoryType**

\*express是一个代表 Range 对象的变量。

### Range.Style

返回或设置指定对象的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

\*express是一个代表 Range 对象的变量。

**说明**

要设置该属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。如果返回包含多个样式的范围的样式，则只返回第一个字符样式或段落样式。

**示例**

| 注释                                                   |
|--------------------------------------------------------|
| **Characters** 集合的每个元素都是一个 **Range** 对象。 |

``` JavaScript
function test()
{
for(let c = 1; c <= Selection.Characters.Count; c++) {
    MsgBox(Selection.Characters.Item(c).Style)
}
}
```

### Range.Subdocuments

返回一个 **Subdocuments** 集合，该集合代表指定范围或文档中的所有子文档。只读。

**语法**

**express.Subdocuments**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例显示嵌入在活动文档中的子文档的数量。*/
MsgBox(ActiveDocument.Range.Subdocuments.Count)
```

### Range.SynonymInfo

返回一个 **SynonymInfo** 对象，该对象包含同义词库中有关某范围的内容的同义词、反义词或相关单词和表达方式的信息。

**语法**

**express.SynonymInfo**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例返回所选内容的第一个含义的同义词列表。*/
function test()
{
let Slist = Selection.Range.SynonymInfo.SynonymList(1)
for(let i = 0;i <= Slist.length - 1;i++) {
    MsgBox(Slist[i])
}
}
```

### Range.Tables

返回一个 **Tables** 集合，该集合代表指定范围内的所有表格。只读。

**语法**

**express.Tables**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例在活动文档中创建一个 5x5 表格，然后对其应用预定义格式。*/
function test()
{
Selection.Collapse(wdCollapseStart)
let myTable = ActiveDocument.Tables.Add(Selection.Range,5,5)
myTable.AutoFormat(wdTableFormatClassic2)
}
```

``` JavaScript
/*以下示例在活动文档中第一个表格的首列中插入数值和文本。*/
function test()
{
let num = 90
let myCells = ActiveDocument.Tables.Item(1).Columns.Item(1).Cells
for(let acell = 1; acell <= myCells.Count; acell++) {
    myCells.Item(acell).Range.Text = num + " Sales"
    num = num + 1
}
}
```

### Range.TableStyle

返回一个 **Variant** 类型的值，该值代表用于设置表格格式的样式。只读。

**语法**

**express.TableStyle**

\*express是一个代表 Range 对象的变量。

### Range.Text

返回或设置指定区域或选定内容中的文本。 **String** 类型，可读写。

**语法**

**express.Text**

\*express是一个代表 Range 对象的变量。

**说明**

**Text** 属性返回该区域的无格式纯文本。如果设置该属性，则将替换该区域中的现有文本。

**示例**

``` JavaScript
/*本示例用“Dear”替换活动文档的第一个词。*/
function test() {
  let myRange = Application.ActiveDocument.Words.Item(1)
  myRange.Text = "Dear "
}
```

### Range.TextRetrievalMode

返回一个 **TextRetrievalMode** 对象，该对象控制从指定 **区域** 检索文字的方式。可读写。

**语法**

**express.TextRetrievalMode**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例检索选定文字（排除隐藏文字），并将其插入活动文档第三段的开始处。*/
function test()
{
if(Selection.Type == wdSelectionNormal) {
    let Range1 = Selection.Range
    Range1.TextRetrievalMode.IncludeHiddenText = false
    let Range2 = ActiveDocument.Paragraphs.Item(2).Range
    Range2.InsertAfter(Range1.Text)
}
}
```

``` JavaScript
/*本示例在大纲视图中检索并显示前三个段落。*/
function test()
{
let myRange = ActiveDocument.Range(ActiveDocument
    .Paragraphs.Item(1).Range.Start, 
    ActiveDocument.Paragraphs.Item(3).Range.End)
myRange.TextRetrievalMode.ViewType = wdOutlineView
MsgBox(myRange.Text)
}
```

``` JavaScript
/*本示例先排除引用选定文字的区域中的域代码和隐藏文字，然后在一个消息框中显示文字。*/
function test()
{
if(Selection.Type == wdSelectionNormal) {
    let aRange = Selection.Range
    aRange.TextRetrievalMode.IncludeHiddenText = false
    aRange.TextRetrievalMode.IncludeFieldCodes = false
    MsgBox(aRange.Text)
}
}
```

### Range.TopLevelTables

返回一个 **Tables** 集合，该集合代表当前范围最外部嵌套层上的表格。只读。

**语法**

**express.TopLevelTables**

\*express是一个代表 Range 对象的变量。

**说明**

此方法返回一个集合，该集合仅包含当前范围的上下文中最外部嵌套层上的表格。这些表格可能不在整套嵌套表格的最外嵌套层中。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

以下示例新建一个文档，创建一个三层嵌套表格，并在每张表格的第一个单元格中填入该表格所在的嵌套层数。接着选定第二层表格的第二列，然后选定所选内容中的顶层表格的第一列。尽管最里面的表格在整套嵌套表格的上下文关系中并非顶层表格，但仍会被选定。

``` JavaScript
function test()
{
Documents.Add()
ActiveDocument.Tables.Add(Selection.Range,
    3, 3, wdWord9TableBehavior, wdAutoFitContent)
let aRange = ActiveDocument.Tables.Item(1).Range
    aRange.Copy()
    aRange.Cells.Item(1).Range.Text = aRange.Cells.Item(1).NestingLevel
    aRange.Cells.Item(5).Range.PasteAsNestedTable()
    let acRange = aRange.Cells.Item(5).Tables.Item(1).Range
        acRange.Cells.Item(1).Range.Text = acRange.Cells.Item(1).NestingLevel
        acRange.Cells.Item(5).Range.PasteAsNestedTable()
        let accRange = acRange.Cells.Item(5).Tables.Item(1).Range
            accRange.Cells.Item(1).Range.Text = 
                accRange.Cells.Item(1).NestingLevel       
        acRange.Columns.Item(2).Select()
        Selection.Range.TopLevelTables(1).Select()
}
```

### Range.TwoLinesInOne

返回或设置 WPS 是否将两行文本合并为一行，并指定括住文本的字符（如果有）。 **WdTwoLinesInOneType** 类型，可读写。

**语法**

**express.TwoLinesInOne**

\*express是一个代表 Range 对象的变量。

**说明**

将 **TwoLinesInOne** 属性设为 **wdTwoLinesInOneNoBrackets** 可将两行文本合并为一行，且不将文本括在任何字符中。将 **TwoLinesInOne** 属性设为 **wdTwoLinesInOneNone** 可将合并为一行的文本还原为两行。

**示例**

``` JavaScript
/*本示例设置当前选定内容的格式，使两行文本合并为一行，并将其括在括号中。*/
Selection.Range.TwoLinesInOne = wdTwoLinesInOneParentheses
```

### Range.Underline

返回或设置应用于范围的下划线的类型。可读写 **WdUnderline** 类型。

**语法**

**express.Underline**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例对活动文档中的第四个单词应用双下划线。*/
Application.ActiveDocument.Words.Item(4).Underline = wdUnderlineDouble
```

### Range.Updates

返回 CoAuthUpdates 集合对象，该对象代表范围中的所有可用更新。只读。

**语法**

**express.Updates**

\*express是一个代表 Range 对象的变量。

**说明**

使用 **Updates** 属性可返回 CoAuthUpdates 集合。

| 注释                                                                                             |
|--------------------------------------------------------------------------------------------------|
| 该属性仅可用于支持共同创作的文档。如果尝试在不支持共同创作的文档中访问该属性，将导致运行时错误。 |

**示例**

``` JavaScript
/*下面的代码示例显示活动文档的第一个段落中的可用更新数量。*/
function test()
{
let countOfUpdates = ActiveDocument.Paragraphs.Item(1).Range.Updates.Count

MsgBox("The number of updates is " + countOfUpdates)
}
```

### Range.WordOpenXML

返回一个 **String** 类型的值，该值以 WPS Open XML 格式表示区域中包含的 XML。只读。

**语法**

**express.WordOpenXML**

\*express是一个代表 Range 对象的变量。

**说明**

此属性只返回文档中为表示指定区域所需的 XML。

### Range.Words

返回一个 **Words** 集合，该集合代表范围中的所有单词。只读。

**语法**

**express.Words**

\*express是一个代表 Range 对象的变量。

**说明**

文档中的标点符号和段落标记包括在 **Words** 集合中。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例显示所选内容的字数。段落标记、部分单词和标点符号都统计在内。*/
MsgBox("There are " + Selection.Words.Count + " words.")
```

``` JavaScript
/*以下示例遍历 myRange（范围为从活动文档的开始到所选内容的末尾）中的单词，并删除该范围中出现的单词“Franklin”（包含尾部空格）。*/
function test()
{
let myRange = ActiveDocument.Range(0,Selection.End)
for(let aWord = 1; aWord <= myRange.Words.Count; aWord++) {
    if(myRange.Words.Item(aWord).Text == "Franklin ") {
        myRange.Words.Item(aWord).Delete()
    }
}
}
```

### Range.XMLNodes

返回一个 **XMLNodes** 集合，该集合代表指定区域中的 XML 元素（包括任何只是部分属于该区域的元素）。只读。

**语法**

**express.XMLNodes**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例访问活动文档的开头 200 个字符中包含的 XML 元素。*/
function test()
{
let objRange = ActiveDocument.Range(1, 200)

let objNode = objRange.XMLNodes.Item(1)
}
```

### Range.XMLParentNode

返回一个 **XMLNode** 对象，该对象代表区域的父级 XML 节点。只读。

**语法**

**express.XMLParentNode**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例返回指定区域的父节点。*/
function test()
{
let objRange = ActiveDocument.Range(1, 200)
let objNode = objRange.XMLParentNode
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Range](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
