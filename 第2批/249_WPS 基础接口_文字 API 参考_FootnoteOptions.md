# FootnoteOptions 对象

## 简介

代表分配给文档中脚注的范围或所选内容的属性。

## 说明

使用 Range 或 Selection 对象可返回 FootnoteOptions 对象。使用 FootnoteOptions 对象可为文档的不同区域分配不同的脚注属性。例如，您可能希望长文档说明部分的脚注以小写字母显示，而文档中其余部分的脚注以星号显示。以下示例使用 NumberingRule 、 NumberStyle 和 StartingNumber 属性设置活动文档的第一节中的尾注格式。

``` JavaScript
function BookIntro(){
    //Sets the range as section one of the active document
    let rngIntro = Application.ActiveDocument.Sections.Item(1).Range
    //Formats the EndnoteOptions properties
    rngIntro.FootnoteOptions.NumberingRule = wdRestartPage
    rngIntro.FootnoteOptions.NumberStyle = wdNoteNumberStyleLowercaseLetter
    rngIntro.FootnoteOptions.StartingNumber = 1
}
```

## 属性

| 序号 | 名称                                              | 说明                                                                                |
|:----:|:--------------------------------------------------|:------------------------------------------------------------------------------------|
|  1   | [Application](#FootnoteOptions.Application)       | 返回一个代表 WPS 应用程序的 Application 对象。                                      |
|  2   | [Creator](#FootnoteOptions.Creator)               | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。        |
|  3   | [Location](#FootnoteOptions.Location)             | 返回或设置所有脚注的位置。 WdFootnoteLocation 类型，可读写。                        |
|  4   | [NumberStyle](#FootnoteOptions.NumberStyle)       | 返回或设置脚注的编号样式。 WdNoteNumberStyle 类型，可读写。                         |
|  5   | [NumberingRule](#FootnoteOptions.NumberingRule)   | 返回或设置在分页符或分节符之后的脚注或尾注的编号方式。可读写 WdNumberingRule 类型。 |
|  6   | [Parent](#FootnoteOptions.Parent)                 | 返回一个 Object 类型值，该值代表指定 FootnoteOptions 对象的父对象。                 |
|  7   | [StartingNumber](#FootnoteOptions.StartingNumber) | 返回或设置脚注的起始编号。 Long 类型，可读写。                                      |

## 成员属性

### FootnoteOptions.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 FootnoteOptions 对象的变量。

### FootnoteOptions.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FootnoteOptions 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### FootnoteOptions.Location

返回或设置所有脚注的位置。 **WdFootnoteLocation** 类型，可读写。

**语法**

**express.Location**

\*express是一个代表 FootnoteOptions 对象的变量。

**示例**

``` JavaScript
/*本示例将脚注置于每页的底端。*/
Application.ActiveDocument.Footnotes.Location = wdBottomOfPage
```

### FootnoteOptions.NumberStyle

返回或设置脚注的编号样式。 **WdNoteNumberStyle** 类型，可读写。

**语法**

**express.NumberStyle**

\*express是一个代表 FootnoteOptions 对象的变量。

**说明**

某些 **WdNoteNumberStyle** 常量可能不可用，具体取决于所选择或安装的语言支持（例如，美国英语）。

### FootnoteOptions.NumberingRule

返回或设置在分页符或分节符之后的脚注或尾注的编号方式。可读写 **WdNumberingRule** 类型。

**语法**

**express.NumberingRule**

\*express是一个代表 FootnoteOptions 对象的变量。

**示例**

``` JavaScript
/*如果将第一节中的脚注编号设置为在每个分节符后重新开始，则以下示例设置每页都重新开始编号。*/
function test() {
  let myRange = Application.ActiveDocument.Sections.Item(1).Range
  if(myRange.Footnotes.NumberingRule == wdRestartSection) {
      myRange.Footnotes.NumberingRule = wdRestartPage
  }
}
```

### FootnoteOptions.Parent

返回一个 **Object** 类型值，该值代表指定 **FootnoteOptions** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 FootnoteOptions 对象的变量。

### FootnoteOptions.StartingNumber

返回或设置脚注的起始编号。 **Long** 类型，可读写。

**语法**

**express.StartingNumber**

\*express是一个代表 FootnoteOptions 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/FootnoteOptions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
