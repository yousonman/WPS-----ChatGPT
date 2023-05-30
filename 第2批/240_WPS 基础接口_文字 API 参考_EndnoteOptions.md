# EndnoteOptions 对象

## 简介

代表指定给文档中尾注的某范围或所选内容的属性。

## 说明

使用 Range 或 Selection 对象的 EndnoteOptions 属性可返回 EndnoteOptions 对象。

使用 EndnoteOptions 对象可为文档的不同区域指定不同的尾注属性。例如，您可能希望将长文档说明部分的尾注用小写罗马数字显示，而文档中其余部分的尾注用阿拉伯数字显示。以下示例使用 NumberingRule 、 NumberStyle 和 StartingNumber 属性设置活动文档中第一节的尾注格式。

``` JavaScript
function BookIntro() {
    //Sets the range as section one of the active document
    let rngIntro = Application.ActiveDocument.Sections.Item(1).Range

    //Formats the EndnoteOptions properties
    let re = rngIntro.EndnoteOptions
        re.NumberingRule = wdRestartSection
        re.NumberStyle = 2
        re.StartingNumber = 1
}
```

## 属性

| 序号 | 名称                                             | 说明                                                                          |
|:----:|:-------------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#EndnoteOptions.Application)       |                                                                               |
|  2   | [Creator](#EndnoteOptions.Creator)               | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。  |
|  3   | [Location](#EndnoteOptions.Location)             | 返回或设置所有尾注的位置。 WdEndnoteLocation 类型，可读写。                   |
|  4   | [NumberStyle](#EndnoteOptions.NumberStyle)       | 返回或设置尾注的编号样式。 WdNoteNumberStyle 类型，可读写。                   |
|  5   | [NumberingRule](#EndnoteOptions.NumberingRule)   | 返回或设置在分页符或分节符之后脚注或尾注的编号方式。可读写 WdNumberingRule 。 |
|  6   | [Parent](#EndnoteOptions.Parent)                 | 返回一个 Object 类型值，该值代表指定 EndnoteOptions 对象的父对象。            |
|  7   | [StartingNumber](#EndnoteOptions.StartingNumber) | 返回或设置尾注的起始编号。 Long 类型，可读写。                                |

## 成员属性

### EndnoteOptions.Application

**语法**

**express.Application**

\*express是一个代表 EndnoteOptions 对象的变量。

### EndnoteOptions.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 EndnoteOptions 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### EndnoteOptions.Location

返回或设置所有尾注的位置。 **WdEndnoteLocation** 类型，可读写。

**语法**

**express.Location**

\*express是一个代表 EndnoteOptions 对象的变量。

**示例**

``` JavaScript
/*本示例将所有的尾注置于节的结尾。*/
Application.ActiveDocument.Endnotes.Location = wdEndOfSection
```

### EndnoteOptions.NumberStyle

返回或设置尾注的编号样式。 **WdNoteNumberStyle** 类型，可读写。

**语法**

**express.NumberStyle**

\*express是一个代表 EndnoteOptions 对象的变量。

**说明**

某些 **WdNoteNumberStyle** 常量可能不可用，具体取决于所选择或安装的语言支持（例如，美国英语）。

### EndnoteOptions.NumberingRule

返回或设置在分页符或分节符之后脚注或尾注的编号方式。可读写 WdNumberingRule 。

**语法**

**express.NumberingRule**

\*express是一个代表 EndnoteOptions 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档的每一分节符后重新开始尾注的编号。*/
Application.ActiveDocument.Endnotes.NumberingRule = wdRestartSection
```

### EndnoteOptions.Parent

返回一个 **Object** 类型值，该值代表指定 **EndnoteOptions** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 EndnoteOptions 对象的变量。

### EndnoteOptions.StartingNumber

返回或设置尾注的起始编号。 **Long** 类型，可读写。

**语法**

**express.StartingNumber**

\*express是一个代表 EndnoteOptions 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/EndnoteOptions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
