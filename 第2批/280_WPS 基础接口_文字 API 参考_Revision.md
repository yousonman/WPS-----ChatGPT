# Revision 对象

## 简介

代表由修订标记所标记的修改。 Revision 对象是 Revisions 集合的一个成员。 Revisions 集合包含某范围或文档中的所有修订标记。

## 方法

| 序号 | 名称                       | 说明                                             |
|:----:|:---------------------------|:-------------------------------------------------|
|  1   | [Accept](#Revision.Accept) | 接受指定修订、删除修订标记并将更改合并到文档中。 |
|  2   | [Reject](#Revision.Reject) | 拒绝接受指定的修订。将删除修订标记，不改变原文   |

## 属性

| 序号 | 名称                                             | 说明                                                                                      |
|:----:|:-------------------------------------------------|:------------------------------------------------------------------------------------------|
|  1   | [Application](#Revision.Application)             | 返回一个代表 WPS 应用程序的 Application 对象。                                            |
|  2   | [Author](#Revision.Author)                       | 返回或设置完成指定修订的用户姓名                                                          |
|  3   | [Cells](#Revision.Cells)                         | 返回一个 Cells 集合，该集合代表已经用修订标记进行标记的表格单元格。只读。                 |
|  4   | [Creator](#Revision.Creator)                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。              |
|  5   | [Date](#Revision.Date)                           | 修订的日期和时间                                                                          |
|  6   | [FormatDescription](#Revision.FormatDescription) | 返回一个 String 类型的值，该值代表对修订中所作格式更改的说明。只读。                      |
|  7   | [Index](#Revision.Index)                         | 返回一个 Long 类型的值，该值代表项目在集合中的位置。只读。                                |
|  8   | [MovedRange](#Revision.MovedRange)               | 返回一个 Range 对象，该对象代表在带修订的文档中从一个位置移至另一个位置的文本区域。只读。 |
|  9   | [Parent](#Revision.Parent)                       | 返回一个 Object 类型值，该值代表指定 Revision 对象的父对象。                              |
|  10  | [Range](#Revision.Range)                         | 返回一个 Range 对象，该对象代表修订标记内包含的文档部分。                                 |
|  11  | [Style](#Revision.Style)                         | 返回一个 Style 对象，该对象代表与一个修订标记相关联的样式。                               |
|  12  | [Type](#Revision.Type)                           | 返回修订的类型。只读 [WdRevisionType]() 类型。                                            |

## 成员方法

### Revision.Accept

接受指定修订、删除修订标记并将更改合并到文档中。

**语法**

**express.Accept ()**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
//接受当前文档的第一个修订
function test()
{
    let rev;
    if (Application.ActiveDocument.Revisions.Count >= 1)
        rev = Application.ActiveDocument.Revisions.Item(1)
    rev.Accept();
}
```

### Revision.Reject

拒绝接受指定的修订。将删除修订标记，不改变原文

**语法**

**express.Reject ()**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
//拒绝当前文档的第二个修订。
function test()
{
    let rev;
    if (Application.ActiveDocument.Revisions.Count >= 2)
        rev = Application.ActiveDocument.Revisions.Item(2)
    rev.Reject();
}
```

## 成员属性

### Revision.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Revision 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Revision.Author

返回或设置完成指定修订的用户姓名

**语法**

**express.Author**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
//显示所选部分中的第一节的第一处修订的作者姓名
function test()
{
    let rngSection = Application.Selection.Sections.Item(1).Range
    alert("Revisions made by " + rngSection.Revisions.Item(1).Author)
}
```

### Revision.Cells

返回一个 **Cells** 集合，该集合代表已经用修订标记进行标记的表格单元格。只读。

**语法**

**express.Cells**

\*express是一个代表 Revision 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Revision.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Revision 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Revision.Date

修订的日期和时间

**语法**

**express.Date**

\*express是一个代表 Revision 对象的变量。

### Revision.FormatDescription

返回一个 **String** 类型的值，该值代表对修订中所作格式更改的说明。只读。

**语法**

**express.FormatDescription**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
/*本示例显示文档中修订的所有格式更改的说明。*/
function FmtChanges() {
    for(let revFmtRev = 1; revFmtRev <= ActiveDocument.Revisions.Count; revFmtRev++) {
        if(!ActiveDocument.Revisions.Item(revFmtRev).FormatDescription) {
            MsgBox("Format changes made : " + ActiveDocument.Revisions.Item(revFmtRev).FormatDescription)
        }
    }
}
```

### Revision.Index

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Revision 对象的变量。

### Revision.MovedRange

返回一个 **Range** 对象，该对象代表在带修订的文档中从一个位置移至另一个位置的文本区域。只读。

**语法**

**express.MovedRange**

\*express是一个代表 Revision 对象的变量。

### Revision.Parent

返回一个 **Object** 类型值，该值代表指定 **Revision** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Revision 对象的变量。

### Revision.Range

返回一个 **Range** 对象，该对象代表修订标记内包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 Revision 对象的变量。

### Revision.Style

返回一个 **Style** 对象，该对象代表与一个修订标记相关联的样式。

**语法**

**express.Style**

\*express是一个代表 Revision 对象的变量。

### Revision.Type

返回修订的类型。只读 **[WdRevisionType]()** 类型。

**语法**

**express.Type**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
//获取当前激活文档的第一个修订的类型
Application.ActiveDocument.Revisions.Item(1).Type
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Revision](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
