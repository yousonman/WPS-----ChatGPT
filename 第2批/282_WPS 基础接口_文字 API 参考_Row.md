# Row 对象

## 简介

代表表格中的一行。 Row 对象是 Rows 集合的一个成员

## 说明

代表表格中的一行。 Row 对象是 Rows 集合的一个成员。 Rows 集合包含指定的所选内容、范围或表格中的所有行。

## 方法

| 序号 | 名称                                | 说明                 |
|:----:|:------------------------------------|:---------------------|
|  1   | [ConvertToText](#Row.ConvertToText) |                      |
|  2   | [Delete](#Row.Delete)               | 删除指定的表格行     |
|  3   | [Select](#Row.Select)               | 选择指定的表格行。   |
|  4   | [SetHeight](#Row.SetHeight)         | 设置表格行的高度。   |
|  5   | [SetLeftIndent](#Row.SetLeftIndent) | 设置表格中行的缩进。 |

## 属性

| 序号 | 名称                                              | 说明                                                                                                                                                                 |
|:----:|:--------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Alignment](#Row.Alignment)                       | 返回或设置一个 WdRowAlignment 常量，该常量代表指定行的对齐方式。可读写。                                                                                             |
|  2   | [AllowBreakAcrossPage](#Row.AllowBreakAcrossPage) | 如果允许分页符拆分表格中一行或多行中的文本，则该属性值为 True 。可读写 Long 类型。                                                                                   |
|  3   | [Application](#Row.Application)                   | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                       |
|  4   | [Borders](#Row.Borders)                           | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                                                                                |
|  5   | [Cells](#Row.Cells)                               | 返回一个 Cells 集合，该集合代表在某一列、行、选定内容或区域中的表格单元格。只读。                                                                                    |
|  6   | [Creator](#Row.Creator)                           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                         |
|  7   | [HeadingFormat](#Row.HeadingFormat)               | 如果为 True ，则将指定一行或数行的格式设置为表格标题。当表格不止一页时，可将多行的格式重复设置为表格标题。可以是 True 、 False 或 wdUndefined 。 Long 类型，可读写。 |
|  8   | [Height](#Row.Height)                             | 返回或设置表格中指定行的高度。Single 类型，可读写。                                                                                                                  |
|  9   | [HeightRule](#Row.HeightRule)                     | 返回或设置确定指定单元格或行高度的规则。 WdRowHeightRule 类型，可读写。                                                                                              |
|  10  | [ID](#Row.ID)                                     | 在文档另存为网页时返回或设置指定表格行的标识标签。可读写 String 类型。                                                                                               |
|  11  | [Index](#Row.Index)                               | 返回 Long 值，该值表示项目在集合中的位置。只读。                                                                                                                     |
|  12  | [IsFirst](#Row.IsFirst)                           | 如果指定的行是表格中的首行，则该属性值为 True 。否则为False                                                                                                          |
|  13  | [IsLast](#Row.IsLast)                             | 如果指定的行是表格中的首行，则该属性值为 True 。否则为False                                                                                                          |
|  14  | [LeftIndent](#Row.LeftIndent)                     | 返回或设置一个 Single 类型的值，该值代表指定表格行的左缩进值（以磅为单位）。可读写。                                                                                 |
|  15  | [NestingLevel](#Row.NestingLevel)                 | 返回指定表格行的嵌套层。只读 Long 类型。                                                                                                                             |
|  16  | [Next](#Row.Next)                                 | 返回一个 Row 对象，该对象代表表格中行的集合中的下一个表格行。只读。                                                                                                  |
|  17  | [Parent](#Row.Parent)                             | 返回一个 Object 类型值，该值代表指定 Row 对象的父对象。                                                                                                              |
|  18  | [Previous](#Row.Previous)                         | 返回一个 Row 对象，该对象代表指定行的前一个表格行。只读。                                                                                                            |
|  19  | [Range](#Row.Range)                               | 返回一个 Range 对象，该对象代表指定表格行内包含的文档部分                                                                                                            |
|  20  | [Shading](#Row.Shading)                           | 返回一个 Shading 对象，该对象代表指定对象的底纹格式。                                                                                                                |
|  21  | [SpaceBetweenColumns](#Row.SpaceBetweenColumns)   | 返回或设置指定的一行或多行中相邻列的文本之间的距离（以磅为单位）。可读写 Single 类型。                                                                               |

## 成员方法

### Row.ConvertToText

**语法**

**express.ConvertToText ( *Separator* , *NestedTables* )**

\*express是一个代表 Row 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                      |
|----------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Separator*    | 可选      | Variant  | 用以分隔被转换列的字符（被转换行由段落标记分隔）。可以是下列任何 WdTableFieldSeparator 常量：wdSeparateByCommas、wdSeparateByDefaultListSeparator、wdSeparateByParagraphs 或 wdSeparateByTabs（默认值）。 |
| *NestedTables* | 可选      | Variant  | 如果要将嵌套的表格转换为文本，则为 True。如果 Separator 不是 wdSeparateByParagraphs，则将忽略此参数。默认值为 True。                                                                                      |

**说明**

将表格转换为文本并返回一个 **Range** 对象，该对象代表带分隔符的文本。

### Row.Delete

删除指定的表格行

**语法**

**express.Delete ()**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//删除选中区域的表格的第一行
Application.Selection.Rows.Item(1).Delete()
```

### Row.Select

选择指定的表格行。

**语法**

**express.Select ()**

\*express是一个代表 Row 对象的变量。

**说明**

使用此方法后，请使用 **Selection** 对象来处理所选行。有关详细信息，请参阅 处理 Selection 对象。

**示例**

``` JavaScript
/*以下示例选择 Report.doc 的第一张表格中的第一行。*/
Documents.Item("Report.doc").Tables.Item(1).Rows.Item(1).Select()
```

### Row.SetHeight

设置表格行的高度。

**语法**

**express.SetHeight ( *RowHeight* , *HeightRule* )**

\*express是一个代表 Row 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型        | 说明                         |
|--------------|-----------|-----------------|------------------------------|
| *RowHeight*  | 必选      | Single          | 行的高度，以磅为单位。       |
| *HeightRule* | 必选      | WdRowHeightRule | 用于确定指定行的高度的规则。 |

**示例**

``` JavaScript
/*以下示例创建一张表格，然后将首行的固定行高设置为 0.5 英寸（36 磅）。*/
function test()
{
let newDoc = Documents.Add()
let aTable = newDoc.Tables.Add(Selection.Range, 3, 3)
aTable.Rows.Item(1).SetHeight(InchesToPoints(0.5), wdRowHeightExactly)
}
```

### Row.SetLeftIndent

设置表格中行的缩进。

**语法**

**express.SetLeftIndent ( *LeftIndent* , *RulerStyle* )**

\*express是一个代表 Row 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型     | 说明                                                               |
|--------------|-----------|--------------|--------------------------------------------------------------------|
| *LeftIndent* | 必选      | Single       | 指定的一行或多行的当前左边缘与所需左边缘之间的距离（以磅为单位）。 |
| *RulerStyle* | 必选      | WdRulerStyle | 改变左缩进时，用来控制 WPS 调整表格的方式。                        |

**说明**

上述 **WdRulerStyle** 行为适用于左对齐的表格。居中和右对齐的表格的 **WdRulerStyle** 行为不可预测；在这些情况下，请谨慎使用 **SetLeftIndent** 方法。

**示例**

``` JavaScript
/*以下示例在新文档中创建一张表格，第一行缩进 0.5 英寸（36 磅）。改变左缩进时，单元格宽度将调整以保持表格的右边缘。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)

tableNew.Rows.Item(1).SetLeftIndent(InchesToPoints(0.5),
    wdAdjustSameWidth)
}
```

## 成员属性

### Row.Alignment

返回或设置一个 **WdRowAlignment** 常量，该常量代表指定行的对齐方式。可读写。

**语法**

**express.Alignment**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*本示例使活动文档的第一个表格中第一行的所有单元格居中。*/
function CenterRows() {
    ActiveDocument.Tables.Item(1).Rows.Item(1).Alignment = wdAlignRowCenter
}
```

### Row.AllowBreakAcrossPage

如果允许分页符拆分表格中一行或多行中的文本，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.AllowBreakAcrossPage**

\*express是一个代表 Row 对象的变量。

**说明**

该属性可以是 **True** 、 **False** 或 **wdUndefined** （仅允许拆分部分指定文本）。

**示例**

``` JavaScript
/*以下示例新建一篇包含 5x5 表格的文档，并防止在分页时拆分表格的第三行。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 5, 5)
tableNew.Rows.Item(3).AllowBreakAcrossPages = false
}
```

``` JavaScript
/*本示例确定分页时是否可以拆分当前表格中的行。如果插入点不在表格中，则显示一个消息框。*/
function test()
{
Selection.Collapse(wdCollapseStart)
if(!Selection.Tables.Count) {
    MsgBox("The insertion point is not in a table.")
}
else {
    let lngAllowBreak = Selection.Rows.AllowBreakAcrossPages
}
}
```

### Row.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Row 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Row.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Row 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Row.Cells

返回一个 **Cells** 集合，该集合代表在某一列、行、选定内容或区域中的表格单元格。只读。

**语法**

**express.Cells**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//返回对应的选区的第一行的单元格数
function test()
{
  let tcels = Application.Selection.Rows.Item(1).Cells
  tcels.Count
}
```

### Row.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Row 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Row.HeadingFormat

如果为 **True** ，则将指定一行或数行的格式设置为表格标题。当表格不止一页时，可将多行的格式重复设置为表格标题。可以是 **True** 、 **False** 或 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.HeadingFormat**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档的开始处创建一个 5x5 表格，然后将表格第一行的格式设置为表格标题。*/
function test()
{
let rngTemp = Selection.Range
let tableNew = ActiveDocument.Tables.Add(rngTemp, 5, 5)

tableNew.Rows.Item(1).HeadingFormat = true
}
```

``` JavaScript
/*本示例判定是否将插入点所在行的格式设置为表格标题。*/
function test()
{
if(Selection.Information(wdWithInTable)) {
    if(Selection.Rows.Item(1).HeadingFormat) { 
        MsgBox("The current row is a table heading")
    }
}
else {
    MsgBox("The insertion point is not in a table.")
}
}
```

### Row.Height

返回或设置表格中指定行的高度。Single 类型，可读写。

**语法**

**express.Height**

\*express是一个代表 Row 对象的变量。

**说明**

如果指定行的 **HeightRule** 属性为 **wdRowHeightAuto** ，则 **Height** 返回 **wdUndefined** ；设置 **Height** 属性会将 **HeightRule** 设置为 **wdRowHeightAtLeast** 。

**示例**

``` JavaScript
/*本示例将活动文档第一张表格的行高度设置为至少 20 磅。*/
ActiveDocument.Tables.Item(1).Rows.Height = 20
```

### Row.HeightRule

返回或设置确定指定单元格或行高度的规则。 **WdRowHeightRule** 类型，可读写。

**语法**

**express.HeightRule**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*本示例在新文档中创建一张 3x3 表格，并设置第二行的最小行高为 24 磅。*/
function test()
{
let newDoc = Documents.Add()
let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
    myTable.Rows.Item(2).Height = 24
    myTable.Rows.Item(2).HeightRule = wdRowHeightAtLeast
}
```

### Row.ID

在文档另存为网页时返回或设置指定表格行的标识标签。可读写 **String** 类型。

**语法**

**express.ID**

\*express是一个代表 Row 对象的变量。

**说明**

可以将标签作为引用其他网页或当前文档中内容的超链接。

### Row.Index

返回 **Long** 值，该值表示项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//获取当前激活文档的第一个表格的第二行的索引号
Application.ActiveDocument.Tables.Item(1).Rows.Item(2).Index

//获取选中区域的第二行所在的索引号
Application.Selection.Rows.Item(2).Index
```

### Row.IsFirst

如果指定的行是表格中的首行，则该属性值为 **True** 。否则为False

**语法**

**express.IsFirst**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//定所选内容的首行是否为表格中的首行。
Application.Selection.Rows.Item(1).IsFirst
```

### Row.IsLast

如果指定的行是表格中的首行，则该属性值为 True 。否则为False

**语法**

**express.IsLast**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
Application.Selection.Rows.Item(1).IsLast
```

### Row.LeftIndent

返回或设置一个 **Single** 类型的值，该值代表指定表格行的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*以下示例设置活动文档中第一个表格的第一行的左缩进。*/
ActiveDocument.Tables.Item(1).Rows.Item(1).LeftIndent = InchesToPoints(1)
```

### Row.NestingLevel

返回指定表格行的嵌套层。只读 **Long** 类型。

**语法**

**express.NestingLevel**

\*express是一个代表 Row 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

### Row.Next

返回一个 **Row** 对象，该对象代表表格中行的集合中的下一个表格行。只读。

**语法**

**express.Next**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*如果所选内容位于表格中，则以下示例选择下一个表格行的内容。*/
function test()
{
if(Selection.Information(wdWithInTable)) {
    Selection.Rows.Item(1).Next.Select()
}
}
```

### Row.Parent

返回一个 **Object** 类型值，该值代表指定 **Row** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Row 对象的变量。

### Row.Previous

返回一个 **Row** 对象，该对象代表指定行的前一个表格行。只读。

**语法**

**express.Previous**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*如果所选内容位于表格中，则以下示例选择前一行的内容。*/
function test()
{
if(Selection.Information(wdWithInTable)) {
    Selection.Rows.Item(1).Previous.Select()
}
}
```

### Row.Range

返回一个 **Range** 对象，该对象代表指定表格行内包含的文档部分

**语法**

**express.Range**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//获取表格 1 中的第一行内容并打印

function test()
{
    let rgText
    if(Application.ActiveDocument.Tables.Count >= 1)
        rgText = Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Range.Text
        alert(rgText)
}
```

### Row.Shading

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*以下示例将横线纹理应用于表格 1 的第一行。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1) {
    let myShading = ActiveDocument.Tables.Item(1).Rows.Item(1).Shading
        myShading.Texture = wdTextureHorizontal
}
}
```

### Row.SpaceBetweenColumns

返回或设置指定的一行或多行中相邻列的文本之间的距离（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.SpaceBetweenColumns**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，然后将第一行中的列间距设置为 0.5 英寸。*/
function test()
{
let newDoc = Documents.Add()
let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
myTable.Rows.Item(1).SpaceBetweenColumns = InchesToPoints(0.5)
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Row](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
