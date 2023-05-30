# Column 对象

## 简介

代表单个表格列。 Column 对象是 Columns 集合的成员。 Columns 集合包括一个表格、选定内容或区域中的所有列。

## 说明

使用 Columns ( *Index* ) 返回单个 Column 对象（其中 *Index* 是索引编号）。索引编号代表该列在 Columns 集合中的位置（从左至右计数）。

以下示例选定活动文档中的表 1 的第一列。

``` JavaScript
Application.ActiveDocument.Tables.Item(1).Columns.Item(1).Select()
```

可将 Column 属性与 Cell 对象配合使用，返回 Column 对象。以下示例删除单元格 1 中的文本，然后插入新文本并排序整个列。

``` JavaScript
function test() {
    let rng = Application.ActiveDocument.Tables.Item(1).Cell(1, 1)
    rng.Range.Delete()
    rng.Range.InsertBefore("Sales")
    rng.Column.Sort()
}
```

使用 Add 方法在表格中添加一列。以下示例在活动文档的第一个表格中添加一列，然后使各列宽度相等。

``` JavaScript
function test() {
    if (Application.ActiveDocument.Tables.Count >= 1) {
        let myTable = Application.ActiveDocument.Tables.Item(1)
        myTable.Columns.Add(myTable.Columns.Item(1))
        myTable.Columns.DistributeWidth
    }
}
```

说明可将 Information 属性与 Selection 对象配合使用，返回当前列编号。以下示例选定当前列并在消息框中显示其列编号。

``` JavaScript
function test() {
    if(Application.Selection.Information(wdWithInTable) == true){
        Application.Selection.Columns.Item(1).Select()
        alert("Column " + Application.Selection.Information(wdStartOfRangeColumnNumber))
    }
}
```

## 方法

| 序号 | 名称                         | 说明                                                               |
|:----:|:-----------------------------|:-------------------------------------------------------------------|
|  1   | [AutoFit](#Column.AutoFit)   | 改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。 |
|  2   | [Delete](#Column.Delete)     | 删除指定列。                                                       |
|  3   | [Select](#Column.Select)     | 选择指定表格列。                                                   |
|  4   | [SetWidth](#Column.SetWidth) | 设置表格列的宽度。                                                 |
|  5   | [Sort](#Column.Sort)         | 对指定表格列进行排序。                                             |

## 属性

| 序号 | 名称                                             | 说明                                                                                                     |
|:----:|:-------------------------------------------------|:---------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Column.Application)               | 返回一个代表 WPS 应用程序的 Application 对象。                                                           |
|  2   | [Borders](#Column.Borders)                       | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                    |
|  3   | [Cells](#Column.Cells)                           | 返回一个 Cells 集合，该集合代表表格列中的表格单元格。只读。                                              |
|  4   | [Creator](#Column.Creator)                       | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                             |
|  5   | [Index](#Column.Index)                           | 返回一个 Number 类型的值，该值代表项目在集合中的位置。只读。                                             |
|  6   | [IsFirst](#Column.IsFirst)                       | 如果指定的列是表格的第一列，则返回 True 。只读 Boolean 类型。                                            |
|  7   | [IsLast](#Column.IsLast)                         | 如果指定的列是表格的最后一列，则返回 True 。只读 Boolean 类型。                                          |
|  8   | [NestingLevel](#Column.NestingLevel)             | 返回指定列的嵌套层。只读 Long 类型。                                                                     |
|  9   | [Next](#Column.Next)                             | 返回表格列集合中的下一列。只读。                                                                         |
|  10  | [Parent](#Column.Parent)                         | 返回一个 Object 类型值，该值代表指定 Column 对象的父对象。                                               |
|  11  | [PreferredWidth](#Column.PreferredWidth)         | 返回或设置指定的单元格、列或表格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 Single 类型，可读写。 |
|  12  | [PreferredWidthType](#Column.PreferredWidthType) | 返回或设置用于指定表格列的宽度的首选度量单位。可读/写 WdPreferredWidthType 类型。                        |
|  13  | [Previous](#Column.Previous)                     | 返回表格列集合中的前一列。只读。                                                                         |
|  14  | [Shading](#Column.Shading)                       | 返回一个引用指定列的底纹格式的 Shading 对象。                                                            |
|  15  | [Width](#Column.Width)                           | 返回或设置指定列的宽度，以磅为单位。可读/写 Long 类型。                                                  |

## 成员方法

### Column.AutoFit

改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。

**语法**

**express.AutoFit ()**

\*express是一个代表 Column 对象的变量。

**说明**

如果表格的宽度已等于从左边界到右边界的距离，则此方法无效。

**示例**

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，然后调整第一列的宽度，使之与文本的宽度相称。*/
function test() {
    let docNew = Application.Documents.Add()
    let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
    tableNew.Cell(1, 1).Range.InsertAfter("First cell")
    tableNew.Columns.Item(1).AutoFit()
}
```

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，然后调整所有列的宽度，使之与文本的宽度相称。*/
function test() {
    let docNew = Application.Documents.Add()
    let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
    tableNew.Cell(1, 1).Range.InsertAfter("First cell")
    tableNew.Cell(1, 2).Range.InsertAfter("This is cell (1,2)")
    tableNew.Cell(1, 3).Range.InsertAfter("(1,3)")
    tableNew.Columns.AutoFit()
}
```

### Column.Delete

删除指定列。

**语法**

**express.Delete ()**

\*express是一个代表 Column 对象的变量。

### Column.Select

选择指定表格列。

**语法**

**express.Select ()**

\*express是一个代表 Column 对象的变量。

### Column.SetWidth

设置表格列的宽度。

**语法**

**express.SetWidth ( *ColumnWidth* , *RulerStyle* )**

\*express是一个代表 Column 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型     | 说明                            |
|---------------|-----------|--------------|---------------------------------|
| *ColumnWidth* | 必选      | Single       | 指定列的宽度，以磅为单位。      |
| *RulerStyle*  | 必选      | WdRulerStyle | 控制 WPS 调整单元格宽度的方式。 |

**说明**

上述 **WdRulerStyle** 行为应用于左对齐的表格。 **WdRulerStyle** 行为应用于中对齐和右对齐的表格时可能会出现未知效果，因此应谨慎使用 **SetWidth** 方法。

### Column.Sort

对指定表格列进行排序。

**语法**

**express.Sort ( *ExcludeHeader* , *SortFieldType* , *SortOrder* , *CaseSensitive* , *BidiSort* , *IgnoreThe* , *IgnoreKashida* , *IgnoreDiacritics* , *IgnoreHe* , *LanguageID* )**

\*express是一个代表 Column 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                                                                            |
|--------------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------|
| *ExcludeHeader*    | 可选      | Variant  | 如果该属性值为 True，则不对首行或首段标题进行排序。默认值为 False。                                                                             |
| *SortFieldType*    | 可选      | Variant  | 列的排序类型。可以是 WdSortFieldType 常量之一。                                                                                                 |
| *SortOrder*        | 可选      | Variant  | 列的排序顺序。可以是 WdSortOrder 常量之一。                                                                                                     |
| *CaseSensitive*    | 可选      | Variant  | 如果该属性值为 True，则排序时区分大小写。默认值为 False。                                                                                       |
| *BidiSort*         | 可选      | Variant  | 如果该属性值为 True，则基于从右向左排列的语言规则进行排序。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。                       |
| *IgnoreThe*        | 可选      | Variant  | 如果该属性值为 True，则在从右向左排列的语言的文本排序中忽略阿拉伯字符“alef lam”。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *IgnoreKashida*    | 可选      | Variant  | 如果该属性值为 True，则在从右向左排列的语言的文本排序中忽略“kashidas”。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。           |
| *IgnoreDiacritics* | 可选      | Variant  | 如果该属性值为 True，则在从右向左排列的语言的文本排序中忽略双向控制字符。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。         |
| *IgnoreHe*         | 可选      | Variant  | 如果该属性值为 True，则在从右向左排列的语言的文本排序中忽略希伯来字符“he”。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。       |
| *LanguageID*       | 可选      | Variant  | 指定排序语言。可以是 WdLanguageID 常量之一。                                                                                                    |

**说明**

如果要对表格单元格中的段落进行排序，则只能包括段落标记，不能包括单元格结束标记；如果在选定内容或区域中包括了结束单元格标记，然后试图对段落进行排序， WPS 将显示一条消息，说明未找到进行排序的有效记录。

## 成员属性

### Column.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Column 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Column.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Column 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Column.Cells

返回一个 **Cells** 集合，该集合代表表格列中的表格单元格。只读。

**语法**

**express.Cells**

\*express是一个代表 Column 对象的变量。

### Column.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Column 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Column.Index

返回一个 **Number** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Column 对象的变量。

### Column.IsFirst

如果指定的列是表格的第一列，则返回 **True** 。只读 **Boolean** 类型。

**语法**

**express.IsFirst**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/* 本示例判断选定部分的第一列是否为表格中的第一列 */
alert(Application.Selection.Columns.Item(1).IsFirst);
```

### Column.IsLast

如果指定的列是表格的最后一列，则返回 **True** 。只读 **Boolean** 类型。

**语法**

**express.IsLast**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/*  本示例判断选定部分的第一列是否为表格的最后一列 */
alert(Application.Selection.Columns.Item(1).IsLast);
```

### Column.NestingLevel

返回指定列的嵌套层。只读 **Long** 类型。

**语法**

**express.NestingLevel**

\*express是一个代表 Column 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

**示例**

``` JavaScript
/*本示例新建一个文档，创建一个三层嵌套表格，并在每个表格的第一个单元格中填入该表格所在的嵌套层数。*/
function test() {
    Application.Documents.Add()
    Application.ActiveDocument.Tables.Add(Selection.Range, 3, 3, wdWord9TableBehavior, wdAutoFitContent)
    let rng = Application.ActiveDocument.Tables.Item(1).Range
    rng.Copy()
    rng.Cells.Item(1).Range.Text = rng.Cells.Item(1).NestingLevel
    rng.Cells.Item(5).Range.PasteAsNestedTable()

    let rng_rng = rng.Cells.Item(5).Tables.Item(1).Range
    rng_rng.Cells.Item(1).Range.Text = rng_rng.Cells.Item(1).NestingLevel
    rng_rng.Cells.Item(5).Range.PasteAsNestedTable()

    let rng_rng_rng = rng_rng.Cells.Item(5).Tables.Item(1).Range
    rng_rng_rng.Cells.Item(1).Range.Text = rng_rng_rng.Cells.Item(1).NestingLevel
}
```

### Column.Next

返回表格列集合中的下一列。只读。

**语法**

**express.Next**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/*如果选定内容位于表格之中，则本示例选定下一表格列的内容。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Columns.Item(1).Next.Select()
    }
}
```

### Column.Parent

返回一个 **Object** 类型值，该值代表指定 **Column** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Column 对象的变量。

### Column.PreferredWidth

返回或设置指定的单元格、列或表格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 **Single** 类型，可读写。

**语法**

**express.PreferredWidth**

\*express是一个代表 Column 对象的变量。

**说明**

如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPoints** ，则 **PreferredWidth** 属性将返回或设置以磅为单位的宽度值；如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent** ，则 **PreferredWidth** 属性将返回或设置用窗口宽度百分比表示的宽度值。

### Column.PreferredWidthType

返回或设置用于指定表格列的宽度的首选度量单位。可读/写 **WdPreferredWidthType** 类型。

**语法**

**express.PreferredWidthType**

\*express是一个代表 Column 对象的变量。

### Column.Previous

返回表格列集合中的前一列。只读。

**语法**

**express.Previous**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/*如果选定内容位于表格中，则本示例选定前一表格列的内容。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Columns.Item(1).Previous.Select()
    }
}
```

### Column.Shading

返回一个引用指定列的底纹格式的 **Shading** 对象。

**语法**

**express.Shading**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/*本示例对活动文档中第一张表格的第一列应用水平线纹理。*/
function test() {
    if (Application.ActiveDocument.Tables.Count >= 1) {
        let fpx = Application.ActiveDocument.Tables.Item(1).Columns.Item(1).Shading
        fpx.Texture = wdTextureHorizontal
    }
}
```

### Column.Width

返回或设置指定列的宽度，以磅为单位。可读/写 **Long** 类型。

**语法**

**express.Width**

\*express是一个代表 Column 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Column](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
