# AxisTitle 对象

## 简介

代表图表坐标轴标题。

## 说明

使用 AxisTitle 属性可返回 AxisTitle 对象。 只有当坐标轴的 HasTitle 属性为 True 时， AxisTitle 对象才存在，从而才能使用该对象。

示例

``` JavaScript
/*下例激活第一个嵌入式图表，设置其数值轴标题文本，将其字体设为 10 磅的“Bookman”，并将单词“millions”设为倾斜。*/
function test()
{
    Application.Worksheets.Item("sheet1").ChartObjects(1).Activate()
    let axes = Application.ActiveChart.Axes(xlValue)
    axes.HasTitle = true

    let axistitle = axes.AxisTitle
    axistitle.Caption = "Revenue (millions)"
    axistitle.Font.Name = "bookman"
    axistitle.Font.Size = 10
    axistitle.Characters(10, 8).Font.Italic = true
}
```

## 方法

| 序号 | 名称                        | 说明       |
|:----:|:----------------------------|:-----------|
|  1   | [Delete](#AxisTitle.Delete) | 删除对象。 |
|  2   | [Select](#AxisTitle.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                                  | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#AxisTitle.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Caption](#AxisTitle.Caption)                         | 返回或设置一个 String 值，它代表坐标轴标题文本。                                                                                                                                                                                |
|  3   | [Characters](#AxisTitle.Characters)                   | 返回 Characters 对象，它代表对象文本内某个区域的字符。使用 Characters 对象可为文本字符串内的字符设置格式。                                                                                                                      |
|  4   | [Creator](#AxisTitle.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Format](#AxisTitle.Format)                           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  6   | [Formula](#AxisTitle.Formula)                         | 获取或设置一个 String 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。                                                                                                                                               |
|  7   | [FormulaLocal](#AxisTitle.FormulaLocal)               | 获取或设置一个 String 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。                                                                                                                                         |
|  8   | [FormulaR1C1](#AxisTitle.FormulaR1C1)                 | 获取或设置一个 String 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。                                                                                                                                             |
|  9   | [FormulaR1C1Local](#AxisTitle.FormulaR1C1Local)       | 获取或设置一个 String 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。                                                                                                                                       |
|  10  | [Height](#AxisTitle.Height)                           | 返回对象的高度（以磅为单位）。只读。                                                                                                                                                                                            |
|  11  | [HorizontalAlignment](#AxisTitle.HorizontalAlignment) | 返回或设置一个 Variant 值，它代表指定对象的水平对齐方式。                                                                                                                                                                       |
|  12  | [IncludeInLayout](#AxisTitle.IncludeInLayout)         | 如果在确定图表布局时轴标题将占用图表布局空间，则为 True 。默认值是 True 。可读/写 Boolean 类型。                                                                                                                                |
|  13  | [Left](#AxisTitle.Left)                               | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  14  | [Name](#AxisTitle.Name)                               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  15  | [Orientation](#AxisTitle.Orientation)                 | 返回或设置一个 Variant 值，它代表文本方向。                                                                                                                                                                                     |
|  16  | [Parent](#AxisTitle.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  17  | [Position](#AxisTitle.Position)                       | 返回或设置图表上轴标题的位置。可读/写 XlChartElementPosition 类型。                                                                                                                                                             |
|  18  | [ReadingOrder](#AxisTitle.ReadingOrder)               | 返回或设置指定对象的阅读次序。可为以下常量之一： xlRTL （从右到左）、 xlLTR （从左到右）或 xlContext 。 Long 类型，可读写。                                                                                                     |
|  19  | [Shadow](#AxisTitle.Shadow)                           | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  20  | [Text](#AxisTitle.Text)                               | 返回或设置指定对象中的文本。 String 型，可读写。                                                                                                                                                                                |
|  21  | [Top](#AxisTitle.Top)                                 | 返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                      |
|  22  | [VerticalAlignment](#AxisTitle.VerticalAlignment)     | 返回或设置一个 Variant 值，它代表指定对象的垂直对齐方式。                                                                                                                                                                       |
|  23  | [Width](#AxisTitle.Width)                             | 返回对象的宽度（以磅为单位）。只读。                                                                                                                                                                                            |

## 成员方法

### AxisTitle.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

返回值 一个代表所选对象的 Variant 值。

## 成员属性

### AxisTitle.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 AxisTitle 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
  if(Application.ActiveWorkbook.Application.Value == "ET" ){
      alert("This is an ET Application object.")
  }
  else{
      alert("This is not an ET Application object.")
  }
}
```

### AxisTitle.Caption

返回或设置一个 **String** 值，它代表坐标轴标题文本。

**语法**

**express.Caption**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Characters

返回 **Characters** 对象，它代表对象文本内某个区域的字符。使用 **Characters** 对象可为文本字符串内的字符设置格式。

**语法**

**express.Characters**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

参数

| **名称** | **必选/可选** | **数据类型** | **说明**                                                                                      |
|----------|---------------|--------------|-----------------------------------------------------------------------------------------------|
| *Start*  | 可选          | **Variant**  | 要返回的第一个字符。如果此参数是 1 或被省略，则此属性返回一个以第一个字符为开头的字符区域。   |
| *Length* | 可选          | **Variant**  | 要返回的字符数。如果省略此参数，则此属性返回字符串的后半部分（ *Start* 字符之后的所有字符）。 |

**Characters** 对象不是集合。

### AxisTitle.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### AxisTitle.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Formula

获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.Formula**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

有关 A1 样式表示法的详细信息，请参阅 <span style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;background-color:#FFFFFF;">操作方法：使用 A1 表示法引用单元格和区域</span> 。

### AxisTitle.FormulaLocal

获取或设置一个 **String** 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaLocal**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

有关 A1 样式表示法的详细信息，请参阅 <span style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;background-color:#FFFFFF;">操作方法：使用 A1 表示法引用单元格和区域</span> 。

### AxisTitle.FormulaR1C1

获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.FormulaR1C1Local

获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1Local**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Height

返回对象的高度（以磅为单位）。只读。

**语法**

**express.Height**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.HorizontalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlLeft**        |
| **xlRight**       |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### AxisTitle.IncludeInLayout

如果在确定图表布局时轴标题将占用图表布局空间，则为 **True** 。默认值是 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.IncludeInLayout**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Orientation

返回或设置一个 **Variant** 值，它代表文本方向。

**语法**

**express.Orientation**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Position

返回或设置图表上轴标题的位置。可读/写 **XlChartElementPosition** 类型。

**语法**

**express.Position**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.ReadingOrder

返回或设置指定对象的阅读次序。可为以下常量之一： **xlRTL** （从右到左）、 **xlLTR** （从左到右）或 **xlContext** 。 **Long** 类型，可读写。

**语法**

**express.ReadingOrder**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Text

返回或设置指定对象中的文本。 **String** 型，可读写。

**语法**

**express.Text**

\*express是一个代表 AxisTitle 对象的变量。

**示例**

``` JavaScript
/*此示例设置 Chart1 中的分类轴标题文本。*/
function test()
{
  Application.Charts.Item("Chart1").Axes(xlCategory).HasTitle = true
  Application.Charts.Item("Chart1").Axes(xlCategory).AxisTitle.Text = "Month"
}
```

### AxisTitle.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.VerticalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlBottom**      |
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlTop**         |

### AxisTitle.Width

返回对象的宽度（以磅为单位）。只读。

**语法**

**express.Width**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

返回值 Double

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AxisTitle](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
