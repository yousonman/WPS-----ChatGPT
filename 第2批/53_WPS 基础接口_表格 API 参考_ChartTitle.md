# ChartTitle 对象

## 简介

代表图表标题。

## 说明

使用 ChartTitle 属性可返回 ChartTitle 对象。

只有图表的 HasTitle 属性为 True 时， ChartTitle 对象才存在，从而才能使用该对象。

``` JavaScript
/*下例向工作表 Sheet1 上嵌入的第一个图表添加标题。*/
function test() {
    let myChart = application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
        myChart.HasTitle = true
        myChart.ChartTitle.Text = "February Sales"
}
```

## 方法

| 序号 | 名称                         | 说明       |
|:----:|:-----------------------------|:-----------|
|  1   | [Delete](#ChartTitle.Delete) | 删除对象。 |
|  2   | [Select](#ChartTitle.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartTitle.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Caption](#ChartTitle.Caption)                         | 返回或设置一个 String 值，它代表图表标题文本的方向。                                                                                                                                                                            |
|  3   | [Characters](#ChartTitle.Characters)                   | 返回 Characters 对象，它代表对象文本内某个区域的字符。使用 Characters 对象可为文本字符串内的字符设置格式。                                                                                                                      |
|  4   | [Creator](#ChartTitle.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Format](#ChartTitle.Format)                           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  6   | [Formula](#ChartTitle.Formula)                         | 获取或设置一个 String 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。                                                                                                                                               |
|  7   | [FormulaLocal](#ChartTitle.FormulaLocal)               | 获取或设置一个 String 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。                                                                                                                                         |
|  8   | [FormulaR1C1](#ChartTitle.FormulaR1C1)                 | 获取或设置一个 String 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。                                                                                                                                             |
|  9   | [FormulaR1C1Local](#ChartTitle.FormulaR1C1Local)       | 获取或设置一个 String 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。                                                                                                                                       |
|  10  | [Height](#ChartTitle.Height)                           | 返回对象的高度（以磅为单位）。只读。                                                                                                                                                                                            |
|  11  | [HorizontalAlignment](#ChartTitle.HorizontalAlignment) | 返回或设置一个 Variant 值，它代表指定对象的水平对齐方式。                                                                                                                                                                       |
|  12  | [IncludeInLayout](#ChartTitle.IncludeInLayout)         | 如果在确定图表布局时图表标题将占用图表布局空间，则为 True 。默认值是 True 。 Boolean 类型，可读写。                                                                                                                             |
|  13  | [Left](#ChartTitle.Left)                               | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  14  | [Name](#ChartTitle.Name)                               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  15  | [Orientation](#ChartTitle.Orientation)                 | 返回或设置一个 Variant 值，它代表文本方向。                                                                                                                                                                                     |
|  16  | [Parent](#ChartTitle.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  17  | [Position](#ChartTitle.Position)                       | 返回或设置图表上图表标题的位置。可读/写 XlChartElementPosition 类型。                                                                                                                                                           |
|  18  | [ReadingOrder](#ChartTitle.ReadingOrder)               | 返回或设置指定对象的阅读次序。可为以下常量之一： xlRTL （从右到左）、 xlLTR （从左到右）或 xlContext 。 Long 类型，可读写。                                                                                                     |
|  19  | [Shadow](#ChartTitle.Shadow)                           | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  20  | [Text](#ChartTitle.Text)                               | 返回或设置指定对象中的文本。 String 型，可读写。                                                                                                                                                                                |
|  21  | [Top](#ChartTitle.Top)                                 | 返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                      |
|  22  | [VerticalAlignment](#ChartTitle.VerticalAlignment)     | 返回或设置一个 Variant 值，它代表指定对象的垂直对齐方式。                                                                                                                                                                       |
|  23  | [Width](#ChartTitle.Width)                             | 返回对象的宽度（以磅为单位）。只读。                                                                                                                                                                                            |

## 成员方法

### ChartTitle.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 ChartTitle 对象的变量。

**返回值**

Variant

### ChartTitle.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 ChartTitle 对象的变量。

**返回值**

Variant

## 成员属性

### ChartTitle.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ChartTitle 对象的变量。

**说明**

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/

function test() {
    let myObject = ActiveWorkbook
        if (myObject.Application.Value == "ET") {
            MsgBox("This is an ET Application object.")
        } else {
            MsgBox("This is not an ET Application object.")
        }
}
```

### ChartTitle.Caption

返回或设置一个 **String** 值，它代表图表标题文本的方向。

**语法**

**express.Caption**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.Characters

返回 **Characters** 对象，它代表对象文本内某个区域的字符。使用 **Characters** 对象可为文本字符串内的字符设置格式。

**语法**

**express.Characters**

\*express是一个代表 ChartTitle 对象的变量。

**说明**

**Characters** 对象不是集合。

### ChartTitle.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 ChartTitle 对象的变量。

**说明**

**ChartFormat** 对象包含图表区的线条、填充、效果和文本格式。

### ChartTitle.Formula

获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.Formula**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.FormulaLocal

获取或设置一个 **String** 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaLocal**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.FormulaR1C1

获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.FormulaR1C1Local

获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1Local**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.Height

返回对象的高度（以磅为单位）。只读。

**语法**

**express.Height**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.HorizontalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 ChartTitle 对象的变量。

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

### ChartTitle.IncludeInLayout

如果在确定图表布局时图表标题将占用图表布局空间，则为 **True** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeInLayout**

\*express是一个代表 ChartTitle 对象的变量。

**说明**

此属性对于图表是否处于自动版式模式无影响。如果用户使用 **“图表上方”** 命令添加标题，则图表将变得较小。如果用户随后删除标题，或者选择一个覆盖标题选项，则图表将变得较大，就好像标题不在图表上一样

### ChartTitle.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.Orientation

返回或设置一个 **Variant** 值，它代表文本方向。

**语法**

**express.Orientation**

\*express是一个代表 ChartTitle 对象的变量。

**说明**

此属性的值可设为

### ChartTitle.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.Position

返回或设置图表上图表标题的位置。可读/写 **XlChartElementPosition** 类型。

**语法**

**express.Position**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.ReadingOrder

返回或设置指定对象的阅读次序。可为以下常量之一： **xlRTL** （从右到左）、 **xlLTR** （从左到右）或 **xlContext** 。 **Long** 类型，可读写。

**语法**

**express.ReadingOrder**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 ChartTitle 对象的变量。

**示例**

``` JavaScript
/*此示例向 myChart 的标题添加阴影。*/
function test() {
    Charts.Item("Chart1").ChartTitle.Shadow = true
}
```

### ChartTitle.Text

返回或设置指定对象中的文本。 **String** 型，可读写。

**语法**

**express.Text**

\*express是一个代表 ChartTitle 对象的变量。

**示例**

``` JavaScript
/*此示例设置 Chart1 的图表标题文本。*/

function test() {
let myChart = Charts.Item("Chart1")
myChart.HasTitle = true
myChart.ChartTitle.Text = "First Quarter Sales"
}
```

### ChartTitle.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 ChartTitle 对象的变量。

### ChartTitle.VerticalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 ChartTitle 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlBottom**      |
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlTop**         |

### ChartTitle.Width

返回对象的宽度（以磅为单位）。只读。

**语法**

**express.Width**

\*express是一个代表 ChartTitle 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartTitle](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
