# Frame 对象

## 简介

可以创建一个图形或控件的功能组。要为多个控件创建组，可先画框架，接着在框架中添加控件。

## 说明

可以创建一个图形或控件的功能组。要为多个控件创建组，可先画框架，接着在框架中添加控件。

## 方法

| 序号 | 名称                | 说明                       |
|:----:|:--------------------|:---------------------------|
|  1   | [Move](#Frame.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                | 说明                                                      |
|:----:|:------------------------------------|:----------------------------------------------------------|
|  1   | [BackColor](#Frame.BackColor)       | 获取或设置指定对象的内部颜色                              |
|  2   | [BorderColor](#Frame.BorderColor)   | 您可以通过该属性指定控件边框的颜色                        |
|  3   | [BorderStyle](#Frame.BorderStyle)   | 指定控件边框的显示方式                                    |
|  4   | [Caption](#Frame.Caption)           | 获取或设置出现在控件中的文本                              |
|  5   | [Enabled](#Frame.Enabled)           | 您可以通过该属性设置或返回是否启用该控件                  |
|  6   | [Font](#Frame.Font)                 | 指定对象的字体，只读                                      |
|  7   | [ForeColor](#Frame.ForeColor)       | 您可以通过该属性指定控件中文本的颜色                      |
|  8   | [Height](#Frame.Height)             | 获取或设置指定对象的高度                                  |
|  9   | [HeightPolicy](#Frame.HeightPolicy) | 获取或设置指定对像的高度策略                              |
|  10  | [Left](#Frame.Left)                 | 您可以通过该属性指定控件在对象或报表中的位置              |
|  11  | [Name](#Frame.Name)                 | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  12  | [TabIndex](#Frame.TabIndex)         | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  13  | [TabStop](#Frame.TabStop)           | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  14  | [Top](#Frame.Top)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  15  | [Visible](#Frame.Visible)           | 返回或设置该对象是否可见                                  |
|  16  | [Width](#Frame.Width)               | 获取或设置指定对象的宽度                                  |
|  17  | [WidthPolicy](#Frame.WidthPolicy)   | 获取或设置指定对象的宽度策略                              |

## 成员方法

### Frame.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Top* , *Left* , *Width* , *Height* )**

\*express是一个代表 Frame 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                               |
|----------|-----------|----------|------------------------------------|
| *Top*    | 必选      | Number   | 对象移动后相对于窗口的左边缘的距离 |
| *Left*   | 可选      | Number   | 对象移动后相对于窗口的上边缘的距离 |
| *Width*  | 可选      | Number   | 对象移动后的预期宽度               |
| *Height* | 可选      | Number   | 对象移动后的预期长度               |

**说明**

将对象移动到参数指定的位置

**示例**

``` JavaScript
/*将Frame1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.Frame1.Move(100,100,100,100)
}
```

## 成员属性

### Frame.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将Frame1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.Frame1.BackColor = 0x665898
}
```

### Frame.BorderColor

您可以通过该属性指定控件边框的颜色

**语法**

**express.BorderColor**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件边框的颜色

**示例**

``` JavaScript
/*将Frame1的边框设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.Frame1.BorderColor = 0x665898
}
```

### Frame.BorderStyle

指定控件边框的显示方式

**语法**

**express.BorderStyle**

\*express是一个代表 Frame 对象的变量。

**说明**

指定控件边框的显示方式，可设置为以下值：

| 设置   | 值  | 描述     |
|--------|-----|----------|
| None   | 0   | 无边框   |
| Single | 1   | 实现边框 |

**示例**

``` JavaScript
/*将Frame1的边框显示方式设置为实线*/
function func1()
{
    UserForm1.Frame1.BorderStyle = 1
}
```

### Frame.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将Frame1中的文本设置为"Frame1"*/
function func1()
{
    UserForm1.Frame1.Caption = "Frame1"
}
```

### Frame.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.Frame1.Enabled = true
}
```

### Frame.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 Frame 对象的变量。

**说明**

指定对象的字体，只读

### Frame.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将Frame1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.Frame1.ForeColor = 0x658978
}
```

### Frame.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将Frame1控件的高度设置为100*/
function func1()
{
    UserForm1.Frame1.Height = 100
}
```

### Frame.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将Frame1的高度设置为固定*/
function func1()
{
    UserForm1.Frame1.HeightPolicy = 2
}
```

### Frame.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将Frame1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.Frame1.Left = 100
}
```

### Frame.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### Frame.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将Frame1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.Frame1.TabIndex = 2
}
```

### Frame.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将Frame1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.Frame1.TabStop = true
}
```

### Frame.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将Frame1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.Frame1.Top = 100
}
```

### Frame.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 Frame 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将Frame1设置为可见*/
function func1()
{
    UserForm1.Frame1.Visible = true
}
```

### Frame.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将Frame1的宽度设置为100*/
function func1()
{
    UserForm1.Frame1.Width = 100
}
```

### Frame.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将Frame1的宽度设置为固定*/
function func1()
{
    UserForm1.Frame1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/Frame](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ChartFormat 对象

## 简介

提供对图表元素艺术字格式的访问。

## 说明

如果使用的属性或方法不适用于 ChartFormat 对象所附加到的对象的类型，则会产生运行时错误。

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                                                                                        |
|:----:|:--------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartFormat.Application)     | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 Application 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 Application 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
|  2   | [Creator](#ChartFormat.Creator)             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                  |
|  3   | [Fill](#ChartFormat.Fill)                   | 返回一个包含图表元素填充格式属性的父图表元素的 FillFormat 对象。只读。                                                                                                                                                      |
|  4   | [Glow](#ChartFormat.Glow)                   | 返回一个包含图表元素发光格式属性的指定图表的 GlowFormat 对象。只读。                                                                                                                                                        |
|  5   | [Line](#ChartFormat.Line)                   | 返回一个包含指定图表元素线条格式属性的 LineFormat 对象。只读。                                                                                                                                                              |
|  6   | [Parent](#ChartFormat.Parent)               | 返回指定对象的父对象。只读。                                                                                                                                                                                                |
|  7   | [PictureFormat](#ChartFormat.PictureFormat) | 返回包含图片的指定图表的 PictureFormat 对象。只读。                                                                                                                                                                         |
|  8   | [Shadow](#ChartFormat.Shadow)               | 返回一个 ShadowFormat 对象，该对象包含图表元素的阴影格式属性。只读。                                                                                                                                                        |
|  9   | [SoftEdge](#ChartFormat.SoftEdge)           | 返回指定图表的 SoftEdgeFormat 对象，该对象包含图表的柔化边缘格式设置属性。只读。                                                                                                                                            |
|  10  | [TextFrame2](#ChartFormat.TextFrame2)       | 返回一个包含指定图表元素文本格式的 TextFrame2 对象。只读。                                                                                                                                                                  |
|  11  | [ThreeD](#ChartFormat.ThreeD)               | 返回一个 ThreeDFormat 对象，该对象包含指定图表的三维效果格式属性。只读。                                                                                                                                                    |

## 成员属性

### ChartFormat.Application

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 ChartFormat 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息*/
function test()
{
    if(ActiveWorkbook.Application.Value == "ET"){
        alert("This is an ET Application object.")
    }
    else{
        alert("This is not an ET Application object.")
    }
}
```

### ChartFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ChartFormat.Fill

返回一个包含图表元素填充格式属性的父图表元素的 **FillFormat** 对象。只读。

**语法**

**express.Fill**

\*express是一个代表 ChartFormat 对象的变量。

### ChartFormat.Glow

返回一个包含图表元素发光格式属性的指定图表的 **GlowFormat** 对象。只读。

**语法**

**express.Glow**

\*express是一个代表 ChartFormat 对象的变量。

**说明**

发光效果为图形添加了色彩鲜明的彩色边缘。

### ChartFormat.Line

返回一个包含指定图表元素线条格式属性的 **LineFormat** 对象。只读。

**语法**

**express.Line**

\*express是一个代表 ChartFormat 对象的变量。

**说明**

对于线条来说， **LineFormat** 对象代表线条本身；而对于带有边框的图表来说， **LineFormat** 对象代表边框。

### ChartFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartFormat 对象的变量。

### ChartFormat.PictureFormat

返回包含图片的指定图表的 **PictureFormat** 对象。只读。

**语法**

**express.PictureFormat**

\*express是一个代表 ChartFormat 对象的变量。

### ChartFormat.Shadow

返回一个 **ShadowFormat** 对象，该对象包含图表元素的阴影格式属性。只读。

**语法**

**express.Shadow**

\*express是一个代表 ChartFormat 对象的变量。

### ChartFormat.SoftEdge

返回指定图表的 **SoftEdgeFormat** 对象，该对象包含图表的柔化边缘格式设置属性。只读。

**语法**

**express.SoftEdge**

\*express是一个代表 ChartFormat 对象的变量。

### ChartFormat.TextFrame2

返回一个包含指定图表元素文本格式的 **TextFrame2** 对象。只读。

**语法**

**express.TextFrame2**

\*express是一个代表 ChartFormat 对象的变量。

### ChartFormat.ThreeD

返回一个 **ThreeDFormat** 对象，该对象包含指定图表的三维效果格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 ChartFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# Connections 对象

## 简介

指定工作簿的 Connection 对象的集合。

## 说明

``` JavaScript
/*以下示例演示如何从现有文件向工作簿添加连接。*/
Application.ActiveWorkbook.Connections.AddFromFile(
    "C:\\Documents and Settings\\myComputer\\My Documents\\My Data Sources\\Northwind 2007 Customers.odc")
```

## 方法

| 序号 | 名称                                    | 说明                   |
|:----:|:----------------------------------------|:-----------------------|
|  1   | [Add](#Connections.Add)                 | 添加到工作簿的新连接。 |
|  2   | [AddFromFile](#Connections.AddFromFile) | 添加从指定文件的连接。 |
|  3   | [Item](#Connections.Item)               | 此方法创建一个连接项。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                               |
|:----:|:----------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Connections.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#Connections.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#Connections.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Parent](#Connections.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员方法

### Connections.Add

添加到工作簿的新连接。

**语法**

**express.Add ( *Name* , *Description* , *ConnectionString* , *CommandText* , *lCmdtype* )**

\*express是一个代表 Connections 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                     |
|--------------------|-----------|----------|--------------------------|
| *Name*             | 必选      | String   | 连接的名称。             |
| *Description*      | 必选      | String   | 连接的简要说明。         |
| *ConnectionString* | 必选      | String   | 连接字符串。             |
| *CommandText*      | 必选      | String   | 用于创建连接的命令文本。 |
| *lCmdtype*         | 可选      | String   | 命令类型。               |

**返回值**

WorkbookConnection

### Connections.AddFromFile

添加从指定文件的连接。

**语法**

**express.AddFromFile ( *Filename* )**

\*express是一个代表 Connections 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明         |
|------------|-----------|----------|--------------|
| *Filename* | 必选      | String   | 文件的名称。 |

**返回值**

WorkbookConnection

### Connections.Item

此方法创建一个连接项。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Connections 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明         |
|---------|-----------|----------|--------------|
| *Index* | 必选      | Variant  | 项的索引值。 |

**返回值**

WorkbookConnection

## 成员属性

### Connections.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Connections 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### Connections.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 Connections 对象的变量。

### Connections.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Connections 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Connections.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Connections 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Connections](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DataLabels 对象

## 简介

由数据系列中所有 DataLabel 对象组成的集合。

## 说明

每个 DataLabel 对象代表一个数据点或趋势线的数据标签。对于没有可定义数据点的数据系列（例如面积图数据系列）， DataLabels 集合包含单个数据标签。

使用 DataLabels 方法可返回单个 DataLabels 集合。

``` JavaScript
/*下例设置第一个图表的第一个数据系列中数据标签的数字格式。*/
function test(){
let myCollection = Application.Charts.Item(1).SeriesCollection(1)
    myCollection.HasDataLabels = true
    myCollection.DataLabels.NumberFormat = "##.##"
}
```

使用 DataLabels ( *index* )（其中 *index* 为数据标签的索引号）可返回单个 DataLabel 对象。

``` JavaScript
/*下例在第一张工作表上嵌入的第一个图表上，设置第一个数据系列中的第五个数据标签的数字格式。*/
Application.Worksheets.Item(1).ChartObjects(1).Chart 
    .SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```

## 方法

| 序号 | 名称                         | 说明                   |
|:----:|:-----------------------------|:-----------------------|
|  1   | [Delete](#DataLabels.Delete) | 删除对象。             |
|  2   | [Item](#DataLabels.Item)     | 从集合中返回一个对象。 |
|  3   | [Select](#DataLabels.Select) | 选择对象。             |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#DataLabels.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AutoText](#DataLabels.AutoText)                       | 如果对象会根据内容自动生成合适的文字，则为 True 。可读/写 Boolean 类型。                                                                                                                                                        |
|  3   | [Count](#DataLabels.Count)                             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  4   | [Creator](#DataLabels.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Format](#DataLabels.Format)                           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  6   | [HorizontalAlignment](#DataLabels.HorizontalAlignment) | 返回或设置一个 Variant 值，它代表指定对象的水平对齐方式。                                                                                                                                                                       |
|  7   | [Name](#DataLabels.Name)                               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  8   | [NumberFormat](#DataLabels.NumberFormat)               | 返回或设置一个 String 值，它代表对象的格式代码。                                                                                                                                                                                |
|  9   | [NumberFormatLinked](#DataLabels.NumberFormatLinked)   | 如果指定数字格式指向单元格（以便当单元格的格式更改时数据标签的格式也作相应的改动），则为 True 。 Boolean 类型，可读写。                                                                                                         |
|  10  | [NumberFormatLocal](#DataLabels.NumberFormatLocal)     | 以采用用户语言字符串的形式返回或设置一个 Variant 值，它代表对象的格式代码。                                                                                                                                                     |
|  11  | [Orientation](#DataLabels.Orientation)                 | 返回或设置一个 Variant 值，它代表文本方向。                                                                                                                                                                                     |
|  12  | [Parent](#DataLabels.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  13  | [Position](#DataLabels.Position)                       | 返回或设置一个 XlDataLabelPosition 值，它代表数据标签的位置。                                                                                                                                                                   |
|  14  | [ReadingOrder](#DataLabels.ReadingOrder)               | 返回或设置指定对象的阅读次序。可为以下常量之一： xlRTL （从右到左）、 xlLTR （从左到右）或 xlContext 。 Long 类型，可读写。                                                                                                     |
|  15  | [Separator](#DataLabels.Separator)                     | 设置或返回一个 Variant 值，它代表用于图表中数据标签的分隔符。可读/写。                                                                                                                                                          |
|  16  | [Shadow](#DataLabels.Shadow)                           | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  17  | [ShowBubbleSize](#DataLabels.ShowBubbleSize)           | 如果为 True ，则在图表中显示数据标签的气泡的大小。如果为 False ，则隐藏气泡的大小。 Boolean 类型，可读写。                                                                                                                      |
|  18  | [ShowCategoryName](#DataLabels.ShowCategoryName)       | 如果为 True ，则在图表中显示数据标签的分类名称。如果为 False ，则隐藏该分类的名称。 Boolean 类型，可读写。                                                                                                                      |
|  19  | [ShowLegendKey](#DataLabels.ShowLegendKey)             | 如果数据标签图例项标示可见，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                 |
|  20  | [ShowPercentage](#DataLabels.ShowPercentage)           | 如果为 True ，则在图表中显示数据标签的百分比数值。如果为 False ，则隐藏该百分比数值。 Boolean 类型，可读写。                                                                                                                    |
|  21  | [ShowSeriesName](#DataLabels.ShowSeriesName)           | 返回或设置一个 Boolean 值，它指明图表上数据标签的系列名称显示方式。如果为 True ，则显示系列名称。如果为 False ，则隐藏系列名称。可读写。                                                                                        |
|  22  | [VerticalAlignment](#DataLabels.VerticalAlignment)     | 返回或设置一个 Variant 值，它代表指定对象的垂直对齐方式。                                                                                                                                                                       |

## 成员方法

### DataLabels.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 DataLabels 对象的变量。

**返回值**

Variant

### DataLabels.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 DataLabels 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 必选      | Variant  | 对象的索引号。 |

**返回值**

包含在集合中的一个 DataLabel 对象。

**示例**

``` JavaScript
/*此示例对数据标签的数字格式进行设置，该数据标签是第一张工作表中第一张嵌入式图表的第一个数据系列的第五个数据标签。*/
function test(){
Application.Worksheets.Item(1).ChartObjects(1).Chart 
    .SeriesCollection(1).DataLabels.Item(5).NumberFormat = "0.000"
}
```

### DataLabels.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 DataLabels 对象的变量。

**返回值**

Variant

## 成员属性

### DataLabels.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 DataLabels 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### DataLabels.AutoText

如果对象会根据内容自动生成合适的文字，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.AutoText**

\*express是一个代表 DataLabels 对象的变量。

**示例**

``` JavaScript
/*此示例对 Chart1 的第一个数据序列的数据标签进行设置，自动生成合适的文字*/
Application.Charts.Item("Chart1").SeriesCollection(1).DataLabels.AutoText = true
```

| 注释                                                                                                                                                                                                                                                                                                                                                                     |
|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 如果在 **“即时窗口”** 中运行 ` ?ActiveChart.SeriesCollection(1).DataLabels.AutoText ` ，您将收到以下结果： ET 2003：不返回任何内容。 ET 2007 及更高版本：只有在所有 **DataLabels** 的 **AutoText** 都为 **True** 时返回 **True** ，如果所有 **DataLabels** 的 **AutoText** 都为 **False** ，或者部分 **DataLabels** 的 **AutoText** 为 **False** ，那么返回 **False** 。 |

### DataLabels.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 DataLabels 对象的变量。

### DataLabels.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DataLabels 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### DataLabels.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 DataLabels 对象的变量。

### DataLabels.HorizontalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 DataLabels 对象的变量。

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

### DataLabels.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 DataLabels 对象的变量。

### DataLabels.NumberFormat

返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

\*express是一个代表 DataLabels 对象的变量。

**说明**

格式代码与 **“设置单元格格式”** 对话框中的 **“格式代码”** 选项是同一个字符串。 **Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### DataLabels.NumberFormatLinked

如果指定数字格式指向单元格（以便当单元格的格式更改时数据标签的格式也作相应的改动），则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.NumberFormatLinked**

\*express是一个代表 DataLabels 对象的变量。

### DataLabels.NumberFormatLocal

以采用用户语言字符串的形式返回或设置一个 **Variant** 值，它代表对象的格式代码。

**语法**

**express.NumberFormatLocal**

\*express是一个代表 DataLabels 对象的变量。

**说明**

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### DataLabels.Orientation

返回或设置一个 **Variant** 值，它代表文本方向。

**语法**

**express.Orientation**

\*express是一个代表 DataLabels 对象的变量。

**说明**

此属性的值可设为

### DataLabels.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DataLabels 对象的变量。

### DataLabels.Position

返回或设置一个 **XlDataLabelPosition** 值，它代表数据标签的位置。

**语法**

**express.Position**

\*express是一个代表 DataLabels 对象的变量。

### DataLabels.ReadingOrder

返回或设置指定对象的阅读次序。可为以下常量之一： **xlRTL** （从右到左）、 **xlLTR** （从左到右）或 **xlContext** 。 **Long** 类型，可读写。

**语法**

**express.ReadingOrder**

\*express是一个代表 DataLabels 对象的变量。

### DataLabels.Separator

设置或返回一个 **Variant** 值，它代表用于图表中数据标签的分隔符。可读/写。

**语法**

**express.Separator**

\*express是一个代表 DataLabels 对象的变量。

**说明**

如果使用字符串，则用一个字符串作为分隔符。如果使用 **xlDataLabelSeparatorDefault** (= 1)，则用默认数据标签分隔符，即逗号或换行符，具体取决于数据标签。

当返回值“1”时，表明用户未更改默认的分隔符逗号“,”。您也可通过传递值“1”将分隔符更改回默认值。

以程序方式访问数据标签之前必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例将第一个图表上的第一个系列的数据标签分隔符设置为分号。此示例假定活动工作表上存在图表。*/
Application.ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1) 
        .DataLabels.Separator = ";"
```

### DataLabels.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 DataLabels 对象的变量。

### DataLabels.ShowBubbleSize

如果为 **True** ，则在图表中显示数据标签的气泡的大小。如果为 **False** ，则隐藏气泡的大小。 **Boolean** 类型，可读写。

**语法**

**express.ShowBubbleSize**

\*express是一个代表 DataLabels 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例显示第一个图表上第一个系列的数据标签的气泡大小。此示例假定活动工作表上存在图表。*/
function test(){
 Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1) 
        .DataLabels.ShowBubbleSize = true
}
```

### DataLabels.ShowCategoryName

如果为 **True** ，则在图表中显示数据标签的分类名称。如果为 **False** ，则隐藏该分类的名称。 **Boolean** 类型，可读写。

**语法**

**express.ShowCategoryName**

\*express是一个代表 DataLabels 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例显示第一个图表上第一个系列的数据标签的分类名称。此示例假定活动工作表上存在图表。*/
function test(){
  Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1) 
        .DataLabels.ShowCategoryName = true
}
```

### DataLabels.ShowLegendKey

如果数据标签图例项标示可见，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowLegendKey**

\*express是一个代表 DataLabels 对象的变量。

**示例**

``` JavaScript
/*此示例设置 Chart1 中系列一的数据标签，以显示数值和图例标示*/
function test(){
let myDataLabel = Application.Charts.Item("Chart1").SeriesCollection(1).DataLabels
    myDataLabel.ShowLegendKey = true
    myDataLabel.Type = xlShowValue
}
```

### DataLabels.ShowPercentage

如果为 **True** ，则在图表中显示数据标签的百分比数值。如果为 **False** ，则隐藏该百分比数值。 **Boolean** 类型，可读写。

**语法**

**express.ShowPercentage**

\*express是一个代表 DataLabels 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例显示了第一个图表上第一个数据系列的数据标签的百分比数值。此示例假定活动工作表上存在图表。*/
function test(){
  Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1) 
        .DataLabels.ShowPercentage = true
}
```

### DataLabels.ShowSeriesName

返回或设置一个 **Boolean** 值，它指明图表上数据标签的系列名称显示方式。如果为 **True** ，则显示系列名称。如果为 **False** ，则隐藏系列名称。可读写。

**语法**

**express.ShowSeriesName**

\*express是一个代表 DataLabels 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例显示了第一个图表上第一个数据系列的数据标签的系列名称。此示例假定活动工作表上存在图表。*/
function test(){
Application.ActiveSheet.ChartObjects(1).Activate
    ActiveChart.SeriesCollection(1) 
        .DataLabels.ShowSeriesName = true
}
```

### DataLabels.VerticalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 DataLabels 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlBottom**      |
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlTop**         |

> WPS 开发文档： [WPS 基础接口/表格 API 参考/DataLabels](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ErrorBars 对象

## 简介

代表图表数据系列上的误差线。

## 说明

误差线表明图表数据的不确定程度。只有二维图表上的面积、条形、柱形、折线和散点图组中的数据系列可以有误差线。只有散点图组中的数据系列可以有 x 误差线和 y 误差线。此对象不是集合。没有代表单个误差线的对象；或者打开系列中所有数据点的 x 误差线或 y 误差线，或者将其全部关闭。

ErrorBar 方法更改误差线的格式和类型。

使用 ErrorBars 属性可返回 ErrorBars 对象。

``` JavaScript
/*下例打开嵌入的第一个图表的第一个数据系列的误差线，并设置误差线的尾部样式。*/
function test(){
Application.Worksheets.Item("sheet1").ChartObjects(1).Activate()
ActiveChart.SeriesCollection(1).HasErrorBars = true
ActiveChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
}
```

## 方法

| 序号 | 名称                                    | 说明                 |
|:----:|:----------------------------------------|:---------------------|
|  1   | [ClearFormats](#ErrorBars.ClearFormats) | 清除对象的格式设置。 |
|  2   | [Delete](#ErrorBars.Delete)             | 删除对象。           |
|  3   | [Select](#ErrorBars.Select)             | 选择对象。           |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ErrorBars.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Border](#ErrorBars.Border)           | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  3   | [Creator](#ErrorBars.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [EndStyle](#ErrorBars.EndStyle)       | 返回或设置误差线的尾端样式。可为以下 XlEndStyleCap 常量之一： xlCap 或 xlNoCap 。 Long 类型，可读写。                                                                                                                           |
|  5   | [Format](#ErrorBars.Format)           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  6   | [Name](#ErrorBars.Name)               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  7   | [Parent](#ErrorBars.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### ErrorBars.ClearFormats

清除对象的格式设置。

**语法**

**express.ClearFormats ()**

\*express是一个代表 ErrorBars 对象的变量。

**返回值**

Variant

### ErrorBars.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 ErrorBars 对象的变量。

**返回值**

Variant

### ErrorBars.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 ErrorBars 对象的变量。

**返回值**

Variant

## 成员属性

### ErrorBars.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ErrorBars 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### ErrorBars.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 ErrorBars 对象的变量。

### ErrorBars.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ErrorBars 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ErrorBars.EndStyle

返回或设置误差线的尾端样式。可为以下 **XlEndStyleCap** 常量之一： **xlCap** 或 **xlNoCap** 。 **Long** 类型，可读写。

**语法**

**express.EndStyle**

\*express是一个代表 ErrorBars 对象的变量。

**示例**

``` JavaScript
/*本示例对 Chart1 中第一个系列的误差线的尾端样式进行设置。本示例必须在其第一个系列带 Y 误差线的二维折线图上运行。*/
Application.Charts.Item("Chart1").SeriesCollection(1).ErrorBars.EndStyle = xlCap
```

### ErrorBars.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 ErrorBars 对象的变量。

### ErrorBars.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 ErrorBars 对象的变量。

### ErrorBars.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ErrorBars 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ErrorBars](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ErrorCheckingOptions 对象

## 简介

代表应用程序的错误检查选项。

## 说明

使用 Application 对象的 ErrorCheckingOptions 属性可返回 ErrorCheckingOptions 对象。

引用 Errors 对象的 Item 属性可查看与错误检查选项关联的索引值列表。

返回 ErrorCheckingOptions 对象后，可使用以下属性设置或返回错误检查选项。这些属性都是 ErrorCheckingOptions 对象的成员。

- BackgroundChecking
- EmptyCellReferences
- EvaluateToError
- InconsistentFormula
- IndicatorColorIndex
- NumberAsText
- OmittedCells
- TextDate
- UnlockedFormulaCells

``` JavaScript
/*下例使用 TextDate 属性启用对两位数年份文本日期的错误检查，并通知用户。*/
function test(){
 let rngFormula = Application.Range("A1")

    Range("A1").Formula = "'April 23, 00"
    Application.ErrorCheckingOptions.TextDate = true

    //Perform check to see if 2 digit year TextDate check is on.
    if(rngFormula.Errors.Item(xlTextDate).Value == true) {
        alert("The text date error checking feature is enabled.")
    }
    else {
        alert("The text date error checking feature is not on.")
    }

}
```

## 属性

| 序号 | 名称                                                                       | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ErrorCheckingOptions.Application)                           | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [BackgroundChecking](#ErrorCheckingOptions.BackgroundChecking)             | 警告用户并指出违背可用错误检查规则的所有单元格。如果该属性设置为 True （默认值），则在违背启用错误的所有单元格旁边将出现 “自动更正选项” 按钮。如果为 False ，则禁用背景错误检查。 Boolean 类型，可读写。                        |
|  3   | [Creator](#ErrorCheckingOptions.Creator)                                   | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [EmptyCellReferences](#ErrorCheckingOptions.EmptyCellReferences)           | 如果该属性设置为 True （默认值），则 ET 使用 “自动更正选项” 按钮识别被选取的单元格，这些单元格中包括引用了空单元格的公式。如果该值为 False ，则禁用空单元格引用检查。 Boolean 类型，可读写。                                    |
|  5   | [EvaluateToError](#ErrorCheckingOptions.EvaluateToError)                   | 如果该属性设置为 True （默认值），则 ET 将使用 “自动更正选项” 按钮识别被选中的单元格，这些单元格中包含计算结果错误的公式。如果该值为 False ，则禁用对计算结果为错误值的单元格的错误检查。 Boolean 类型，可读写。                |
|  6   | [InconsistentFormula](#ErrorCheckingOptions.InconsistentFormula)           | 如果该值为 True （默认值），则 ET 将识别包含不一致公式的单元格区域。如果该值为 False ，则禁用不一致公式的检查。 Boolean 类型，可读写。                                                                                          |
|  7   | [InconsistentTableFormula](#ErrorCheckingOptions.InconsistentTableFormula) | 如果表公式不一致，则返回 True 。可读/写 Boolean 类型。                                                                                                                                                                          |
|  8   | [IndicatorColorIndex](#ErrorCheckingOptions.IndicatorColorIndex)           | 返回或设置错误检查选项的指示器的颜色。 XlColorIndex 类型，可读写。                                                                                                                                                              |
|  9   | [ListDataValidation](#ErrorCheckingOptions.ListDataValidation)             | 如果在列表中启用了数据有效性验证，则该属性值为 Boolean 值 True 。 Boolean 类型，可读写。                                                                                                                                        |
|  10  | [NumberAsText](#ErrorCheckingOptions.NumberAsText)                         | 如果该值为 True （默认值），则 ET 将用 “自动更正选项” 按钮识别被选定的单元格，这些单元格包含文本格式的数字。如果该值为 False ，则禁用对文本格式数字的错误检查。 Boolean 类型，可读写。                                          |
|  11  | [OmittedCells](#ErrorCheckingOptions.OmittedCells)                         | 如果该值为 True （默认值），则 ET 将用“自动更正选项”按钮识别包含公式的选定单元格，其中该公式引用的区域应包括相邻单元格，但这些单元格却被遗漏掉了。如果该值为 False ，则禁用对被遗漏单元格的错误检查。 Boolean 类型，可读写。    |
|  12  | [Parent](#ErrorCheckingOptions.Parent)                                     | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  13  | [TextDate](#ErrorCheckingOptions.TextDate)                                 | 如果设置为 True （默认值），则 ET 将使用 “自动更正选项” 按钮识别包含用两位数字表示年份的文本日期的单元格。如果该属性值为 False ，则禁用此类单元格的错误检查。 Boolean 类型，可读写。                                            |
|  14  | [UnlockedFormulaCells](#ErrorCheckingOptions.UnlockedFormulaCells)         | 如果该值为 True （默认值），则 ET 将识别未锁定并包含一个公式的选定单元格。如果该值为 False ，则禁用包含公式的未锁定单元格的错误检查。 Boolean 类型，可读写。                                                                    |

## 成员属性

### ErrorCheckingOptions.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}   }
}
```

### ErrorCheckingOptions.BackgroundChecking

警告用户并指出违背可用错误检查规则的所有单元格。如果该属性设置为 **True** （默认值），则在违背启用错误的所有单元格旁边将出现 **“自动更正选项”** 按钮。如果为 **False** ，则禁用背景错误检查。 **Boolean** 类型，可读写。

**语法**

**express.BackgroundChecking**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**说明**

请参考 **ErrorCheckingOptions** 对象以查阅该对象的可用成员列表。

**示例**

``` JavaScript
/*在下例中，当用户选择单元格 A1（该单元格包含引用空单元格的公式）时，将出现“自动更正选项”按钮。*/
function test(){
 //Simulate an error by referring to empty cells.
    Application.ErrorCheckingOptions.BackgroundChecking = true
    Range("A1").Select()
    ActiveCell.Formula = "=A2/A3"
}
```

### ErrorCheckingOptions.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ErrorCheckingOptions.EmptyCellReferences

如果该属性设置为 **True** （默认值），则 ET 使用 **“自动更正选项”** 按钮识别被选取的单元格，这些单元格中包括引用了空单元格的公式。如果该值为 **False** ，则禁用空单元格引用检查。 **Boolean** 类型，可读写。

**语法**

**express.EmptyCellReferences**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**示例**

``` JavaScript
/*在下例中，单元格 A1 旁出现“自动更正选项”按钮，该单元格包括引用空单元格的公式。*/
function test(){
 Application.ErrorCheckingOptions.EmptyCellReferences = true
    Range("A1").Formula = "=A2+A3"
}
```

### ErrorCheckingOptions.EvaluateToError

如果该属性设置为 **True** （默认值），则 ET 将使用 **“自动更正选项”** 按钮识别被选中的单元格，这些单元格中包含计算结果错误的公式。如果该值为 **False** ，则禁用对计算结果为错误值的单元格的错误检查。 **Boolean** 类型，可读写。

**语法**

**express.EvaluateToError**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**示例**

``` JavaScript
/*在本示例中，单元格 A3 旁出现“自动更正选项”按钮，该单元格包括一个被零除的错误*/
function test(){

  //Simulate a divide-by-zero error.
    Application.ErrorCheckingOptions.EvaluateToError = true
    Range("A1").Value2 = 1
    Range("A2").Value2 = 0
    Range("A3").Formula = "=A1/A2"
}
```

### ErrorCheckingOptions.InconsistentFormula

如果该值为 **True** （默认值），则 ET 将识别包含不一致公式的单元格区域。如果该值为 **False** ，则禁用不一致公式的检查。 **Boolean** 类型，可读写。

**语法**

**express.InconsistentFormula**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**说明**

单元格区域中相一致的公式必须位于包含不一致的公式的单元格的左边、右边或上边、下边，以便 **InconsistentFormula** 属性正常工作。

**示例**

``` JavaScript
/*在下例中，当用户选择单元格 B4（包含不一致的公式）时，将显示“自动更正选项”按钮。*/
function test(){

 Application.ErrorCheckingOptions.InconsistentFormula = true

    Range("A1:A3").Value2 = 1
    Range("B1:B3").Value2 = 2
    Range("C1:C3").Value2 = 3

    Range("A4").Formula = "=SUM(A1:A3)"  //Consistent formula.
    Range("B4").Formula = "=SUM(B1:B2)"  //Inconsistent formula.
    Range("C4").Formula = "=SUM(C1:C3)"  //Consistent formula.
}
```

### ErrorCheckingOptions.InconsistentTableFormula

如果表公式不一致，则返回 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.InconsistentTableFormula**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

### ErrorCheckingOptions.IndicatorColorIndex

返回或设置错误检查选项的指示器的颜色。 **XlColorIndex** 类型，可读写。

**语法**

**express.IndicatorColorIndex**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**说明**

|                                                         |
|---------------------------------------------------------|
| **XlColorIndex** 可以是下列 **XlColorIndex** 常量之一。 |
| **xlColorIndexAutomatic** 为默认系统颜色。              |
| **xlColorIndexNone** 不与该属性一起使用。               |

通过输入相应的索引值可为指示器指定特殊的颜色。可用 **Colors** 属性返回当前的调色板。

**示例**

``` JavaScript
/*在本示例中，ET 将查看错误检查的指示器颜色是否设置为默认系统颜色，并通知给用户。*/
function test(){
 if(Application.ErrorCheckingOptions.IndicatorColorIndex == xlColorIndexAutomatic) {
        alert("Your indicator color for error checking is set to the default system color.")
    }
    else {
        alert("Your indicator color for error checking is not set to the default system color.")
    }
}
```

### ErrorCheckingOptions.ListDataValidation

如果在列表中启用了数据有效性验证，则该属性值为 **Boolean** 值 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ListDataValidation**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

### ErrorCheckingOptions.NumberAsText

如果该值为 **True** （默认值），则 ET 将用 **“自动更正选项”** 按钮识别被选定的单元格，这些单元格包含文本格式的数字。如果该值为 **False** ，则禁用对文本格式数字的错误检查。 **Boolean** 类型，可读写。

**语法**

**express.NumberAsText**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**示例**

``` JavaScript
/*在下例中，包含文本格式数字的单元格 A1 旁出现“自动更正选项”按钮。*/
function test(){
 //Simulate an error by referencing a number stored as text.
    Application.ErrorCheckingOptions.NumberAsText = true
    Range("A1").Value = "'1"
}
```

### ErrorCheckingOptions.OmittedCells

如果该值为 **True** （默认值），则 ET 将用“自动更正选项”按钮识别包含公式的选定单元格，其中该公式引用的区域应包括相邻单元格，但这些单元格却被遗漏掉了。如果该值为 **False** ，则禁用对被遗漏单元格的错误检查。 **Boolean** 类型，可读写。

**语法**

**express.OmittedCells**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**示例**

``` JavaScript
/*下例中，“自动更正选项”按钮出现在包含公式的单元格 A4 旁。*/
function test(){
 Application.ErrorCheckingOptions.OmittedCells = true
    Range("A1").Value2 = 1
    Range("A2").Value2 = 2
    Range("A3").Value2 = 3
    Range("A4").Formula = "=Sum(A1:A2)"
}
```

### ErrorCheckingOptions.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

### ErrorCheckingOptions.TextDate

如果设置为 **True** （默认值），则 ET 将使用 **“自动更正选项”** 按钮识别包含用两位数字表示年份的文本日期的单元格。如果该属性值为 **False** ，则禁用此类单元格的错误检查。 **Boolean** 类型，可读写。

**语法**

**express.TextDate**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**示例**

``` JavaScript
/*下例中，“自动更正选项”按钮将出现在单元格 A1 旁，该单元格包含用两位数字表示年份的文本日期。*/

function test(){

 //Simulate an error by referencing a text date with a two-digit year.
    Application.ErrorCheckingOptions.TextDate = true
    Range("A1").Formula = "'April 23, 00"
}
```

### ErrorCheckingOptions.UnlockedFormulaCells

如果该值为 **True** （默认值），则 ET 将识别未锁定并包含一个公式的选定单元格。如果该值为 **False** ，则禁用包含公式的未锁定单元格的错误检查。 **Boolean** 类型，可读写。

**语法**

**express.UnlockedFormulaCells**

\*express是一个代表 ErrorCheckingOptions 对象的变量。

**示例**

``` JavaScript
/*本示例中，“自动更正选项”按钮出现在单元格 A3 旁，该单元格是一个包含公式的未锁定单元格。*/
function test(){
 Application.ErrorCheckingOptions.UnlockedFormulaCells = true
    Range("A1").Value2 = 1
    Range("A2").Value2 = 2
    Range("A3").Formula = "=A1+A2"
    Range("A3").Locked = false
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ErrorCheckingOptions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListColumns 对象

## 简介

指定的 ListObject 对象中所有 ListColumn 对象的集合。

## 说明

每个 ListColumn 对象都代表表格中的一列。

| 注释                                               |
|----------------------------------------------------|
| 该列的名称会自动生成。在添加完该列后可更改其名称。 |

使用 ListObject 对象的 ListColumns 属性可返回 ListColumns 集合。

``` JavaScript
/*下例给工作簿的第一张工作表的默认 ListObject 对象添加一个新列。由于未指定位置，因此在最右边添加一个新列。*/
let myNewColumn = Application.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Add()
```

## 方法

| 序号 | 名称                    | 说明                   |
|:----:|:------------------------|:-----------------------|
|  1   | [Add](#ListColumns.Add) | 向列表对象中添加新列。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ListColumns.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#ListColumns.Count)             | 返回一个 Integer 值，它代表集合中对象的数量。                                                                                                                                                                                   |
|  3   | [Creator](#ListColumns.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Item](#ListColumns.Item)               | 从集合中返回一个对象。                                                                                                                                                                                                          |
|  5   | [Parent](#ListColumns.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### ListColumns.Add

向列表对象中添加新列。

**语法**

**express.Add ( *Position* )**

\*express是一个代表 ListColumns 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                      |
|------------|-----------|----------|---------------------------------------------------------------------------|
| *Position* | 可选      | Variant  | Integer 类型。从 1 开始指定新列的相对位置。以前位于此位置的列则向后移动。 |

**返回值**

一个代表新列的 ListColumn 对象。

**说明**

如果不指定 *Position* ，就在最右边添加一列，并自动为该列生成名称。该列的名称可在添加后更改。

**示例**

``` JavaScript
/*下例给工作簿的第一张工作表的默认 ListObject 对象添加一个新列。由于未指定位置，因此在最右边添加一个新列。*/
function test(){
let myNewColumn = Application.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Add()
}
```

## 成员属性

### ListColumns.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ListColumns 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### ListColumns.Count

返回一个 **Integer** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 ListColumns 对象的变量。

### ListColumns.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ListColumns 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ListColumns.Item

从集合中返回一个对象。

**语法**

**express.Item**

\*express是一个代表 ListColumns 对象的变量。

### ListColumns.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ListColumns 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ListColumns](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# MultiThreadedCalculation 对象

## 简介

返回或设置并发计算模式。

## 属性

| 序号 | 名称                                                 | 说明                                                                                                                                                               |
|:----:|:-----------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#MultiThreadedCalculation.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#MultiThreadedCalculation.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Enabled](#MultiThreadedCalculation.Enabled)         | 使用 Enabled 属性可以在运行时启用或禁用 MultiThreadedCalculation 对象。可读/写。                                                                                   |
|  4   | [Parent](#MultiThreadedCalculation.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  5   | [ThreadCount](#MultiThreadedCalculation.ThreadCount) | 获取进程的线程总数，这些线程是指定的 MultiThreadedCalculation 对象的一部分。                                                                                       |
|  6   | [ThreadMode](#MultiThreadedCalculation.ThreadMode)   | 返回或设置指定的 MultiThreadedCalculation 对象的线程模式。可读/写 XlThreadMode 类型。                                                                              |

## 成员属性

### MultiThreadedCalculation.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

### MultiThreadedCalculation.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### MultiThreadedCalculation.Enabled

使用 **Enabled** 属性可以在运行时启用或禁用 **MultiThreadedCalculation** 对象。可读/写。

**语法**

**express.Enabled**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

### MultiThreadedCalculation.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

### MultiThreadedCalculation.ThreadCount

获取进程的线程总数，这些线程是指定的 **MultiThreadedCalculation** 对象的一部分。

**语法**

**express.ThreadCount**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

### MultiThreadedCalculation.ThreadMode

返回或设置指定的 **MultiThreadedCalculation** 对象的线程模式。可读/写 **XlThreadMode** 类型。

**语法**

**express.ThreadMode**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

**说明**

线程模式可以设置为 **XlThreadModeAutomatic** 或 **XlThreadModeManual** 。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/MultiThreadedCalculation](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotFields 对象

## 简介

指定的数据透视表中所有 PivotField 对象的集合。

## 说明

下列存取器方法返回数据透视表字段的子集，在某些情况下，使用这些属性会更为方便：

- ColumnFields 属性
- DataFields 属性
- HiddenFields 属性
- PageFields 属性
- RowFields 属性
- VisibleFields 属性

使用 PivotFields 对象的 PivotTable 方法可返回 PivotFields 集合。下例列举了 Sheet3 上第一张数据透视表中的字段名称。

``` JavaScript
let pTab = Worksheets.Item("sheet3").PivotTables(1)
for(let i=1;i <= pTab.PivotFields().Count;i++) {
    MsgBox(pTab.PivotFields(i).Name)
}
```

使用 PivotFields ( *index* )（其中 *index* 是字段名称或索引号）可返回一个 PivotField 对象。下例使 Sheet3 上第一张数据透视表中的字段“Year”成为行字段。

``` JavaScript
Worksheets.Item("sheet3").PivotTables(1).PivotFields("year").Orientation = xlRowField
```

## 方法

| 序号 | 名称                      | 说明                   |
|:----:|:--------------------------|:-----------------------|
|  1   | [Item](#PivotFields.Item) | 从集合中返回一个对象。 |

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFields.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFields.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 Long 值，它代表集合中对象的数量。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFields.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFields.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。只读。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotFields.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 PivotFields 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

一个 Object 值，它代表集合包含的某个对象。

**说明**

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

此示例将工作表 Sheet3 中第一个数据透视表上的 Year 字段设置为行字段。

``` JavaScript
Worksheets.Item("sheet3").PivotTables(1).PivotFields().Item("year").Orientation = xlRowField
```

## 成员属性

### PivotFields.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读

**语法**

**express.Application**

\*express是一个代表 PivotFields 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### PivotFields.Count

table { word-break:break-all; }

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 PivotFields 对象的变量。

### PivotFields.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotFields 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotFields.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotFields 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotFields](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# QueryTable 对象

## 简介

代表一个利用从外部数据源（如 SQL Server 或 Microsoft Access 数据库）返回的数据生成的工作表表格。

## 说明

QueryTable 对象是 QueryTables 集合的成员。

## 方法

| 序号 | 名称                                       | 说明                                                                                                      |
|:----:|:-------------------------------------------|:----------------------------------------------------------------------------------------------------------|
|  1   | [CancelRefresh](#QueryTable.CancelRefresh) | 取消指定查询表的所有后台查询。使用 Refreshing 属性可以确定当前是否正在进行后台查询。                      |
|  2   | [Delete](#QueryTable.Delete)               | 删除对象。                                                                                                |
|  3   | [Refresh](#QueryTable.Refresh)             | 更新外部数据区域 ( QueryTable )。                                                                         |
|  4   | [ResetTimer](#QueryTable.ResetTimer)       | 重新设置指定的查询表或数据透视表的刷新计时器，使其时间间隔等于上次使用 RefreshPeriod 属性设置的时间间隔。 |
|  5   | [SaveAsODC](#QueryTable.SaveAsODC)         | 将查询表缓存的源保存为“Microsoft Office 数据连接”文件。                                                   |

## 属性

| 序号 | 名称                                                                       | 说明                                                                                                                                                                                                                                                |
|:----:|:---------------------------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AdjustColumnWidth](#QueryTable.AdjustColumnWidth)                         | 如果每次刷新指定的查询表时列宽都会自动调整为最适合的宽度，则为 True 。如果每次刷新时列宽不进行自动调整，则为 False 。默认值为 True 。 Boolean 类型，可读写。                                                                                        |
|  2   | [Application](#QueryTable.Application)                                     | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象（可对 OLE 自动操作对象使用此属性来返回该对象的应用程序）。只读。                         |
|  3   | [BackgroundQuery](#QueryTable.BackgroundQuery)                             | 如果查询表的查询是异步执行（在后台执行）的，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                     |
|  4   | [CommandText](#QueryTable.CommandText)                                     | 返回或设置指定数据源的命令串。 Variant 型，可读写。                                                                                                                                                                                                 |
|  5   | [CommandType](#QueryTable.CommandType)                                     | 返回或设置备注部分的下表中列出的 XlCmdType 常量之一。返回或设置的常量用于描述 CommandText 属性的值。默认值为 xlCmdSQL 。 XlCmdType 类型，可读写。                                                                                                   |
|  6   | [Connection](#QueryTable.Connection)                                       | 返回或设置包含下列某项的字符串：允许 ET 连接到 OLE DB 数据源的 OLE DB 设置；允许 ET 连接到 ODBC 数据源的 ODBC 设置；允许 ET 连接到 Web 数据源的 URL；或者文本文件的名称或路径，或是指定某个数据库或 Web 查询的文件名称或路径。 Variant 型，可读写。 |
|  7   | [Creator](#QueryTable.Creator)                                             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                                          |
|  8   | [Destination](#QueryTable.Destination)                                     | 返回查询表目标区域（查询结果表放置的区域）的左上角单元格。目标区域必须位于包含 QueryTable 对象的工作表中。 Range 类型，只读。                                                                                                                       |
|  9   | [EditWebPage](#QueryTable.EditWebPage)                                     | 返回或设置用于 Web 查询的网页的 统一资源定位符 (URL) 。 Variant 类型，可读写。                                                                                                                                                                      |
|  10  | [EnableEditing](#QueryTable.EnableEditing)                                 | 如果允许用户对指定查询表进行编辑，则该值为 True ；如果用户只能刷新查询表，则该值为 False 。 Boolean 类型，可读写。                                                                                                                                  |
|  11  | [EnableRefresh](#QueryTable.EnableRefresh)                                 | 如果用户可刷新数据透视表高速缓存或查询表，则为 True 。默认值为 True 。 Boolean 类型，可读写。                                                                                                                                                       |
|  12  | [FetchedRowOverflow](#QueryTable.FetchedRowOverflow)                       | 如果上次使用 Refresh 方法返回的行数比工作表中可用行数大，则该值为 True 。 Boolean 类型，只读。                                                                                                                                                      |
|  13  | [FieldNames](#QueryTable.FieldNames)                                       | 如果数据源的字段名称作为返回数据的列标题显示，则该值为 True 。默认值为 True 。 Boolean 类型，可读写。                                                                                                                                               |
|  14  | [FillAdjacentFormulas](#QueryTable.FillAdjacentFormulas)                   | 如果数据源的字段名称作为返回数据的列标题显示，则该值为 True 。默认值为 True 。 Boolean 类型，可读写。                                                                                                                                               |
|  15  | [ListObject](#QueryTable.ListObject)                                       | 为 QueryTable 对象返回一个 ListObject 对象。 ListObject 对象类型，只读。                                                                                                                                                                            |
|  16  | [MaintainConnection](#QueryTable.MaintainConnection)                       | 如果从刷新数据开始直至关闭工作簿，都一直保留指向指定数据源的连接，则为 True 。默认值是 True 。 Boolean 类型，可读写。                                                                                                                               |
|  17  | [Name](#QueryTable.Name)                                                   | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                                        |
|  18  | [Parameters](#QueryTable.Parameters)                                       | 返回一个 Parameters 集合，该集合表示查询表参数。只读。                                                                                                                                                                                              |
|  19  | [Parent](#QueryTable.Parent)                                               | 返回指定对象的父对象。只读。                                                                                                                                                                                                                        |
|  20  | [PostText](#QueryTable.PostText)                                           | 返回或设置用于 post 方法的字符串，post 方法用于向 Web 服务器输入数据以从 Web 查询中返回数据。 String 类型，可读写。                                                                                                                                 |
|  21  | [PreserveColumnInfo](#QueryTable.PreserveColumnInfo)                       | 如果每次刷新查询表时，列排序、筛选和布局信息都会保留，则该值为 True 。默认值为 False 。 Boolean 类型，可读写。                                                                                                                                      |
|  22  | [PreserveFormatting](#QueryTable.PreserveFormatting)                       | 如果将数据前五行的任何常用格式设置应用到查询表的新数据行，则为 True 。对未使用的单元格不进行格式设置。如果将应用到查询表的最新一次自动套用格式应用于新数据行，则属性为 False 。默认值是 True 。                                                     |
|  23  | [QueryType](#QueryTable.QueryType)                                         | 表示 ET 填充查询表时所使用的查询类型。 XlQueryType 类型，只读。                                                                                                                                                                                     |
|  24  | [Recordset](#QueryTable.Recordset)                                         | 返回或设置一个 Recordset 对象，该对象用作指定查询表的数据源。                                                                                                                                                                                       |
|  25  | [Refreshing](#QueryTable.Refreshing)                                       | 如果指定的查询表正在进行后台查询，则为 True 。可读/写 Boolean 类型。                                                                                                                                                                                |
|  26  | [RefreshOnFileOpen](#QueryTable.RefreshOnFileOpen)                         | 如果每次打开工作簿时，数据透视表高速缓存或查询表自动更新，则为 True 。默认值为 False 。 Boolean 类型，可读写。                                                                                                                                      |
|  27  | [RefreshPeriod](#QueryTable.RefreshPeriod)                                 | 返回或设置两次刷新之间的时间间隔。 Long 型，可读写。                                                                                                                                                                                                |
|  28  | [RefreshStyle](#QueryTable.RefreshStyle)                                   | 返回或设置为了容纳查询返回的记录集中的行数而在指定工作表中插入或删除行时所使用的方式。 XlCellInsertionMode 类型。可读写。                                                                                                                           |
|  29  | [ResultRange](#QueryTable.ResultRange)                                     | 返回一个 Range 对象，该对象代表指定查询表所覆盖的工作表区域。只读。                                                                                                                                                                                 |
|  30  | [RobustConnect](#QueryTable.RobustConnect)                                 | 返回或设置数据透视表缓存与其数据源连接的方式。 XlRobustConnect 类型，可读写。                                                                                                                                                                       |
|  31  | [RowNumbers](#QueryTable.RowNumbers)                                       | 如果行号作为第一列添加到指定查询表，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                     |
|  32  | [SaveData](#QueryTable.SaveData)                                           | 如果将数据透视表的数据随工作簿一起保存，则为 True 。如果仅保存数据透视表的定义，则为 False 。 Boolean 类型，可读写。                                                                                                                                |
|  33  | [SavePassword](#QueryTable.SavePassword)                                   | 如果将 ODBC 连接字符串中的密码信息与指定查询一起保存，则为 True 。如果不保存密码信息，则该值为 False 。 Boolean 类型，可读写。                                                                                                                      |
|  34  | [Sort](#QueryTable.Sort)                                                   | 返回查询表范围的排序条件。只读。                                                                                                                                                                                                                    |
|  35  | [SourceConnectionFile](#QueryTable.SourceConnectionFile)                   | 返回或设置一个 String 值，它指明用于创建查询表的 Microsoft Office 数据连接文件或类似文件。可读写。                                                                                                                                                  |
|  36  | [SourceDataFile](#QueryTable.SourceDataFile)                               | 返回或设置一个 String 值，它指明查询表的源数据文件。                                                                                                                                                                                                |
|  37  | [TextFileColumnDataTypes](#QueryTable.TextFileColumnDataTypes)             | 返回或设置一个有序的常量数组，用于指定文本文件中相应列的数据类型，而该文本文件则是正要导入查询表中的文本文件。每一列的默认常量为 xlGeneral 。 Variant 类型，可读写。                                                                                |
|  38  | [TextFileCommaDelimiter](#QueryTable.TextFileCommaDelimiter)               | 如果将文本文件导入查询表中时，以逗号作为分隔符，则该值为 True 。如果以其他字符作为分隔符，则该值为 False 。默认值为 False 。 Boolean 类型，可读写。                                                                                                 |
|  39  | [TextFileDecimalSeparator](#QueryTable.TextFileDecimalSeparator)           | 返回或设置小数分隔符，在将文本文件导入查询表中时，ET 将使用小数分隔符。默认值为系统小数分隔符。 String 类型，可读写。                                                                                                                               |
|  40  | [TextFileFixedColumnWidths](#QueryTable.TextFileFixedColumnWidths)         | 返回或设置一个整数数组，该数组对应于正要向查询表中导入的文本文件的列宽（按字符）。有效的宽度为 1 到 32767 个字符。 Variant 类型，可读写。                                                                                                           |
|  41  | [TextFileOtherDelimiter](#QueryTable.TextFileOtherDelimiter)               | 返回或设置在向查询表中导入文本文件时用作分隔符的字符。默认值为 null 。 String 类型，可读写。                                                                                                                                                        |
|  42  | [TextFileParseType](#QueryTable.TextFileParseType)                         | 返回或设置要导入查询表的文本文件中数据的列格式。 XlTextParsingType 类型，可读写。                                                                                                                                                                   |
|  43  | [TextFilePlatform](#QueryTable.TextFilePlatform)                           | 返回或设置正向查询表中导入的文本文件的原始格式。该属性确定在数据导入过程中使用何种代码页。 XlPlatform 类型，可读写。                                                                                                                                |
|  44  | [TextFilePromptOnRefresh](#QueryTable.TextFilePromptOnRefresh)             | 如果每次刷新查询表时都要指定导入文本文件的名称，则该属性值为 True 。 “导入文本文件” 对话框允许用户指定路径和文件名。默认值为 False 。 Boolean 类型，可读写。                                                                                        |
|  45  | [TextFileSemicolonDelimiter](#QueryTable.TextFileSemicolonDelimiter)       | 如果在将文本文件导入查询表时使用分号作为分隔符，并且 TextFileParseType 属性的值为 xlDelimited ，则为 True 。默认值为 False 。 Boolean 类型，可读写。                                                                                                |
|  46  | [TextFileSpaceDelimiter](#QueryTable.TextFileSpaceDelimiter)               | 如果向查询表中导入文本文件时，使用空格字符作为分隔符，则该值为 True 。默认值为 False 。 Boolean 类型，可读写。                                                                                                                                      |
|  47  | [TextFileStartRow](#QueryTable.TextFileStartRow)                           | 返回或设置向查询表中导入文本文件时进行文本分列的起始行号。其有效值为 1 到 32767 之间的整数。默认值为 1。 Long 类型，可读写。                                                                                                                        |
|  48  | [TextFileTabDelimiter](#QueryTable.TextFileTabDelimiter)                   | 如果向查询表中导入文本文件时使用 Tab 作为分隔符，则该值为 True 。默认值为 False 。 Boolean 类型，可读写。                                                                                                                                           |
|  49  | [TextFileTextQualifier](#QueryTable.TextFileTextQualifier)                 | 返回或设置向查询表中导入文本文件时的文本识别符。文本识别符用于指定包含的数据是文本格式。 XlTextQualifier 类型，可读写。                                                                                                                             |
|  50  | [TextFileThousandsSeparator](#QueryTable.TextFileThousandsSeparator)       | 返回或设置向查询表中导入文本文件时 ET 所使用的千位分隔符。默认为系统千位分隔符。 String 类型，可读写。                                                                                                                                              |
|  51  | [TextFileTrailingMinusNumbers](#QueryTable.TextFileTrailingMinusNumbers)   | 如果为 True ，则表示 ET 将导入的数字作为以“-”符号开头的文本，“-”符号为负号。如果为 False ，则表示 ET 将导入的数字作为以“-”符号开头的文本，“-”符号为文本。 Boolean 类型，可读写。                                                                    |
|  52  | [TextFileVisualLayout](#QueryTable.TextFileVisualLayout)                   | 返回或设置一个 XlTextVisualLayoutType 枚举，该枚举表示导入的文本的可见布局是从左到右还是从右到左。                                                                                                                                                  |
|  53  | [WebConsecutiveDelimitersAsOne](#QueryTable.WebConsecutiveDelimitersAsOne) | 当从网页的 HTML \<PRE\> 标记中向查询表导入数据时，如果将连续多个分隔符看作单个分隔符，并且数据将被分列，则该值为 True 。如果将连续多个分隔符看作多个分隔符，则该值为 False 。默认值是 True 。可读/写 Boolean 类型。                                 |
|  54  | [WebDisableDateRecognition](#QueryTable.WebDisableDateRecognition)         | 向查询表中导入网页时，如果将类似日期的数据当作文本进行处理，则该值为 True 。如果使用了日期识别，则该值为 False 。默认值为 False 。 Boolean 类型，可读写。                                                                                           |
|  55  | [WebDisableRedirections](#QueryTable.WebDisableRedirections)               | 如果对 QueryTable 对象禁用 Web 查询重定向，则该属性的值为 True 。默认值为 False 。 Boolean 类型，可读写。                                                                                                                                           |
|  56  | [WebFormatting](#QueryTable.WebFormatting)                                 | 返回或设置一个值，该值确定向查询表中导入网页时网页中应用了多少格式设置（如果有）。 XlWebFormatting 类型，可读写。                                                                                                                                   |
|  57  | [WebPreFormattedTextToColumns](#QueryTable.WebPreFormattedTextToColumns)   | 返回或设置向查询表中导入网页时，是否对网页 HTML \<PRE\> 标记内的数据进行分列。默认值为 True 。 Boolean 类型，可读写。                                                                                                                               |
|  58  | [WebSelectionType](#QueryTable.WebSelectionType)                           | 返回或设置一个值，该值决定是向查询表中导入整个网页、网页上的所有表格还是仅网页上的特定表格。 XlWebSelectionType 类型，可读写。                                                                                                                      |
|  59  | [WebSingleBlockTextImport](#QueryTable.WebSingleBlockTextImport)           | 向查询表中导入网页时，如果位于指定网页的 HTML \<PRE\> 标记中的数据是同时进行处理的，则该值为 True 。如果数据是以连续行的数据块方式导入的，以便能识别标题行，则该值为 False 。默认值为 False 。 Boolean 类型，可读写。                               |
|  60  | [WebTables](#QueryTable.WebTables)                                         | 向查询表中导入网页时，返回或设置由逗号分隔的表格名称或表格索引号的列表。 String 类型，可读写。                                                                                                                                                      |
|  61  | [WorkbookConnection](#QueryTable.WorkbookConnection)                       | 返回查询表所使用的 WorkbookConnection 对象。只读。/p\>                                                                                                                                                                                              |

## 成员方法

### QueryTable.CancelRefresh

取消指定查询表的所有后台查询。使用 **Refreshing** 属性可以确定当前是否正在进行后台查询。

**语法**

**express.CancelRefresh ()**

\*express是一个代表 QueryTable 对象的变量。

**示例**

``` JavaScript
//本示例取消查询表的刷新操作。
function test()
{
    let Qtable = Application.Worksheets.Item(1).QueryTables.Item(1)
    if(Qtable.Refreshing)
    {
      Qtable.CancelRefresh()
    }
}
```

### QueryTable.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 QueryTable 对象的变量。

### QueryTable.Refresh

更新外部数据区域 ( **QueryTable** )。

**语法**

**express.Refresh ( *BackgroundQuery* )**

\*express是一个代表 QueryTable 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                 |
|-------------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *BackgroundQuery* | 可选      | Variant  | 只用于基于 SQL 查询结果的 QueryTables。如果为 True，则在数据库建立连接并提交查询之后，将控制返回给过程。QueryTable 在后台进行更新。如果为 False，则在所有数据被取回到工作表之后，将控制返回给过程。如果没有指定该参数，则由 BackgroundQuery 属性的设置决定查询模式。 |

**返回值**

Boolean

**说明**

下列说明适用于基于 SQL 查询结果的 **QueryTable** 对象。

**Refresh** 方法使 ET 连接到 **QueryTable** 对象的数据源，执行 SQL 查询，并将数据返回到基于 **QueryTable** 对象的区域。除非调用该方法，否则 **QueryTable** 对象不与数据源通信。

当与 OLE DB 或 ODBC 数据源建立连接时，ET 使用由 **Connection** 属性指定的连接字符串。如果指定的连接字符串缺少必需的值，将显示对话框，提示用户提供必需的信息。如果 **DisplayAlerts** 属性为 **False** ，则不显示对话框， **Refresh** 方法失败并导致“连接信息无效”异常。

在 ET 建立一个成功的连接之后，将存储完整的连接字符串，这样，以后在同一编辑会话中调用 **Refresh** 方法时就不会再显示提示。通过检查 **Connection** 属性的值可以获得完整的连接字符串。

完成数据库连接后，将检查 SQL 查询的有效性。如果该查询无效， **Refresh** 方法将失败并导致“SQL 语法错误”异常。

如果查询需要参数，则必须在调用 **Refresh** 方法之前，用参数绑定信息初始化 **Parameters** 集合。如果未绑定足够的参数， **Refresh** 方法将失败并导致“参数错误”异常。如果将参数设置为提示用户输入参数值，则无论 **DisplayAlerts** 属性的设置如何，都会向用户显示对话框。如果用户取消参数对话框， **Refresh** 将停止并返回 **False** 。如果对 **Parameters** 绑定了额外的参数，则这些额外参数将被忽略。

如果成功地完成或启动查询，则 **Refresh** 方法返回 **True** ；如果用户取消连接或参数对话框，该方法返回 **False** 。

要查看取回的行数是否超过了工作表中的可用行数，请检查 **FetchedRowOverflow** 属性。每次调用 **Refresh** 方法之前，该属性都将被初始化。

### QueryTable.ResetTimer

重新设置指定的查询表或数据透视表的刷新计时器，使其时间间隔等于上次使用 **RefreshPeriod** 属性设置的时间间隔。

**语法**

**express.ResetTimer ()**

\*express是一个代表 QueryTable 对象的变量。

**示例**

``` JavaScript
//此示例重新设置当前活动工作表上第一张查询表的刷新计时器。
Application.ActiveSheet.QueryTables.Item(1).ResetTimer()
```

### QueryTable.SaveAsODC

将查询表缓存的源保存为“Microsoft Office 数据连接”文件。

**语法**

**express.SaveAsODC ( *ODCFileName* , *Description* , *Keywords* )**

\*express是一个代表 QueryTable 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                   |
|---------------|-----------|----------|----------------------------------------|
| *ODCFileName* | 必选      | String   | 保存文件的位置。                       |
| *Description* | 可选      | Variant  | 要保存到文件中的说明。                 |
| *Keywords*    | 可选      | Variant  | 可用于搜索该文件的以空格分隔的关键字。 |

## 成员属性

### QueryTable.AdjustColumnWidth

如果每次刷新指定的查询表时列宽都会自动调整为最适合的宽度，则为 **True** 。如果每次刷新时列宽不进行自动调整，则为 **False** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AdjustColumnWidth**

\*express是一个代表 QueryTable 对象的变量。

**说明**

最大列宽为屏幕宽度的三分之二。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **AdjustColumnWidth** 属性。

**示例**

``` JavaScript
//此示例关闭对新增查询表（第一个工作簿中的第一张工作表上）的列宽自动调整。
function test()
{
    let qyTab = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Add(varDBConnStr,Range("B1"),"Select Price From CurrentStocks Where Symbol = 'MSFT'")
    qyTab.AdjustColumnWidth = false
    qyTab.Refresh()
}
```

### QueryTable.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象（可对 OLE 自动操作对象使用此属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Application** 属性。

**示例**

``` JavaScript
//本示例显示一条有关创建 myObject 的应用程序的消息。
function test()
{
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") 
    {
    alert("This is an ET Application object.")
    }
    else 
    {
    alert("This is not an ET Application object.")
    }
}
```

### QueryTable.BackgroundQuery

如果查询表的查询是异步执行（在后台执行）的，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.BackgroundQuery**

\*express是一个代表 QueryTable 对象的变量。

**说明**

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性是只读的，并且始终返回 **False** 。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **BackgroundQuery** 属性。

### QueryTable.CommandText

返回或设置指定数据源的命令串。 **Variant** 型，可读写。

**语法**

**express.CommandText**

\*express是一个代表 QueryTable 对象的变量。

**说明**

对于 OLE DB 源， **CommandType** 属性描述 **CommandText** 属性的值。

对于 ODBC 源，设置 **CommandText** 将导致刷新数据。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **CommandText** 属性。

包含相应查询表的工作表必须为活动状态，才可访问此属性。

**示例**

``` JavaScript
//此示例设置第一张查询表的 ODBC 数据源的命令串。注意，该命令串是一个 SQL 语句。
function test()
{
    let qtQtrResults = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)
    qtQtrResults.CommandType = xlCmdSql
    qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10"
    qtQtrResults.Refresh()
}
```

### QueryTable.CommandType

返回或设置备注部分的下表中列出的 **XlCmdType** 常量之一。返回或设置的常量用于描述 **CommandText** 属性的值。默认值为 **xlCmdSQL** 。 **XlCmdType** 类型，可读写。

**语法**

**express.CommandType**

\*express是一个代表 QueryTable 对象的变量。

**说明**

|                                                                                                                                                                                                                                                                                                                                                                                                 |
|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **XlCmdType** 可为以下这些 **XlCmdType** 常量之一。                                                                                                                                                                                                                                                                                                                                             |
| **xlCmdCube** 。包含 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的 多维数据集?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。） 名称。 |
| **xlCmdDefault** 。包含 OLE DB 提供者可识别的命令文本。                                                                                                                                                                                                                                                                                                                                         |
| **xlCmdSql** 。包含一个 SQL 语句。                                                                                                                                                                                                                                                                                                                                                              |
| **xlCmdTable** 。包含用于访问 OLE DB 数据源的表名称。                                                                                                                                                                                                                                                                                                                                           |

只有当查询表或数据透视表缓存的 **QueryType** 属性值为 **xlOLEDBQuery** 时，才可以设置 **CommandType** 属性。

当 **CommandType** 属性的值为 **xlCmdCube** 时，如果没有与查询表相关联的数据透视表，则不能更改该值。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **CommandType** 属性。

**示例**

``` JavaScript
//此示例设置第一张查询表的 ODBC 数据源的命令串。注意，该命令串是一个 SQL 语句。
function test()
{
    let qtQtrResults = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)
    qtQtrResults.CommandType = xlCmdSql
    qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10"
    qtQtrResults.Refresh()
}
```

### QueryTable.Connection

返回或设置包含下列某项的字符串：允许 ET 连接到 OLE DB 数据源的 OLE DB 设置；允许 ET 连接到 ODBC 数据源的 ODBC 设置；允许 ET 连接到 Web 数据源的 URL；或者文本文件的名称或路径，或是指定某个数据库或 Web 查询的文件名称或路径。 **Variant** 型，可读写。

**语法**

**express.Connection**

\*express是一个代表 QueryTable 对象的变量。

**说明**

设置 **Connection** 属性不会立即与数据源建立连接。必须使用 **Refresh** 方法建立连接并检索数据。

有关连接字符串语法的详细信息，请参阅 **QueryTables** 集合的 **Add** 方法。

另外，也可以通过选择 Microsoft ActiveX 数据对象 (ADO) 库直接访问数据源。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Connection** 属性。

**示例**

``` JavaScript
//此示例为第一张工作表上的第一个查询表提供新的 ODBC 连接信息。
Application.Worksheets.Item(1).QueryTables.Item(1).Connection = "ODBC;DSN=96SalesData;UID=Rep21;PWD=NUyHwYQI;"

//此示例指定一个文本文件。
Application.Worksheets.Item(1).QueryTables.Item(1).Connection = "TEXT;C:\\My Documents\\19980331.txt"
```

### QueryTable.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据则将作为 **ListObject** 对象导入。您可以使用 **ListObject** 的 **QueryTable** 属性访问 **Creator** 属性。

### QueryTable.Destination

返回查询表目标区域（查询结果表放置的区域）的左上角单元格。目标区域必须位于包含 **QueryTable** 对象的工作表中。 **Range** 类型，只读。

**语法**

**express.Destination**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Destination** 属性。

**示例**

``` JavaScript
//本示例滚动活动窗口，将第一张查询表的左上角单元格滚动至窗口的左上角。
function test()
{
    let addr = Application.Worksheets.Item(1).QueryTables.Item(1).Destination
    Application.ActiveWindow.ScrollColumn = addr.Column
    Application.ActiveWindow.ScrollRow = addr.Row
}
```

### QueryTable.EditWebPage

返回或设置用于 Web 查询的网页的 统一资源定位符 (URL) 。 **Variant** 类型，可读写。

**语法**

**express.EditWebPage**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果 **EditWebPage** 属性没有设置，则返回 **null** 。只有当查询类型为 Web 或 OLE 时， **EditWebPage** 属性才有意义。

如果 **EditWebPage** 不是空值，则忽略用于刷新的 **WebTables** 属性。因此，XML 查询和 **WebTable** 属性将引用原始网页中的表，并且只用于编辑情况以预先填充 **“Web 查询”** 对话框。

如果使用用户界面导入数据，则会将来自 Web 查询或文本查询的数据作为 **QueryTable** 对象导入，而将所有其他外部数据作为 **ListObject** 对象导入。

如果使用对象模型导入数据，则必须将来自 Web 查询或文本查询的数据作为 **QueryTable** 导入，同时可将所有其他外部数据作为 **ListObject** 或 **QueryTable** 导入。

**EditWebPage** 属性只应用于 **QueryTable** 对象。

**示例**

``` JavaScript
//在本示例中，ET 将一个网页的 URL 显示给用户。本示例假定单元格 A1 中的 QueryTable 对象位于活动的工作表上，并且名为“MyhomePage.htm”的文件位于 C: 驱动器上。
function test()
{
    //Set the EditWebPage property to a source.
    Range("A1").QueryTable.EditWebPage = "C:\\MyHomepage.htm"

    //Display the source to the user.
    alert(Range("A1").QueryTable.EditWebPage)
}
```

### QueryTable.EnableEditing

如果允许用户对指定查询表进行编辑，则该值为 **True** ；如果用户只能刷新查询表，则该值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableEditing**

\*express是一个代表 QueryTable 对象的变量。

**说明**

本示例对第一张查询表进行设置，使用户不能对其进行编辑。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **EnableEditing** 属性。

**示例**

``` JavaScript
Application.Worksheets.Item(1).QueryTables.Item(1).EnableEditing = false
```

### QueryTable.EnableRefresh

如果用户可刷新数据透视表高速缓存或查询表，则为 **True** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableRefresh**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果 **EnableRefresh** 属性设置为 **False** ，则忽略 **RefreshOnFileOpen** 属性。

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，将此属性设置为 **False** 会禁用更新。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **EnableRefresh** 属性。

### QueryTable.FetchedRowOverflow

如果上次使用 **Refresh** 方法返回的行数比工作表中可用行数大，则该值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.FetchedRowOverflow**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **FetchedRowOverflow** 属性。

**示例**

``` JavaScript
//本示例对第一张查询表进行刷新。如果由查询返回的行数超过工作表中的可用行数，则显示一条错误消息。
function test()
{
    let Qtable = Application.Worksheets.Item(1).QueryTables.Item(1)
    Qtable.Refresh()
    if(Qtable.FetchedRowOverflow) 
    {
       alert("Query too large: please redefine.")
    }
}
```

### QueryTable.FieldNames

如果数据源的字段名称作为返回数据的列标题显示，则该值为 **True** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.FieldNames**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**FieldNames** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例对第一张查询表进行设置，使表中不显示字段名称。
Application.Worksheets.Item(1).QueryTables.Item(1).FieldNames = false
```

### QueryTable.FillAdjacentFormulas

如果数据源的字段名称作为返回数据的列标题显示，则该值为 **True** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.FillAdjacentFormulas**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**FillAdjacentFormulas** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例对第一张查询表进行设置，每当查询表刷新时就自动更新右侧的公式。
Application.Sheets.Item("sheet1").QueryTables.Item(1).FillAdjacentFormulas = true
```

### QueryTable.ListObject

为 **QueryTable** 对象返回一个 **ListObject** 对象。 **ListObject** 对象类型，只读。

**语法**

**express.ListObject**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**ListObject** 属性仅适用于 **ListObject** 对象。

### QueryTable.MaintainConnection

如果从刷新数据开始直至关闭工作簿，都一直保留指向指定数据源的连接，则为 **True** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MaintainConnection**

\*express是一个代表 QueryTable 对象的变量。

**说明**

只有当查询表或数据透视表缓存的 **QueryType** 属性值为 **xlOLEDBQuery** 时，才可以设置 **MaintainConnection** 属性。

如果预计会频繁对服务器进行查询，则可将此属性设置为 **True** ，这样能减少重新连接的时间因而可提高性能。将此属性设置为 **False** ，将会关闭一个打开的连接。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **MaintainConnection** 属性。

**示例**

``` JavaScript
//本示例在活动工作表的 A3 单元格上新建一个基于 OLAP 提供程序
//（OLAP 提供程序：对特定类型的 OLAP 数据库提供访问功能的一组软件。该软件包括数据源驱动程序以及与数据库连接所必需的其他客户端软件。）
//的数据透视表缓存，然后基于该缓存新建一个数据透视表。本示例将在初始刷新后终止连接。
function test()
{
    let PTcache = Application.ActiveWorkbook.PivotCaches().Add(xlExternal)
    PTcache.Connection = "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National"
    PTcache.MaintainConnection = false
    PTcache.CreatePivotTable(Range("A3"),"PivotTable1")

    let PT = Application.ActiveSheet.PivotTables("PivotTable1") 
    PT.SmallGrid = false
    PT.PivotCache.RefreshPeriod = 0
    PT.CubeFields.Item("[state]").Orientation = xlColumnField
    PT.CubeFields.Item("[state]").Position = 0
    PT.CubeFields.Item("[Measures].[Count Of au_id]").Orientation = xlDataField
    PT.CubeFields.Item("[Measures].[Count Of au_id]").Position = 0
}
```

### QueryTable.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 QueryTable 对象的变量。

**说明**

下表列出了 **Name** 属性和相关属性的示例值，同时假设有一个具有唯一名称“\[Europe\].\[France\].\[Paris\]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，还有一个包含项名称“Paris”的非 OLAP 数据源。

| 属性       | 值（OLAP 数据源）                       | 值（非 OLAP 数据源）        |
|------------|-----------------------------------------|-----------------------------|
| **题注**   | Paris                                   | Paris                       |
| **名称**   | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |
| **源名称** | \[Europe\].\[France\].\[Paris\]（只读） | （与 SQL 属性值相同，只读） |
| **值**     | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**Name** 属性仅适用于 **QueryTable** 对象。

### QueryTable.Parameters

返回一个 **Parameters** 集合，该集合表示查询表参数。只读。

**语法**

**express.Parameters**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Parameters** 属性。

**示例**

``` JavaScript
//本示例从现有带参数查询中返回 Parameters 集合。如果第一个参数使用字符数据类型，则用户只须在提示对话框中输入字符。
function test()
{
    let Qtable = Application.Sheets.Item("Sheet1").QueryTables.Item(1).Parameters.Item(1)
    if(Qtable.DataType == xlParamTypeVarChar) 
    {
       Qtable.SetParam(xlPrompt, "Enter a character only")
    }
}
```

### QueryTable.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Parent** 属性。

### QueryTable.PostText

返回或设置用于 post 方法的字符串，post 方法用于向 Web 服务器输入数据以从 Web 查询中返回数据。 **String** 类型，可读写。

**语法**

**express.PostText**

\*express是一个代表 QueryTable 对象的变量。

**说明**

ET 带有 Web 查询示例，使用 WordPad 或其他文本编辑器可对其中的 HTML 代码进行更改。查询示例在 Microsoft Office 的安装文件夹的“Queries”文件夹中。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**PostText** 属性仅适用于 **QueryTable** 对象。

### QueryTable.PreserveColumnInfo

如果每次刷新查询表时，列排序、筛选和布局信息都会保留，则该值为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.PreserveColumnInfo**

\*express是一个代表 QueryTable 对象的变量。

**说明**

只有当查询表使用数据库连接时，该属性才有效。

可以将该属性设置为 **False** ，以便与 ET 的早期版本兼容。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **PreserveColumnInfo** 属性。

**示例**

``` JavaScript
//本示例保留列排序、筛选和布局信息，以便与 ET 的早期版本兼容。
function test()
{
    let cnnConnect = Application.ADODB.Connection
    cnnConnect.Open "Provider=SQLOLEDB;" & _
    "Data Source=srvdata;" & _
    "User ID=wadet;Password=4me2no;"

    let rstRecordset = Application.ADODB.Recordset
    rstRecordset.Open _
    Source:="Select Name, Quantity, Price From Products", _
    ActiveConnection:=cnnConnect, _
    CursorType:=adOpenDynamic, _
    LockType:=adLockReadOnly, _
    Options:=adCmdText

    let worsheet = Application.ActiveSheet.QueryTables.Add( Connection:=rstRecordset, Destination:=Range("A1"))
    worsheet.Name = "Contact List"
    worsheet.FieldNames = True
    worsheet.RowNumbers = False
    worsheet.FillAdjacentFormulas = False
    worsheet.PreserveFormatting = True
    worsheet.RefreshOnFileOpen = False
    worsheet.BackgroundQuery = True
    worsheet.RefreshStyle = xlInsertDeleteCells
    worsheet.SavePassword = True
    worsheet.SaveData = True
    worsheet.AdjustColumnWidth = True
    worsheet.RefreshPeriod = 0
    worsheet.PreserveColumnInfo = True
    worsheet.Refresh BackgroundQuery:=False

}
```

### QueryTable.PreserveFormatting

如果将数据前五行的任何常用格式设置应用到查询表的新数据行，则为 **True** 。对未使用的单元格不进行格式设置。如果将应用到查询表的最新一次自动套用格式应用于新数据行，则属性为 **False** 。默认值是 **True** 。

**语法**

**express.PreserveFormatting**

\*express是一个代表 QueryTable 对象的变量。

**说明**

对于数据库查询表，默认的格式设置为 **xlSimple** 。

刷新查询表时，将对查询表应用新的自动套用格式样式。只要 **PreserveFormatting** 的值为 **False** ，则 AutoFormat（自动套用格式）就会被设置为 **None** 。这样，任何在 **PreserveFormatting** 被设置为 **False** 或在查询表刷新之前设置的自动套用格式都不会起作用，且相应产生的查询表也不会被应用任何格式。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **PreserveFormatting** 属性。

**示例**

``` JavaScript
//此示例保留第一张工作表上的第一个数据透视表的格式。
Application.Worksheets.Item(1).PivotTables("Pivot1").PreserveFormatting = true

//此示例演示了将 PreserveFormatting 设置为 False 后，如何操作才能将“自动套用格式”设置为 xlRangeAutoFormatNone，而并不是指定的 xlRangeAutoFormatColor1 格式。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)
    Qtable.Range.AutoFormat = xlRangeAutoFormatColor1
    Qtable.PreserveFormatting = false
    Qtable.Refresh()
}
```

### QueryTable.QueryType

表示 ET 填充查询表时所使用的查询类型。 **XlQueryType** 类型，只读。

**语法**

**express.QueryType**

\*express是一个代表 QueryTable 对象的变量。

**说明**

|                                                                                                                                                                                    |
|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **XlQueryType** 可为以下这些 **XlQueryType** 常量之一。                                                                                                                            |
| **xlTextImport** 。基于文本文件，仅用于查询表                                                                                                                                      |
| **xlOLEDBQuery** 。基于 OLE DB 查询，包括 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源 |
| **xlWebQuery** 。基于网页，仅用于查询表                                                                                                                                            |
| **xlADORecordset** 。基于 ADO 记录集查询                                                                                                                                           |
| **xlDAORecordSet** 。基于 DAO 记录集查询，只用于查询表                                                                                                                             |
| **xlODBCQuery** 。基于 ODBC 数据源                                                                                                                                                 |

在 **Connection** 属性值的前缀中指定数据源。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **QueryType** 属性。

**示例**

``` JavaScript
//如果表是基于网页的，则本示例将刷新第一张工作表上的第一个查询表。
function test()
{
    let qtQtrResults = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)
    if(qtQtrResults.QueryType == xlWebQuery)
    {
       qtQtrResults.Refresh()
    }
}
```

### QueryTable.Recordset

返回或设置一个 **Recordset** 对象，该对象用作指定查询表的数据源。

**语法**

**express.Recordset**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用此属性来覆盖现有记录集，则更改将在运行 **Refresh** 方法时生效。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RecordSet** 属性。

**示例**

``` JavaScript
//此示例对用于第一张查询表（第一张工作表上）的 Recordset 对象进行更改，然后刷新该查询表。
function test()
{
    Application.Worksheets.Item(1).QueryTables.Item(1).Recordset = Workbooks.OpenDatabase("C:\\Nwind.mdb").OpenRecordset("employees")
    Application.Worksheets.Item(1).QueryTables.Item(1).Refresh()
}
```

### QueryTable.Refreshing

如果指定的查询表正在进行后台查询，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.Refreshing**

\*express是一个代表 QueryTable 对象的变量。

**说明**

使用 **CancelRefresh** 方法可以取消后台查询。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Refreshing** 属性。

**示例**

``` JavaScript
//本示例在查询表一正在进行后台查询的情况下显示了一条信息。
function test()
{
    if(Application.Worksheets.Item(1).QueryTables.Item(1).Refreshing)
    {
        alert("Query is currently refreshing: please wait")
    }
    else
    {
        Application.Worksheets.Item(1).QueryTables.Item(1).Refresh(false)
        Application.Worksheets.Item(1).QueryTables.Item(1).ResultRange.Select()
    }
}
```

### QueryTable.RefreshOnFileOpen

如果每次打开工作簿时，数据透视表高速缓存或查询表自动更新，则为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.RefreshOnFileOpen**

\*express是一个代表 QueryTable 对象的变量。

**说明**

在 Visual Basic 中使用 **Open** 方法打开工作簿时，不会自动刷新查询表和数据透视表。请在打开工作簿后，使用 **Refresh** 方法刷新数据。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RefreshOnFileOpen** 属性。

### QueryTable.RefreshPeriod

返回或设置两次刷新之间的时间间隔。 **Long** 型，可读写。

**语法**

**express.RefreshPeriod**

\*express是一个代表 QueryTable 对象的变量。

**说明**

将周期设置为 0（零），则会禁用自动定时刷新，并且等同于将该属性设置为 **Null** 。

**RefreshPeriod** 属性的值可以是一个 0 到 32767 之间的整数。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RefreshPeriod** 属性。

### QueryTable.RefreshStyle

返回或设置为了容纳查询返回的记录集中的行数而在指定工作表中插入或删除行时所使用的方式。 **XlCellInsertionMode** 类型。可读写。

**语法**

**express.RefreshStyle**

\*express是一个代表 QueryTable 对象的变量。

**说明**

|                                                                                     |
|-------------------------------------------------------------------------------------|
| **XlCellInsertionMode** 可为以下 **XlCellInsertionMode** 常量之一。                 |
| **xlInsertDeleteCells** 插入或者删除部分行以适应新记录集所需要的实际行数。          |
| **xlOverwriteCells** 不向工作表添加新的单元格或行。如果溢出则覆盖周围单元格的数据。 |
| **xlInsertEntireRows** 必要时插入所有行以允许溢出。不从工作表删除单元格或行。       |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RefreshStyle** 属性。

**示例**

``` JavaScript
//本示例向 Sheet1 添加查询表，并对 RefreshStyle 属性进行设置，插入两个新行以放置查询结果数据。
function test()
{
    let Qtable = Application.Sheets.Item("sheet1").QueryTables.Add("Finder;c:\\myfile.dqy",Range("sheet1!a1"))
    Qtable.RefreshStyle = xlInsertEntireRows
    Qtable.Refresh()
}
```

### QueryTable.ResultRange

返回一个 **Range** 对象，该对象代表指定查询表所覆盖的工作表区域。只读。

**语法**

**express.ResultRange**

\*express是一个代表 QueryTable 对象的变量。

**说明**

该区域不包括字段名称行或行号列。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **ResultRange** 属性。

**示例**

``` JavaScript
//本示例对查询表一中的第一列数据进行汇总，并在数据区域下方显示第一列数据的总和。
function test()
{
    let Qtable = Application.Sheets.Item("sheet1").QueryTables.Item(1).ResultRange.Columns.Item(1)
    Qtable.Name = "Column1"
    Qtable.End(xlDown).Offset(2, 0).Formula = "=sum(Column1)"
}
```

### QueryTable.RobustConnect

返回或设置数据透视表缓存与其数据源连接的方式。 **XlRobustConnect** 类型，可读写。

**语法**

**express.RobustConnect**

\*express是一个代表 QueryTable 对象的变量。

**说明**

|                                                                                                                   |
|-------------------------------------------------------------------------------------------------------------------|
| **XlRobustConnect** 可为以下这些 **XlRobustConnect** 常量之一。                                                   |
| **xlAlways** 。缓存始终使用外部源信息（由 **SourceConnectionFile** 或 **SourceDataFile** 属性定义）进行重新连接。 |
| **xlAsRequired** 。缓存通过 **Connection** 属性使用外部源信息进行重新连接。                                       |
| **xlNever** 。缓存从不使用源信息进行重新连接。                                                                    |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RobustConnect** 属性。

### QueryTable.RowNumbers

如果行号作为第一列添加到指定查询表，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.RowNumbers**

\*express是一个代表 QueryTable 对象的变量。

**说明**

设置该属性为 **True** 并不会使行号立即显示。在查询表下一次刷新时行号将显示，而且在查询表每次刷新时将重新配置这些行号。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RowNumbers** 属性。

**示例**

``` JavaScript
//本示例为查询表添加行号和字段名。
function test()
{
    let Qtable = Application.Worksheets.Item(1).QueryTables.Item("ExternalData1")
    Qtable.RowNumbers = true
    Qtable.FieldNames = true
    Qtable.Refresh()
}
```

### QueryTable.SaveData

如果将数据透视表的数据随工作簿一起保存，则为 **True** 。如果仅保存数据透视表的定义，则为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.SaveData**

\*express是一个代表 QueryTable 对象的变量。

**说明**

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性始终设置为 **False** 。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **SaveData** 属性。

### QueryTable.SavePassword

如果将 ODBC 连接字符串中的密码信息与指定查询一起保存，则为 **True** 。如果不保存密码信息，则该值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.SavePassword**

\*express是一个代表 QueryTable 对象的变量。

**说明**

此属性由数据透视表和查询表在 ODBC 和 OLEDB 查询中使用。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **SavePassword** 属性。

如果 **ListObject** 已连接到 SharePoint 列表，此属性将被忽略。

**示例**

``` JavaScript
//此示例在每次保存查询表一时，都不保存 ODBC 连接字符串的密码信息。
Application.Worksheets.Item(1).QueryTables.Item(1).SavePassword = false
```

### QueryTable.Sort

返回查询表范围的排序条件。只读。

**语法**

**express.Sort**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将导入为 **QueryTable** 对象，而所有其他数据将导入为 **ListObject** object.

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Sort** 属性。

**示例**

``` JavaScript
//下面的示例将刷新查询表，并获取排序条件。
function test()
{
    Application.QueryTable.Refresh()
    let Qtable = Application.QueryTable.Sort
}
```

### QueryTable.SourceConnectionFile

返回或设置一个 **String** 值，它指明用于创建查询表的 Microsoft Office 数据连接文件或类似文件。可读写。

**语法**

**express.SourceConnectionFile**

\*express是一个代表 QueryTable 对象的变量。

**说明**

来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。可以使用 **ListObject** 的 **QueryTable** 属性访问 **SourceConnectionFile** 属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **SourceConnectionFile** 属性。

### QueryTable.SourceDataFile

返回或设置一个 **String** 值，它指明查询表的源数据文件。

**语法**

**express.SourceDataFile**

\*express是一个代表 QueryTable 对象的变量。

**说明**

对于基于文件的数据源（如 Access）， **SourceDataFile** 属性包含源数据文件的完全限定路径。对于基于服务器的数据源（如 SQL Server），此属性设置为 **Null** 。如果通过编程方式更改 **Connection** 属性，则 **SourceDataFile** 属性设置为 **Null** 。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **SourceDataFile** 属性。

### QueryTable.TextFileColumnDataTypes

返回或设置一个有序的常量数组，用于指定文本文件中相应列的数据类型，而该文本文件则是正要导入查询表中的文本文件。每一列的默认常量为 **xlGeneral** 。 **Variant** 类型，可读写。

**语法**

**express.TextFileColumnDataTypes**

\*express是一个代表 QueryTable 对象的变量。

**说明**

可以使用下表列出的 **xlColumnDataType** 常量来指定在数据导入过程中所使用的列数据类型或执行的操作。

| 常量                | 说明               |
|---------------------|--------------------|
| **xlGeneralFormat** | 概要               |
| **xlTextFormat**    | 文本               |
| **xlSkipColumn**    | 跳过列             |
| **xlDMYFormat**     | “日-月-年”日期格式 |
| **xlDYMFormat**     | “日-年-月”日期格式 |
| **xlEMDFormat**     | EMD 日期           |
| **xlMDYFormat**     | “月-日-年”日期格式 |
| **xlMYDFormat**     | “月-年-日”日期格式 |
| **xlYDMFormat**     | “年-日-月”日期格式 |
| **xlYMDFormat**     | “年-月-日”日期格式 |

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）时，才使用此属性。

如果在含有多列的数组中指定了多个元素，则将忽略那些值。

只有在安装并选定了中国台湾地区语言时才可以使用 **xlEMDFormat** 。 **xlEMDFormat** 常量指定使用中国台湾地区纪元日期。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileColumnDataTypes** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例向第一个工作簿中第一张工作表上的新查询表中导入一个固定宽度的文本文件。该文本文件的第一列为五个字符宽度，作为文本导入。第二列为四个字符宽度，被跳过。该文本文件的剩余部分被导入第三列，并对其应用常规格式。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlFixedWidth
    Qtable1.TextFileFixedColumnWidths = [5, 4]
    Qtable1.TextFileColumnDataTypes = [xlTextFormat, xlSkipColumn, xlGeneralFormat]
    Qtable1.Refresh()
}
```

### QueryTable.TextFileCommaDelimiter

如果将文本文件导入查询表中时，以逗号作为分隔符，则该值为 **True** 。如果以其他字符作为分隔符，则该值为 **False** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.TextFileCommaDelimiter**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）并且 **TextFileParseType** 属性的值为 **xlDelimited** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileCommaDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为逗号，然后刷新该查询表。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlDelimited
    Qtable1.TextFileCommaDelimiter = true
    Qtable1.Refresh()
}
```

### QueryTable.TextFileDecimalSeparator

返回或设置小数分隔符，在将文本文件导入查询表中时，ET 将使用小数分隔符。默认值为系统小数分隔符。 **String** 类型，可读写。

**语法**

**express.TextFileDecimalSeparator**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）并且该文件包含的小数位分隔符和千位分隔符与计算机上使用的不同（由于使用不同的语言设置）时，才使用此属性。

下表显示当使用不同分隔符向 ET 中导入文本时得到的不同结果。数字结果显示在最右边的列中。

| 系统小数分隔符 | 系统千位分隔符 | TextFileDecimalSeparator 值 | TextFileThousandsSeparator 值 | 导入的文本 | 单元格的值（数据类型） |
|----------------|----------------|-----------------------------|-------------------------------|------------|------------------------|
| 句号           | 逗号           | 逗号                        | 句号                          | 123.123,45 | 123,123.45（数字）     |
| 句号           | 逗号           | 逗号                        | 逗号                          | 123.123,45 | 123.123,45（文本）     |
| 逗号           | 句号           | 逗号                        | 句号                          | 123,123.45 | 123,123.45（数字）     |
| 句号           | 逗号           | 句号                        | 逗号                          | 123 123.45 | 123 123.45（文本）     |
| 句号           | 逗号           | 句号                        | 空格                          | 123 123.45 | 123,123.45（数字）     |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileDecimalSeparator** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例为工作表 Sheet1 上第一份查询表保存最初的小数分隔符，并将其设置为逗号，以准备将一个法语文本文件（举例）导入 ET 的美国英语版中。
function test()
{
    let Qtable = Application.Worksheets.Item("Sheet1").QueryTables.Item(1).TextFileDecimalSeparator
    Application.Worksheets.Item("Sheet1").QueryTables.Item(1).TextFileDecimalSeparator = ","
}
```

### QueryTable.TextFileFixedColumnWidths

返回或设置一个整数数组，该数组对应于正要向查询表中导入的文本文件的列宽（按字符）。有效的宽度为 1 到 32767 个字符。 **Variant** 类型，可读写。

**语法**

**express.TextFileFixedColumnWidths**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）并且 **TextFileParseType** 属性的值为 **xlFixedWidth** 时，才使用此属性。

必须指定一个有效的非负值列宽。如果指定的列超出了文本文件的宽度，则会忽略那些值。如果文本文件的宽度大于指定的列的总宽度，则文本文件的剩余部分将会导入到一个附加列中。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileFixedColumnWidths** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例向第一个工作簿中第一张工作表上的新查询表中导入一个固定宽度的文本文件。该文本文件的第一列为五个字符宽度，作为文本导入。第二列为四个字符宽度，被跳过。该文本文件的剩余部分被导入第三列，并对其应用常规格式。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlFixedWidth
    Qtable1.TextFileFixedColumnWidths = [5, 4]
    Qtable1.TextFileColumnDataTypes = [xlTextFormat, xlSkipColumn, xlGeneralFormat]
    Qtable1.Refresh()
}
```

### QueryTable.TextFileOtherDelimiter

返回或设置在向查询表中导入文本文件时用作分隔符的字符。默认值为 **null** 。 **String** 类型，可读写。

**语法**

**express.TextFileOtherDelimiter**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）并且 **TextFileParseType** 属性的值为 **xlDelimited** 时，才使用此属性。

如果在字符串中指定了多个字符，则只使用第一个字符。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileOtherDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为井号 (#)，然后刷新该查询表。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlDelimited
    Qtable1.TextFileOtherDelimiter = "#"
    Qtable1.Refresh()
}
```

### QueryTable.TextFileParseType

返回或设置要导入查询表的文本文件中数据的列格式。 **XlTextParsingType** 类型，可读写。

**语法**

**express.TextFileParseType**

\*express是一个代表 QueryTable 对象的变量。

**说明**

|                                                                 |
|-----------------------------------------------------------------|
| **XlTextParsingType** 可为以下 **XlTextParsingType** 常量之一。 |
| **xlFixedWidth** 表示将文件中的数据安排在固定宽度的列中。       |
| **xlDelimited** 默认值表示文件由分隔符分隔。                    |

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileParseType** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例向第一个工作簿中第一张工作表上的新查询表中导入一个固定宽度的文本文件。该文本文件的第一列为五个字符宽度，作为文本导入。第二列为四个字符宽度，被跳过。该文本文件的剩余部分被导入第三列，并对其应用常规格式。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlFixedWidth
    Qtable1.TextFileFixedColumnWidths = [5, 4]
    Qtable1.TextFileColumnDataTypes = [xlTextFormat, xlSkipColumn, xlGeneralFormat]
    Qtable1.Refresh()
}
```

### QueryTable.TextFilePlatform

返回或设置正向查询表中导入的文本文件的原始格式。该属性确定在数据导入过程中使用何种代码页。 **XlPlatform** 类型，可读写。

**语法**

**express.TextFilePlatform**

\*express是一个代表 QueryTable 对象的变量。

**说明**

默认值是在 **“文本文件导入向导”** 的 **“文件原始格式”** 选项中的当前设置。

|                                                   |
|---------------------------------------------------|
| **XlPlatform** 可为以下 **XlPlatform** 常量之一。 |
| **xlMacintosh**                                   |
| **xlMSDOS**                                       |
| **xlWindows**                                     |

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）时才使用该属性。

如果使用用户界面导入数据，则会将来自 Web 查询或文本查询的数据作为 **QueryTable** 对象导入，而将所有其他外部数据作为 **ListObject** 对象导入。

如果使用对象模型导入数据，则必须将来自 Web 查询或文本查询的数据作为 **QueryTable** 导入，同时可将所有其他外部数据作为 **ListObject** 或 **QueryTable** 导入。

**TextFilePlatform** 属性只应用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例向第一个工作簿中第一张工作表上的查询表中导入一个 MS-DOS 文本文件，然后刷新该查询表。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFilePlatform = xlMSDOS
    Qtable1.TextFileParseType = xlDelimited
    Qtable1.TextFileTabDelimiter = true
    Qtable1.Refresh()
}
```

### QueryTable.TextFilePromptOnRefresh

如果每次刷新查询表时都要指定导入文本文件的名称，则该属性值为 **True** 。 **“导入文本文件”** 对话框允许用户指定路径和文件名。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.TextFilePromptOnRefresh**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）时才使用该属性。

如果该属性值为 **True** ，则第一次刷新查询表时，并不显示对话框。

在用户界面中，其默认值为 **True** 。

如果使用用户界面导入数据，则会将来自 Web 查询或文本查询的数据作为 **QueryTable** 对象导入，而将所有其他外部数据作为 **ListObject** 对象导入。

如果使用对象模型导入数据，则必须将来自 Web 查询或文本查询的数据作为 **QueryTable** 导入，同时可将所有其他外部数据作为 **ListObject** 或 **QueryTable** 导入。

**TextFilePromptOnRefresh** 属性只应用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例在每次刷新第一个工作簿中第一张工作表上的查询表时，都提示用户输入文本文件的名称。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlDelimited
    Qtable1.TextFilePromptOnRefresh = true
    Qtable1.TextFileTabDelimiter = true
    Qtable1.Refresh()
}
```

### QueryTable.TextFileSemicolonDelimiter

如果在将文本文件导入查询表时使用分号作为分隔符，并且 **TextFileParseType** 属性的值为 **xlDelimited** ，则为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.TextFileSemicolonDelimiter**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileSemicolonDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为分号，然后刷新该查询表。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlDelimited
    Qtable1.TextFileSemicolonDelimiter = true
    Qtable1.Refresh()
}
```

### QueryTable.TextFileSpaceDelimiter

如果向查询表中导入文本文件时，使用空格字符作为分隔符，则该值为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.TextFileSpaceDelimiter**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）并且 **TextFileParseType** 属性的值为 **xlDelimited** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileSpaceDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例将空格字符设置为第一个工作簿中第一张工作表上的查询表的分隔符，然后刷新该查询表。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlDelimited
    Qtable1.TextFileSemicolonDelimiter = true
    Qtable1.Refresh()
}
```

### QueryTable.TextFileStartRow

返回或设置向查询表中导入文本文件时进行文本分列的起始行号。其有效值为 1 到 32767 之间的整数。默认值为 1。 **Long** 类型，可读写。

**语法**

**express.TextFileStartRow**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileStartRow** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例将第一个工作簿中第一张工作表上查询表的文本分列起始行设置为第五行，然后刷新该查询表。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlDelimited
    Qtable1.TextFileStartRow = 5
    Qtable1.TextFileTabDelimiter = true
    Qtable1.Refresh()
}
```

### QueryTable.TextFileTabDelimiter

如果向查询表中导入文本文件时使用 Tab 作为分隔符，则该值为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.TextFileTabDelimiter**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）并且 **TextFileParseType** 属性的值为 **xlDelimited** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileTabDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为 Tab，然后刷新该查询表。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlDelimited
    Qtable1.TextFileTabDelimiter = true
    Qtable1.Refresh()
}
```

### QueryTable.TextFileTextQualifier

返回或设置向查询表中导入文本文件时的文本识别符。文本识别符用于指定包含的数据是文本格式。 **XlTextQualifier** 类型，可读写。

**语法**

**express.TextFileTextQualifier**

\*express是一个代表 QueryTable 对象的变量。

**说明**

|                                                             |
|-------------------------------------------------------------|
| **XlTextQualifier** 可为以下 **XlTextQualifier** 常量之一。 |
| **xlTextQualifierNone**                                     |
| **xlTextQualifierDoubleQuote** 默认值                       |
| **xlTextQualifierSingleQuote**                              |

仅当查询表基于文本文件数据（ **QueryType** 属性设置为 **xlTextImport** ）时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileTextQualifier** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例将第一个工作簿中第一张工作表上的查询表的文本识别符设置为单引号。
function test()
{
    let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)
    let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))
    Qtable1.TextFileParseType = xlDelimited
    Qtable1.TextFileTextQualifier = xlTextQualifierSingleQuote
    Qtable1.Refresh()
}
```

### QueryTable.TextFileThousandsSeparator

返回或设置向查询表中导入文本文件时 ET 所使用的千位分隔符。默认为系统千位分隔符。 **String** 类型，可读写。

**语法**

**express.TextFileThousandsSeparator**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表基于文本文件数据（ **QueryTable** 属性设置为 **xlTextImport** ）时，尤其是当文件包含的小数位分隔符和千位分隔符与计算机上使用的不同（由于使用不同的语言设置）时，才使用此属性。

下表显示当使用不同分隔符向 ET 中导入文本时得到的不同结果。数字结果显示在最右边的列中。

| 系统小数分隔符 | 系统千位分隔符 | TextFileDecimalSeparator 值 | TextFileThousandsSeparator 值 | 导入的文本 | 单元格的值（数据类型） |
|----------------|----------------|-----------------------------|-------------------------------|------------|------------------------|
| 句号           | 逗号           | 逗号                        | 句号                          | 123.123,45 | 123,123.45（数字）     |
| 句号           | 逗号           | 逗号                        | 逗号                          | 123.123,45 | 123.123,45（文本）     |
| 逗号           | 句号           | 逗号                        | 句号                          | 123,123.45 | 123,123.45（数字）     |
| 句号           | 逗号           | 句号                        | 逗号                          | 123 123.45 | 123 123.45（文本）     |
| 句号           | 逗号           | 句号                        | 空格                          | 123 123.45 | 123,123.45（数字）     |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileThousandsSeparator** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例保存工作表 Sheet1 上第一个查询表最初的千位分隔符，并将其设置为句号，以准备将法语文本文件（举例）导入 ET 的美国英语版中。
function test()
{
    let strDecSep = Application.Worksheets.Item("Sheet1").QueryTables.Item(1).TextFileThousandsSeparator
    Application.Worksheets.Item("Sheet1").QueryTables.Item(1).TextFileThousandsSeparator = "."
}
```

### QueryTable.TextFileTrailingMinusNumbers

如果为 **True** ，则表示 ET 将导入的数字作为以“-”符号开头的文本，“-”符号为负号。如果为 **False** ，则表示 ET 将导入的数字作为以“-”符号开头的文本，“-”符号为文本。 **Boolean** 类型，可读写。

**语法**

**express.TextFileTrailingMinusNumbers**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileTrailingMinusNumbers** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//在本示例中，ET 确定单元格 A1 的设置，然后将导入的数值处理为以“-”符号开头的文本。本示例假定在活动工作表上存在 QueryTable 对象。
function test()
{
// Determine setting for TextFileTrailingMinusNumbers
    if(Range("A1").QueryTable.TextFileTrailingMinusNumbers == true)
    {
        alert("Numbers imported as text that begin with a '-' symbol will be treated as a negative symbol.")
    }
    else 
    {
        alert("Numbers imported as text that begin with a '-' symbol will not be treated as a negative symbol.")
    }
}
```

### QueryTable.TextFileVisualLayout

返回或设置一个 **XlTextVisualLayoutType** 枚举，该枚举表示导入的文本的可见布局是从左到右还是从右到左。

**语法**

**express.TextFileVisualLayout**

\*express是一个代表 QueryTable 对象的变量。

**说明**

|                                                                           |
|---------------------------------------------------------------------------|
| **XlTextVisualLayoutType** 可为以下 **XlTextVisualLayoutType** 常量之一。 |
| **xlTextVisualLTR**                                                       |
| **xlTextVisualRTL**                                                       |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileVisualLayout** 属性仅适用于 **QueryTable** 对象。

### QueryTable.WebConsecutiveDelimitersAsOne

当从网页的 HTML \<PRE\> 标记中向查询表导入数据时，如果将连续多个分隔符看作单个分隔符，并且数据将被分列，则该值为 **True** 。如果将连续多个分隔符看作多个分隔符，则该值为 **False** 。默认值是 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.WebConsecutiveDelimitersAsOne**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** ，查询返回 HTML 文档，并且 **WebPreFormattedTextToColumns** 属性设置为 **True** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebConsecutiveDelimitersAsOne** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为空格字符，然后刷新该查询表。连续多个空格作为单个空格处理。
function test()
{
    let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)
    let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))
    qtQtrResults.WebConsecutiveDelimitersAsOne = true
    qtQtrResults.Refresh()
}
```

### QueryTable.WebDisableDateRecognition

向查询表中导入网页时，如果将类似日期的数据当作文本进行处理，则该值为 **True** 。如果使用了日期识别，则该值为 **False** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.WebDisableDateRecognition**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebDisableDateRecognition** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例关闭日期识别，以便类似于日期的网页数据将作为文本导入。然后刷新该查询表。
function test()
{
    let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1) 
    let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))
    qtQtrResults.WebDisableDateRecognition = true
    qtQtrResults.Refresh()
}
```

### QueryTable.WebDisableRedirections

如果对 **QueryTable** 对象禁用 Web 查询重定向，则该属性的值为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.WebDisableRedirections**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebDisableRedirections** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//在本示例中，ET 确定工作簿中第一张工作表的 Web 查询重定向的设置。本示例假定 QueryTable 对象位于第一张工作表上，否则将出现运行错误。
function test()
{
    let wksSheet = Application.ActiveSheet
    alert(wksSheet.QueryTables.Item(1).WebDisableRedirections)
}
```

### QueryTable.WebFormatting

返回或设置一个值，该值确定向查询表中导入网页时网页中应用了多少格式设置（如果有）。 **XlWebFormatting** 类型，可读写。

**语法**

**express.WebFormatting**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

|                                                     |
|-----------------------------------------------------|
| XlWebFormatting 可为以下 XlWebFormatting 常量之一。 |
| **xlWebFormattingAll**                              |
| **xlWebFormattingRTF**                              |
| **xlWebFormattingNone** 默认值                      |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebFormatting** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表，并导入应用于数据的所有网页格式，然后刷新该查询表。
function test()
{
    let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)
    let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))
    qtQtrResults.WebFormatting = xlWebFormattingAll
    qtQtrResults.Refresh()
}
```

### QueryTable.WebPreFormattedTextToColumns

返回或设置向查询表中导入网页时，是否对网页 HTML \<PRE\> 标记内的数据进行分列。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.WebPreFormattedTextToColumns**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebPreFormattedTextToColumns** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表。注意，本示例并未对 HTML <PRE> 标记之间的数据进行分列。
function test()
{
    let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)
    let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))
    qtQtrResults.WebFormatting = xlNone
    qtQtrResults.WebPreFormattedTextToColumns = false
    qtQtrResults.Refresh()
}
```

### QueryTable.WebSelectionType

返回或设置一个值，该值决定是向查询表中导入整个网页、网页上的所有表格还是仅网页上的特定表格。 **XlWebSelectionType** 类型，可读写。

**语法**

**express.WebSelectionType**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

如果此属性的值为 **xlSpecifiedTables** ，则可以使用 **WebTables** 属性指定要导入的表格。

|                                                           |
|-----------------------------------------------------------|
| XlWebSelectionType 可为以下 XlWebSelectionType 常量之一。 |
| **xlEntirePage**                                          |
| **xlAllTables** 默认值                                    |
| **xlSpecifiedTables**                                     |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebSelectionType** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表，然后导入网页中第一个和第二个表格中的数据。
function test()
{
    let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)
    let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))
    qtQtrResults.WebFormatting = xlNone
    qtQtrResults.WebSelectionType = xlSpecifiedTables
    qtQtrResults.WebTables = "1,2"
    qtQtrResults.Refresh()
}
```

### QueryTable.WebSingleBlockTextImport

向查询表中导入网页时，如果位于指定网页的 HTML \<PRE\> 标记中的数据是同时进行处理的，则该值为 **True** 。如果数据是以连续行的数据块方式导入的，以便能识别标题行，则该值为 **False** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.WebSingleBlockTextImport**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebSingleBlockTextImport** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表，然后同时导入 HTML <PRE> 标记中的所有数据。
function test()
{
    let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)
    let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))
    qtQtrResults.WebSingleBlockTextImport = true
    qtQtrResults.Refresh()
}
```

### QueryTable.WebTables

向查询表中导入网页时，返回或设置由逗号分隔的表格名称或表格索引号的列表。 **String** 类型，可读写。

**语法**

**express.WebTables**

\*express是一个代表 QueryTable 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** ，查询返回 HTML 文档，并且 **WebSelectionType** 属性的值为 **xlSpecifiedTables** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebTables** 属性仅适用于 **QueryTable** 对象。

**示例**

``` JavaScript
//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表，然后导入网页中第一个和第二个表格中的数据。
function test()
{
    let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)
    let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))
    qtQtrResults.WebFormatting = xlNone
    qtQtrResults.WebSelectionType = xlSpecifiedTables
    qtQtrResults.WebTables = "1,2"
    qtQtrResults.Refresh()
}
```

### QueryTable.WorkbookConnection

返回查询表所使用的 **WorkbookConnection** 对象。只读。/p\>

**语法**

**express.WorkbookConnection**

\*express是一个代表 QueryTable 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WorkbookConnection** 属性仅适用于 **QueryTable** 对象。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/QueryTable](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TextFrame2 对象

## 简介

代表 Shape 、 ShapeRange 或 ChartFormat 对象的文本框。

## 说明

该对象包含文本框中的文本，还包含控制文本框对齐方式和位置的属性和方法。使用 TextFrame2 属性可返回 TextFrame2 对象。

示例

下例向 ` myDocument ` 添加一个矩形，向矩形添加文本，然后设置文本框的边距。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let rng = myDocument.Shapes.AddShape(Application.Enum.msoShapeRectangle, 0, 0, 250, 140).TextFrame2
    rng.TextRange.Text = "Here is some test text"
    rng.MarginBottom = 10
    rng.MarginLeft = 10
    rng.MarginRight = 10
    rng.MarginTop = 10
}
```

## 方法

| 序号 | 名称                                 | 说明                                       |
|:----:|:-------------------------------------|:-------------------------------------------|
|  1   | [DeleteText](#TextFrame2.DeleteText) | 删除文本框中的文字以及所有关联的文字属性。 |

## 属性

| 序号 | 名称                                             | 说明                                                                                                |
|:----:|:-------------------------------------------------|:----------------------------------------------------------------------------------------------------|
|  1   | [Application](#TextFrame2.Application)           | 返回 Application 对象。只读。                                                                       |
|  2   | [AutoSize](#TextFrame2.AutoSize)                 | 指定的对象的大小，可自动调整大小以适应其中包含的文字。可读/写 MsoAutoSize 。                        |
|  3   | [Column](#TextFrame2.Column)                     | 返回代表文本框内的列的 TextColumn2 对象。只读。                                                     |
|  4   | [Creator](#TextFrame2.Creator)                   | 返回一个 32 位整数，该整数指示创建对象的应用程序。只读 Long 类型。                                  |
|  5   | [HasText](#TextFrame2.HasText)                   | 返回指定文本框是否有文本。只读 MsoTriState 类型。                                                   |
|  6   | [HorizontalAnchor](#TextFrame2.HorizontalAnchor) | 返回或设置指定文本的水平定位类型。可读/写 MsoHorizontalAnchor 类型。                                |
|  7   | [MarginBottom](#TextFrame2.MarginBottom)         | 返回或设置从文本框底端到包含文本的形状的内接矩形底端的距离（以磅为单位）。可读/写 Single 类型。     |
|  8   | [MarginLeft](#TextFrame2.MarginLeft)             | 以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形的左边界的距离。可读/写 Single 类型。   |
|  9   | [MarginRight](#TextFrame2.MarginRight)           | 返回或设置从文本框右边缘到包含文本的形状的内接矩形右边缘的距离（以磅为单位）。可读/写 Single 类型。 |
|  10  | [MarginTop](#TextFrame2.MarginTop)               | 返回或设置从文本框顶端到包含文本的形状的内接矩形顶端的距离（以磅为单位）。可读/写 Single 类型。     |
|  11  | [NoTextRotation](#TextFrame2.NoTextRotation)     | 返回或设置在指定的对象旋转时，文本是否保持扁平。可读/写                                             |
|  12  | [Orientation](#TextFrame2.Orientation)           | 返回或设置一个值，该值代表文本框方向。可读/写 MsoTextOrientation 类型。                             |
|  13  | [Parent](#TextFrame2.Parent)                     | 返回指定对象的父对象。只读。                                                                        |
|  14  | [PathFormat](#TextFrame2.PathFormat)             | 返回或设置指定文本框的路径类型。可读/写 MsoPathFormat 类型。                                        |
|  15  | [Ruler](#TextFrame2.Ruler)                       | 返回一个 Ruler 对象，该对象代表指定文本的标尺。只读。                                               |
|  16  | [TextRange](#TextFrame2.TextRange)               | 返回 TextRange2 对象，该对象代表对象中的文本。只读                                                  |
|  17  | [ThreeD](#TextFrame2.ThreeD)                     | 返回一个 ThreeDFormat 对象，该对象包含指定文本的三维效果格式属性。只读。                            |
|  18  | [VerticalAnchor](#TextFrame2.VerticalAnchor)     | 返回或设置指定文本的垂直定位类型。可读/写 MsoVerticalAnchor 类型。                                  |
|  19  | [WarpFormat](#TextFrame2.WarpFormat)             | 返回或设置指定文本框的弯曲类型。可读/写 MsoWarpFormat 。                                            |
|  20  | [WordArtformat](#TextFrame2.WordArtformat)       | 返回或设置指定文本框的艺术字类型。可读/写 MsoPresetTextEffect 类型。                                |
|  21  | [WordWrap](#TextFrame2.WordWrap)                 | 返回或设置形状边界内外的文本换行。可读/写 MsoTriState 类型。                                        |

## 成员方法

### TextFrame2.DeleteText

删除文本框中的文字以及所有关联的文字属性。

**语法**

**express.DeleteText ()**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

关联的文字属性包括 **Font** 属性，例如加粗、下划线等等。

**示例**

如果文本框包含文字，则此示例将删除文本框中的文字。

``` JavaScript
function test() {
    let deleteText = Application.ActiveSheet.Shapes.Item(1).TextFrame2
    if (deleteText.HasText) {
        deleteText.DeleteText()
    }
}
```

## 成员属性

### TextFrame2.Application

返回 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

如果不与对象识别符一起使用，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可将此属性与一个 OLE 自动化对象一起使用以返回该对象的应用程序）。

### TextFrame2.AutoSize

指定的对象的大小，可自动调整大小以适应其中包含的文字。可读/写 **MsoAutoSize** 。

**语法**

**express.AutoSize**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.Column

返回代表文本框内的列的 **TextColumn2** 对象。只读。

**语法**

**express.Column**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.Creator

返回一个 32 位整数，该整数指示创建对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

**示例**

本示例显示一条有关 ET 工作簿的创建者的消息。

``` JavaScript
function test() {
    let myObject = Application.ActiveWorkbook.ActiveSheet.Shapes.Item(1)
    if (myObject.TextFrame2.Creator == 1480803660) {
        alert("This is a ET object.")
    }
    else {
        alert("This is not a ET object.")
    }
}
```

### TextFrame2.HasText

返回指定文本框是否有文本。只读 **MsoTriState** 类型。

**语法**

**express.HasText**

\*express是一个代表 TextFrame2 对象的变量。

**示例**

本示例设置文本框中文本的格式（如果文本框包含文本）。

``` JavaScript
function test() {
    let rng = Application.ActiveSheet.Shapes.Item(1).TextFrame2
    if (rng.HasText) {
        rng.TextRange2.Font.Name = "Arial"
    }
}
```

### TextFrame2.HorizontalAnchor

返回或设置指定文本的水平定位类型。可读/写 **MsoHorizontalAnchor** 类型。

**语法**

**express.HorizontalAnchor**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.MarginBottom

返回或设置从文本框底端到包含文本的形状的内接矩形底端的距离（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.MarginBottom**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.MarginLeft

以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形的左边界的距离。可读/写 **Single** 类型。

**语法**

**express.MarginLeft**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.MarginRight

返回或设置从文本框右边缘到包含文本的形状的内接矩形右边缘的距离（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.MarginRight**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.MarginTop

返回或设置从文本框顶端到包含文本的形状的内接矩形顶端的距离（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.MarginTop**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.NoTextRotation

返回或设置在指定的对象旋转时，文本是否保持扁平。可读/写

**语法**

**express.NoTextRotation**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

如果文本框旋转时，文本仍保持扁平，则为 **msoTrue** (-1)；否则为 **msoFalse** (0)。

### TextFrame2.Orientation

返回或设置一个值，该值代表文本框方向。可读/写 **MsoTextOrientation** 类型。

**语法**

**express.Orientation**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

此属性值可设置为从 -90 到 90 度的整数值，或者以下 **MsoTextOrientation** 常量之一。

### TextFrame2.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.PathFormat

返回或设置指定文本框的路径类型。可读/写 **MsoPathFormat** 类型。

**语法**

**express.PathFormat**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.Ruler

返回一个 **Ruler** 对象，该对象代表指定文本的标尺。只读。

**语法**

**express.Ruler**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.TextRange

返回 **TextRange2** 对象，该对象代表对象中的文本。只读

**语法**

**express.TextRange**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.ThreeD

返回一个 **ThreeDFormat** 对象，该对象包含指定文本的三维效果格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.VerticalAnchor

返回或设置指定文本的垂直定位类型。可读/写 **MsoVerticalAnchor** 类型。

**语法**

**express.VerticalAnchor**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.WarpFormat

返回或设置指定文本框的弯曲类型。可读/写 **MsoWarpFormat** 。

**语法**

**express.WarpFormat**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.WordArtformat

返回或设置指定文本框的艺术字类型。可读/写 **MsoPresetTextEffect** 类型。

**语法**

**express.WordArtformat**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.WordWrap

返回或设置形状边界内外的文本换行。可读/写 **MsoTriState** 类型。

**语法**

**express.WordWrap**

\*express是一个代表 TextFrame2 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/TextFrame2](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Top10 对象

## 简介

代表条件格式规则的前十项。通过对某一区域应用颜色，有助于查看相对于其他单元格的单元格的值。

## 说明

代表条件格式规则的前十项。通过对某一区域应用颜色，有助于查看相对于其他单元格的单元格的值。

以下示例通过条件格式规则生成一个动态数据集并对前 10 个值应用颜色。

``` JavaScript
function test() {                                                                   
    //Building data
    Application.Range("A1").Value2 = "Name"
    Application.Range("B1").Value2 = "Number"
    Application.Range("A2").Value2 = "Agent1"
    Application.Range("A2").AutoFill(Application.Range("A2:A26"), Application.Enum.xlFillDefault)
    Application.Range("B2:B26").FormulaArray = "=INT(RAND()*101)"
    Application.Range("B2:B26").Select()

    //Applying Conditional Formatting Top 10
    Application.Selection.FormatConditions.AddTop10()
    Application.Selection.FormatConditions.Item(Application.Selection.FormatConditions.Count).SetFirstPriority()
    let ft = Application.Selection.FormatConditions.Item(1)
    ft.TopBottom = Application.Enum.xlTop10Top
    ft.Rank = 10
    ft.Percent = false

    //Applying color fill
    let rng = Application.Selection.FormatConditions.Item(1).Font
    rng.Color = -16752384
    rng.TintAndShade = 0
    let fpx = Application.Selection.FormatConditions.Item(1).Interior
    fpx.PatternColorIndex = Application.Enum.xlAutomatic
    fpx.Color = 13561798
    fpx.TintAndShade = 0
}
```

## 方法

| 序号 | 名称                                                | 说明                                                                              |
|:----:|:----------------------------------------------------|:----------------------------------------------------------------------------------|
|  1   | [Delete](#Top10.Delete)                             | 删除指定的条件格式规则对象。                                                      |
|  2   | [ModifyAppliesToRange](#Top10.ModifyAppliesToRange) | 设置此格式规则所应用于的单元格区域。                                              |
|  3   | [SetFirstPriority](#Top10.SetFirstPriority)         | 将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则。 |
|  4   | [SetLastPriority](#Top10.SetLastPriority)           | 为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。        |

## 属性

| 序号 | 名称                                | 说明                                                                                                                                                             |
|:----:|:------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Top10.Application)   | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读 |
|  2   | [AppliesTo](#Top10.AppliesTo)       | 返回一个 Range 对象，指定格式规则将应用于的单元格区域。                                                                                                          |
|  3   | [Borders](#Top10.Borders)           | 返回一个 Borders 集合，该集合在条件格式规则的计算结果为 True 时指定单元格边框的格式。只读。                                                                      |
|  4   | [CalcFor](#Top10.CalcFor)           | 返回或设置 XlCalcFor 枚举的常量之一，该常量指定透视数据表中的条件格式的计算方式。                                                                                |
|  5   | [Creator](#Top10.Creator)           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                       |
|  6   | [Font](#Top10.Font)                 | 返回一个 Font 对象，该对象在条件格式规则的计算结果为 True 时指定字体格式。只读。                                                                                 |
|  7   | [Interior](#Top10.Interior)         | 返回一个 Interior 对象，该对象在条件格式规则的计算结果为 True 时指定单元格的内部属性。只读。                                                                     |
|  8   | [NumberFormat](#Top10.NumberFormat) | 在条件格式规则的计算结果为 True 时返回或设置应用于单元格的数字格式。 Variant 型，可读写。                                                                        |
|  9   | [Parent](#Top10.Parent)             | 返回指定对象的父对象。只读。                                                                                                                                     |
|  10  | [Percent](#Top10.Percent)           | 返回或设置一个 Boolean 值，指定是否按百分比值确定排位。                                                                                                          |
|  11  | [Priority](#Top10.Priority)         | 返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。                                                                       |
|  12  | [PTCondition](#Top10.PTCondition)   | 返回一个 Boolean 值，指明是否将条件格式应用于数据透视表图表。只读。                                                                                              |
|  13  | [Rank](#Top10.Rank)                 | 返回或设置一个 Long 值，为条件格式规则指定排位值的数值或百分比。                                                                                                 |
|  14  | [ScopeType](#Top10.ScopeType)       | 返回或设置 XlPivotConditionScope 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。                                                              |
|  15  | [StopIfTrue](#Top10.StopIfTrue)     | 返回或设置一个 Boolean 值，该值确定在当前规则的计算结果为 True 时是否应计算单元格上的其他格式规则。                                                              |
|  16  | [TopBottom](#Top10.TopBottom)       | 返回或设置 XlTopBottom 枚举的常量之一，该常量决定是从顶端还是从底端开始计算排位。                                                                                |
|  17  | [Type](#Top10.Type)                 | 返回 XlFormatConditionType 枚举的常量之一，该常量指定条件格式的类型。只读。                                                                                      |

## 成员方法

### Top10.Delete

删除指定的条件格式规则对象。

**语法**

**express.Delete ()**

\*express是一个代表 Top10 对象的变量。

### Top10.ModifyAppliesToRange

设置此格式规则所应用于的单元格区域。

**语法**

**express.ModifyAppliesToRange ( *Range* )**

\*express是一个代表 Top10 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                       |
|---------|-----------|----------|----------------------------|
| *Range* | 必选      | Range    | 此格式规则将应用于的区域。 |

**说明**

该区域必须采用 A1 引用样式，并全部包含在 **FormatConditions** 集合的父工作表内。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可以使用货币符号，但会被忽略。

您也可以在区域的任意部分中使用局部定义名称，但该名称必须使用宏语言。

### Top10.SetFirstPriority

将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则。

**语法**

**express.SetFirstPriority ()**

\*express是一个代表 Top10 对象的变量。

**说明**

当工作表中有多个条件格式规则时，如果之前未将此规则设置为优先级“1”，此方法将导致工作表上的所有其他现有规则的优先级都增加 1。

| 注释                                     |
|------------------------------------------|
| 将在工作表级别应用条件格式规则的优先级。 |

### Top10.SetLastPriority

为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。

**语法**

**express.SetLastPriority ()**

\*express是一个代表 Top10 对象的变量。

**示例**

优先级的实际值将等于工作表上条件格式规则的总数。当工作表中有多个条件格式规则时，此方法将导致优先级值大于此规则的规则的优先级增加 1。

| 注释                                     |
|------------------------------------------|
| 将在工作表级别应用条件格式规则的优先级。 |

## 成员属性

### Top10.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读

**语法**

**express.Application**

\*express是一个代表 Top10 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### Top10.AppliesTo

返回一个 **Range** 对象，指定格式规则将应用于的单元格区域。

**语法**

**express.AppliesTo**

\*express是一个代表 Top10 对象的变量。

### Top10.Borders

返回一个 **Borders** 集合，该集合在条件格式规则的计算结果为 **True** 时指定单元格边框的格式。只读。

**语法**

**express.Borders**

\*express是一个代表 Top10 对象的变量。

**说明**

对于条件格式对象，您只能为单元格的上边框、下边框和侧边框设置属性。

### Top10.CalcFor

返回或设置 **XlCalcFor** 枚举的常量之一，该常量指定透视数据表中的条件格式的计算方式。

**语法**

**express.CalcFor**

\*express是一个代表 Top10 对象的变量。

**说明**

只有在条件格式应用于数据透视表中的数据时，此属性才适用。

只有在 **Top10.ScopeType** 属性设置为 **xlFieldsScope** 时，才可将此属性设置为 **xlAllValues** 、 **xlColGroups** 或 **xlRowGroups** 。

### Top10.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Top10 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Top10.Font

返回一个 **Font** 对象，该对象在条件格式规则的计算结果为 **True** 时指定字体格式。只读。

**语法**

**express.Font**

\*express是一个代表 Top10 对象的变量。

**说明**

条件格式对象并不支持 **Font** 对象的所有属性。您可以设置字体样式、下划线、颜色和删除线属性。

### Top10.Interior

返回一个 **Interior** 对象，该对象在条件格式规则的计算结果为 **True** 时指定单元格的内部属性。只读。

**语法**

**express.Interior**

\*express是一个代表 Top10 对象的变量。

### Top10.NumberFormat

在条件格式规则的计算结果为 **True** 时返回或设置应用于单元格的数字格式。 **Variant** 型，可读写。

**语法**

**express.NumberFormat**

\*express是一个代表 Top10 对象的变量。

**说明**

数字格式是使用 **“单元格格式”** 对话框的 **“数字”** 选项卡上显示的相同格式代码指定的。您可以使用内置的数字格式，例如 ` "General" ` 或者 创建自定义数字格式 。

### Top10.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Top10 对象的变量。

### Top10.Percent

返回或设置一个 **Boolean** 值，指定是否按百分比值确定排位。

**语法**

**express.Percent**

\*express是一个代表 Top10 对象的变量。

### Top10.Priority

返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。

**语法**

**express.Priority**

\*express是一个代表 Top10 对象的变量。

**说明**

设置优先级时，值必须为介于 1 和工作表上条件格式规则总数之间的正整数。对于工作表上的所有规则，优先级必须为唯一值，因此，如果更改指定条件格式规则的优先级，将可能会导致工作表上其他规则的优先级值移位。

### Top10.PTCondition

返回一个 **Boolean** 值，指明是否将条件格式应用于数据透视表图表。只读。

**语法**

**express.PTCondition**

\*express是一个代表 Top10 对象的变量。

### Top10.Rank

返回或设置一个 **Long** 值，为条件格式规则指定排位值的数值或百分比。

**语法**

**express.Rank**

\*express是一个代表 Top10 对象的变量。

### Top10.ScopeType

返回或设置 **XlPivotConditionScope** 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。

**语法**

**express.ScopeType**

\*express是一个代表 Top10 对象的变量。

**说明**

默认值为 **xlSelectionScope** ，它使用 **AppliesTo** 属性设置范围。

### Top10.StopIfTrue

返回或设置一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则。

**语法**

**express.StopIfTrue**

\*express是一个代表 Top10 对象的变量。

**说明**

为了支持向后兼容性，此属性的默认值为 **True** ，而这与在用户界面中创建的格式规则的行为相反。

### Top10.TopBottom

返回或设置 **XlTopBottom** 枚举的常量之一，该常量决定是从顶端还是从底端开始计算排位。

**语法**

**express.TopBottom**

\*express是一个代表 Top10 对象的变量。

**说明**

### Top10.Type

返回 **XlFormatConditionType** 枚举的常量之一，该常量指定条件格式的类型。只读。

**语法**

**express.Type**

\*express是一个代表 Top10 对象的变量。

**说明**

此属性将始终返回 **Long** 值“5”，等价于 **xlTop10** 。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Top10](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Workbooks 对象

## 简介

ET 应用程序中当前打开的所有 Workbook 对象的集合。

## 说明

有关使用一个 Workbook 对象的详细信息，请参阅 Workbook 对象。

## 示例

使用 Workbooks 属性可返回 Workbooks 集合。下例关闭所有打开的工作簿。

``` JavaScript
Application.Workbooks.Close()
```

使用 Add 方法可创建一个新空工作簿并将它添加到集合。下例给 ET 添加一个新空工作簿。

``` JavaScript
Application.Workbooks.Add()
```

使用 Open 方法可打开一个文件。这样会为打开的文件创建一个新工作簿。下例将文件 Array.xls 打开为只读工作簿。

``` JavaScript
Application.Workbooks.Open("Array.xls",null,true)
```

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#Workbooks.Add">Add</a></td>
<td style="text-align: left;" data-valian="middle"><p>新建一个工作表。新工作表将成为活动工作表。</p>
<p>返回值： 一个代表新工作簿的 Workbook 对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#Workbooks.Close">Close</a></td>
<td style="text-align: left;" data-valian="middle"><p>关闭对象。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#Workbooks.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>从集合中返回一个对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#Workbooks.Open">Open</a></td>
<td style="text-align: left;" data-valian="middle"><p>打开一个工作簿。 返回值： 一个代表打开的工作簿的 Workbook 对象。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#Workbooks.OpenText">OpenText</a></td>
<td style="text-align: left;" data-valian="middle"><p>载入一个文本文件，并将其作为包含单个工作表的新工作簿进行分列处理，然后在此工作表中放入经过分列处理的文本文件数据。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Workbooks.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Workbooks.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |

## 成员方法

### Workbooks.Add

新建一个工作表。新工作表将成为活动工作表。

返回值： 一个代表新工作簿的 Workbook 对象。

**语法**

**express.Add ( *Template* )**

\*express是一个代表 Workbooks 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                                                                                    |
|------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Template* | 可选      | Variant  | 确定如何创建新工作簿。如果此参数为指定现有 ET 文件名的字符串，那么创建新工作簿将以该指定的文件作为模板。如果此参数为常量，新工作簿将包含一个指定类型的工作表。可为以下 XlWBATemplate 常量之一：xlWBATChart、xlWBATExcel4IntlMacroSheet、xlWBATExcel4MacroSheet 或 xlWBATWorksheet。如果省略此参数，ET 将创建包含一定数目空白工作表的新工作簿（该数目由 SheetsInNewWorkbook 属性设置）。 |

**说明**

如果 *Template* 参数指定的是文件，则该文件名可包含路径。

### Workbooks.Close

关闭对象。

**语法**

**express.Close ()**

\*express是一个代表 Workbooks 对象的变量。

**示例**

``` JavaScript
/*此示例关闭所有打开的工作簿。如果某个打开的工作簿有改动，ET 将显示询问是否保存更改的对话框和相应提示。*/
Application.Workbooks.Close()
```

### Workbooks.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Workbooks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
/*此示例将变量 wb 设置为 Myaddin.xla 的工作簿。*/
let wb = Application.Workbooks.Item("myaddin.xla")
```

### Workbooks.Open

打开一个工作簿。 返回值： 一个代表打开的工作簿的 Workbook 对象。

**语法**

**express.Open ( *FileName* , *UpdateLinks* , *ReadOnly* , *Format* , *Password* , *WriteResPassword* , *IgnoreReadOnlyRecommended* , *Origin* , *Delimiter* , *Editable* , *Notify* , *Converter* , *AddToMru* , *Local* , *CorruptLoad* )**

\*express是一个代表 Workbooks 对象的变量。

**参数**

| 名称                        | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                                                         |
|-----------------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName*                  | 可选      | Variant  | String. 要打开的工作簿的文件名。                                                                                                                                                                                                                                                                                                                             |
| *UpdateLinks*               | 可选      | Variant  | 指定更新文件中外部引用（链接）的方式，如下面的公式 =SUM(\[Budget.xls\]Annual!C10:C25) 中对 Budget.xls 工作簿中某个区域的引用。如果省略此参数，则提示用户指定链接的更新方式。有关此参数所用值的详细信息，请参阅“说明”部分。如果 ET 正在打开 WKS、WK1 或 WK3 格式的文件，并且 UpdateLinks 参数为 0，则不创建任何图表；否则 ET 根据附加于该文件的图形生成图表。 |
| *ReadOnly*                  | 可选      | Variant  | 如果为 True，则以只读模式打开工作簿。                                                                                                                                                                                                                                                                                                                        |
| *Format*                    | 可选      | Variant  | 如果 ET 正在打开文本文件，则由此参数指定分隔符。如果省略此参数，则使用当前的分隔符。有关此参数值的详细信息，请参阅“备注”部分。                                                                                                                                                                                                                               |
| *Password*                  | 可选      | Variant  | 一个字符串，包含打开受保护工作簿所需的密码。如果省略此参数并且工作簿已设置密码，则提示用户输入密码。                                                                                                                                                                                                                                                         |
| *WriteResPassword*          | 可选      | Variant  | {description} 一个字符串，包含写入受保护工作簿所需的密码。如果省略此参数并且工作簿已设置密码，则提示用户输入密码。                                                                                                                                                                                                                                           |
| *IgnoreReadOnlyRecommended* | 可选      | Variant  | 如果为 True，则不让 ET 显示只读的建议消息（如果该工作簿以“建议只读”选项保存）。                                                                                                                                                                                                                                                                              |
| *Origin*                    | 可选      | Variant  | 如果该文件为文本文件，则此参数用于指示该文件来源于何种操作系统（以便正确映射代码页和回车/换行符 (CR/LF)）。可为以下 XlPlatform 常量之一：xlMacintosh、xlWindows 或 xlMSDOS。如果省略此参数，则使用当前操作系统。                                                                                                                                             |
| *Delimiter*                 | 可选      | Variant  | 如果该文件为文本文件并且 Format 参数为 6，则此参数是一个字符串，指定用作分隔符的字符。例如，可使用 Chr(9) 代表制表符，使用“,”代表逗号，使用“;”代表分号，或者使用自定义字符。只使用字符串的第一个字符。                                                                                                                                                       |
| *Editable*                  | 可选      | Variant  | 如果文件为 ET 4.0 加载宏，则此参数为 True 时可打开该加载宏以使其在窗口中可见。如果此参数为 False 或被省略，则以隐藏方式打开加载宏，并且无法设为可见。本选项不能应用于由 ET 5.0 或更高版本的 ET 创建的加载宏。如果文件是 ET 模板，则参数值为 True 时，会打开指定模板进行编辑。参数值为 False 时，可根据指定模板打开新的工作簿。默认值为 False。               |
| *Notify*                    | 可选      | Variant  | 当文件不能以可读写模式打开时，如果此参数为 True，则可将该文件添加到文件通知列表。ET 将以只读模式打开该文件并轮询文件通知列表，并在文件可用时向用户发出通知。如果此参数为 False 或被省略，则不请求任何通知，并且不能打开任何不可用的文件。                                                                                                                    |
| *Converter*                 | 可选      | Variant  | 打开文件时试用的第一个文件转换器的索引。首先试用的是指定的文件转换器；如果该转换器不能识别此文件，则试用所有其他转换器。转换器索引由 FileConverters 属性返回的转换器行号组成。                                                                                                                                                                               |
| *AddToMru*                  | 可选      | Variant  | 如果为 True，则将该工作簿添加到最近使用的文件列表中。默认值为 False。                                                                                                                                                                                                                                                                                        |
| *Local*                     | 可选      | Variant  | 如果为 True，则以 ET（包括控制面板设置）的语言保存文件。如果为 False（默认值），则以 示例代码 (VBA) 语言保存文件。VBA 通常为美国英语版本，除非从中运行 Workbooks.Open 的 VBA 项目是旧的国际化 XL5/95 VBA 项目。                                                                                                                                              |
| *CorruptLoad*               | 可选      | Variant  | 可为以下常量之一：xlNormalLoad、xlRepairFile 和 xlExtractData。如果未指定任何值，则默认行为是 xlNormalLoad，并且当通过 OM 启动时不尝试恢复状态。                                                                                                                                                                                                             |

**说明**

默认情况下，以编程方式打开文件时将启用宏。使用 AutomationSecurity 属性可设置以编程方式打开文件时所用的宏安全模式。

可在 *UpdateLinks* 参数中指定下面的一个值，以确定在工作簿打开时是否更新外部引用（链接）：

| 值  | 含义                                 |
|-----|--------------------------------------|
| 0   | 工作簿打开时不更新外部引用（链接）。 |
| 3   | 工作簿打开时更新外部引用（链接）。   |

您可在 *Format* 参数中指定下面的一个值，以确定文件的分隔字符：

| Value | 分隔符                              |
|-------|-------------------------------------|
| 1     | 标签                                |
| 2     | 逗号                                |
| 3     | 空格                                |
| 4     | 分号                                |
| 5     | Nothing                             |
| 6     | 自定义字符（请参阅 Delimiter 参数） |

**示例**

``` JavaScript
/*本示例打开 Analysis.xls 工作簿，然后运行它的 Auto_Open 宏。*/
function test()
{
    Application.Workbooks.Open("ANALYSIS.XLS")
    Application.ActiveWorkbook.RunAutoMacros(xlAutoOpen)
}
```

### Workbooks.OpenText

载入一个文本文件，并将其作为包含单个工作表的新工作簿进行分列处理，然后在此工作表中放入经过分列处理的文本文件数据。

**语法**

**express.OpenText ( *Filename* , *Origin* , *StartRow* , *DataType* , *TextQualifier* , *ConsecutiveDelimiter* , *Tab* , *Semicolon* , *Comma* , *Space* , *Other* , *OtherChar* , *FieldInfo* , *TextVisualLayout* , *DecimalSeparator* , *ThousandsSeparator* , *TrailingMinusNumbers* , *Local* )**

\*express是一个代表 Workbooks 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型        | 说明                                                                                                                                                                                                                                                                        |
|------------------------|-----------|-----------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Filename*             | 必选      | String          | 指定要打开和分列的文本文件的名称。                                                                                                                                                                                                                                          |
| *Origin*               | 可选      | Variant         | 指定文本文件来源。可为以下 XlPlatform 常量之一：xlMacintosh、xlWindows 或 xlMSDOS。此外，它还可以是一个整数，表示所需代码页的代码页编号。例如，“1256”指定源文本文件的编码是阿拉伯语 (Windows)。如果省略该参数，则此方法将使用“文本导入向导”中“文件原始格式”选项的当前设置。 |
| *StartRow*             | 可选      | Variant         | 文本分列处理的起始行号。默认值为 1。                                                                                                                                                                                                                                        |
| *DataType*             | 可选      | Variant         | 指定文件中数据的列格式。可为以下 XlTextParsingType 常量之一：xlDelimited 或 xlFixedWidth。如果未指定该参数，则 ET 将尝试在打开文件时确定列格式。                                                                                                                            |
| *TextQualifier*        | 可选      | XlTextQualifier | 指定文本识别符号。                                                                                                                                                                                                                                                          |
| *ConsecutiveDelimiter* | 可选      | Variant         | 如果为 True，则将连续分隔符视为一个分隔符。默认值为 False。                                                                                                                                                                                                                 |
| *Tab*                  | 可选      | Variant         | 如果为 True，则将制表符用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。                                                                                                                                                                                          |
| *Semicolon*            | 可选      | Variant         | 如果为 True，则将分号用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。                                                                                                                                                                                            |
| *Comma*                | 可选      | Variant         | 如果为 True，则将逗号用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。                                                                                                                                                                                            |
| *Space*                | 可选      | Variant         | 如果为 True，则将空格用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。                                                                                                                                                                                            |
| *Other*                | 可选      | Variant         | 如果为 True，则将 OtherChar 参数指定的字符用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。                                                                                                                                                                       |
| *OtherChar*            | 可选      | Variant         | （如果 Other 为 True，则为必选项）。当 Other 为 True 时，指定分隔符。如果指定了多个字符，则仅使用字符串中的第一个字符而忽略剩余字符。                                                                                                                                       |
| *FieldInfo*            | 可选      | Variant         | 包含单列数据相关分列信息的数组。对该参数的解释取决于 DataType 的值。如果此数据由分隔符分隔，则该参数为由两元素数组组成的数组，其中每个两元素数组指定一个特定列的转换选项。第一个元素为列标（从 1 开始），第二个元素是 XlColumnDataType 的常量之一，用于指定分列方式。       |
| *TextVisualLayout*     | 可选      | Variant         | 文本的可视布局。                                                                                                                                                                                                                                                            |
| *DecimalSeparator*     | 可选      | Variant         | 识别数字时，ET 使用的小数分隔符。默认设置为系统设置。                                                                                                                                                                                                                       |
| *ThousandsSeparator*   | 可选      | Variant         | 识别数字时， ET 使用的千位分隔符。默认设置为系统设置。                                                                                                                                                                                                                      |
| *TrailingMinusNumbers* | 可选      | Variant         | 如果应将结尾为减号字符的数字视为负数处理，则指定为 True。如果为 False 或省略该参数，则将结尾为减号字符的数字视为文本处理。                                                                                                                                                  |
| *Local*                | 可选      | Variant         | 如果分隔符、数字和数据格式应使用计算机的区域设置，则指定为 True。                                                                                                                                                                                                           |

**说明**

**FieldInfo 参数信息**

只有在安装并选定了中国台湾地区语言支持时才可使用 **xlEMDFormat** 。 **xlEMDFormat** 常量指定使用中国台湾地区纪元日期。

列说明符可为任意顺序。输入数据中如果某列没有列说明符，则用常规设置对该列进行分列处理。

**注释：**

- 在指定要跳过的列时，必须明确说明其余各列的类型，否则数据不能正确拆分。
- 如果数据中包含可识别的日期，则工作表中单元格的格式就会设置为“日期”，即便为该列设置的格式为“常规”。此外，如果您为某一列指定了以上日期格式中的一种格式，而数据中并不包含可识别的日期，那么工作表中单元格的格式为“常规”。

如果源数据为定宽列，则每个两元数组的第一个元素指定起始元素在列中的位置，（用整数表示，第一个字符为 0 (零)），第二个元素用 0 到 9 的数字指定分列选项，如前表所示。

**ThousandsSeparator 参数信息**

下表显示了使用不同的导入设置向 ET 中导入文本时的结果。数字结果显示在最右边的列中。

| 系统小数分隔符 | 系统千位分隔符 | 小数分隔符值 | 千位分隔符值 | 导入的文本 | 单元格的值（数据类型） |
|----------------|----------------|--------------|--------------|------------|------------------------|
| 句号           | 逗号           | 逗号         | 句号         | 123.123,45 | 123,123.45（数字）     |
| 句号           | 逗号           | 逗号         | 逗号         | 123.123,45 | 123.123,45（文本）     |
| 逗号           | 句号           | 句号         | 逗号         | 123,123.45 | 123,123.45（数字）     |
| 句号           | 逗号           | 句号         | 逗号         | 123 123.45 | 123 123.45（文本）     |
| 句号           | 逗号           | 句号         | 空格         | 123 123.45 | 123,123.45（数字）     |

**示例**

``` JavaScript
/*本示例打开 Data.txt 文件并将制表符作为分隔符对此文件进行分列处理，将其转换成为工作表。*/
Application.Workbooks.OpenText("C:\\DATA.TXT",null,null,xlDelimited,null,null,true)
```

## 成员属性

### Workbooks.Application

如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Workbooks 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }else {
        alert("This is not an ET Application object.")
    }
}
```

### Workbooks.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Workbooks 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Workbooks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Worksheet 对象

## 简介

代表一个工作表。

## 说明

Worksheet 对象是 Worksheets 集合的成员。 Worksheets 集合包含某个工作簿中所有的 Worksheet 对象。

Worksheet 对象也是 Sheets 集合的成员。 Sheets 集合包含工作簿中所有的工作表（图表工作表和工作表）。

## 示例

使用 Worksheets ( *index* )（其中 *index* 是工作表索引号或名称）可返回一个 Worksheet 对象。下例隐藏活动工作簿中的工作表一。

``` JavaScript
Application.Worksheets.Item(1).Visible = false
```

工作表索引号指示该工作表在工作簿的标签栏上的位置。 ` Worksheets(1) ` 是工作簿中第一个（最左边的）工作表，而 ` Worksheets(Worksheets.Count) ` 是最后一个。所有工作表均包括在索引计数中，即便是隐藏工作表也是如此。

工作表名称显示在该工作表的标签上。使用 Name 属性可设置或返回工作表名称。下例保护 Sheet1 上的方案。

``` JavaScript
let strPassword = prompt("Enter the password for the worksheet")
Application.Worksheets.Item("Sheet1").Protect(strPassword,null,null,true)
```

当工作表处于活动状态时，可以使用ActiveSheet 属性来引用它。下例使用 Activ ate 方法激活 Sheet1，将页面方向设置为横向，然后打印该工作表。

``` JavaScript
Application.Worksheets.Item("Sheet1").Activate()
Application.ActiveSheet.PageSetup.Orientation = xlLandscape
Application.ActiveSheet.PrintOut()
```

## 方法

| 序号 | 名称                              | 说明                                                                         |
|:----:|:----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Copy](#Worksheet.Copy)           | 将工作表复制到工作簿的另一位置。                                             |
|  2   | [Delete](#Worksheet.Delete)       | 删除对象。                                                                   |
|  3   | [Move](#Worksheet.Move)           | 将工作表移到工作簿中的其他位置。                                             |
|  4   | [Protect](#Worksheet.Protect)     | 保护工作表使其不能被修改。                                                   |
|  5   | [Range](#Worksheet.Range)         | 返回一个 Range 对象，它代表一个单元格或单元格区域。                          |
|  6   | [Select](#Worksheet.Select)       | 选择对象。                                                                   |
|  7   | [Unprotect](#Worksheet.Unprotect) | 取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Worksheet.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Cells](#Worksheet.Cells)             | 返回一个 Range 对象，它代表工作表中的所有单元格（不仅仅是当前使用的单元格）。                                                                                                                                                   |
|  3   | [Columns](#Worksheet.Columns)         | 返回一个 Range 对象，它代表活动工作表中的所有列。如果活动文档不是工作表，则 Columns 属性失效。                                                                                                                                  |
|  4   | [Index](#Worksheet.Index)             | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                    |
|  5   | [Name](#Worksheet.Name)               | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                    |
|  6   | [Names](#Worksheet.Names)             | 返回一个 Names 集合，它代表所有特定于工作表的名称（使用“WorksheetName!”前缀定义的名称）。 Names 对象，只读。                                                                                                                    |
|  7   | [Next](#Worksheet.Next)               | 返回代表下一个工作表的 Worksheet 对象。                                                                                                                                                                                         |
|  8   | [Previous](#Worksheet.Previous)       | 返回代表下一个工作表的 Worksheet 对象。                                                                                                                                                                                         |
|  9   | [Protection](#Worksheet.Protection)   | 返回一个 Protection 对象，该对象表示工作表的保护选项。                                                                                                                                                                          |
|  10  | [Rows](#Worksheet.Rows)               | 返回一个 Range 对象，它代表指定工作表中的所有行。 Range 对象，只读。                                                                                                                                                            |
|  11  | [Shapes](#Worksheet.Shapes)           | 返回一个 Shapes 集合，它代表工作表上的所有形状。只读。                                                                                                                                                                          |
|  12  | [Type](#Worksheet.Type)               | 返回一个代表工作表类型的 XlSheetType 值。                                                                                                                                                                                       |
|  13  | [UsedRange](#Worksheet.UsedRange)     | 返回一个 Range 对象，该对象表示指定工作表上所使用的区域。只读。                                                                                                                                                                 |
|  14  | [Visible](#Worksheet.Visible)         | 返回或设置一个 XlSheetVisibility 值，它确定对象是否可见。                                                                                                                                                                       |

## 成员方法

### Worksheet.Copy

将工作表复制到工作簿的另一位置。

**语法**

**express.Copy ( *Before* , *After* )**

\*express是一个代表 Worksheet 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                        |
|----------|-----------|----------|-----------------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 将要在其之前放置所复制工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选      | Variant  | 将要在其之后放置所复制工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After* ，则 ET 将新建一个工作簿，其中包含复制的工作表。

**示例**

``` JavaScript
/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之后。*/
Application.Worksheets.Item("Sheet1").Copy(null,Application.Worksheets.Item("Sheet3"))
```

### Worksheet.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Worksheet 对象的变量。

**说明**

在删除 **Worksheet** 时，此方法默认情况下显示一个对话框，用于提示用户确认是否删除。

### Worksheet.Move

将工作表移到工作簿中的其他位置。

**语法**

**express.Move ( *Before* , *After* )**

\*express是一个代表 Worksheet 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                  |
|----------|-----------|----------|-----------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 在其之前放置移动工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选      | Variant  | 在其之后放置移动工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After* ，ET 将新建一个工作簿，其中包含所移动的工作表。

**示例**

``` JavaScript
/*此示例将当前活动工作簿的 Sheet1 移到 Sheet3 之后。*/
Application.Worksheets.Item("Sheet1").Move(null,Application.Worksheets.Item("Sheet3"))
```

### Worksheet.Protect

保护工作表使其不能被修改。

**语法**

**express.Protect ( *Password* , *DrawingObjects* , *Contents* , *Scenarios* , *UserInterfaceOnly* , *AllowFormattingCells* , *AllowFormattingColumns* , *AllowFormattingRows* , *AllowInsertingColumns* , *AllowInsertingRows* , *AllowInsertingHyperlinks* , *AllowDeletingColumns* , *AllowDeletingRows* , *AllowSorting* , *AllowFiltering* , *AllowUsingPivotTables* )**

\*express是一个代表 Worksheet 对象的变量。

**参数**

| 名称                       | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                     |
|----------------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Password*                 | 可选      | Variant  | 一个字符串，该字符串为工作表或工作簿指定区分大小写的密码。如果省略此参数，不用密码就可以取消对工作表或工作簿的保护。否则，必须指定密码才能取消对工作表或工作簿的保护。如果忘记了密码，就无法取消对工作表或工作簿的保护。 |
| *DrawingObjects*           | 可选      | Variant  | 如果为 True，则保护形状。默认值是 True。                                                                                                                                                                                 |
| *Contents*                 | 可选      | Variant  | 如果为 True，则保护内容。对于图表，这样会保护整个图表。对于工作表，这样会保护锁定的单元格。默认值是 True。                                                                                                               |
| *Scenarios*                | 可选      | Variant  | 如果为 True，则保护方案。此参数仅对工作表有效。默认值是 True。                                                                                                                                                           |
| *UserInterfaceOnly*        | 可选      | Variant  | 如果为 True，则保护用户界面，但不保护宏。如果省略此参数，则既保护宏也保护用户界面。                                                                                                                                      |
| *AllowFormattingCells*     | 可选      | Variant  | 如果为 True，则允许用户为受保护的工作表上的任意单元格设置格式。默认值是 False。                                                                                                                                          |
| *AllowFormattingColumns*   | 可选      | Variant  | 如果为 True，则允许用户为受保护的工作表上的任意列设置格式。默认值是 False。                                                                                                                                              |
| *AllowFormattingRows*      | 可选      | Variant  | 如果为 True，则允许用户为受保护的工作表上的任意行设置格式。默认值是 False。                                                                                                                                              |
| *AllowInsertingColumns*    | 可选      | Variant  | 如果为 True，则允许用户在受保护的工作表上插入列。默认值是 False。                                                                                                                                                        |
| *AllowInsertingRows*       | 可选      | Variant  | 如果为 True，则允许用户在受保护的工作表上插入行。默认值是 False。                                                                                                                                                        |
| *AllowInsertingHyperlinks* | 可选      | Variant  | 如果为 True，则允许用户在受保护的工作表中插入超链接。默认值是 False。                                                                                                                                                    |
| *AllowDeletingColumns*     | 可选      | Variant  | 如果为 True，则允许用户在受保护的工作表上删除列，要删除的列中的每个单元格都被解除锁定。默认值是 False。                                                                                                                  |
| *AllowDeletingRows*        | 可选      | Variant  | 如果为 True，则允许用户在受保护的工作表上删除行，要删除的行中的每个单元格都被解除锁定。默认值是 False。                                                                                                                  |
| *AllowSorting*             | 可选      | Variant  | 如果为 True，则允许用户在受保护的工作表上进行排序。排序区域中的每个单元格必须是解除锁定的或取消保护的。默认值是 False。                                                                                                  |
| *AllowFiltering*           | 可选      | Variant  | 如果为 True，则允许用户在受保护的工作表上设置筛选。用户可以更改筛选条件，但是不能启用或禁用自动筛选功能。用户也可以在已有的自动筛选功能上设置筛选。默认值是 False。                                                      |
| *AllowUsingPivotTables*    | 可选      | Variant  | 如果为 True，则允许用户在受保护的工作表上使用数据透视表。默认值是 False。                                                                                                                                                |

**说明**

要在受保护的工作表上做更改，如果提供密码，则可在受保护的工作表上使用 **Protect** 方法。另一种方法是：取消工作表保护，对工作表做一些必要的更改，然后再次保护工作表。

**注释：** “取消保护”的意思是单元格可以被锁定（ “设置单元格格式” 对话框），但在 “允许用户编辑区域” 对话框中定义的单元格区域内，并且用户已通过密码取消了对单元格区域的保护或已通过 NT 权限的验证。

### Worksheet.Range

返回一个 Range 对象，它代表一个单元格或单元格区域。

**语法**

**express.Range ( *Cell1* , *Cell2* )**

\*express是一个代表 Worksheet 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                       |
|---------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Cell1* | 必选      | Variant  | 区域名称。必须为采用宏语言的 A1 样式引用。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可包括货币符号，但它们会被忽略掉。您可以在区域中任一部分使用局部定义名称。如果使用名称，则假定该名称使用的是宏语言。 |
| *Cell2* | 可选      | Variant  | 区域左上角和右下角的单元格。可以是一个包含单个单元格、整列或整行的 Range 对象，或者也可以是一个用宏语言为单个单元格命名的字符串。                                                                                                          |

**说明**

如果在没有对象识别符时使用，则该属性是 ` ActiveSheet.Range ` 的快捷方式（它返回活动表的一个区域，如果活动表不是一张工作表，则该属性无效）。

当应用于 **Range** 对象时，该属性与 **Range** 对象相关。例如，如果选中单元格 C3，那么 ` Selection.Range("B1") ` 返回单元格 D3，因为它同 **Selection** 属性返回的 **Range** 对象相关。此外，代码 ` ActiveSheet.Range("B1") ` 总是返回单元格 B1。

**示例**

``` JavaScript
/*此示例将 Sheet1 上 A1 单元格的值设置为 3.14159。*/
Application.Worksheets.Item("Sheet1").Range("A1").Value2 = 3.14159

/*此示例在 Sheet1 的 A1 单元格中创建一个公式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Formula = "=10*RAND()"

/*此示例在 Sheet1 上的单元格区域 A1:D10 中进行循环。如果某个单元格的值小于 0.001，则此代码将用 0（零）来取代该值。*/
function test()
{
    for(let i = 1; i <= Application.Worksheets.Item("Sheet1").Range("A1:D10").Count; i++){
        if(Application.Worksheets.Item("Sheet1").Range("A1:D10").Item(i).Value2 < .001 ){
            Application.Worksheets.Item("Sheet1").Range("A1:D10").Item(i).Value2 = 0
        }
    }
}

/*此示例在名为“TestRange”的区域上进行循环，并显示该区域中空白单元格的个数。*/
function test()
{
    let numBlanks = 0
    for(let i = 1; i <= Application.Range("TestRange").Count; i++){
        if(Application.Range("TestRange").Item(i).Value2 == null){
            numBlanks++
        }
    }
    alert("There are " + numBlanks + " empty cells in this range")
}

/*此示例将 Sheet1 中单元格区域 A1:C5 上的字体样式设置为斜体。此示例使用 Range 属性的语法 2。*/
Application.Worksheets.Item("Sheet1").Range(Application.Cells.Item(1, 1), Application.Cells.Item(5, 3)).Font.Italic = true
```

### Worksheet.Select

选择对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 Worksheet 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                              |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

**说明**

要选择单元格或单元格区域，请使用 **Select** 方法。要使单个单元格成为活动单元格，请使用 **Activate** 方法。

### Worksheet.Unprotect

取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。

**语法**

**express.Unprotect ( *Password* )**

\*express是一个代表 Worksheet 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                            |
|------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Password* | 可选      | Variant  | 一个字符串，指定用于解除工作表或工作簿保护的密码，此密码是区分大小写的。如果工作表或工作簿不设密码保护，则省略此参数。如果对工作表省略此参数，而该工作表又设有密码保护，ET 将提示您输入密码。如果对工作簿省略此参数，而该工作簿又设有密码保护，则该方法将失效。 |

**说明**

如果您忘记了密码，将不能取消工作表或工作簿的保护。建议将密码和对应文档名妥善保存。

**示例**

``` JavaScript
/*此示例取消当前活动工作表的保护。*/
Application.ActiveSheet.Unprotect()
```

## 成员属性

### Worksheet.Application

如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Worksheet 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }else {
        alert("This is not an ET Application object.")
    }
}  
```

### Worksheet.Cells

返回一个 Range 对象，它代表工作表中的所有单元格（不仅仅是当前使用的单元格）。

**语法**

**express.Cells**

\*express是一个代表 Worksheet 对象的变量。

**说明**

因为 **Item** 属性是 Range 对象的 <span class="glossary" "appendpopup(this,'defdefaultproperty_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none;">默认属性</span> ，所以可以在 **Cells** 关键字后面紧接着指定行和列索引。有关详细信息，请参阅 **Item** 属性和本主题的示例。

在不使用对象识别符的情况下，使用此属性将返回一个 Range 对象，它代表活动工作表中所有的单元格。

**示例**

``` JavaScript
/*本示例将 Sheet1 中单元格 C5 的字号设置为 14 磅。*/
Application.Worksheets.Item("Sheet1").Cells.Item(5, 3).Font.Size = 14

/*本示例清除 Sheet1 上第一个单元格的公式。*/
Application.Worksheets.Item("Sheet1").Cells.Item(1).ClearContents()

/*本示例将 Sheet1 上所有单元格的字体设置为 8 磅“Arial”字体*/
function test()
{
    Application.Worksheets.Item("Sheet1").Cells.Font.Name = "Arial"
    Application.Worksheets.Item("Sheet1").Cells.Font.Size = 8
}
```

### Worksheet.Columns

返回一个 Range 对象，它代表活动工作表中的所有列。如果活动文档不是工作表，则 **Columns** 属性失效。

**语法**

**express.Columns**

\*express是一个代表 Worksheet 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ` ActiveSheet.Columns ` 。

此属性在应用于一个是多重选定区域的 **Range** 对象时，会只从该区域的第一个子区域中返回列。例如，如果 **Range** 对象有两个子区域 A1:B2 和 C3:D4，那么， ` Selection.Columns.Count ` 的返回值是 2，而不是 4。若要对一个可能包含多重选定区域的区域使用此属性，请测试 ` Areas.Count ` 以确定此区域内是否包含多个子区域。如果包含，请对此区域内的每个子区域进行循环。

**示例**

``` JavaScript
/*此示例将 Sheet1 的第一列（A 列）的字体设置为加粗。*/
Application.Worksheets.Item("Sheet1").Columns.Item(1).Font.Bold = true
```

### Worksheet.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 Worksheet 对象的变量。

**示例**

### Worksheet.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Worksheet 对象的变量。

### Worksheet.Names

返回一个 Names 集合，它代表所有特定于工作表的名称（使用“WorksheetName!”前缀定义的名称）。 **Names** 对象，只读。

**语法**

**express.Names**

\*express是一个代表 Worksheet 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ` ActiveWorkbook.Names ` 。

**示例**

``` JavaScript
/*本示例是将 Sheet1 中的 A1 单元格的名称定义为“myName”。*/
Application.ActiveWorkbook.Names.Add("myName", null, null, null, null, null, null, null, null, "=Sheet1!R1C1")
```

### Worksheet.Next

返回代表下一个工作表的 Worksheet 对象。

**语法**

**express.Next**

\*express是一个代表 Worksheet 对象的变量。

**说明**

如果该对象为区域，则此属性会模拟 Tab 键，但会返回下一个单元格，而不选中它。

在处于保护状态的工作表中，此属性返回下一个未锁定单元格。在未保护的工作表中，此属性总是返回紧靠指定单元格右边的单元格。

**示例**

``` JavaScript
/*此示例选定 sheet1 中下一个未锁定单元格。如果 sheet1 未保护，选定的单元格将是紧靠活动单元格右边的单元格。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Next.Select()
}
```

### Worksheet.Previous

返回代表下一个工作表的 Worksheet 对象。

**语法**

**express.Previous**

\*express是一个代表 Worksheet 对象的变量。

**说明**

如果指定对象为区域，则此属性的作用是仿效 Shift+Tab；但此属性只是返回上一单元格，并不选定它。

在保护工作表上，该属性返回上一个未锁定的单元格。在未保护的工作表上，该属性通常返回指定单元格左侧相邻的单元格。

**示例**

``` JavaScript
/*此示例选定 sheet1 中上一个未锁定单元格。如果 sheet1 未保护，选定的单元格将是紧靠活动单元格左边的单元格。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Previous.Select()
}
```

### Worksheet.Protection

返回一个 Protection 对象，该对象表示工作表的保护选项。

**语法**

**express.Protection**

\*express是一个代表 Worksheet 对象的变量。

**示例**

``` JavaScript
/*本示例对活动工作表进行保护，并判断是否能在受保护的工作表中插入列，然后将此状态通知用户。*/
function test()
{
    Application.ActiveSheet.Protect()

    // Check the ability to insert columns on a protected sheet.
    // Notify the user of this status.
    if(Application.ActiveSheet.Protection.AllowInsertingColumns == true){
        alert("The insertion of columns is allowed on this protected worksheet.")
    }
    else{
        alert("The insertion of columns is not allowed on this protected worksheet.")
    }
}
```

### Worksheet.Rows

返回一个 Range 对象，它代表指定工作表中的所有行。 **Range** 对象，只读。

**语法**

**express.Rows**

\*express是一个代表 Worksheet 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ` ActiveSheet.Rows ` 。

该属性在应用于是多个选定区域的 **Range** 对象时，只从该区域中第一个子区域内返回行。例如，如果 **Range** 对象有两个子区域：A1:B2 和 C3:D4，则 ` Selection.Rows.Count ` 返回 2 而不是 4。若要在一个可能包含多个选定区域的区域中使用此属性，请测试 ` Areas.Count ` 以确定该区域是否包含多个选择区域。如果是，请对此区域内的每个子区域进行循环，如第 3 个示例所示。

**示例**

``` JavaScript
/*此示例删除 Sheet1 的第三行。*/
Application.Worksheets.Item("Sheet1").Rows.Item(3).Delete()

/*此示例显示 Sheet1 选定区域中的行数。如果选择了多个子区域，此示例将对每一个子区域进行循环。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    let areaCount = Application.Selection.Areas.Count
    if(areaCount <= 1){
       alert("The selection contains " + Application.Selection.Rows.Count + " rows.")
    }
    else{
        for(let i = 1; i <= Application.Selection.Areas.Count; i++){
            alert("Area " + i + " of the selection contains " + Application.Selection.Areas.Rows.Count + " rows.")
        }
    }
}
```

### Worksheet.Shapes

返回一个 Shapes 集合，它代表工作表上的所有形状。只读。

**语法**

**express.Shapes**

\*express是一个代表 Worksheet 对象的变量。

**示例**

``` JavaScript
/*此示例向第一张工作表中添加兰色的虚线。*/
function test()
{
    let Wshapes = Application.Worksheets.Item(1).Shapes.AddLine(10, 10, 250, 250).Line
    Wshapes.DashStyle = Application.Enum.msoLineDashDotDot
    Wshapes.ForeColor.RGB = (50, 0, 128)
}
```

### Worksheet.Type

返回一个代表工作表类型的 XlSheetType 值。

**语法**

**express.Type**

\*express是一个代表 Worksheet 对象的变量。

### Worksheet.UsedRange

返回一个 Range 对象，该对象表示指定工作表上所使用的区域。只读。

**语法**

**express.UsedRange**

\*express是一个代表 Worksheet 对象的变量。

**示例**

``` JavaScript
/*本示例选定 Sheet1 中的已用区域。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveSheet.UsedRange.Select()
}
```

### Worksheet.Visible

返回或设置一个 XlSheetVisibility 值，它确定对象是否可见。

**语法**

**express.Visible**

\*express是一个代表 Worksheet 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏 Sheet1。*/
Application.Worksheets.Item("Sheet1").Visible = false

/*本示例将 Sheet1 设置为可见。*/
Application.Worksheets.Item("Sheet1").Visible = true

/*本示例使活动工作簿中的每一张工作表可见。*/
for(let i = 1; i <= Sheets.Count; i++){
    Application.Sheets.Visible = true
}

/*本示例新建一张工作表，然后将其 Visible 属性设为 xlVeryHidden。要引用该工作表，可使用其对象变量 newSheet，如本示例最后一行所示。若要在另一过程中使用 newSheet 对象变量，必须在模块的第一行中将其声明为公用变量 (Public newSheet As Object)，然后才能执行 Sub 或 Function 过程。*/
function test()
{
    let newSheet = Application.Worksheets.Add()
    newSheet.Visible = Application.Enum.xlVeryHidden
    newSheet.Range("A1:D4").Formula = "=RAND()"
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Worksheet](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# WorksheetFunction 对象

## 简介

用作可从 Visual Basic 中调用的 ET 工作表函数的容器。

## 说明

使用 WorksheetFunction 属性可返回 WorksheetFunction 对象。下例显示给区域 A1:A10 应用 Min 工作表函数的结果。

``` JavaScript
function test(){
let myRange = Application.Worksheets.Item("Sheet1").Range("A1:C10")
let answer = Application.WorksheetFunction.Min(myRange)
alert(answer)
}
```

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.AccrInt">AccrInt</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回定期付息有价证券的应计利息。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.AccrIntM">AccrIntM</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回到期一次性付息有价证券的应计利息。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Acos">Acos</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回数字的反余弦值。反余弦值是余弦值为 <em>Arg1</em> 的角度。返回的角度以弧度给出，范围从 0（零）到 pi。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Acosh">Acosh</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回数字的反双曲余弦值。数字必须大于或等于 1。反双曲余弦值的双曲余弦值为 <em>Arg1</em> ，因此 Acosh(Cosh(number)) 等于 <em>Arg1</em> 。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Aggregate">Aggregate</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回列表或数据库中的一个聚合。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.AmorDegrc">AmorDegrc</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回每个结算期的折旧值。此函数为法国会计系统提供。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.AmorLinc">AmorLinc</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回每个结算期的折旧值。此函数为法国会计系统提供。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.And">And</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果其所有参数都为 TRUE，则返回 TRUE；如果一个或多个参数为 FALSE，则返回 FALSE。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Asc">Asc</a></td>
<td style="text-align: left;" data-valian="middle"><p>对于双字节字符集 (DBCS) 语言，将全角（双字节）字符更改为半角（单字节）字符。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Asin">Asin</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个数字的反正弦值。反正弦值是一个角度，这个角度的正弦值为 <em>Arg1</em> 。返回的角度采用弧度的形式，其范围从 -pi/2 到 pi/2。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Asinh">Asinh</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回数字的反双曲正弦值。反双曲正弦值的双曲正弦值为 <em>Arg1</em> ，因此 Asinh(Sinh(number)) 等于 <em>Arg1</em> 。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Atan2">Atan2</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定的 x 和 y 坐标值的反正切值。反正切值是角度，是从 x 轴到通过原点 (0, 0) 和坐标点 (x_num, y_num) 的直线之间的夹角。该角度用弧度给出，介于 -pi 和 pi 之间（不包括 -pi）。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Atanh">Atanh</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回数字的反双曲正切值。数字必须介于 -1 和 1 之间（不包括 -1 和 1）。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.AveDev">AveDev</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回多个数据点与其平均值的绝对偏差的平均值。AveDev 是对数据集中可变性的度量。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Average">Average</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回参数的平均值（算术平均值）。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.AverageIf">AverageIf</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回区域内满足给定条件的所有单元格的平均值（算术平均值）。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.AverageIfs">AverageIfs</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回满足多个条件的所有单元格的平均值（算术平均值）。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.BahtText">BahtText</a></td>
<td style="text-align: left;" data-valian="middle"><p>将数字转换为泰语文本并添加后缀“泰铢”。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.BesselI">BesselI</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回修正的 Bessel 函数，它等效于计算纯虚数参数值的 Bessel 函数。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.BesselJ">BesselJ</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Bessel 函数。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">21</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.BesselK">BesselK</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回修正 Bessel 函数，它等效于根据纯虚数参数计算的 Bessel 函数。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">22</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.BesselY">BesselY</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Bessel 函数，该函数也称为 Weber 函数或 Neumann 函数。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">23</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Beta_Dist">Beta_Dist</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Beta 累积分布函数。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">24</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Beta_Inv">Beta_Inv</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定的 Beta 分布的累积分布函数的反函数。即，如果 probability = Beta_Dist(x,...)，则 Beta_Inv(probability,...) = x。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">25</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.BetaDist">BetaDist</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Beta 累积分布函数。</p>
<p>要点 ??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。</p>
<p>有关新函数的详细信息，请参阅 <span>Beta_Dist</span> 方法。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">26</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.BetaInv">BetaInv</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定的 Beta 分布的累积分布函数的反函数。即，如果 probability = BetaDist(x,...)，则 BetaInv(probability,...) = x。</p>
<p>要点 ??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。</p>
<p>有关新函数的详细信息，请参阅 <span>Beta_Inv</span> 方法。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">27</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Bin2Dec">Bin2Dec</a></td>
<td style="text-align: left;" data-valian="middle"><p>将二进制数转换为十进制数。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">28</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Bin2Hex">Bin2Hex</a></td>
<td style="text-align: left;" data-valian="middle"><p>将二进制数转换为十六进制数。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">29</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Bin2Oct">Bin2Oct</a></td>
<td style="text-align: left;" data-valian="middle"><p>将二进制数转换为八进制数。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">30</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Binom_Dist">Binom_Dist</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一元二项式分布的概率。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">31</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Binom_Inv">Binom_Inv</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一元二项式概率分布的反函数。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">32</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.BinomDist">BinomDist</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一元二项式分布的概率。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">33</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Ceiling">Ceiling</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回向上舍入（远离零）到最接近的 significance 的倍数的 number。</p>
<p>要点 ??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。</p>
<p>有关新函数的详细信息，请参阅 <span>Ceiling_Precise</span> 方法。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">34</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.Ceiling_Precise">Ceiling_Precise</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回舍入到最接近的有效位倍数的指定数字。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">35</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.ChiDist">ChiDist</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 χ2 分布的单尾概率。</p>
<p>此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。</p>
<p>有关新函数的详细信息，请参阅 <span>ChiSq_Dist_RT</span> 和 <span>ChiSq_Dist</span> 方法。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">36</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.ChiInv">ChiInv</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 χ2 分布的单尾概率的反函数。</p>
<p>此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。</p>
<p>有关新函数的详细信息，请参阅 <span>ChiSq_Inv_RT</span> 和 <span>ChiSq_Inv</span> 方法。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">37</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.ChiSq_Dist">ChiSq_Dist</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 χ2 分布。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">38</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorksheetFunction.ChiSq_Dist_RT">ChiSq_Dist_RT</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 χ2 分布的右尾概率。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                                                                                      |
|:----:|:----------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#WorksheetFunction.Application) | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 Application 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 Application 对象。可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序。只读。 |
|  2   | [Creator](#WorksheetFunction.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                |
|  3   | [Parent](#WorksheetFunction.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                              |

## 成员方法

### WorksheetFunction.AccrInt

返回定期付息有价证券的应计利息。

**语法**

**express.AccrInt ( *Arg1* , *Arg2* , *Arg3* , *Arg4* , *Arg5* , *Arg6* , *Arg7* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                |
|--------|-----------|----------|-------------------------------------|
| *Arg1* | 必选      | Variant  | Issue date - 债券的发行日期。       |
| *Arg2* | 必选      | Variant  | First interest - 债券的首次计息日。 |
| *Arg3* | 必选      | Variant  | Settlement - 债券的结算日。         |
| *Arg4* | 必选      | Variant  | Rate - 债券的年息票利率。           |
| *Arg5* | 必选      | Variant  | Par - 债券的票面价值。              |
| *Arg6* | 必选      | Variant  | Frequency - 每年支付息票的次数。    |
| *Arg7* | 可选      | Variant  | Basis - 要使用的日计数基准类型。    |

**返回值**

Double

**说明**

日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

下表描述了可用于 *Arg5* 的值。

| 基准     | 日计数基准                       |
|----------|----------------------------------|
| 0 或省略 | 美国（美国证券交易商协会）30/360 |
| 1        | 实际天数/实际天数                |
| 2        | 实际天数/360                     |
| 3        | 实际天数/365                     |
| 4        | 欧洲 30/360                      |

### WorksheetFunction.AccrIntM

返回到期一次性付息有价证券的应计利息。

**语法**

**express.AccrIntM ( *Arg1* , *Arg2* , *Arg3* , *Arg4* , *Arg5* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                   |
|--------|-----------|----------|--------------------------------------------------------|
| *Arg1* | 必选      | Variant  | 债券的发行日。                                         |
| *Arg2* | 必选      | Variant  | 债券的到期日。                                         |
| *Arg3* | 必选      | Variant  | 债券的年息票利率。                                     |
| *Arg4* | 必选      | Variant  | 债券的票面价值。如果省略 par，ACCRINTM 则使用 \$1000。 |
| *Arg5* | 可选      | Variant  | 要使用的日计数基准类型。                               |

**返回值**

Double

**说明**

**要点：** 日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

下表描述了可用于 *Arg5* 的值。

| 基准     | 日计数基准                       |
|----------|----------------------------------|
| 0 或省略 | 美国（美国证券交易商协会）30/360 |
| 1        | 实际天数/实际天数                |
| 2        | 实际天数/360                     |
| 3        | 实际天数/365                     |
| 4        | 欧洲 30/360                      |

### 以下列表包含在使用 **ACCRINTM** 时要注意的信息

- ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日是序列号 1，2008 年 1 月 1 日是序列号 39,448，这是因为它距 1900 年 1 月 1 日有 39,448 天。

- Issue、maturity 和 basis 将被截尾取整。

- 如果 issue 或 maturity 是无效的日期，则 ACCRINTM 将生成一个错误。

- 如果 rate ≤ 0 或 par ≤ 0，则 ACCRINTM 将产生一个错误。

- 如果 basis \< 0 或 basis \> 4，则 ACCRINTM 将产生一个错误。

- 如果 issue ≥ maturity，则 ACCRINTM 将产生一个错误。

- ACCRINTM 的计算公式如下：

  其中：

  A = 按月计算的应计天数。在计算到期付息的利息时指发行日与到期日之间的天数。

  D = 年基准数。

### WorksheetFunction.Acos

返回数字的反余弦值。反余弦值是余弦值为 *Arg1* 的角度。返回的角度以弧度给出，范围从 0（零）到 pi。

**语法**

**express.Acos ( *Arg1* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                      |
|--------|-----------|----------|-------------------------------------------|
| *Arg1* | 必选      | Double   | 所需角度的余弦值，必须介于 -1 到 1 之间。 |

**说明**

如果要将结果从弧度转换为度，则请将结果乘以 180/PI()，或者使用 Degrees 方法。

### WorksheetFunction.Acosh

返回数字的反双曲余弦值。数字必须大于或等于 1。反双曲余弦值的双曲余弦值为 *Arg1* ，因此 Acosh(Cosh(number)) 等于 *Arg1* 。

**语法**

**express.Acosh ( *Arg1* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                      |
|--------|-----------|----------|---------------------------|
| *Arg1* | 必选      | Double   | 等于或大于 1 的任意实数。 |

**返回值**

Double

### WorksheetFunction.Aggregate

返回列表或数据库中的一个聚合。

**语法**

**express.Aggregate ( *Arg1* , *Arg2* , *Arg3* , *Arg4 - Arg 30* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
|-----------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Arg1*          | 必选      | Double   | Function_num - 一个介于 1 到 19 之间的数字，用于指定要使用的函数。 Function_num 函数 1 AVERAGE 2 COUNT 3 COUNTA 4 MAX 5 MIN 6 PRODUCT 7 STDEV.S 8 STDEV.P 9 SUM 10 VAR.S 11 VAR.P 12 MEDIAN 13 MODE.SNGL 14 LARGE 15 SMALL 16 PERCENTILE.INC 17 QUARTILE.INC 18 PERCENTILE.EXC 19 QUARTILE.EXC Arg2 必选 Double 选项 - 一个数值，用于确定要在函数的求值范围中忽略的值。 选项 行为 0 或省略 忽略嵌套的 SUBTOTAL 和 AGGREGATE 函数 1 忽略隐藏行以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 2 忽略错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 3 忽略隐藏行、错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 4 忽略空值 5 忽略隐藏行 6 忽略错误值 7 忽略隐藏行和错误值 |
| *Arg2*          | 必选      | Double   | 选项 - 一个数值，用于确定要在函数的求值范围中忽略的值。 选项 行为 0 或省略 忽略嵌套的 SUBTOTAL 和 AGGREGATE 函数 1 忽略隐藏行以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 2 忽略错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 3 忽略隐藏行、错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 4 忽略空值 5 忽略隐藏行 6 忽略错误值 7 忽略隐藏行和错误值                                                                                                                                                                                                                                                                                                                 |
| *Arg3*          | 必选      | Range    | Ref1 - 带有多个要对其计算聚合值的数值参数的函数的第一个数值参数。                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| *Arg4 - Arg 30* | 可选      | Variant  | Ref2 - Ref30 - 要对其计算聚合值的 2 到 30 个数值参数。                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |

**返回值**

Double

**说明**

- 以下约束将会根据 **Function_num** 值应用于 Ref 参数 ( *Arg3 - Arg 30* )。
  <table>
  <colgroup>
  <col style="width: 25%" />
  <col style="width: 25%" />
  <col style="width: 25%" />
  <col style="width: 25%" />
  </colgroup>
  <thead>
  <tr class="header">
  <th>Function_num</th>
  <th>Ref1</th>
  <th>Ref2</th>
  <th>Ref3、Ref4…</th>
  </tr>
  </thead>
  <tbody>
  <tr class="odd">
  <td>1-13</td>
  <td><strong>有效类型：</strong>
  <ul>
  <li>任何单元格引用</li>
  <li>并集</li>
  <li>交集</li>
  <li>定义的名称</li>
  <li>结构化引用</li>
  </ul>
  <strong>无效类型：</strong>
  <ul>
  <li>实际数据</li>
  <li>数组</li>
  </ul></td>
  <td><strong>有效类型：</strong>
  <ul>
  <li>任何单元格引用</li>
  <li>并集</li>
  <li>交集</li>
  <li>定义的名称</li>
  <li>结构化引用</li>
  </ul>
  <strong>无效类型：</strong>
  <ul>
  <li>实际数据</li>
  <li>数组</li>
  </ul></td>
  <td><strong>有效类型：</strong>
  <ul>
  <li>任何单元格引用</li>
  <li>并集</li>
  <li>交集</li>
  <li>定义的名称</li>
  <li>结构化引用</li>
  </ul>
  <strong>无效类型：</strong>
  <ul>
  <li>实际数据</li>
  <li>数组</li>
  </ul></td>
  </tr>
  <tr class="even">
  <td>14-17</td>
  <td><strong>有效类型：</strong>
  <ul>
  <li>任何单元格引用</li>
  <li>并集</li>
  <li>交集</li>
  <li>定义的名称</li>
  <li>结构化引用</li>
  <li>实际数据</li>
  <li>数组</li>
  </ul></td>
  <td><strong>有效类型：</strong>
  <ul>
  <li>任何单元格引用</li>
  <li>并集</li>
  <li>交集</li>
  <li>定义的名称</li>
  <li>结构化引用</li>
  <li>实际数据</li>
  <li>数组</li>
  </ul></td>
  <td><strong>不允许使用引用</strong></td>
  </tr>
  <tr class="odd">
  <td></td>
  <td></td>
  <td></td>
  <td></td>
  </tr>
  </tbody>
  </table>
- 如果需要但未提供第二个 Ref 参数，AGGREGATE 将会返回错误 \#VALUE!。
- 如果有一个或多个引用为三维引用，AGGREGATE 将会返回错误值 \#VALUE!。

### WorksheetFunction.AmorDegrc

返回每个结算期的折旧值。此函数为法国会计系统提供。

**语法**

**express.AmorDegrc ( *Arg1* , *Arg2* , *Arg3* , *Arg4* , *Arg5* , *Arg6* , *Arg7* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                     |
|--------|-----------|----------|--------------------------|
| *Arg1* | 必选      | Variant  | 资产原值。               |
| *Arg2* | 必选      | Variant  | 资产的购买日。           |
| *Arg3* | 必选      | Variant  | 第一个结算期的结束日期。 |
| *Arg4* | 必选      | Variant  | 资产在寿命结束时的残值。 |
| *Arg5* | 必选      | Variant  | 结算期。                 |
| *Arg6* | 必选      | Variant  | 折旧率。                 |
| *Arg7* | 可选      | Variant  | 要使用的年基准数。       |

**说明**

如果某项资产是在结算期的中期购入的，则应考虑按比例折旧。该方法与 AmorLinc 相似，不同之处在于该方法中用于计算的折旧系数取决于资产的寿命。

下表描述了 *Arg7* 中使用的值。

| Basis    | 日期系统                |
|----------|-------------------------|
| 0 或省略 | 360 天（NASD 方法）     |
| 1        | 实际天数                |
| 3        | 一年 365 天             |
| 4        | 一年 360 天（欧洲方法） |

- ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。ET for the Macintosh 使用另外一个默认日期系统。
- 此函数将返回折旧值，截止到资产寿命的最后一个期间，或直到累积折旧值大于资产原值减去残值后所得的结果。
- 折旧系数为：
  | 资产的寿命 (1/rate) | 折旧系数 |
  |---------------------|----------|
  | 3 到 4 年           | 1.5      |
  | 5 到 6 年           | 2        |
  | 6 年以上            | 2.5      |
- 最后一个期间之前的那个期间的折旧率将增加到 50%，最后一个期间的折旧率将增加到 100%。
- 如果资产的寿命在 0（零）到 1、1 到 2、2 到 3 或 4 到 5 之间，则将返回错误值 \#NUM!。

### WorksheetFunction.AmorLinc

返回每个结算期的折旧值。此函数为法国会计系统提供。

**语法**

**express.AmorLinc ( *Arg1* , *Arg2* , *Arg3* , *Arg4* , *Arg5* , *Arg6* , *Arg7* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                     |
|--------|-----------|----------|--------------------------|
| *Arg1* | 必选      | Variant  | 资产原值。               |
| *Arg2* | 必选      | Variant  | 资产的购买日。           |
| *Arg3* | 必选      | Variant  | 第一个结算期的结束日期。 |
| *Arg4* | 必选      | Variant  | 资产在寿命结束时的残值。 |
| *Arg5* | 必选      | Variant  | 结算期。                 |
| *Arg6* | 必选      | Variant  | 折旧率。                 |
| *Arg7* | 可选      | Variant  | 要使用的年基准数。       |

**返回值**

Double

**说明**

如果某项资产是在结算期的中期购入的，则应考虑按比例折旧。

下表描述了 *Arg7* 所使用的值。

| Basis    | 日期系统                |
|----------|-------------------------|
| 0 或省略 | 360 天（NASD 方法）     |
| 1        | 实际天数                |
| 3        | 一年 365 天             |
| 4        | 一年 360 天（欧洲方法） |

**要点** ??日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。ET for the Macintosh 使用另外一个默认日期系统。

### WorksheetFunction.And

如果其所有参数都为 TRUE，则返回 TRUE；如果一个或多个参数为 FALSE，则返回 FALSE。

**语法**

**express.And ( *Arg1 - Arg30* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                               |
|----------------|-----------|----------|----------------------------------------------------|
| *Arg1 - Arg30* | 必选      | Variant  | 1 到 30 个需要进行检验的条件，可为 TRUE 或 FALSE。 |

**返回值**

Boolean

**说明**

- 参数的计算结果必须为逻辑值（如 TRUE 或 FALSE），或者参数必须是包含逻辑值的数组或引用。
- 如果数组或引用参数中包含文本或空单元格，则这些值将被忽略。
- 如果指定区域内不包含逻辑值，则此方法将生成一个错误值。

### WorksheetFunction.Asc

对于双字节字符集 (DBCS) 语言，将全角（双字节）字符更改为半角（单字节）字符。

**语法**

**express.Asc ( *Arg1* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                   |
|--------|-----------|----------|----------------------------------------------------------------------------------------|
| *Arg1* | 必选      | String   | 文本或对包含要更改的文本的单元格的引用。如果文本中不包含任何全角字母，则文本不会更改。 |

**返回值**

String

### WorksheetFunction.Asin

返回一个数字的反正弦值。反正弦值是一个角度，这个角度的正弦值为 *Arg1* 。返回的角度采用弧度的形式，其范围从 -pi/2 到 pi/2。

**语法**

**express.Asin ( *Arg1* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                      |
|--------|-----------|----------|-------------------------------------------|
| *Arg1* | 必选      | Double   | 所需角度的正弦值，必须介于 -1 到 1 之间。 |

**返回值**

Double

**说明**

要用度表示反正弦值，请将结果乘以 180/PI( ) 或使用 Degrees 方法。

### WorksheetFunction.Asinh

返回数字的反双曲正弦值。反双曲正弦值的双曲正弦值为 *Arg1* ，因此 Asinh(Sinh(number)) 等于 *Arg1* 。

**语法**

**express.Asinh ( *Arg1* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|------------|
| *Arg1* | 必选      | Double   | 任意实数。 |

**返回值**

Double

### WorksheetFunction.Atan2

返回指定的 x 和 y 坐标值的反正切值。反正切值是角度，是从 x 轴到通过原点 (0, 0) 和坐标点 (x_num, y_num) 的直线之间的夹角。该角度用弧度给出，介于 -pi 和 pi 之间（不包括 -pi）。

**语法**

**express.Atan2 ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明          |
|--------|-----------|----------|---------------|
| *Arg1* | 必选      | Double   | 点的 x 坐标。 |
| *Arg2* | 必选      | Double   | 点的 y 坐标。 |

**返回值**

Double

**说明**

- 结果为正表示从 x 轴逆时针旋转的角度；结果为负表示从 x 轴顺时针旋转的角度。
- Atan2(a,b) 等于 Atan(b/a)，不同之处在于 Atan2 中的 a 可等于 0。
- 如果 x_num 和 y_num 为 0，则 Atan2 将返回一个错误值。
- 要用度表示反正切值，请将结果乘以 180/PI( ) 或使用 Degrees 方法。

### WorksheetFunction.Atanh

返回数字的反双曲正切值。数字必须介于 -1 和 1 之间（不包括 -1 和 1）。

**语法**

**express.Atanh ( *Arg1* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                          |
|--------|-----------|----------|-------------------------------|
| *Arg1* | 必选      | Double   | 介于 1 和 -1 之间的任意实数。 |

**返回值**

Double

**说明**

反双曲正切值的双曲正切值为 *Arg1* ，因此 Atanh(Tanh(number)) 等于 *Arg1* 。

### WorksheetFunction.AveDev

返回多个数据点与其平均值的绝对偏差的平均值。AveDev 是对数据集中可变性的度量。

**语法**

**express.AveDev ( *Arg1 - Arg 30* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                              |
|-----------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------|
| *Arg1 - Arg 30* | 必选      | Variant  | 要计算其绝对偏差平均值的 1 到 30 个参数。也可以不使用这种用逗号分隔参数的形式，而使用一个数组或一个对数组的引用。 |

**返回值**

Double

**说明**

- AveDev 受输入数据中的度量单位的影响。
- 参数要么是数字，要么是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 平均偏差的计算公式为：

### WorksheetFunction.Average

返回参数的平均值（算术平均值）。

**语法**

**express.Average ( *Arg1 - Arg30* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                |
|----------------|-----------|----------|-------------------------------------|
| *Arg1 - Arg30* | 必选      | Variant  | 要求其平均值的 1 到 30 个数值参数。 |

**返回值**

Double

**说明**

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。
- 如果要对引用中的逻辑值和数字的文本表示进行计算，请使用 AVERAGEA 函数。

| 注释                                                                                         |
|----------------------------------------------------------------------------------------------|
| Average 方法衡量趋中性，趋中性是统计分布中一组数字的中心位置。三种最常见的趋中性衡量方式为： |

- **平均值** 即算术平均值，是通过将一组数字相加，然后除以数字个数得到的。例如，2、3、3、5、7 和 10 的平均值为 30 除以 6，即为 5。
- **中值** 是一组数字中居于中间的数，即在这组数字中，有一半的数字的值比它大，有一半的数字的值比它小。例如，2、3、3、5、7 和 10 的中值为 4。
- **众值** 即在一组数字中出现次数最多的数字。例如，2、3、3、5、7 和 10 的众值为 3。

对于对称分布的一组数字，这三种趋中性衡量方式完全相同。对于偏态分布的一组数字，这些衡量方式可能会不同。

**提示** 对单元格求平均值时，应牢记空单元格与含零值单元格之间的区别，尤其在已经清除了 **“视图”** 选项卡上的 **“零值”** 复选框（ **“工具”** 菜单上的 **“选项”** 命令）的情况下更应如此。空单元格不计算在内，但会计算含零值的单元格。

### WorksheetFunction.AverageIf

返回区域内满足给定条件的所有单元格的平均值（算术平均值）。

**语法**

**express.AverageIf ( *Arg1* , *Arg2* , *Arg3* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                                  |
|--------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------|
| *Arg1* | 必选      | Range    | 要求其平均值的一个或多个单元格。                                                                                                      |
| *Arg2* | 必选      | Variant  | 定义将对哪些单元格求平均值的条件，其形式可以为数字、表达式、单元格引用或文本。例如，条件可以表示为 32、"32"、"\>32"、"apples" 或 B4。 |
| *Arg3* | 可选      | Variant  | 要求其平均值的实际单元格集合。如果省略，则使用 range。                                                                                |

**返回值**

Double

**说明**

- range 内包含 TRUE 或 FALSE 的单元格将被忽略。
- 如果 range 或 average_range 中的某个单元格为空单元格，则 AverageIf 会将其忽略。
- 如果条件中的单元格为空，则 AverageIf 会将其作为 0 值处理。
- 如果区域中没有满足条件的单元格，则 AverageIf 将生成一个错误值。
- 可以在条件中使用通配符，包括问号 (?) 和星号 (\*)。问号可匹配任意的单个字符；星号可匹配任意一串字符。如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。
- Average_range 的大小和形状不必与 range 相同。求其平均值的实际单元格的确定方法如下：使用 average_range 中左上角的单元格作为起始单元格，然后将与 range 的大小和形状对应的所有单元格包含到其中。例如：
  | 如果 range 为 | 并且 average_range 为 | 则实际纳入计算的单元格为 |
  |---------------|-----------------------|--------------------------|
  | A1:A5         | B1:B5                 | B1:B5                    |
  | A1:A5         | B1:B3                 | B1:B5                    |
  | A1:B4         | C1:D4                 | C1:D4                    |
  | A1:B4         | C1:C2                 | C1:D4                    |

| 注释                                                                                           |
|------------------------------------------------------------------------------------------------|
| AverageIf 方法衡量趋中性，趋中性是统计分布中一组数字的中心位置。三种最常见的趋中性衡量方式为： |

- **平均值** 即算术平均值，是通过将一组数字相加，然后除以数字个数得到的。例如，2、3、3、5、7 和 10 的平均值为 30 除以 6，即为 5。
- **中值** 是一组数字中居于中间的数，即在这组数字中，有一半的数字的值比它大，有一半的数字的值比它小。例如，2、3、3、5、7 和 10 的中值为 4。
- **众值** 即在一组数字中出现次数最多的数字。例如，2、3、3、5、7 和 10 的众值为 3。

对于对称分布的一组数字，这三种趋中性衡量方式完全相同。对于偏态分布的一组数字，这些衡量方式可能会不同。

### WorksheetFunction.AverageIfs

返回满足多个条件的所有单元格的平均值（算术平均值）。

**语法**

**express.AverageIfs ( *Arg1 - Arg30* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                 |
|----------------|-----------|----------|--------------------------------------|
| *Arg1 - Arg30* | 必选      | Range    | 在其中计算相关条件的一个或多个区域。 |

**返回值**

Double

**说明**

- 如果 average_range 中的某个单元格为空单元格，则 AverageIfs 会将其忽略。
- 如果条件区域中的某个单元格为空，则 AverageIfs 会将其作为 0 值处理。
- range 中包含 TRUE 的单元格作为 1 计算；range 中包含 FALSE 的单元格作为 0（零）计算。
- 仅在每个单元格中指定的对应条件都为 True 时，才会在平均值计算过程中使用 average_range 中的该单元格。
- 如果 average_range 中的单元格均为空，或包含无法转换为数字的文本值，则 AverageIfs 将生成一个错误。
- 如果没有满足所有条件的单元格，则 AverageIfs 将生成一个错误值。
- 可以在条件中使用通配符，包括问号 (?) 和星号 (\*)。问号可匹配任意的单个字符；星号可匹配任意一串字符。如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。
- 每个 criteria_range 的大小和形状不必与 average_range 相同。计算其平均值的实际单元格的确定方法如下：使用 criteria_range 中左上角的单元格作为起始单元格，然后将与 range 的大小和形状对应的所有单元格包含到其中。例如：
  | 如果 average_range 为 | 并且 criteria_range 为 | 则实际纳入计算的单元格为 |
  |-----------------------|------------------------|--------------------------|
  | A1:A5                 | B1:B5                  | B1:B5                    |
  | A1:A5                 | B1:B3                  | B1:B5                    |
  | A1:B4                 | C1:D4                  | C1:D4                    |
  | A1:B4                 | C1:C2                  | C1:D4                    |

| 注释                                                                                            |
|-------------------------------------------------------------------------------------------------|
| AverageIfs 函数衡量趋中性，趋中性是统计分布中一组数字的中心位置。三种最常见的趋中性衡量方式为： |

- **平均值** 即算术平均值，是通过将一组数字相加，然后除以数字个数得到的。例如，2、3、3、5、7 和 10 的平均值为 30 除以 6，即为 5。
- **中值** 是一组数字中居于中间的数，即在这组数字中，有一半的数字的值比它大，有一半的数字的值比它小。例如，2、3、3、5、7 和 10 的中值为 4。
- **众值** 即在一组数字中出现次数最多的数字。例如，2、3、3、5、7 和 10 的众值为 3。

对于对称分布的一组数字，这三种趋中性衡量方式完全相同。对于偏态分布的一组数字，这些衡量方式可能会不同。

### WorksheetFunction.BahtText

将数字转换为泰语文本并添加后缀“泰铢”。

**语法**

**express.BahtText ( *Arg1* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                 |
|--------|-----------|----------|----------------------------------------------------------------------|
| *Arg1* | 必选      | Double   | 要转换成文本的数字、对包含数字的单元格的引用或计算结果为数字的公式。 |

**返回值**

String

**说明**

在 ET for Windows 中，可以使用 **“控制面板”** 中的 “区域设置”或 **“区域选项”** 将泰铢格式更改为其他样式。

在 ET for the Macintosh 中，可以使用 **“数字控制面板”** 将泰铢数字格式更改为其他样式。

### WorksheetFunction.BesselI

返回修正的 Bessel 函数，它等效于计算纯虚数参数值的 Bessel 函数。

**语法**

**express.BesselI ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                               |
|--------|-----------|----------|----------------------------------------------------|
| *Arg1* | 必选      | Variant  | 要计算的函数值。                                   |
| *Arg2* | 必选      | Variant  | Bessel 函数的阶。如果 n 不是整数，则将被截尾取整。 |

**返回值**

Double

**说明**

- 如果 x 是非数值，BesselI 将返回错误值 \#VALUE!。
- 如果 n 是非数值，BesselI 将生成一个错误值。
- 如果 n \< 0，BesselI 将生成一个错误值。
- 变量 x 的 n 阶修正 Bessel 函数为：

### WorksheetFunction.BesselJ

返回 Bessel 函数。

**语法**

**express.BesselJ ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                               |
|--------|-----------|----------|----------------------------------------------------|
| *Arg1* | 必选      | Variant  | 计算的函数值。                                     |
| *Arg2* | 必选      | Variant  | Bessel 函数的阶。如果 n 不是整数，则将被截尾取整。 |

**返回值**

Double

**说明**

- 如果 x 为非数值型，则 BesselJ 将生成一个错误值。

- 如果 n 为非数值型，则 BesselJ 将返回一个错误值。

- 如果 n \< 0，则 BesselJ 将生成一个错误值。

- 变量 x 的 n 阶 Bessel 函数为：

  其中：

  为 γ 函数。

### WorksheetFunction.BesselK

返回修正 Bessel 函数，它等效于根据纯虚数参数计算的 Bessel 函数。

**语法**

**express.BesselK ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                        |
|--------|-----------|----------|---------------------------------------------|
| *Arg1* | 必选      | Variant  | 计算的函数值。                              |
| *Arg2* | 必选      | Variant  | 函数的阶。如果 n 不是整数，则将被截尾取整。 |

**返回值**

Double

**说明**

- 如果 x 为非数值型，则 BesselK 将生成一个错误值。

- 如果 n 为非数值型，则 BesselK 将生成一个错误值。

- 如果 n \< 0，则 BesselK 将生成一个错误值。

- 变量 x 的 n 阶修正 Bessel 函数为：

  其中 Jn 和 Yn 分别为 J 和 Y 的 Bessel 函数。

### WorksheetFunction.BesselY

返回 Bessel 函数，该函数也称为 Weber 函数或 Neumann 函数。

**语法**

**express.BesselY ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                        |
|--------|-----------|----------|---------------------------------------------|
| *Arg1* | 必选      | Variant  | 计算的函数值。                              |
| *Arg2* | 必选      | Variant  | 函数的阶。如果 n 不是整数，则将被截尾取整。 |

**返回值**

Double

**说明**

- 如果 x 为非数值型，则 BesselY 将生成一个错误值。
- 如果 n 为非数值型，则 BesselY 将生成一个错误值。
- 如果 n \< 0，则 BesselY 将生成一个错误值。
- 变量 x 的 n 阶 Bessel 函数为：

### WorksheetFunction.Beta_Dist

返回 Beta 累积分布函数。

**语法**

**express.Beta_Dist ( *Arg1* , *Arg2* , *Arg3* , *Arg4* , *Arg5* , *Arg6* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                                  |
|--------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------|
| *Arg1* | 必选      | Double   | 介于 A 和 B 之间的值，用来计算函数的值。                                                                                              |
| *Arg2* | 必选      | Double   | 该分布的 Alpha 参数。                                                                                                                 |
| *Arg3* | 必选      | Double   | 该分布的 Beta 参数。                                                                                                                  |
| *Arg4* | 可选      | Variant  | cumulative - 一个决定函数形式的逻辑值。如果 cumulative 为 TRUE，则 BETA.DIST 将返回累积分布函数；如果为 FALSE，则将返回概率密度函数。 |
| *Arg5* | 可选      | Variant  | x 所属区间的可选下界。                                                                                                                |
| *Arg6* | 可选      | Variant  | x 所属区间的可选上界。                                                                                                                |

**返回值**

Double

**说明**

Beta 分布通常用于研究样本中某些事物变化的百分比，例如人们一天中用来看电视的时间所占的比例：

- 如果任一参数为非数值型，则 Beta_Dist 将返回错误值 \#VALUE!。
- 如果 Alpha ≤ 0 或 Beta ≤ 0，则 Beta_Dist 将生成一个错误值。
- 如果 x \< A、x \> B 或 A = B，则 Beta_Dist 将生成一个错误值。
- 如果忽略 A 和 B 的值（即忽略下界和上界），则 Beta_Dist 将使用标准累积 Beta 分布，使 A = 0 且 B = 1。

### WorksheetFunction.Beta_Inv

返回指定的 Beta 分布的累积分布函数的反函数。即，如果 probability = Beta_Dist(x,...)，则 Beta_Inv(probability,...) = x。

**语法**

**express.Beta_Inv ( *Arg1* , *Arg2* , *Arg3* , *Arg4* , *Arg5* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                     |
|--------|-----------|----------|--------------------------|
| *Arg1* | 必选      | Double   | 与 Beta 分布相关的概率。 |
| *Arg2* | 必选      | Double   | 该分布的 Alpha 参数。    |
| *Arg3* | 必选      | Double   | 该分布的 Beta 参数。     |
| *Arg4* | 可选      | Variant  | x 所属区间的可选下界。   |
| *Arg5* | 可选      | Variant  | x 所属区间的可选上界。   |

**返回值**

Double

**说明**

在项目计划中，如果给定预期完成时间和变化率，则可使用 Beta 分布来建立可能完成时间的模型：

- 如果任一参数为非数值型，则 Beta_Inv 将生成一个错误值。
- 如果 Alpha ≤ 0 或 Beta ≤ 0，则 Beta_Inv 将生成一个错误值。
- 如果 probability ≤ 0 或 probability \> 1，则 Beta_Inv 将生成一个错误值。
- 如果忽略 A 和 B 的值（即忽略下界和上界），则 Beta_Inv 将使用标准累积 Beta 分布，使 A = 0 且 B = 1。

如果给定 probability 的值，Beta_Inv 将使用 Beta_Dist(x, alpha, beta, TRUE, A, B) = probability 求解值 x。因此，Beta_Inv 的精度取决于 Beta_Dist 的精度。Beta_Inv 使用迭代搜索技术。

### WorksheetFunction.BetaDist

返回 Beta 累积分布函数。

**要点** ??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 Beta_Dist 方法。

**语法**

**express.BetaDist ( *Arg1* , *Arg2* , *Arg3* , *Arg4* , *Arg5* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                     |
|--------|-----------|----------|------------------------------------------|
| *Arg1* | 必选      | Double   | 介于 A 和 B 之间的值，用来计算函数的值。 |
| *Arg2* | 必选      | Double   | 分布参数。                               |
| *Arg3* | 必选      | Double   | 分布参数。                               |
| *Arg4* | 可选      | Variant  | x 所属区间的可选下界。                   |
| *Arg5* | 可选      | Variant  | x 所属区间的可选上界。                   |

**返回值**

Double

**说明**

Beta 分布通常用于研究样本中某些事物变化的百分比，例如人们一天中用来看电视的时间所占的比例。

- 如果任一参数为非数值型，则 BetaDist 将返回错误值 \#VALUE!。
- 如果 alpha ≤ 0 或 beta ≤ 0，则 BetaDist 将生成一个错误值。
- 如果 x \< A、x \> B 或 A = B，则 BetaDist 将生成一个错误值。
- 如果省略 A 和 B 的值，则 BetaDist 将使用标准累积 Beta 分布，使 A = 0，B = 1。

### WorksheetFunction.BetaInv

返回指定的 Beta 分布的累积分布函数的反函数。即，如果 probability = BetaDist(x,...)，则 BetaInv(probability,...) = x。

**要点** ??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 Beta_Inv 方法。

**语法**

**express.BetaInv ( *Arg1* , *Arg2* , *Arg3* , *Arg4* , *Arg5* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                     |
|--------|-----------|----------|--------------------------|
| *Arg1* | 必选      | Double   | 与 Beta 分布相关的概率。 |
| *Arg2* | 必选      | Double   | 该分布的 Alpha 参数。    |
| *Arg3* | 必选      | Double   | 该分布的 Beta 参数。     |
| *Arg4* | 可选      | Variant  | x 所属区间的可选下界。   |
| *Arg5* | 可选      | Variant  | x 所属区间的可选上界。   |

**返回值**

Double

**说明**

在项目计划中，如果已知预期完成时间和变化率，可以使用 beta 分布来建立可能完成时间的模型。

- 如果任一参数是非数值的，BetaInv 就会生成一个错误值。
- 如果 alpha ≤ 0 或 beta ≤ 0，BetaInv 将生成一个错误值。
- 如果 probability ≤ 0 或 probability \> 1，BetaInv 将生成一个错误值。
- 如果省略 A 或 B 的值，BetaInv 将使用标准累积 beta 分布，因此，A = 0，B = 1。

如果已知 probability 的值，BetaInv 将求出 x 值以使等式 BetaDist(x, alpha, beta, A, B) = probability 成立。因此，BetaInv 的精度取决于 BetaDist 的精度。BetaInv 使用迭代搜索技术。如果搜索在 100 次迭代之后尚未收敛，该函数将生成一个错误值。

### WorksheetFunction.Bin2Dec

将二进制数转换为十进制数。

**语法**

**express.Bin2Dec ( *Arg1* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                |
|--------|-----------|----------|---------------------------------------------------------------------------------------------------------------------|
| *Arg1* | 必选      | Variant  | 要转换的二进制数。Number 不能多于 10 个字符（10 位）。最高位为符号位，其余 9 位为数字位。负数用二进制数的补码表示。 |

**返回值**

String

**说明**

如果 number 不是有效的二进制数或多于 10 个字符（10 位），则 Bin2Dec 将生成一个错误值。

### WorksheetFunction.Bin2Hex

将二进制数转换为十六进制数。

**语法**

**express.Bin2Hex ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                           |
|--------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------|
| *Arg1* | 必选      | Variant  | 要转换的二进制数。Number 不能多于 10 个字符（10 位）。最高位为符号位，其余 9 位为数字位。负数用二进制数的补码表示。            |
| *Arg2* | 可选      | Variant  | 要使用的字符数。如果省略 places，Bin2Hex 将用能表示此数的最少字符来表示。当需要为返回的值填充前导 0（零）时，places 尤其有用。 |

**返回值**

String

**说明**

- 如果 number 不是有效的二进制数或多于 10 个字符（10 位），则 Bin2Hex 将生成一个错误。
- 如果 number 为负数，则 Bin2Hex 将忽略 places，并返回一个以 10 个字符表示的十六进制数。
- 如果 Bin2Hex 所需的字符数比 places 指定的字符数多，则将生成一个错误。
- 如果 places 不是整数，则将被截尾取整。
- 如果 places 为非数值型，则 Bin2Hex 将生成一个错误。
- 如果 places 为负数，则 Bin2Hex 将生成一个错误。

### WorksheetFunction.Bin2Oct

将二进制数转换为八进制数。

**语法**

**express.Bin2Oct ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                           |
|--------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------|
| *Arg1* | 必选      | Variant  | 要转换的二进制数。Number 不能多于 10 个字符（10 位）。最高位为符号位，其余 9 位为数字位。负数用二进制数的补码表示。            |
| *Arg2* | 可选      | Variant  | 要使用的字符数。如果省略 places，Bin2Oct 将用能表示此数的最少字符来表示。当需要为返回的值填充前导 0（零）时，places 尤其有用。 |

**返回值**

String

**说明**

- 如果 number 不是有效的二进制数或多于 10 个字符（10 位），则 Bin2Oct 将生成一个错误。
- 如果 number 为负数，则 Bin2Oct 将忽略 places，并返回一个以 10 个字符表示的八进制数。
- 如果 Bin2Oct 所需的字符数比 places 指定的字符数多，则将生成一个错误。
- 如果 places 不是整数，则将被截尾取整。
- 如果 places 为非数值型，则 Bin2Oct 将生成一个错误。
- 如果 places 为负数，则 Bin2Oct 将生成一个错误。

### WorksheetFunction.Binom_Dist

返回一元二项式分布的概率。

**语法**

**express.Binom_Dist ( *Arg1* , *Arg2* , *Arg3* , *Arg4* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                         |
|--------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Arg1* | 必选      | Double   | Number_s - 试验成功的次数。                                                                                                                                                                                  |
| *Arg2* | 必选      | Double   | Trials - 独立试验的次数。                                                                                                                                                                                    |
| *Arg3* | 必选      | Double   | Probability_s - 每次试验成功的概率。                                                                                                                                                                         |
| *Arg4* | 必选      | Boolean  | Cumulative - 一个逻辑值，用于确定函数的形式。如果 cumulative 为 True，则 Binom_Dist 方法返回累积分布函数，即最多存在 number_s 次成功的概率；如果为 False，则返回概率密度函数，即存在 number_s 次成功的概率。 |

**返回值**

Double

**说明**

当任何试验的结果仅包含成功或失败两种情况，当试验是独立试验，以及当整个试验过程中成功的概率固定不变时， **Binom_Dist** 方法适用于固定次数的检验或试验。例如， **Binom_Dist** 方法可以计算出生的三个婴儿中两个是男孩的概率。

- Number_s 和 trials 将被截尾取整。

- 如果 number_s、trials 或 probability_s 为非数值，则 **Binom_Dist** 方法将生成错误。

- 如果 number_s \< 0 或 number_s \> trials，则 **Binom_Dist** 方法将生成错误。

- 如果 probability_s \< 0 或 probability_s \> 1，则 **Binom_Dist** 方法将生成错误。

- 二项式概率密度函数的计算公式为：

  其中：

  为 COMBIN(n,x)。

  累积二项式分布的计算公式为：

### WorksheetFunction.Binom_Inv

返回一元二项式概率分布的反函数。

**语法**

**express.Binom_Inv ( *Arg1* , *Arg2* , *Arg3* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                 |
|--------|-----------|----------|--------------------------------------|
| *Arg1* | 必选      | Double   | Trials - 贝努利试验次数。            |
| *Arg2* | 必选      | Double   | Probability_s - 每次试验成功的概率。 |
| *Arg3* | 必选      | Double   | Alpha - 临界值。                     |

**返回值**

Double

**说明**

- 如果 Trials、Probability_s 或 Alpha 为非数值型， **Binom_Inv** 方法将生成错误。
- 如果 Trials 不是整数，则截尾取整。
- 如果 Trials \< 0，则 **Binom_Inv** 方法生成错误。
- 如果 Probability_s \< 0 或 Probability_s \> 1，则 **Binom_Inv** 方法生成错误。
- 如果 Alpha \< 0 或 Alpha \> 1，则 **Binom_Inv** 方法生成错误。

### WorksheetFunction.BinomDist

返回一元二项式分布的概率。

**语法**

**express.BinomDist ( *Arg1* , *Arg2* , *Arg3* , *Arg4* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                       |
|--------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Arg1* | 必选      | Double   | 试验成功的次数。                                                                                                                                                                           |
| *Arg2* | 必选      | Double   | 独立试验的次数                                                                                                                                                                             |
| *Arg3* | 必选      | Double   | 每次试验成功的概率。                                                                                                                                                                       |
| *Arg4* | 必选      | Boolean  | 一个逻辑值，用于确定函数的形式。如果 cumulative 为 TRUE，则 BinomDist 返回累积分布函数，即最多存在 number_s 次成功的概率；如果为 FALSE，则返回概率密度函数，即存在 number_s 次成功的概率。 |

**返回值**

Double

**说明**

当任何试验的结果仅包含成功或失败两种情况，当试验是独立试验，以及当整个试验过程中成功的概率固定不变时，BinomDist 适用于固定次数的检验或试验。例如，BinomDist 可以计算出生的三个婴儿中两个是男孩的概率。

- Number_s 和 trials 将被截尾取整。

- 如果 number_s、trials 或 probability_s 为非数值型，则 BinomDist 将生成一个错误。

- 如果 number_s \< 0 或 number_s \> trials，则 BinomDist 将生成一个错误。

- 如果 probability_s \< 0 或 probability_s \> 1，则 BinomDist 将生成一个错误。

- 二项式概率密度函数的计算公式为：

  其中：

  为 COMBIN(n,x)。

  累积二项式分布的计算公式为：

### WorksheetFunction.Ceiling

返回向上舍入（远离零）到最接近的 significance 的倍数的 number。

**要点** ??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 **Ceiling_Precise** 方法。

**语法**

**express.Ceiling ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| *Arg1* | 必选      | Double   | Number - 要舍入的值。                   |
| *Arg2* | 必选      | Double   | Significance - 用以进行舍入计算的倍数。 |

**返回值**

Double

**说明**

例如，如果要避免在价格中使用“分”，而产品价格为 ￥4.42, 那么可使用公式 ` Ceiling(4.42,0.05) ` 将价格进位到最近的“角”。

- 如果任一参数为非数值， **Ceiling** 将生成错误。
- 不论 number 的符号如何，向远离零的方向调整时，值都会向上舍入。如果 number 正好是 significance 的倍数，则无需进行任何舍入处理。
- 如果 number 和 significance 的符号不同，则 **Ceiling** 将生成错误。

### WorksheetFunction.Ceiling_Precise

返回舍入到最接近的有效位倍数的指定数字。

**语法**

**express.Ceiling_Precise ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| *Arg1* | 必选      | Double   | Number - 要舍入的值。                   |
| *Arg2* | 可选      | Variant  | Significance - 用以进行舍入计算的倍数。 |

**返回值**

Double

**说明**

例如，如果要避免在价格中使用“分”，而产品价格为 ￥4.42, 那么可使用公式 ` Ceiling(4.42,0.05) ` 将价格进位到最近的“角”。

根据 number 和 significance 参数的符号， **Ceiling_Precise** 方法离零舍入或向零舍入。

| 符号 (Arg1/Arg2) | 舍入       |
|------------------|------------|
| -/-              | 向零舍入。 |
| +/+              | 离零舍入。 |
| -/+              | 向零舍入。 |
| +/-              | 离零舍入。 |

- 如果任一参数为非数值， **Ceiling_Precise** 将生成错误。
- 如果 number 正好是 significance 的倍数，则无需进行任何舍入处理。

### WorksheetFunction.ChiDist

返回 χ2 分布的单尾概率。

此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 ChiSq_Dist_RT 和 ChiSq_Dist 方法。

**语法**

**express.ChiDist ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明               |
|--------|-----------|----------|--------------------|
| *Arg1* | 必选      | Double   | 用来计算分布的值。 |
| *Arg2* | 必选      | Double   | 自由度数。         |

**返回值**

Double

**说明**

χ <sup>2</sup> 分布与 χ <sup>2</sup> 检验相关。使用 χ <sup>2</sup> 检验可以比较观察值和期望值。

例如，某项遗传学实验可能假设下一代植物将呈现出某一组颜色。通过对观察到的结果和所期望的结果进行比较，可以判定初始假设是否有效。

- 如果两个参数中的任意一个为非数值型，则 ChiDist 将生成一个错误。
- 如果 x 为负数，则 ChiDist 将生成一个错误。
- 如果 degrees_freedom 不是整数，则将被截尾取整。
- 如果 degrees_freedom \< 1 或 degrees_freedom \> 10^10，则 ChiDist 将生成一个错误。
- ChiDist 按 ChiDist =P(X\>x) 计算，其中 X 为 χ <sup>2</sup> 随机变量。

### WorksheetFunction.ChiInv

返回 χ2 分布的单尾概率的反函数。

此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 ChiSq_Inv_RT 和 ChiSq_Inv 方法。

**语法**

**express.ChiInv ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| *Arg1* | 必选      | Double   | 与 χ2 分布相关的概率。 |
| *Arg2* | 必选      | Double   | 自由度数。             |

**返回值**

Double

**说明**

如果 probability = ChiDist(x,...)，则 ChiInv(probability,...) = x。使用此函数可对观测到的结果和所期望的结果进行比较，以判定初始假设是否有效。

- 如果两个参数中的任意一个为非数值型，则 ChiInv 将生成一个错误。
- 如果 probability \< 0 或 probability \> 1，则 ChiInv 将生成一个错误。
- 如果 degrees_freedom 不是整数，则将被截尾取整。
- 如果 degrees_freedom \< 1 或 degrees_freedom ≥ 10^10，则 ChiInv 将生成一个错误。

如果已给定概率值，则 ChiInv 将使用 ChiDist(x, degrees_freedom) = probability 求解值 x。因此，ChiInv 的精度取决于 ChiDist 的精度。ChiInv 使用迭代搜索技术。如果搜索在 64 次迭代之后没有收敛，则该函数将生成一个错误。

### WorksheetFunction.ChiSq_Dist

返回 χ2 分布。

**语法**

**express.ChiSq_Dist ( *Arg1* , *Arg2* , *Arg3* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                                   |
|--------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------|
| *Arg1* | 必选      | Double   | x - 要在计算分布时使用的值。                                                                                                           |
| *Arg2* | 必选      | Double   | deg_freedom - 自由度数。                                                                                                               |
| *Arg3* | 可选      | Variant  | cumulative - 一个决定函数形式的逻辑值。如果 cumulative 为 TRUE，则 CHISQ_DIST 将返回累积分布函数；如果为 FALSE，则将返回概率密度函数。 |

**返回值**

Double

**说明**

- 如果任一参数为非数值型，则 CHISQ_DIST 将返回错误值 \#VALUE!。
- 如果 x 为负数，则 CHISQ_DIST 将返回错误值 \#NUM!。
- 如果 deg_freedom 不是整数，则将被截尾取整。

### WorksheetFunction.ChiSq_Dist_RT

返回 χ2 分布的右尾概率。

**语法**

**express.ChiSq_Dist_RT ( *Arg1* , *Arg2* )**

\*express是一个代表 WorksheetFunction 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明               |
|--------|-----------|----------|--------------------|
| *Arg1* | 必选      | Double   | 用来计算分布的值。 |
| *Arg2* | 必选      | Double   | 自由度数。         |

**说明**

χ <sup>2</sup> 分布与 χ <sup>2</sup> 检验相关。使用 χ <sup>2</sup> 检验可以比较观察值和期望值。

例如，某项遗传学实验可能假设下一代植物将呈现出某一组颜色。通过对观察到的结果和所期望的结果进行比较，可以判定最初的假设是否有效：

- 如果任一参数为非数值型，则 ChiSq_Dist_RT 将生成一个错误。
- 如果 x 为负数，则 ChiSq_Dist_RT 将生成一个错误。
- 如果 degrees_freedom 不是整数，则将被截尾取整。
- 如果 degrees_freedom \< 1 或 degrees_freedom \> 10^10，则 ChiSq_Dist_RT 将生成一个错误。
- ChiSq_Dist_RT 的计算公式为 ChiSq_Dist_RT = P(X\>x)，其中 X 为 χ <sup>2</sup> 随机变量。

## 成员属性

### WorksheetFunction.Application

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象。可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序。只读。

**语法**

**express.Application**

\*express是一个代表 WorksheetFunction 对象的变量。

**示例**

``` JavaScript
function test(){
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
}
```

### WorksheetFunction.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 WorksheetFunction 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### WorksheetFunction.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 WorksheetFunction 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/WorksheetFunction](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CheckBox 对象

## 简介

代表一个复选框窗体域。

## 说明

使用 FormFields ( *Index* ) 返回单个 FormField 对象（其中 *Index* 是与复选框关联的索引编号或书签名称）。可将 CheckBox 属性与 FormField 对象配合使用，返回 CheckBox 对象。以下示例从活动文档中选择名为“Check1”的复选框窗体域。

``` JavaScript
Application.ActiveDocument.FormFields.Item("Check1").CheckBox.Value = true
```

索引编号代表窗体域在 FormFields 集合中的位置。以下示例检查第一个窗体域的类型，如果是复选框，则选中该复选框。

``` JavaScript
function test() {
    if (Application.ActiveDocument.FormFields.Item(1).Type == wdFieldFormCheckBox) {
        Application.ActiveDocument.FormFields.Item(1).CheckBox.Value = true
    }
}
```

以下示例先判断 *ffield* 对象是否有效，然后再将复选框大小修改为 14 磅。

``` JavaScript
function test() {
    let ffield = Application.ActiveDocument.FormField.Item(1).CheckBox
    if (ffield.Valid == true) {
        ffield.AutoSize = false
        ffield.Size = 14
    }
    else {
        alert("First field is not a check box")
    }
}
```

可将 Add 方法与 FormFields 对象配合使用，添加一个复选框窗体域。以下示例在活动文档的开头添加一个复选框，并将其命名为“Color”，然后选定该复选框。

``` JavaScript
function test() {
    let rng = Application.ActiveDocument.FormFields.Add(ActiveDocument.Range(0, 0), wdFieldFormCheckBox)
    rng.Name = "Color"
    rng.CheckBox.Value = true
}
```

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                             |
|:----:|:-------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CheckBox.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                                                                             |
|  2   | [AutoSize](#CheckBox.AutoSize)       | 如果为 True ，则根据环绕文字的字体大小调整复选框或文本框的大小。如果为 False ，则根据 Size 属性调整复选框或文本框的大小。 Boolean 类型，可读写。 |
|  3   | [Creator](#CheckBox.Creator)         | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。                                                                     |
|  4   | [Default](#CheckBox.Default)         | 返回或设置默认的复选框值。如果选中默认值，则该属性值为 True 。 Boolean 类型，可读写。                                                            |
|  5   | [Parent](#CheckBox.Parent)           | 返回一个 Object 类型的值，该值代表指定 CheckBox 对象的父对象。                                                                                   |
|  6   | [Size](#CheckBox.Size)               | 返回或设置复选框的大小。 Single 类型，可读写。                                                                                                   |
|  7   | [Valid](#CheckBox.Valid)             | 如果指定的窗体域对象为有效的复选框窗体域，则该属性值为 True 。 Boolean 类型，只读。                                                              |
|  8   | [Value](#CheckBox.Value)             | 如果选中此复选框，则该属性值为 True 。 Boolean 类型，可读写。                                                                                    |

## 成员属性

### CheckBox.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 CheckBox 对象的变量。

### CheckBox.AutoSize

如果为 **True** ，则根据环绕文字的字体大小调整复选框或文本框的大小。如果为 **False** ，则根据 **Size** 属性调整复选框或文本框的大小。 **Boolean** 类型，可读写。

**语法**

**express.AutoSize**

\*express是一个代表 CheckBox 对象的变量。

### CheckBox.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 CheckBox 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>该值也可用常量 <strong>wdCreatorCode</strong> 表示。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

### CheckBox.Default

返回或设置默认的复选框值。如果选中默认值，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Default**

\*express是一个代表 CheckBox 对象的变量。

**示例**

``` JavaScript
/*如果活动文档的第一个窗体域是复选框类型，本示例将检索其默认值。*/
function test() {
    let blnDefault
    if (Application.ActiveDocument.FormFields.Item(1).Type == wdFieldFormCheckBox) {
        blnDefault = Application.ActiveDocument.FormFields.Item(1).CheckBox.Default
    }
}
```

### CheckBox.Parent

返回一个 **Object** 类型的值，该值代表指定 **CheckBox** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 CheckBox 对象的变量。

### CheckBox.Size

返回或设置复选框的大小。 **Single** 类型，可读写。

**语法**

**express.Size**

\*express是一个代表 CheckBox 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中名为“Check1”的复选框的大小设置为 14 磅，然后选中该复选框。*/
function test() {
    let rng = Application.ActiveDocument.FormFields.Item("Check1").CheckBox
    rng.AutoSize = false
    rng.Size = 14
    rng.Value = true
}
```

### CheckBox.Valid

如果指定的窗体域对象为有效的复选框窗体域，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.Valid**

\*express是一个代表 CheckBox 对象的变量。

**示例**

``` JavaScript
/*本示例在插入点添加一个文字型窗体域。由于 myFormField 是一个文字输入域而不是一个复选框，所以消息框将显示“False”。*/
function test() {
    Application.Selection.Collapse(wdCollapseStart)
    let myFormField = Application.ActiveDocument.FormFields.Add(Application.Selection.Range, wdFieldFormTextInput)
    alert(myFormField.CheckBox.Valid)
}
```

### CheckBox.Value

如果选中此复选框，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Value**

\*express是一个代表 CheckBox 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/CheckBox](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FillFormat 对象

## 简介

代表形状的填充格式。形状可以有单色、渐变、纹理、图案、图片或半透明填充。

## 说明

使用 Fill 属性可返回 FillFormat 对象。以下示例向活动文档添加一个矩形，然后设置该矩形填充的渐变和颜色。

``` JavaScript
function test() {
    let shp = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80);
    let fill = shp.Fill;
    fill.ForeColor.RGB = 0x00FFFF
    fill.OneColorGradient(msoGradientHorizontal, 1, 1);
}
```

## 方法

| 序号 | 名称                                             | 说明                         |
|:----:|:-------------------------------------------------|:-----------------------------|
|  1   | [OneColorGradient](#FillFormat.OneColorGradient) | 将指定填充设置为单色渐变。   |
|  2   | [Patterned](#FillFormat.Patterned)               | 将指定填充设置为一种图案     |
|  3   | [Solid](#FillFormat.Solid)                       | 将指定填充设置为统一的颜色。 |
|  4   | [TwoColorGradient](#FillFormat.TwoColorGradient) | 将指定填充设为双色渐变。     |
|  5   | [UserPicture](#FillFormat.UserPicture)           | 用图像填充指定的形状。       |

## 属性

| 序号 | 名称                                               | 说明                                                                             |
|:----:|:---------------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [BackColor](#FillFormat.BackColor)                 | 返回或设置一个 [ColorFormat]() 对象，该对象代表填充的背景色，可读写。            |
|  2   | [ForeColor](#FillFormat.ForeColor)                 | 返回或设置一个 [ColorFormat]() 对象，该对象代表填充的前景色。可读写。            |
|  3   | [GradientAngle](#FillFormat.GradientAngle)         | 返回或设置指定填充格式的渐变填充的角度。可读/写。                                |
|  4   | [GradientColorType](#FillFormat.GradientColorType) | 返回指定填充的渐变颜色的样式。只读 MsoGradientColorType 类型。                   |
|  5   | [GradientDegree](#FillFormat.GradientDegree)       | 返回值表明单色渐变填充的深浅程度。 Number 类型，只读。                           |
|  6   | [Pattern](#FillFormat.Pattern)                     | 返回或设置一个 MsoPatternType 常量，该常量代表应用于指定填充或线条的图案。只读。 |
|  7   | [Type](#FillFormat.Type)                           | 返回图形填充格式的类型。只读 MsoFillType 类型。                                  |

## 成员方法

### FillFormat.OneColorGradient

将指定填充设置为单色渐变。

**语法**

**express.OneColorGradient ( *Style* , *Variant* , *Degree* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型         | 说明                                                                                                                                                                                                        |
|-----------|-----------|------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*   | 必选      | MsoGradientStyle | 渐变样式。可以是任意 MsoGradientStyle 常量（msoGradientFromTitle 除外，它仅应用于 WPP）。                                                                                                                   |
| *Variant* | 必选      | Number           | 渐变变量。取值范围为从 1 到 4，对应于“填充效果”对话框中“渐变”选项卡上的四种变形。如果 Style 为 msoGradientFromCenter，则该参数可为 1 或 2。 Degree 必选 Single 渐变深度。值的范围从 0.0（深）到 1.0（浅）。 |
| *Degree*  | 必选      | Number           | 渐变深度。值的范围从 0.0（深）到 1.0（浅）。                                                                                                                                                                |

**示例**

``` JavaScript
/* 本示例将在活动文档中添加一个具有单色渐变填充效果的矩形 */
function test() {
    let shp = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80);
    let fill = shp.Fill;
    fill.ForeColor.RGB = 0x00FFFF
    fill.OneColorGradient(msoGradientFromCenter, 3, 1);
}
```

### FillFormat.Patterned

将指定填充设置为一种图案

**语法**

**express.Patterned ( *Pattern* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型       | 说明                     |
|-----------|-----------|----------------|--------------------------|
| *Pattern* | 必选      | MsoPatternType | 用于指定填充方式的图案。 |

**说明**

可用 **BackColor** 和 **ForeColor** 属性设置图案颜色。

**示例**

``` JavaScript
/* 本示例在活动文档中添加一个作为填充图案的矩形。 */
function test() {
    let shp = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80);
    let fill = shp.Fill;
    fill.Patterned(msoPatternDarkVertical);
}
```

### FillFormat.Solid

将指定填充设置为统一的颜色。

**语法**

**express.Solid ()**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用该方法将渐变、纹理、图案或背景填充转换为纯色填充。

**示例**

``` JavaScript
/* 将当前文档第一个图片设置为绿色单色填充 */
function test() {
    let firstShape = Application.ActiveDocument.Shapes.Item(1);
    firstShape.Fill.Solid()
    firstShape.Fill.ForeColor.RGB = 0x00ff00
}
```

### FillFormat.TwoColorGradient

将指定填充设为双色渐变。

**语法**

**express.TwoColorGradient ( *Style* , *Variant* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型         | 说明                                                                                                                                                                  |
|-----------|-----------|------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*   | 必选      | MsoGradientStyle | 渐变样式。                                                                                                                                                            |
| *Variant* | 必选      | Number           | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 Style 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |

### FillFormat.UserPicture

用图像填充指定的形状。

**语法**

**express.UserPicture ( *PictureFile* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明         |
|---------------|-----------|----------|--------------|
| *PictureFile* | 必选      | String   | 图片文件名。 |

## 成员属性

### FillFormat.BackColor

返回或设置一个 **[ColorFormat]()** 对象，该对象代表填充的背景色，可读写。

**语法**

**express.BackColor**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例添加一个矩形，并返回填充背景色 */
Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle,10,10,20,30).Fill.BackColor
```

### FillFormat.ForeColor

返回或设置一个 **[ColorFormat]()** 对象，该对象代表填充的前景色。可读写。

**语法**

**express.ForeColor**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/* 设置矩形填充的前景色 */
Application.ActiveDocument.Shapes.Item(1).Fill.ForeColor.RGB=0x0000ff
```

### FillFormat.GradientAngle

返回或设置指定填充格式的渐变填充的角度。可读/写。

**语法**

**express.GradientAngle**

\*express是一个代表 FillFormat 对象的变量。

**说明**

可以在图表中各种元素（形状）的格式设置中指定渐变填充。例如，可以使用 **“数据系列格式”** 对话框将 **“柱形图”** 图表中列的格式设置为渐变填充。在这种情况下， **GradientAngle** 属性与 **“数据系列格式”** 对话框的 **“填充”** 类别中 **“角度”** 框的设置相对应。 **GradientAngle** 属性的有效值范围为从 0 到 359.9。

### FillFormat.GradientColorType

返回指定填充的渐变颜色的样式。只读 **MsoGradientColorType** 类型。

**语法**

**express.GradientColorType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性为只读属性。可用 **[OneColorGradient]()** 、 **[PresetGradient]()** 或 **[TwoColorGradient]()** 方法设置填充的渐变类型。

**示例**

``` JavaScript
/* 本示例将活动文档中所有图形的双色渐变填充更改为预设的渐变填充 */
function test() {
    let shps = Application.ActiveDocument.Shapes;
    for (let i = 1; i <= shps.Count; i++) {
        if (shps.Item(i).Fill.GradientColorType === msoGradientColorMixed) {
            shps.Item(i).Fill.PresetGradient(msoGradientFromCenter,1,msoGradientBrass);
        }
    }
}
```

### FillFormat.GradientDegree

返回值表明单色渐变填充的深浅程度。 **Number** 类型，只读。

**语法**

**express.GradientDegree**

\*express是一个代表 FillFormat 对象的变量。

**说明**

0 表示黑色与图形前景色混色形成渐变；1 表示白色混色；介于 0 和 1 之间的值表示混入前景色偏深或偏白的阴影。

可用 **[OneColorGradient]()** 方法设置填充的渐变程度。

### FillFormat.Pattern

返回或设置一个 **MsoPatternType** 常量，该常量代表应用于指定填充或线条的图案。只读。

**语法**

**express.Pattern**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/* 获取激活文档的第一个填充图案 */
Application.ActiveDocument.Shapes.Item(1).Fill.Pattern
```

### FillFormat.Type

返回图形填充格式的类型。只读 **MsoFillType** 类型。

**语法**

**express.Type**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/* 获取当前文档填充类型 */
Application.ActiveDocument.Shapes.Item(1).Fill.Type
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/FillFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Font 对象

## 简介

包含对象的字体属性（如字体名称、字号、颜色等）。

## 说明

使用 Font 属性可返回 Font 对象。以下指令对所选内容应用加粗格式。

``` JavaScript
Application.Selection.Font.Bold = true
```

以下示例将活动文档中第一个段落的格式设置为 24 磅、Arial 和加粗。

``` JavaScript
function test() {
    let myRange = ActiveDocument.Paragraphs.Item(1).Range
    myRange.Font.Bold = true
    myRange.Font.Name = "Arial"
    myRange.Font.Size = 24
}
```

以下示例将活动文档中“标题 2”样式的格式更改为 Arial 字体和斜体。

``` JavaScript
function test() {
    Application.ActiveDocument.Styles.Item(wdStyleHeading2).Font.Name = "Arial"
    Application.ActiveDocument.Styles.Item(wdStyleHeading2).Font.Italic = true
}
```

您可以使用 n ew 关键字来创建一个独立的新 Font 对象。以下示例创建一个 Font 对象，设置一些格式属性，然后将该 Font 对象应用于活动文档的第一个段落。

``` JavaScript
function test() {
    let myFont = Application.ActiveDocument.Paragraphs(1).Range.Font
    myFont.Bold = true
    myFont.Name = "Arial"
} 
```

## 方法

| 序号 | 名称                                               | 说明                                                                                                                                          |
|:----:|:---------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Grow](#Font.Grow)                                 | 将字号增大为下一个可用的字号。                                                                                                                |
|  2   | [Reset](#Font.Reset)                               | 删除手动字符格式（不使用样式应用的格式）。例如，如果手动将单词设置为加粗格式，但基本样式为纯文本（不是粗体），则可用 Reset 方法删除加粗格式。 |
|  3   | [SetAsTemplateDefault](#Font.SetAsTemplateDefault) | 将指定的字体格式设置为活动文档和所有基于活动模板创建的新文档的默认格式。                                                                      |
|  4   | [Shrink](#Font.Shrink)                             | 将字号减小为下一个可用的字号。                                                                                                                |

## 属性

| 序号 | 名称                                                         | 说明                                                                                  |
|:----:|:-------------------------------------------------------------|:--------------------------------------------------------------------------------------|
|  1   | [AllCaps](#Font.AllCaps)                                     | 如果字体格式为全部字母大写，则该属性值为 True 。 Long 类型，可读写。                  |
|  2   | [Animation](#Font.Animation)                                 | 返回或设置应用于字体的动态效果类型。 WdAnimation 类型，可读写。                       |
|  3   | [Application](#Font.Application)                             | 返回一个代表 WPS 应用程序的 Application 对象。                                        |
|  4   | [Bold](#Font.Bold)                                           | 如果字体格式为加粗，则该属性值为 True 。 Long 类型，可读写。                          |
|  5   | [BoldBi](#Font.BoldBi)                                       | 如果字体格式为加粗，则该属性值为 True 。 Long 类型，可读写。                          |
|  6   | [Borders](#Font.Borders)                                     | 返回一个 Borders 集合，该集合代表指定字体的所有边框。                                 |
|  7   | [ColorIndex](#Font.ColorIndex)                               | 返回或设置一个 WdColorIndex 常量，该常量代表指定字体的颜色。可读写。                  |
|  8   | [ColorIndexBi](#Font.ColorIndexBi)                           | 返回或设置从右向左语言文档中指定 Font 对象的颜色。 WdColorIndex 类型，可读写。        |
|  9   | [ContextualAlternates](#Font.ContextualAlternates)           | 指定是否为指定字体启用上下文替换。可读/写 Long 类型。                                 |
|  10  | [Creator](#Font.Creator)                                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。          |
|  11  | [DiacriticColor](#Font.DiacriticColor)                       | 返回或设置用于指定 Font 对象的音调符号的 24 位颜色。                                  |
|  12  | [DisableCharacterSpaceGrid](#Font.DisableCharacterSpaceGrid) | 如果 WPS 忽略相应 Font 对象的每行字符数，则该属性值为 True 。 Boolean 类型，可读写。  |
|  13  | [DoubleStrikeThrough](#Font.DoubleStrikeThrough)             | 如果将指定字体的格式设置为双删除线文本，则该属性值为 True 。                          |
|  14  | [Duplicate](#Font.Duplicate)                                 | 返回一个只读 Font 对象，该对象代表指定字体的字符格式。                                |
|  15  | [Emboss](#Font.Emboss)                                       | 如果指定字体格式为阳文，则该属性值为 True 。 Long 类型，可读写。                      |
|  16  | [EmphasisMark](#Font.EmphasisMark)                           | 返回或设置 WdEmphasisMark 常量，该常量代表字符或指定字符串的着重号。可读写。          |
|  17  | [Engrave](#Font.Engrave)                                     | 如果字体格式为阴文，则该属性值为 True 。 Long 类型，可读写。                          |
|  18  | [Fill](#Font.Fill)                                           | 返回 FillFormat 对象，该对象包含指定范围文本使用的字体的填充格式属性。只读。          |
|  19  | [Glow](#Font.Glow)                                           | 返回 GlowFormat 对象，该对象代表指定范围文本使用的字体的发光格式。只读。              |
|  20  | [Hidden](#Font.Hidden)                                       | 如果字体格式为隐藏文字，则该属性值为 True 。 Long 类型，可读写。                      |
|  21  | [Italic](#Font.Italic)                                       | 如果将字体或范围设置为倾斜格式，则该属性值为 True 。 Long 类型，可读写。              |
|  22  | [ItalicBi](#Font.ItalicBi)                                   | 如果将字体或范围设置为倾斜格式，则该属性值为 True 。可读写 Long 类型。                |
|  23  | [Kerning](#Font.Kerning)                                     | 返回或设置 WPS 根据字号自动调整字距时的最小字号。 Single 类型，可读写。               |
|  24  | [Ligatures](#Font.Ligatures)                                 | 返回或设置指定 Font 对象的连字设置。可读/写 WdLigatures 。                            |
|  25  | [Line](#Font.Line)                                           | 返回 LineFormat 对象，该对象指定行的格式。可读/写。                                   |
|  26  | [Name](#Font.Name)                                           | 返回或设置指定对象的名称。 String 类型，可读写。                                      |
|  27  | [NameAscii](#Font.NameAscii)                                 | 返回或设置拉丁语（字符代码从 0（零）到 127 的字符）所用的字体。 String 类型，可读写。 |
|  28  | [NameBi](#Font.NameBi)                                       | 返回或设置使用从右向左语言的文档中的字体的名称。 String 类型，可读写。                |
|  29  | [NameFarEast](#Font.NameFarEast)                             | 返回或设置一种东亚字体名称。 String 类型，可读写。                                    |
|  30  | [NameOther](#Font.NameOther)                                 | 返回或设置字符代码从 128 到 255 的字符所用的字体。 String 类型，可读写。              |
|  31  | [NumberForm](#Font.NumberForm)                               | 返回或设置 OpenType 字体的数字形式设置。可读/写 WdNumberForm 。                       |
|  32  | [NumberSpacing](#Font.NumberSpacing)                         | 返回或设置字体的数字间距设置。可读/写 WdNumberSpacing 。                              |
|  33  | [Outline](#Font.Outline)                                     | 如果字体格式为镂空，则该属性值为 True 。 Long 类型，可读写。                          |
|  34  | [Parent](#Font.Parent)                                       | 返回一个 Object 类型值，该值代表指定 Font 对象的父对象。                              |
|  35  | [Position](#Font.Position)                                   | 返回或设置相对于基准线的文本位置（以磅为单位）。可读写 Long 类型。                    |
|  36  | [Reflection](#Font.Reflection)                               | 返回一个 ReflectionFormat 对象，该对象代表形状的反射格式。只读。                      |
|  37  | [Scaling](#Font.Scaling)                                     | 返回或设置应用于字体的缩放百分比。 Long 类型，可读写。                                |
|  38  | [Shading](#Font.Shading)                                     | 返回一个 Shading 对象，该对象代表指定字体的底纹格式。                                 |
|  39  | [Shadow](#Font.Shadow)                                       | 如果将指定字体设置为阴影格式，则该属性值为 True 。可读写 Long 类型。                  |
|  40  | [Size](#Font.Size)                                           | 返回或设置字号（以磅为单位）。 Single 类型，可读写。                                  |
|  41  | [SizeBi](#Font.SizeBi)                                       | 返回或设置字体大小（以磅为单位）。 Single 类型，可读写。                              |
|  42  | [SmallCaps](#Font.SmallCaps)                                 | 如果字体格式为小型大写字母，则该属性值为 True 。 Long 类型，可读写。                  |
|  43  | [Spacing](#Font.Spacing)                                     | 返回或设置字符的间距（以磅为单位）。可读写 Single 类型。                              |
|  44  | [StrikeThrough](#Font.StrikeThrough)                         | 如果字体格式为带有删除线，则该属性值为 True 。 Long 类型，可读写。                    |
|  45  | [StylisticSet](#Font.StylisticSet)                           | 指明指定字体的样式集。可读/写 WdStylisticSet 。                                       |
|  46  | [Subscript](#Font.Subscript)                                 | 如果字体格式为下标，则该属性值为 True 。 Long 类型，可读写。                          |
|  47  | [Superscript](#Font.Superscript)                             | 如果字体格式为上标，则该属性值为 True 。 Long 类型，可读写。                          |
|  48  | [TextColor](#Font.TextColor)                                 | 返回一个代表指定字体的颜色的 ColorFormat 对象。只读。                                 |
|  49  | [TextShadow](#Font.TextShadow)                               | 返回 ShadowFormat 对象，该对象指明指定字体的阴影格式。                                |
|  50  | [ThreeD](#Font.ThreeD)                                       | 返回 ThreeDFormat 对象，该对象包含指定字体的三维效果格式属性。只读。                  |
|  51  | [Underline](#Font.Underline)                                 | 返回或设置应用于字体的下划线类型。可读写 WdUnderline 类型。                           |
|  52  | [UnderlineColor](#Font.UnderlineColor)                       | 返回或设置指定 Font 对象的下划线的 24 位颜色。                                        |

## 成员方法

### Font.Grow

将字号增大为下一个可用的字号。

**语法**

**express.Grow ()**

\*express是一个代表 Font 对象的变量。

**说明**

如果所选内容或范围包含多种字号，则将每一个字号都增大为下一个可用的设置。

**示例**

``` JavaScript
/*以下示例增大新文档中第四个单词的字号。*/
function test() {
  let rngTemp = Application.Documents.Add().Content
  rngTemp.InsertAfter("This is a test of the Grow method.")
  alert("Click OK to increase the font size of the fourth word.")
  rngTemp.Words.Item(4).Font.Grow()
}

/*以下示例增大选定文本的字号。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      Application.Selection.Font.Grow()
  }
  else{
      alert("You need to select some text.")
  }
}
```

### Font.Reset

删除手动字符格式（不使用样式应用的格式）。例如，如果手动将单词设置为加粗格式，但基本样式为纯文本（不是粗体），则可用 **Reset** 方法删除加粗格式。

**语法**

**express.Reset ()**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/*本示例删除选定部分中的手动设定的格式。*/
Application.Selection.Font.Reset()
```

### Font.SetAsTemplateDefault

将指定的字体格式设置为活动文档和所有基于活动模板创建的新文档的默认格式。

**语法**

**express.SetAsTemplateDefault ()**

\*express是一个代表 Font 对象的变量。

**说明**

默认字体格式存储在“正文”样式中。

**示例**

``` JavaScript
/*以下示例将所选内容中的字符格式设置为默认格式。*/
Selection.Font.SetAsTemplateDefault()
```

### Font.Shrink

将字号减小为下一个可用的字号。

**语法**

**express.Shrink ()**

\*express是一个代表 Font 对象的变量。

**说明**

如果所选内容或范围中包括多种字号，则将每一个字号都减小为下一个可用的设置。

**示例**

``` JavaScript
/*以下示例在新文档中插入一行字号逐渐减小的字母 Z。*/
function test()
{
let myRange = Documents.Add().Content
myRange.Font.Size = 45

for(let i = 1; i <= 5; i++){
    myRange.InsertAfter("Z")

    for(let r = 1; r <= 3; r++){
        myRange.Characters.Item(i).Font.Shrink()
    }
}
}
```

## 成员属性

### Font.AllCaps

如果字体格式为全部字母大写，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.AllCaps**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 wdUndefined（当返回值既可为 **True** ，也可为 **False** 时取该值）。该属性可设置为 **True** 、 **False** 或 **wdToggle** （与当前设置相反）。

如果设置 **AllCaps** 为 **True** ，则 **SmallCaps** 属性的值为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例检查活动文档第三段的文本格式是否设置为全部字母大写。*/
function test()
{
if(ActiveDocument.Paragraphs.Item(3).Range.Font.AllCaps == true){ 
    MsgBox("Text is all caps.")
}
else{
    MsgBox("Text is not all caps.")
}
}
```

``` JavaScript
/*本示例将选定文本的字体格式设置为全部字母大写。*/
function test()
{
if(Selection.Type == wdSelectionNormal){
    Selection.Font.AllCaps = true
}
else{
    MsgBox("You need to select some text.")
}
}
```

### Font.Animation

返回或设置应用于字体的动态效果类型。 **WdAnimation** 类型，可读写。

**语法**

**express.Animation**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/*本示例为新文档中的文本设置动态效果。*/
function test()
{
let docNew = Documents.Add()
docNew.Content.InsertAfter("This is a test of animation.")
docNew.Content.Font.Animation = wdAnimationLasVegasLights
}
```

``` JavaScript
/*本示例为选定文本设置动态效果。*/
function test()
{
if(Selection.Type == wdSelectionNormal) {
    Selection.Font.Animation = wdAnimationShimmer
}
else{
    MsgBox("You need to select some text.")
}
}
```

### Font.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Font 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Font.Bold

如果字体格式为加粗，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Bold**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

**示例**

``` JavaScript
/*本示例实现的功能是：如果选定内容中有一部分字体为加粗格式，则将全部选定内容的格式都设为加粗格式。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal) {
      if(Application.Selection.Font.Bold == wdUndefined) {
          Application.Selection.Font.Bold = true
      }
  }
  else{
      alert("You need to select some text.")
  }
}
```

### Font.BoldBi

如果字体格式为加粗，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.BoldBi**

\*express是一个代表 Font 对象的变量。

**说明**

该属性返回 **True** 、 **False** 或 **wdUndefined** （用于加粗和非加粗混合的文本）。该属性可设为 **True** 、 **False** 或 **wdToggle** 。

**BoldBi** 属性应用于从右向左语言的文字。

**示例**

``` JavaScript
/*本示例将从右向左语言的活动文档中第一段的格式设为加粗。*/
ActiveDocument.Paragraphs.Item(1).Range.Font.BoldBi = true
```

### Font.Borders

返回一个 **Borders** 集合，该集合代表指定字体的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Font 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Font.ColorIndex

返回或设置一个 **WdColorIndex** 常量，该常量代表指定字体的颜色。可读写。

**语法**

**express.ColorIndex**

\*express是一个代表 Font 对象的变量。

**说明**

**wdByAuthor** 常量不是有效的字体颜色。

**示例**

``` JavaScript
/*本示例更改活动文档首段的文字颜色。*/
Application.ActiveDocument.Paragraphs.Item(1).Range.Font.ColorIndex = wdGreen

/*本示例将所选文本的格式设置为红色。*/
Application.Selection.Font.ColorIndex = wdRed
```

### Font.ColorIndexBi

返回或设置从右向左语言文档中指定 **Font** 对象的颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.ColorIndexBi**

\*express是一个代表 Font 对象的变量。

**说明**

**wdByAuthor** 常量对 **Font** 对象是无效的。

**示例**

``` JavaScript
/*本示例将 Font 对象的颜色设置为蓝绿色。*/
Selection.Font.ColorIndexBi = wdTeal
```

### Font.ContextualAlternates

指定是否为指定字体启用上下文替换。可读/写 **Long** 类型。

**语法**

**express.ContextualAlternates**

\*express是一个代表 Font 对象的变量。

**说明**

上下文替换是根据单个字符周围的字母（其上下文）应用于这些字符的连字。上下文替换还可以应用于特定上下文中的所有词，例如常常用于标题中的词（如“of”和“the”）。为字体启用上下文替换后，将在字体设计器定义的那些上下文中使用上下文替换，而不使用标准连字。

设置该属性与选中 **“使用上下文替换”** （位于 WPS 2015 中 **“字体”** 对话框中 **“高级”** 选项卡上的 **“OpenType 功能”** 组中）旁边的复选框具有相同作用。

**示例**

``` JavaScript
/*下面的代码示例为活动文档中的字体启用上下文替换。*/
ActiveDocument.Range.Font.ContextualAlternates = true
```

### Font.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Font 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Font.DiacriticColor

返回或设置用于指定 **Font** 对象的音调符号的 24 位颜色。

**语法**

**express.DiacriticColor**

\*express是一个代表 Font 对象的变量。

**说明**

此属性可以是任何有效的 **WdColor** 常量或 Visual Basic 的 **RGB** 函数所返回的值。必须将 **UseDiffDiacColor** 属性的值设为 **True** 才能使用此属性。

**示例**

``` JavaScript
/*本示例将当前选定内容中的音调符号设为蓝色。*/
function test()
{
if(Options.UseDiffDiacColor == true) {
    Selection.Font.DiacriticColor = wdColorBlue
}
}
```

### Font.DisableCharacterSpaceGrid

如果 WPS 忽略相应 **Font** 对象的每行字符数，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisableCharacterSpaceGrid**

\*express是一个代表 Font 对象的变量。

**说明**

如果只将某些指定文字的 **DisableCharacterSpaceGrid** 属性设为 **True** ，则该属性会返回 **wdUndefined** 。

**示例**

``` JavaScript
/*该示例设置 WPS 忽略选定文本每行中的字符数。*/
Selection.Font.DisableCharacterSpaceGrid = true
```

### Font.DoubleStrikeThrough

如果将指定字体的格式设置为双删除线文本，则该属性值为 **True** 。

**语法**

**express.DoubleStrikeThrough**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （既可为 **True** ，也可为 **False** ）。可设置为 **True** 、 **False** 或 **wdToggle** 。 **Long** 类型，可读写。要设置或返回单删除线格式，可使用 **StrikeThrough** 属性。如果将 **DoubleStrikeThrough** 设置为 **True** ，则 **StrikeThrough** 将设置为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例对所选文字应用双删除线格式。*/
function test()
{
if(Selection.Type == wdSelectionNormal) {
    Selection.Font.DoubleStrikeThrough = true
}
else{
    MsgBox("You need to select some text.")
}
}
```

``` JavaScript
/*本示例取消活动文档中第一个单词的双删除线格式，并将该单词的第一个字母大写。*/
function test()
{
ActiveDocument.Words.Item(1).Font.DoubleStrikeThrough = false
ActiveDocument.Words.Item(1).Case = wdTitleSentence
}
```

### Font.Duplicate

返回一个只读 **Font** 对象，该对象代表指定字体的字符格式。

**语法**

**express.Duplicate**

\*express是一个代表 Font 对象的变量。

**说明**

可以使用 **Duplicate** 属性获取重复的 **Font** 对象所有属性的设置。可将 **Duplicate** 属性返回的对象赋给另一个 **Font** 对象，从而一次应用该对象的所有设置。在将重复对象赋给另一个对象之前，可更改重复对象的任何属性，而不会影响源对象。

**示例**

``` JavaScript
/*本示例将 MyDupFont 变量设为选定内容的字符格式，从 MyDupFont 中删除加粗格式，并添加倾斜格式。然后，本示例会创建一个新文档，插入文本，再将 MyDupFont 中保存的格式应用于该文本。*/
function test()
{
let myDupFont = Selection.Font.Duplicate
myDupFont.Bold = false
myDupFont.Italic = true

Documents.Add()
Selection.InsertAfter("This is some text.")
Selection.Font = myDupFont
}
```

### Font.Emboss

如果指定字体格式为阳文，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Emboss**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** 。该属性可设为 **True** 、 **False** 或 **wdToggle** 。如果将 **Emboss** 设置为 **True** ，则 **Engrave** 将设置为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例将新文档中第二个句子用阳文显示。*/
function test()
{
let con = Documents.Add().Content
con.InsertAfter("This is the first sentence. ")
con.InsertAfter("This is the second sentence. ")
con.Sentences.Item(2).Font.Emboss = true
}
```

``` JavaScript
/*本示例将选定的文本用阳文显示。*/
function test()
{
if(Selection.Type == wdSelectionNormal) {
    Selection.Font.Emboss = true
}
else{
    MsgBox("You need to select some text.")
}
}
```

### Font.EmphasisMark

返回或设置 **WdEmphasisMark** 常量，该常量代表字符或指定字符串的着重号。可读写。

**语法**

**express.EmphasisMark**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：在活动文档的第四个单词上面设置着重号：逗号。*/
ActiveDocument.Words.Item(4).EmphasisMark = wdEmphasisMarkOverComma
```

### Font.Engrave

如果字体格式为阴文，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Engrave**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。如果将 **Engrave** 设置为 **True** ，则 **Emboss** 将设置为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例将活动文档中的第一个字母的格式设置为阴文。*/
function test()
{
let rngTemp = ActiveDocument.Characters.Item(1)
rngTemp.Font.Size = 20
rngTemp.Font.Engrave = true
}
```

``` JavaScript
/*本示例将选定内容的格式设置为阴文。*/
function test()
{
if(Selection.Type == wdSelectionNormal){
    Selection.Font.Engrave = true
}
else{
    MsgBox("You need to select some text.")
}
}
```

### Font.Fill

返回 FillFormat 对象，该对象包含指定范围文本使用的字体的填充格式属性。只读。

**语法**

**express.Fill**

\*express是一个代表 Font 对象的变量。

### Font.Glow

返回 GlowFormat 对象，该对象代表指定范围文本使用的字体的发光格式。只读。

**语法**

**express.Glow**

\*express是一个代表 Font 对象的变量。

### Font.Hidden

如果字体格式为隐藏文字，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Hidden**

\*express是一个代表 Font 对象的变量。

**说明**

该属性返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

若要控制隐藏文字的显示，请使用 **View** 对象的 **ShowHiddenText** 属性。

当不显示隐藏文字时，用 **TextRetrievalMode** 对象的 **IncludeHiddenText** 属性可控制返回 **Range** 对象的属性或方法是否包含隐藏文字。

**示例**

``` JavaScript
/*本示例检查选定内容中的隐藏文字。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      if(Application.Selection.Font.Hidden == wdUndefined || Application.Selection.Font.Hidden == true){
          alert("There is hidden text in the selection.")
      }
      else{
          alert("No hidden text in the selection.")
      }
  }
  else{
      alert("You need to select some text.")
  }
}

/*本示例显示活动窗口中的所有隐藏文字并将选定内容的格式设置为隐藏文字。*/
function test() {
  Application.ActiveDocument.ActiveWindow.View.ShowHiddenText = true
  
  if(Application.Selection.Type == wdSelectionNormal){
      Application.Selection.Font.Hidden = true
  }
}
```

### Font.Italic

如果将字体或范围设置为倾斜格式，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Italic**

\*express是一个代表 Font 对象的变量。

**说明**

该属性返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

**示例**

``` JavaScript
/*以下示例检查所选内容中的倾斜格式并删除所选内容中存在的倾斜格式。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      let mySel = Application.Selection.Font.Italic
  
      if(mySel == wdUndefined || mySel == true){
          alert("there is italic text in selection. " + "Click OK to remove.")
          Application.Selection.Font.Italic = false
      }
      else{
          alert("No italic text in the selection.")
      }
  }
  else{
      alert("You need to select some text.")
  }
}
```

### Font.ItalicBi

如果将字体或范围设置为倾斜格式，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.ItalicBi**

\*express是一个代表 Font 对象的变量。

**说明**

该属性返回 **True** 、 **False** 或 **wdUndefined** （用于倾斜和非倾斜混合格式的文本）。可设置为 **True** 、 **False** 或 **wdToggle** 。

**ItalicBi** 属性用于从右向左的语言。

**示例**

``` JavaScript
/*以下示例将从右向左语言的活动文档中的第一段设置为倾斜格式。*/
ActiveDocument.Paragraphs.Item(1).Range.ItalicBi = true
```

### Font.Kerning

返回或设置 WPS 根据字号自动调整字距时的最小字号。 **Single** 类型，可读写。

**语法**

**express.Kerning**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档进行自动字距调整的最小字号设为 12 磅或更大。*/
ActiveDocument.Content.Font.Kerning = 12
```

``` JavaScript
/*本示例显示 WPS 根据字号自动调整选定文本的字距时的最小字号。*/
function test()
{
if(Selection.Type == wdSelectionNormal) {
    MsgBox(Selection.Font.Kerning)
}
else {
    MsgBox("You need to select some text.")
}
}
```

### Font.Ligatures

返回或设置指定 **Font** 对象的连字设置。可读/写 WdLigatures 。

**语法**

**express.Ligatures**

\*express是一个代表 Font 对象的变量。

**说明**

OpenType 字体支持使用连字。 **Ligatures** 属性指定要应用于 OpenType 字体的 WPS 2015 连字设置。

下表列出了连字的四个基本值。

| 值     | 说明                                                                         |
|--------|------------------------------------------------------------------------------|
| 标准   | 设计用来增强可读性和吸引力。拉丁语言中的标准连字包括“fi”、“fl”和“ff”。       |
| 上下文 | 设计用来通过根据周围文本提供最佳连字选择来增强可读性和吸引力。               |
| 历史   | 较老的、装饰用连字，对于现代读者可能看起来比较古老。不是专为可读性而设计的。 |
| 任意   | 旨在起装饰作用，不是为了可读性。                                             |

这四个基本值的组合构成了 **Ligatures** 属性的可用值集合。该组值在 WdLigatures 枚举中提供。

**示例**

``` JavaScript
/*下面的代码示例将任意连字应用于活动文档中的字体。*/
ActiveDocument.Range.Font.Ligatures = wdLigaturesDiscretional
```

### Font.Line

返回 LineFormat 对象，该对象指定行的格式。可读/写。

**语法**

**express.Line**

\*express是一个代表 Font 对象的变量。

### Font.Name

返回或设置指定对象的名称。 **String** 类型，可读写。

**语法**

**express.Name**

\*express是一个代表 Font 对象的变量。

**说明**

本示例将选定内容设置为 Arial 字体的加粗格式。

**示例**

``` JavaScript
/*本示例将选定内容设置为 Arial 字体的加粗格式。*/
function test() {
  Application.Selection.Font.Name = "Arial"
  Application.Selection.Font.Bold = true
}
```

### Font.NameAscii

返回或设置拉丁语（字符代码从 0（零）到 127 的字符）所用的字体。 **String** 类型，可读写。

**语法**

**express.NameAscii**

\*express是一个代表 Font 对象的变量。

**说明**

在 WPS 美国英语版中，该属性的默认值是 Times New Roman。使用 Name 属性可以更改应用于所有文本的字体和 **“格式”** 工具栏 **“字体”** 框中显示的字体。

**示例**

``` JavaScript
/*本示例设置拉丁语所用的字体。*/
Application.Selection.Font.NameAscii = "Century"
```

### Font.NameBi

返回或设置使用从右向左语言的文档中的字体的名称。 **String** 类型，可读写。

**语法**

**express.NameBi**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/*本示例将选定内容的格式设为 Arial 字体。*/
Selection.Font.NameBi = "Arial"
```

### Font.NameFarEast

返回或设置一种东亚字体名称。 **String** 类型，可读写。

**语法**

**express.NameFarEast**

\*express是一个代表 Font 对象的变量。

**说明**

在英语（美国）版的 WPS 中，本属性的默认值是 Times New Roman。推荐使用本属性返回或设置文档中的亚洲文本所用的字体，该文档是用 WPS 的亚洲版创建的。

**示例**

``` JavaScript
/*本示例显示应用于选定内容的东亚字体名称。*/
MsgBox(Selection.Font.NameFarEast)
```

### Font.NameOther

返回或设置字符代码从 128 到 255 的字符所用的字体。 **String** 类型，可读写。

**语法**

**express.NameOther**

\*express是一个代表 Font 对象的变量。

**说明**

在 WPS 美国英语版中，该属性的默认值是 Times New Roman。使用 **Name** 属性可以更改应用于所有文本的字体和 **“格式”** 工具栏 **“字体”** 框中显示的字体。

**示例**

``` JavaScript
/*本示例设置字符代码从 128 到 255 的字符所用的字体。*/
Selection.Font.NameOther = "Century"
```

### Font.NumberForm

返回或设置 OpenType 字体的数字形式设置。可读/写 WdNumberForm 。

**语法**

**express.NumberForm**

\*express是一个代表 Font 对象的变量。

**说明**

可以通过一致的高度配合文本基线（称为“内衬”）来显示 OpenType 字体中的数字，也可以通过变动高度（称为“悬挂”或“旧样式”）来显示这些数字，其中数字显示在文本基线之上或之下。使用 **NumberForm** 属性可指定是使用内衬还是旧样式来显示数字。

设置此属性与选择 **“数字形式:”** （ WPS 2015 中 **“字体”** 对话框的 **“高级”** 选项卡上的 **“OpenType 功能”** 组）旁边的下拉框中的项具有相同作用。

**示例**

``` JavaScript
/*下面的代码示例将活动文档中字体的数字形式设置为“旧样式”。*/
ActiveDocument.Range.Font.NumberForm = wdNumberFormOldStyle
```

### Font.NumberSpacing

返回或设置字体的数字间距设置。可读/写 WdNumberSpacing 。

**语法**

**express.NumberSpacing**

\*express是一个代表 Font 对象的变量。

**说明**

OpenType 字体支持成比例和表格图表功能来控制数字间距。成比例数字间距按不同的宽度处理每个数字。例如，“1”的显示要比“5”窄。表格数字间距按宽度相等处理数字，从而使它们能够垂直对齐，以增加可读性，特别是对于财务信息。

**示例**

``` JavaScript
/*下面的代码示例将活动文档中字体的数字间距设置为成比例。*/
ActiveDocument.Range.Font.NumberSpacing = wdNumberSpacingProportional
```

### Font.Outline

如果字体格式为镂空，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Outline**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

**示例**

``` JavaScript
/*本示例为活动文档前三个词设置镂空字体格式。*/
function test()
{
let myRange = ActiveDocument.Range(ActiveDocument.Words.Item(1).Start, ActiveDocument.Words.Item(3).End)
myRange.Font.Outline = true
}
```

``` JavaScript
/*本示例为选定文本切换镂空格式。*/
Selection.Font.Outline = wdToggle
```

``` JavaScript
/*当选定内容中的部分文本应用了镂空格式时，本示例清除选定内容中的镂空字体格式。*/
function test()
{
let myFont = Selection.Font
if(myFont.Outline == wdUndefined){
    myFont.Outline = false
}
}
```

### Font.Parent

返回一个 **Object** 类型值，该值代表指定 **Font** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Font 对象的变量。

### Font.Position

返回或设置相对于基准线的文本位置（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Position**

\*express是一个代表 Font 对象的变量。

**说明**

正值将文本提升，负值将文本降低。

**示例**

``` JavaScript
/*以下示例将所选文本降低 2 磅。*/
Selection.Font.Position = -2
```

### Font.Reflection

返回一个 ReflectionFormat 对象，该对象代表形状的反射格式。只读。

**语法**

**express.Reflection**

\*express是一个代表 Font 对象的变量。

### Font.Scaling

返回或设置应用于字体的缩放百分比。 **Long** 类型，可读写。

**语法**

**express.Scaling**

\*express是一个代表 Font 对象的变量。

**说明**

本属性以当前字体大小的百分比水平拉长或压缩文字（缩放范围从 1 到 600）。

**示例**

``` JavaScript
/*本示例将活动文档中的文字水平拉长到原大小的 110%。*/
ActiveDocument.Content.Font.Scaling = 110
```

``` JavaScript
/*本示例将 Sales.doc 中第一段的文字压缩到原大小的 90%。*/
function test()
{
let myFont = Documents.Item("Sales.doc").Paragraphs.Item(1).Range.Font
myFont.Scaling = 90
myFont.Bold = false
}
```

### Font.Shading

返回一个 **Shading** 对象，该对象代表指定字体的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Font 对象的变量。

### Font.Shadow

如果将指定字体设置为阴影格式，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.Shadow**

\*express是一个代表 Font 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例对所选内容应用阴影和加粗格式。*/
function test()
{
if(Selection.Type == wdSelectionNormal){
    Selection.Font.Shadow = true
    Selection.Font.Bold = true
}
else {
    MsgBox("You need to select some text.")
}
}
```

### Font.Size

返回或设置字号（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.Size**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/*以下示例插入文本并将该文本的第七个单词的字号设置为 20 磅。*/
function test() {
  Application.Selection.Collapse(wdCollapseEnd)
  let rng = Application.Selection.Range
  rng.Font.Reset()
  rng.InsertBefore( "This is a demonstration of font size.")
  rng.Words.Item(7).Font.Size = 20
}

/*以下示例确定所选文本的字号。*/
function test() {
  let mySel = Application.Selection.Font.Size
  if(mySel == wdUndefined){
      alert("there is a mix of font sizes in the selection.")
  }
  else {
      alert(mySel + " points")
  }
}
```

### Font.SizeBi

返回或设置字体大小（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.SizeBi**

\*express是一个代表 Font 对象的变量。

**说明**

**SizeBi** 属性适用于使用从右向左语言的文本。

**示例**

``` JavaScript
/*本示例将第一个单词的字体大小设置为 20 磅。*/
ActiveDocument.Paragraphs.Item(1).Range.Words.Item(1).Font.SizeBi = 20
```

### Font.SmallCaps

如果字体格式为小型大写字母，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.SmallCaps**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

如果将 **SmallCaps** 属性设为 **True** ，则 **AllCaps** 属性会设为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例区分新文档中小型大写字母和全部大写字母之间的差别。*/
function test()
{
let myRange = Documents.Add().Content
    myRange.InsertAfter("This is a demonstration of SmallCaps.")
    myRange.Words.Item(6).Font.SmallCaps = true
    myRange.InsertParagraphAfter()
    myRange.InsertAfter("This is a demonstration of AllCaps.")
    myRange.Words.Item(14).Font.AllCaps = true
}
```

``` JavaScript
/*如果部分选定内容中的字体格式已设置为小型大写字母格式，本示例将整个选定内容的字体格式设置为小型大写字母格式。*/
function test()
{
if(Selection.Type == wdSelectionNormal) { 
    let mySel = Selection.Font.SmallCaps
    if(mySel == wdUndefined) {
        Selection.Font.SmallCaps = true
    }
}
else {
    MsgBox("You need to select some text.")
}
}
```

### Font.Spacing

返回或设置字符的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.Spacing**

\*express是一个代表 Font 对象的变量。

**说明**

------------------------------------------------------------------------

**示例**

``` JavaScript
/*以下示例演示活动文档开头的两种不同字符间距。*/
function test()
{
let myRange = ActiveDocument.Range(0, 0)
    myRange.InsertAfter("Demonstration of no character spacing.")
    myRange.InsertParagraphAfter()
    myRange.InsertAfter("Demonstration of character spacing (1.5pt).")
    myRange.InsertParagraphAfter()

ActiveDocument.Paragraphs.Item(2).Range.Font.Spacing = 1.5
}
```

``` JavaScript
/*以下示例将所选文本的字符间距设置为 2 磅。*/
function test()
{
if(Selection.Type == wdSelectionNormal){
    Selection.Font.Spacing = 2
}
else{
    MsgBox("You need to select some text.")
}
}
```

### Font.StrikeThrough

如果字体格式为带有删除线，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.StrikeThrough**

\*express是一个代表 Font 对象的变量。

**说明**

**StrikeThrough** 属性返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

使用 **DoubleStrikeThrough** 属性可设置或返回双删除线格式。

**示例**

``` JavaScript
/*本示例为活动文档中的前三个单词设置删除线格式。*/
function test()
{
let myDoc = ActiveDocument
let myRange = myDoc.Range(myDoc.Words.Item(1).Start, myDoc.Words.Item(3).End)
myRange.Font.StrikeThrough = true
}
```

``` JavaScript
/*本示例为选定文本设置删除线格式。*/
function test()
{
if(Selection.Type == wdSelectionNormal){
    Selection.Font.StrikeThrough = true
}
else{
    MsgBox("You need to select some text.")
}
}
```

### Font.StylisticSet

指明指定字体的样式集。可读/写 WdStylisticSet 。

**语法**

**express.StylisticSet**

\*express是一个代表 Font 对象的变量。

**说明**

一些 OpenType 字体提供了样式集。样式集定义字体内要一起使用的一组字符，通常是为了获得视觉上的协调性，例如用于标题中。

**示例**

``` JavaScript
/*下面的代码示例将活动文档的字体设置为 Gabriola，然后应用 Gabriola 字体提供的第六个样式集。*/
function test()
{
ActiveDocument.Range.Font.Name = "Gabriola"
ActiveDocument.Range.Font.StylisticSet = wdStylisticSet06
}
```

### Font.Subscript

如果字体格式为下标，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Subscript**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

如果将 **Subscript** 属性设为 **True** ，则 **Superscript** 属性会设为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例在活动文档的开头插入文本，并将第十个字符设为下标。*/
function test()
{
let myRange = ActiveDocument.Range(0, 0)
myRange.InsertAfter("Water = H20")
myRange.Characters.Item(10).Font.Subscript = true
}
```

``` JavaScript
/*本示例检查所选文本的下标格式。*/
function test()
{
if(Selection.Type == wdSelectionNormal){
    let mySel = Selection.Font.Subscript
    if(mySel == wdUndefined || mySel == true){
        MsgBox("Subscript text exists in the selection.")
    }
    else{
        MsgBox("No subscript text in the selection.")
    }
}
else{
    MsgBox("You need to select some text.")
}
}
```

### Font.Superscript

如果字体格式为上标，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Superscript**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

如果将 **Superscript** 属性设为 **True** ，则 **Subscript** 属性会设为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例在活动文档的开头插入文本，并将第四个单词的两个字符设为上标。*/
function test()
{
let myRange = ActiveDocument.Range(0, 0)
myRange.InsertAfter("Superscript in the 4th word.")
ActiveDocument.Range(20, 22).Font.Superscript = true
}
```

``` JavaScript
/*本示例将所选文本设为上标。*/
function test()
{
if(Selection.Type == wdSelectionNormal){
    Selection.Font.Superscript = true
}
else{
    MsgBox("You need to select some text.")
}
}
```

### Font.TextColor

返回一个代表指定字体的颜色的 ColorFormat 对象。只读。

**语法**

**express.TextColor**

\*express是一个代表 Font 对象的变量。

### Font.TextShadow

返回 ShadowFormat 对象，该对象指明指定字体的阴影格式。

**语法**

**express.TextShadow**

\*express是一个代表 Font 对象的变量。

### Font.ThreeD

返回 ThreeDFormat 对象，该对象包含指定字体的三维效果格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 Font 对象的变量。

### Font.Underline

返回或设置应用于字体的下划线类型。可读写 **WdUnderline** 类型。

**语法**

**express.Underline**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/*以下示例给所选内容添加单下划线。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      Application.Selection.Font.Underline = wdUnderlineSingle
  }
  else{
      alert("You need to select some text.")
  }
}
```

### Font.UnderlineColor

返回或设置指定 **Font** 对象的下划线的 24 位颜色。

**语法**

**express.UnderlineColor**

\*express是一个代表 Font 对象的变量。

**说明**

此属性可以是任何有效的 **WdColor** 常量，也可以是 Visual Basic 的 **RGB** 函数返回的值。将 **UnderlineColor** 属性设置为 **wdColorAutomatic** 会将下划线的颜色重新设置为其上文本的颜色。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Font](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Page 对象

## 简介

代表文档中的页面。使用 Page 对象及相关方法和属性可通过编程方式定义文档的页面版式。

## 说明

使用 Item 方法可访问文档中的特定页面。以下示例访问活动文档的第一页。

``` JavaScript
let objPage = ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1)
```

要访问页码，请使用 Range 或 Selection 对象的 Information 属性，或者 Break 对象（该对象属于指定 Page 对象的 Breaks 集合）的 PageIndex 属性。

Page 对象的 Top 和 Left 属性总是返回 0（零），代表页面的左上角。 Height 和 Width 属性返回在“页面设置”对话框中指定的或通过 PageSetup 对象指定的页面大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面， Height 属性返回 792， Width 属性返回 612。所有这四个属性都是只读的。

## 属性

| 序号 | 名称                                     | 说明                                                                                               |
|:----:|:-----------------------------------------|:---------------------------------------------------------------------------------------------------|
|  1   | [Application](#Page.Application)         | 返回一个代表 WPS 应用程序的 Application 对象。                                                     |
|  2   | [Breaks](#Page.Breaks)                   | 返回一个 Breaks 集合，该集合代表页面上的分隔符。                                                   |
|  3   | [Creator](#Page.Creator)                 | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                       |
|  4   | [EnhMetaFileBits](#Page.EnhMetaFileBits) | 返回一个 Variant 类型的值，该值代表文本页显示方式的图片表示形式。只读。                            |
|  5   | [Height](#Page.Height)                   | 返回一个 Long 类型的值，该值代表页面的高度（以像素为单位）。                                       |
|  6   | [Left](#Page.Left)                       | 返回一个 Long 类型的值，该值代表页面的左边缘。只读。                                               |
|  7   | [Parent](#Page.Parent)                   | 返回一个 Object 类型值，该值代表指定 Page 对象的父对象。                                           |
|  8   | [Rectangles](#Page.Rectangles)           | 返回一个 Rectangles 集合，该集合代表文档页面上的文字或图形部分。                                   |
|  9   | [Top](#Page.Top)                         | 返回一个 Long 类型的值，该值代表页面的上边缘。只读。                                               |
|  10  | [Width](#Page.Width)                     | 返回一个 Long 类型的值，该值代表在 “页面设置” 对话框中定义的纸张宽度，以磅为单位。只读 Long 类型。 |

## 成员属性

### Page.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 Page 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Page.Breaks

返回一个 **Breaks** 集合，该集合代表页面上的分隔符。

**语法**

**express.Breaks**

\*express是一个代表 Page 对象的变量。

**说明**

**Breaks** 集合包含分页符、分栏符和分节符。使用 **Breaks** 集合以及相关对象和属性可通过编程方式定义文档的页面版式。

**示例**

``` JavaScript
/*以下示例返回活动文档中第一页上的分隔符。*/
let objBreaks = ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1).Breaks
```

### Page.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Page 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Page.EnhMetaFileBits

返回一个 **Variant** 类型的值，该值代表文本页显示方式的图片表示形式。只读。

**语法**

**express.EnhMetaFileBits**

\*express是一个代表 Page 对象的变量。

**说明**

**EnhMetaFileBits** 属性返回一个字节数组，该数组可在 Visual Basic 或 Microsoft C++ 开发环境中用于 Microsoft Windows 32 应用程序编程接口。

### Page.Height

返回一个 **Long** 类型的值，该值代表页面的高度（以像素为单位）。

**语法**

**express.Height**

\*express是一个代表 Page 对象的变量。

**说明**

**Page** 对象的 **Top** 和 **Left** 属性总是返回 0（零），以指示页面的左上角。 **Height** 和 **Width** 属性返回在“页面设置”对话框中指定的或通过 PageSetup 对象指定的纸张大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面， **Height** 属性返回 792， **Width** 属性返回 612。所有这四个属性都是只读的。

**示例**

``` JavaScript
Page 对象的 Top 和 Left 属性总是返回 0（零），以指示页面的左上角。
Height 和 Width 属性返回在“页面设置”对话框中指定的或通过 PageSetup 对象指定的纸张大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。
例如，对于纵向模式下的 8-1/2 × 11 英寸页面，Height 属性返回 792，Width 属性返回 612。所有这四个属性都是只读的。
```

### Page.Left

返回一个 **Long** 类型的值，该值代表页面的左边缘。只读。

**语法**

**express.Left**

\*express是一个代表 Page 对象的变量。

**说明**

**Page** 对象的 **Top** 和 **Left** 属性总是返回 0（零），代表页面的左上角。 **Height** 和 **Width** 属性返回在“页面设置”对话框中指定的或通过 **PageSetup** 对象指定的页面大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面， **Height** 属性返回 792， **Width** 属性返回 612。所有这四个属性都是只读的。

### Page.Parent

返回一个 **Object** 类型值，该值代表指定 **Page** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Page 对象的变量。

### Page.Rectangles

返回一个 **Rectangles** 集合，该集合代表文档页面上的文字或图形部分。

**语法**

**express.Rectangles**

\*express是一个代表 Page 对象的变量。

**说明**

使用 **Rectangles** 集合及相关对象和属性可通过编程方式定义文档的页面版式。矩形对应于文档页面上的文字或图形部分。

**示例**

``` JavaScript
/*以下示例返回活动文档中第一页的 Rectangles 集合。*/
let objRectangles = ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1).Rectangles
```

### Page.Top

返回一个 **Long** 类型的值，该值代表页面的上边缘。只读。

**语法**

**express.Top**

\*express是一个代表 Page 对象的变量。

**说明**

**Page** 对象的 **Top** 和 **Left** 属性总是返回 0（零），代表页面的左上角。 **Height** 和 **Width** 属性返回在“页面设置”对话框中指定的或通过 **PageSetup** 对象指定的页面大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面， **Height** 属性返回 792， **Width** 属性返回 612。所有这四个属性都是只读的。

### Page.Width

返回一个 **Long** 类型的值，该值代表在 **“页面设置”** 对话框中定义的纸张宽度，以磅为单位。只读 **Long** 类型。

**语法**

**express.Width**

\*express是一个代表 Page 对象的变量。

**说明**

**Page** 对象的 **Top** 和 **Left** 属性总是返回 0（零），代表页面的左上角。 **Height** 和 **Width** 属性返回在“页面设置”对话框中指定的或通过 **PageSetup** 对象指定的页面大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面， **Height** 属性返回 792， **Width** 属性返回 612。所有这四个属性都是只读的。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Page](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Reviewers 对象

## 简介

Reviewer 对象 [的集合](https://docs.microsoft.com/zh-cn/office/vba/api/word.reviewer) ，这些对象代表一个或多个文档的审阅者。 审阅者 集合中包含的所有审阅者审阅或编辑在一台计算机上打开的文档的名称。

## 说明

使用 Reviewers (Index) （其中 Index 是审阅者的名称或索引号）可返回 Reviewers 集合中的单个审阅者。

## 方法

| 序号 | 名称                    | 说明                           |
|:----:|:------------------------|:-------------------------------|
|  1   | [Item](#Reviewers.Item) | 返回集合中的单个 Reviewer 对象 |

## 属性

| 序号 | 名称                      | 说明                                                     |
|:----:|:--------------------------|:---------------------------------------------------------|
|  1   | [Count](#Reviewers.Count) | 返回一 个 Long ，代表集合中审阅者的数量。 此为只读属性。 |

## 成员方法

### Reviewers.Item

返回集合中的单个 **Reviewer** 对象

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Reviewers 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                              |
|---------|-----------|----------|-----------------------------------------------------------------------------------|
| *Index* | 可选      | Variant  | 要返回的单个对象。 可以是 Long 类型，表示序号位置或 字符串 ，表示单个对象的名称。 |

**示例**

``` JavaScript
//获取当前激活窗口第一个审阅者
Application.ActiveWindow.View.Reviewers.Item(1)
```

## 成员属性

### Reviewers.Count

返回一 个 Long ，代表集合中审阅者的数量。 此为只读属性。

**语法**

**express.Count**

\*express是一个代表 Reviewers 对象的变量。

**示例**

``` JavaScript
//获取审阅者的数量。 
Application.ActiveWindow.View.Reviewers.Count
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Reviewers](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Shape 对象

## 简介

代表绘制层中的对象，如自选图形、任意多边形、OLE 对象、ActiveX 控件或图片。 Shape 对象是 Shapes 集合的一个成员，该集合包含了某个文档的正文或个文档的所有页眉和页脚中的所有形状。

## 说明

形状总是属于某个锁定范围。可以将形状置于包含锁定标记的页面上的任何位置。 有三种代表形状的对象：

1.  Shapes 集合，代表文档中所有的形状
2.  ShapeRange 对象，代表文档中指定的形状子集（例如， ShapeRange 对象可以代表文档中的形状一和形状四，也可以代表文档中的所有选定形状）；
3.  Shape 对象，代表文档中的一个形状。如果需要同时处理若干个形状或处理所选内容中的若干个形状，可使用 ShapeRange 集合。

使用 Shapes ( *Index* ) 可返回一个 Shape 对象，其中 *Index* 为名称或索引号。以下示例水平翻转活动文档中的形状一。

``` JavaScript
Application.ActiveDocument.Shapes.Item(1).Flip(msoFlipHorizontal)
```

以下示例水平翻转活动文档中名为“Rectangle 1”的形状。

``` JavaScript
Application.ActiveDocument.Shapes.Item("Rectangle 1").Flip(msoFlipHorizontal)
```

在创建每个形状时为其指定一个默认名称。例如，如果向文档中添加了三个不同的形状，则这三个形状的名称可以是“矩形 2”、“文本框 3”和“椭圆 4”。要使形状的名称更有意义，请设置 Name 属性。

使用 ShapeRange ( *Index* ) 可返回代表所选内容中某个形状的 Shape 对象，其中 *Index* 为名称或索引号。假定所选内容包含至少一个形状，则以下示例为所选内容中的第一个形状设置填充效果。

``` JavaScript
Selection.ShapeRange.Item(1).Fill.ForeColor.RGB = (255, 0, 0)
```

假定所选内容至少包含一个形状，则以下示例为该区域的所有形状设置填充效果。

``` JavaScript
Application.Selection.ShapeRange.Item(1).Fill.ForeColor.RGB = (255, 0, 0)
```

要将 Shape 对象添加到指定文档的形状集合，并返回一个代表新建形状的 Shape 对象，可使用 Shapes 集合的下列方法之一： AddCallout 、 AddCurve 、 AddLabel 、 AddLine 、 AddOleControl 、 AddOleObject 、 AddPolyline 、 AddShape 、 AddTextbox 、 AddTextEffect 或 BuildFreeForm 。以下示例向活动文档添加一个矩形。

``` JavaScript
Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)
```

使用 GroupItems ( *Index* ) 可返回一个代表组合形状中的单个形状的 Shape 对象，其中 *Index* 为形状的名称或形状在该组中的索引号。

使用 Group 或 Regroup 方法可对一系列形状进行组合，并返回一个代表新组合的 Shape 对象。形成一个组合之后，便可以像处理任何其他形状一样处理该组合。

每个 Shape 对象都锁定到某个文本范围。将某个形状锁定到包含锁定范围的第一段的开头。该形状总是与其锁定标记在同一页上。

将 ShowObjectAnchors 属性设置为 True 可以查看锁定标记本身。形状的 Top 和 Left 属性决定了其垂直和水平位置。形状的 RelativeHorizontalPosition 和 RelativeVerticalPosition 属性决定了形状的位置是从锁定段落、包含锁定段落的栏、页边距还是页边缘开始计算。

如果将形状的 LockAnchor 属性设置为 True ，则无法将锁定标记从其所在位置拖动到页面上。

使用 Fill 属性可返回 FillFormat 对象，该对象包含用于设置闭合形状的填充格式的所有属性和方法。 Shadow 属性返回 ShadowFormat 对象，该对象用于设置阴影的格式。使用 Line 属性可返回 LineFormat 对象，该对象包含用于设置线条格式和箭头格式的各种属性和方法。 TextEffect 属性返回 TextEffectFormat 对象，该对象用于设置艺术字的格式。 Callout 属性返回 CalloutFormat 对象，该对象用于设置线形标注的格式。 WrapFormat 属性返回 WrapFormat 对象，该对象用于定义文字环绕形状的方式。 ThreeD 属性返回 ThreeDFormat 对象，该对象用于创建三维形状。可使用 PickUp 和 Apply 方法将一个形状的格式设置传送给另一个形状。

对 SetShapesDefaultProperties 对象使用 Shape 方法可设置文档的默认形状的格式。新形状将继承默认形状的许多属性。

使用 Type 属性可指定形状的类型；例如，任意多边形、自选图形、OLE 对象、标注或链接图片。使用 AutoShapeType 属性可指定自选图形的类型；例如，椭圆、矩形或气球。

使用 Width 和 Height 属性可指定形状的大小。

TextFrame 属性返回 TextFrame 对象，该对象包含用于将文本附加到形状和链接文本框之间的文本的所有属性和方法。

Shape 对象锁定到某一文本范围，但可以自由浮动，并且可以放置在页面上的任何位置。 InlineShape 对象被视为字符，并作为字符置于文本行中。可以使用 ConvertToInlineShape 方法和 ConvertToShape 方法来转换形状的类型。只能将图片、OLE 对象和 ActiveX 控件转换为内嵌形状。

## 方法

| 序号 | 名称                                                            | 说明                                                                                                                                                          |
|:----:|:----------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Apply](#Shape.Apply)                                           | 应用于使用 PickUp 方法复制的特定图形格式。                                                                                                                    |
|  2   | [CanvasCropBottom](#Shape.CanvasCropBottom)                     | 从绘图画布的底部裁剪一定百分比的 绘图画布 高度。                                                                                                              |
|  3   | [CanvasCropLeft](#Shape.CanvasCropLeft)                         | 从绘图画布左侧裁剪一定百分比的 绘图画布 宽度。                                                                                                                |
|  4   | [CanvasCropRight](#Shape.CanvasCropRight)                       | 从绘图画布右侧裁剪一定百分比的 绘图画布 宽度。                                                                                                                |
|  5   | [CanvasCropTop](#Shape.CanvasCropTop)                           | 从绘图画布顶部裁剪一定百分比的 绘图画布 高度。                                                                                                                |
|  6   | [ConvertToInlineShape](#Shape.ConvertToInlineShape)             | 将文档绘图层的指定图形转换为文字层的嵌入式图形。只能转换代表图片、OLE 对象或 ActiveX 控件的图形。此方法返回一个 InlineShape 对象，该对象代表图片或 OLE 对象。 |
|  7   | [Delete](#Shape.Delete)                                         | 删除指定的图形节点。                                                                                                                                          |
|  8   | [Duplicate](#Shape.Duplicate)                                   | 创建指定的 Shape 对象的副本，以标准的偏移将新图形从原图形添加至 Shapes 集合，然后返回新的 Shape 对象。                                                        |
|  9   | [Flip](#Shape.Flip)                                             | 水平或垂直翻转一个图形。                                                                                                                                      |
|  10  | [IncrementLeft](#Shape.IncrementLeft)                           | 将指定形状水平移动指定的磅数。                                                                                                                                |
|  11  | [IncrementRotation](#Shape.IncrementRotation)                   | 使指定的形状绕 Z 轴旋转指定的角度。                                                                                                                           |
|  12  | [IncrementTop](#Shape.IncrementTop)                             | 以指定磅数垂直移动指定形状。                                                                                                                                  |
|  13  | [PickUp](#Shape.PickUp)                                         | 复制指定形状的格式。                                                                                                                                          |
|  14  | [ScaleHeight](#Shape.ScaleHeight)                               | 以指定的比例缩放形状的高度。                                                                                                                                  |
|  15  | [ScaleWidth](#Shape.ScaleWidth)                                 | 按指定的比例缩放形状的宽度。                                                                                                                                  |
|  16  | [Select](#Shape.Select)                                         | 选择指定的形状。                                                                                                                                              |
|  17  | [SetShapesDefaultProperties](#Shape.SetShapesDefaultProperties) | 将文档中默认形状的格式应用于指定的形状。                                                                                                                      |
|  18  | [Ungroup](#Shape.Ungroup)                                       | 取消组合指定形状中的所有组合形状。                                                                                                                            |
|  19  | [ZOrder](#Shape.ZOrder)                                         | 将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 Z 顺序中的位置）。                                                                                 |

## 属性

| 序号 | 名称                                                            | 说明                                                                                                                                                                                                               |
|:----:|:----------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Adjustments](#Shape.Adjustments)                               | 返回一个 Adjustments 对象，该对象包含所有对指定 Shape 对象（代表自选图形或艺术字）进行调整操作的调整值。只读。                                                                                                     |
|  2   | [AlternativeText](#Shape.AlternativeText)                       | 返回或设置与网页的图形相关联的 可选文字?（可选文字：由 Web 浏览器用于在图像下载过程中显示文本，其对象是那些关闭图形的用户，以及那些依赖屏幕阅读软件将屏幕上的图形转换为可读语言的用户。） 。 String 类型，可读写。 |
|  3   | [Anchor](#Shape.Anchor)                                         | 返回一个 Range 对象，该对象代表指定图形或图形区域的锁定范围。只读。                                                                                                                                                |
|  4   | [Application](#Shape.Application)                               | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                                                                     |
|  5   | [AutoShapeType](#Shape.AutoShapeType)                           | 返回或设置指定的 Shape 对象的图形类型，该对象不是代表线条或任意多边形，而是代表自选图形。 MsoAutoShapeType 类型，可读写。                                                                                          |
|  6   | [BackgroundStyle](#Shape.BackgroundStyle)                       | 设置或返回指定形状的背景样式。可读/写 MsoBackgroundStyleIndex 。                                                                                                                                                   |
|  7   | [Callout](#Shape.Callout)                                       | 返回 CalloutFormat 对象，该对象包含指定图形的标注格式属性。只读。                                                                                                                                                  |
|  8   | [CanvasItems](#Shape.CanvasItems)                               | 返回一个 CanvasShapes 对象，该对象代表绘图画布上图形的集合。                                                                                                                                                       |
|  9   | [Chart](#Shape.Chart)                                           | 返回一个 Chart 对象，该对象代表文档中形状集合内的图表。只读。                                                                                                                                                      |
|  10  | [Child](#Shape.Child)                                           | 如果图形是子图形或位于图形区域的所有图形都是同一父图形的子图形，则该属性值为 True 。 MsoTriState 类型，只读。                                                                                                      |
|  11  | [Creator](#Shape.Creator)                                       | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                                                                       |
|  12  | [Fill](#Shape.Fill)                                             | 返回一个 FillFormat 对象，该对象包含指定图形的填充格式属性。只读。                                                                                                                                                 |
|  13  | [Glow](#Shape.Glow)                                             | 返回一个 GlowFormat 对象，该对象代表形状的发光格式。只读。                                                                                                                                                         |
|  14  | [GroupItems](#Shape.GroupItems)                                 | 返回一个 GroupShapes 对象，该对象代表指定图形组中的单个图形。只读。                                                                                                                                                |
|  15  | [HasChart](#Shape.HasChart)                                     | 如果指定的形状具有图表，则为 True 。只读。                                                                                                                                                                         |
|  16  | [HasSmartArt](#Shape.HasSmartArt)                               | 如果形状上显示 SmartArt 图表则返回 True 。只读。                                                                                                                                                                   |
|  17  | [Height](#Shape.Height)                                         | 返回或设置指定图形的高度。 Single 类型，可读写。                                                                                                                                                                   |
|  18  | [HeightRelative](#Shape.HeightRelative)                         | 返回或设置一个 Single 类型的值，该值代表形状高度的相对百分比。可读写。                                                                                                                                             |
|  19  | [HorizontalFlip](#Shape.HorizontalFlip)                         | 表示形状进行过水平翻转。 MsoTriState 类型，只读。                                                                                                                                                                  |
|  20  | [Hyperlink](#Shape.Hyperlink)                                   | 返回一个 Hyperlink 对象，该对象代表与 Shape 对象相关联的超链接。只读。                                                                                                                                             |
|  21  | [ID](#Shape.ID)                                                 | 返回指定形状的标识类型。只读 Long 类型。                                                                                                                                                                           |
|  22  | [LayoutInCell](#Shape.LayoutInCell)                             | 返回一个 Long 类型的值，该值表示表格中的形状显示在表格内部还是表格外部。                                                                                                                                           |
|  23  | [Left](#Shape.Left)                                             | 返回或设置一个 Single 类型的值，该值代表指定形状或形状范围的水平位置（以磅为单位）。也可以是任何有效的 WdShapePosition 常量。可读写。                                                                              |
|  24  | [LeftRelative](#Shape.LeftRelative)                             | 返回或设置一个 Single 类型的值，该值代表形状左侧的相对位置。可读写。                                                                                                                                               |
|  25  | [Line](#Shape.Line)                                             | 返回一个 LineFormat 对象，该对象包含指定形状的线条格式属性。只读。                                                                                                                                                 |
|  26  | [LinkFormat](#Shape.LinkFormat)                                 | 返回一个 LinkFormat 对象，该对象代表链接到文件的形状的链接选项。只读。                                                                                                                                             |
|  27  | [LockAnchor](#Shape.LockAnchor)                                 | 如果 Shape 对象的锁定标记锁定到锁定范围，则该属性值为 True 。可读写 Long 类型。                                                                                                                                    |
|  28  | [LockAspectRatio](#Shape.LockAspectRatio)                       | 如果在调整指定形状的大小时保留其最初比例，则该属性值为 MsoTrue ；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 MsoFalse 。可读写 MsoTriState 类型。                                                     |
|  29  | [Name](#Shape.Name)                                             | 返回或设置指定对象的名称。 String 类型，可读写。                                                                                                                                                                   |
|  30  | [Nodes](#Shape.Nodes)                                           | 返回一个 ShapeNodes 集合，该集合代表指定形状的几何描述。                                                                                                                                                           |
|  31  | [OLEFormat](#Shape.OLEFormat)                                   | 返回一个 OLEFormat 对象，该对象代表指定形状、内嵌形状或域的 OLE 特性（链接除外）。只读。                                                                                                                           |
|  32  | [Parent](#Shape.Parent)                                         | 返回一个 Object 类型值，该值代表指定 Shape 对象的父对象。                                                                                                                                                          |
|  33  | [ParentGroup](#Shape.ParentGroup)                               | 返回一个 Shape 对象，该对象代表子形状或子形状范围的通用父形状。                                                                                                                                                    |
|  34  | [PictureFormat](#Shape.PictureFormat)                           | 返回一个 PictureFormat 对象，该对象包含指定对象的图片格式属性。                                                                                                                                                    |
|  35  | [Reflection](#Shape.Reflection)                                 | 返回一个 ReflectionFormat 对象，该对象代表形状的反射格式。只读。                                                                                                                                                   |
|  36  | [RelativeHorizontalPosition](#Shape.RelativeHorizontalPosition) | 指定形状的相对水平位置。可读写 WdRelativeHorizontalPosition 类型。                                                                                                                                                 |
|  37  | [RelativeHorizontalSize](#Shape.RelativeHorizontalSize)         | 返回或设置一个 WdRelativeVerticalSize 常量，该常量代表形状区域相对的对象。可读写。                                                                                                                                 |
|  38  | [RelativeVerticalPosition](#Shape.RelativeVerticalPosition)     | 指定形状的相对垂直位置。可读写 WdRelativeVerticalPosition 类型。                                                                                                                                                   |
|  39  | [RelativeVerticalSize](#Shape.RelativeVerticalSize)             | 返回或设置一个 WdRelativeVerticalSize 常量，该常量代表形状的相对垂直大小。可读写。                                                                                                                                 |
|  40  | [Rotation](#Shape.Rotation)                                     | 返回或设置指定图形绕 Z 轴旋转的度数。正数表示按顺时针方向进行旋转，负值表示按逆时针方向旋转。 Single 类型，可读写。                                                                                                |
|  41  | [Script](#Shape.Script)                                         | 返回一个 Script 对象，该对象代表网页中图像的一段脚本或代码。                                                                                                                                                       |
|  42  | [Shadow](#Shape.Shadow)                                         | 返回一个 ShadowFormat 对象，该对象代表指定形状的阴影格式。                                                                                                                                                         |
|  43  | [ShapeStyle](#Shape.ShapeStyle)                                 | 返回或设置指定形状的形状样式。可读/写 MsoShapeStyleIndex。                                                                                                                                                         |
|  44  | [SmartArt](#Shape.SmartArt)                                     | 返回 SmartArt 对象，该对象提供一种处理与指定形状相关联的 SmartArt 的方式。只读。                                                                                                                                   |
|  45  | [SoftEdge](#Shape.SoftEdge)                                     | 返回一个 SoftEdgeFormat 对象，该对象代表形状的软边缘格式。只读。                                                                                                                                                   |
|  46  | [TextEffect](#Shape.TextEffect)                                 | 返回一个 TextEffectFormat 对象，该对象包含指定形状的文本效果格式属性。只读。                                                                                                                                       |
|  47  | [TextFrame](#Shape.TextFrame)                                   | 返回一个 TextFrame 对象，该对象包含指定形状的文字。                                                                                                                                                                |
|  48  | [TextFrame2](#Shape.TextFrame2)                                 | 返回一个 TextFrame2 对象，该对象包含指定形状的文本。只读。                                                                                                                                                         |
|  49  | [ThreeD](#Shape.ThreeD)                                         | 返回一个 ThreeDFormat 对象，该对象包含指定形状的三维格式属性。只读。                                                                                                                                               |
|  50  | [Title](#Shape.Title)                                           | 返回或设置包含指定形状的标题的 String 类型值。可读/写。                                                                                                                                                            |
|  51  | [Top](#Shape.Top)                                               | 返回或设置指定形状或形状范围的垂直位置（以磅为单位）。可读写 Single 类型。                                                                                                                                         |
|  52  | [TopRelative](#Shape.TopRelative)                               | 返回或设置一个 Single 类型的值，该值代表形状顶部的相对位置。可读写。                                                                                                                                               |
|  53  | [Type](#Shape.Type)                                             | 返回内嵌形状的类型。只读 MsoShapeType 类型。                                                                                                                                                                       |
|  54  | [VerticalFlip](#Shape.VerticalFlip)                             | 如果指定形状围绕垂直轴进行翻转，则该属性值为 True 。 MsoTriState 类型，只读。                                                                                                                                      |
|  55  | [Vertices](#Shape.Vertices)                                     | 该属性以一系列 坐标对 ?（坐标对：一对值，表示两维数组中存储的点的 x 和 y 坐标，该数组中包含许多点的坐标。） 的形式返回指定任意多边形图形顶点（和贝赛尔曲线的控点）的坐标。只读 Variant 类型。                      |
|  56  | [Visible](#Shape.Visible)                                       | 如果指定对象或应用于该对象的格式是可见的，则该属性值为 True 。 MsoTriState 类型，可读写。                                                                                                                          |
|  57  | [Width](#Shape.Width)                                           | 返回或设置指定形状的宽度（以磅为单位）。可读写 Long 类型。                                                                                                                                                         |
|  58  | [WidthRelative](#Shape.WidthRelative)                           | 返回或设置一个代表形状的相对宽度的 Single 。可读写。                                                                                                                                                               |
|  59  | [WrapFormat](#Shape.WrapFormat)                                 | 返回一个 WrapFormat 对象，该对象包含在指定的形状四周文字环绕的属性。只读。                                                                                                                                         |
|  60  | [ZOrderPosition](#Shape.ZOrderPosition)                         | 返回一个 Long 类型的值，该值代表指定的形状在 Z 顺序中的位置。只读。                                                                                                                                                |

## 成员方法

### Shape.Apply

应用于使用 **PickUp** 方法复制的特定图形格式。

**语法**

**express.Apply ()**

\*express是一个代表 Shape 对象的变量。

**说明**

如果以前没有使用 **PickUp** 方法复制 **Shape** 对象的格式，则使用 **Apply** 方法将导致出错。

**示例**

``` JavaScript
/*本示例复制活动文档中第一个图形的格式，然后将其应用到同一文档的第二个图形。*/
function test()
{
ActiveDocument.Shapes.Item(1).PickUp()
ActiveDocument.Shapes.Item(2).Apply()
}
```

### Shape.CanvasCropBottom

从绘图画布的底部裁剪一定百分比的 绘图画布 高度。

**语法**

**express.CanvasCropBottom ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                       |
|-------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 裁剪后所需保留的绘图画布高度的百分比数值。输入 0.9 作为增量从底部裁剪绘图画布高度的百分之十。输入 0.1 从底部裁剪绘图画布高度的百分之九十。 |

**示例**

``` JavaScript
/*假定活动文档中的第一个图形为绘图画布的情况下，则本示例从活动文档中第一块绘图画布的底部裁剪绘图画布高度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在文档中添加绘图画布。*/
function CropCanvasBottom() {
    let shpCanvas = ActiveDocument.Shapes.Item(1)
    shpCanvas.CanvasCropBottom(0.75)
}
```

### Shape.CanvasCropLeft

从绘图画布左侧裁剪一定百分比的 绘图画布 宽度。

**语法**

**express.CanvasCropLeft ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                               |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 裁剪绘图画布后所需保留的绘图画布宽度的百分比数值。输入 0.9 作为增量从左侧裁剪绘图画布宽度的百分之十。输入 0.1 从左侧裁剪绘图画布宽度的百分之九十。 |

**示例**

``` JavaScript
/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从活动文档中第一块绘图画布的左侧裁剪绘图画布宽度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在文档中添加绘图画布。*/
function CropCanvasLeft() {
    let shpCanvas = ActiveDocument.Shapes.Item(1)
    shpCanvas.CanvasCropLeft(0.75)
}
```

### Shape.CanvasCropRight

从绘图画布右侧裁剪一定百分比的 绘图画布 宽度。

**语法**

**express.CanvasCropRight ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                               |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 裁剪绘图画布后所需保留的绘图画布宽度的百分比数值。输入 0.9 作为增量从右侧裁剪绘图画布宽度的百分之十。输入 0.1 从右侧裁剪绘图画布宽度的百分之九十。 |

**示例**

``` JavaScript
/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从右侧裁剪活动文档中第一块绘图画布宽度的百分之二十五。如果不是这种情况，需要使用 AddCanvas 方法在文档中添加绘图画布。*/
function CropCanvasRight() {
    let shpCanvas = ActiveDocument.Shapes.Item(1)
    shpCanvas.CanvasCropRight(0.75)
}
```

### Shape.CanvasCropTop

从绘图画布顶部裁剪一定百分比的 绘图画布 高度。

**语法**

**express.CanvasCropTop ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                               |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 裁剪绘图画布后所需保留的绘图画布高度的百分比数值。输入 0.9 作为增量从顶部裁剪绘图画布高度的百分之十。输入 0.1 从顶部裁剪绘图画布高度的百分之九十。 |

**示例**

``` JavaScript
/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从活动文档中第一块绘图画布的顶部裁剪绘图画布高度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在活动文档中添加绘图画布。 */
function CropCanvasTop(){
    let shpCanvas = ActiveDocument.Shapes.Item(1)
    shpCanvas.CanvasCropTop(0.75)
}
```

### Shape.ConvertToInlineShape

将文档绘图层的指定图形转换为文字层的嵌入式图形。只能转换代表图片、OLE 对象或 ActiveX 控件的图形。此方法返回一个 **InlineShape** 对象，该对象代表图片或 OLE 对象。

**语法**

**express.ConvertToInlineShape ()**

\*express是一个代表 Shape 对象的变量。

**说明**

支持附加文字的图形不能转换为嵌入式图形。对这样的图形请使用 **ConvertToFrame** 方法。

如果对包含多个图形的 **ShapeRange** 对象使用此方法，将导致出错。

**示例**

``` JavaScript
/*本示例将 MyDoc.doc 中的每张图片转换为嵌入式图形。*/
function test()
{
for(let i = 1; i <= Documents.Item("MyDoc.doc").Shapes.Count; i++){
    if(Documents.Item("MyDoc.doc").Shapes.Item(i).Type == msoPicture){
       Documents.Item("MyDoc.doc").Shapes.Item(i).ConvertToInlineShape()
    }
}
}
```

### Shape.Delete

删除指定的图形节点。

**语法**

**express.Delete ( *Index* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                     |
|---------|-----------|----------|--------------------------|
| *Index* | 必选      | Long     | 要删除的图形节点的数目。 |

**示例**

``` JavaScript
/*删除文档形状集合中的形状一*/
Application.ActiveDocument.Shapes.Item(1).Delete()
```

### Shape.Duplicate

创建指定的 **Shape** 对象的副本，以标准的偏移将新图形从原图形添加至 **Shapes** 集合，然后返回新的 **Shape** 对象。

**语法**

**express.Duplicate ()**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例创建活动文档第一个图形的副本，然后改变新图形的填充格式。*/
function test()
{
let newShape = ActiveDocument.Shapes.Item(1).Duplicate()
    newShape.Fill.PresetGradient(msoGradientVertical, 1, msoGradientGold)
}
```

### Shape.Flip

水平或垂直翻转一个图形。

**语法**

**express.Flip ( *FlipCmd* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型   | 说明       |
|-----------|-----------|------------|------------|
| *FlipCmd* | 必选      | MsoFlipCmd | 翻转方向。 |

**示例**

``` JavaScript
/*本示例首先向活动文档添加一个三角形，然后复制此三角形，再垂直翻转复制的三角形并将其设为红色。*/
function FlipShape(){
    let aShape = ActiveDocument.Shapes.AddShape(msoShapeRightTriangle, 150, 150, 50, 50).Duplicate()
    aShape.Fill.ForeColor.RGB = (255, 0, 0)
    aShape.Fill(msoFlipVertical)
}
```

### Shape.IncrementLeft

将指定形状水平移动指定的磅数。

**语法**

**express.IncrementLeft ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                         |
|-------------|-----------|----------|------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定形状水平移动的距离，以磅为单位。为正值时将形状右移；为负值时将形状左移。 |

**示例**

``` JavaScript
/*以下示例复制 myDocument 上的第一个形状，并设置复制的填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度。*/
function test()
{
let myDocument = ActiveDocument
    myDocument.Shapes.Item(1).Duplicate
    myDocument.Fill.PresetTextured(msoTextureGranite)
    myDocument.IncrementLeft(70)
    myDocument.IncrementTop(-50)
    myDocument.IncrementRotation(30)
}
```

### Shape.IncrementRotation

使指定的形状绕 Z 轴旋转指定的角度。

**语法**

**express.IncrementRotation ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                               |
|-------------|-----------|----------|------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定形状的水平旋转量，以度为单位。为正值时顺时针旋转形状，为负值时逆时针旋转形状。 |

**说明**

使用 **Rotation** 属性可以设置形状的绝对旋转角度。要绕 X 轴或 Y 轴旋转三维形状，请使用 **ThreeDFormat** 对象的 **IncrementRotationX** 或 **IncrementRotationY** 方法。

**示例**

``` JavaScript
/*以下示例复制 myDocument 上的第一个形状，并设置复制的填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度。*/
function test()
{
let myDocument = ActiveDocument
    myDocument.Shapes.Item(1).Duplicate
    myDocument.Fill.PresetTextured(msoTextureGranite)
    myDocument.IncrementLeft(70)
    myDocument.IncrementTop(-50)
    myDocument.IncrementRotation(30)
}
```

### Shape.IncrementTop

以指定磅数垂直移动指定形状。

**语法**

**express.IncrementTop ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                         |
|-------------|-----------|----------|------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定对象的垂直移动距离，以磅为单位。为正值时将形状下移；为负值时将形状上移。 |

**示例**

``` JavaScript
/*以下示例复制 myDocument 上的第一个形状，并设置复制的填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度。*/
function test()
{
let myDocument = ActiveDocument
    myDocument.Shapes.Item(1).Duplicate
    myDocument.Fill.PresetTextured(msoTextureGranite)
    myDocument.IncrementLeft(70)
    myDocument.IncrementTop(-50)
    myDocument.IncrementRotation(30)
}
```

### Shape.PickUp

复制指定形状的格式。

**语法**

**express.PickUp ()**

\*express是一个代表 Shape 对象的变量。

**说明**

用 **Apply** 方法可将复制的格式应用于另一个形状。

**示例**

``` JavaScript
/*以下示例首先复制 myDocument 上第一个形状的格式，然后将复制的形状格式应用于第二个形状。*/
function test()
{
let myDocument = ActiveDocument
    myDocument.Shapes.Item(1).PickUp()
    myDocument.Shapes.Item(2).Apply()
}
```

### Shape.ScaleHeight

以指定的比例缩放形状的高度。

**语法**

**express.ScaleHeight ( *Factor* , *RelativeToOriginalSize* , *Scale* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型     | 说明                                                                                                                                                        |
|--------------------------|-----------|--------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Single       | 指定形状调整后的高度与当前或原始高度的比例。例如，要将一个矩形放大百分之五十，请将此参数指定为 1.5。                                                        |
| *RelativeToOriginalSize* | 必选      | MsoTriState  | 如果该参数值为 True，则相对于原始大小缩放形状。如果该参数值为 False，则相对于当前大小缩放形状。仅当指定的形状为图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选      | MsoScaleFrom | 在缩放形状时，形状中位置不变的部分。                                                                                                                        |

**说明**

对于图片和 OLE 对象，您可以说明是相对于原始大小还是相对于当前大小缩放形状。对于图片和 OLE 对象以外的形状总是相对于当前高度缩放。

**示例**

``` JavaScript
/*以下示例将 myDocument 上的所有图片和 OLE 对象放大至原始高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/
function test()
{
let myDocument = ActiveDocument
let s
for(let i = 1; i <= myDocument.Shapes.Count;i++){
    s = myDocument.Shapes.Item(i)
    switch(s.Type){
        case msoEmbeddedOLEObject:
        case msoLinkedOLEObject:
        case msoOLEControlObject:
        case msoLinkedPicture:
        case msoPicture:
            s.ScaleHeight(1.75, true)
            s.ScaleWidth(1.75, true)
            break
        default:
            s.ScaleHeight(1.75, false)
            s.ScaleWidth(1.75, false)
    }
}
}
```

### Shape.ScaleWidth

按指定的比例缩放形状的宽度。

**语法**

**express.ScaleWidth ( *Factor* , *RelativeToOriginalSize* , *Scale* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型     | 说明                                                                                                                                                        |
|--------------------------|-----------|--------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Single       | 指定形状调整后的宽度与当前或原始宽度的比例。例如，要将一个矩形放大百分之五十，请将此参数指定为 1.5。                                                        |
| *RelativeToOriginalSize* | 必选      | MsoTriState  | 如果该参数值为 True，则相对于原始大小缩放形状。如果该参数值为 False，则相对于当前大小缩放形状。仅当指定的形状为图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选      | MsoScaleFrom | 在缩放形状时，形状中位置不变的部分。                                                                                                                        |

**说明**

对于图片和 OLE 对象，您可以说明是相对于原始大小还是相对于当前大小缩放形状。图片和 OLE 对象以外的形状总是相对于当前宽度缩放。

**示例**

``` JavaScript
/*以下示例将 myDocument 上的所有图片和 OLE 对象放大至原始高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/
function test()
{
let myDocument = ActiveDocument
let s
for(let i = 1; i <= myDocument.Shapes.Count;i++){
    s = myDocument.Shapes.Item(i)
    switch(s.Type){
        case msoEmbeddedOLEObject:
        case msoLinkedOLEObject:
        case msoOLEControlObject:
        case msoLinkedPicture:
        case msoPicture:
            s.ScaleHeight(1.75, true)
            s.ScaleWidth(1.75, true)
            break
        default:
            s.ScaleHeight(1.75, false)
            s.ScaleWidth(1.75, false)
    }
}
}
```

### Shape.Select

选择指定的形状。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                  |
|-----------|-----------|----------|-------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | 在添加形状时，如果该参数值为 True ，则替换所选内容。如果该参数值为 False ，则将新形状添加到所选内容。 |

**示例**

``` JavaScript
/*选中文档形状集合中的形状一*/
Application.ActiveDocument.Shapes.Item(1).Select()
```

### Shape.SetShapesDefaultProperties

将文档中默认形状的格式应用于指定的形状。

**语法**

**express.SetShapesDefaultProperties ()**

\*express是一个代表 Shape 对象的变量。

**说明**

新形状将继承默认形状的许多属性。

**示例**

``` JavaScript
/*以下示例将一个矩形添至 myDocument，设置矩形的填充格式，将矩形的格式应用于默认形状，然后将另一个（较小的）矩形添至文档。后一个矩形与前一个矩形有相同的填充格式。*/
function test()
{
let mydocument = ActiveDocument
let Shape = mydocument.Shapes
let newShape1 = Shape.AddShape(msoShapeRectangle, 5, 5, 80, 60)
    let newShape2 = newShape1.Fill
        newShape2.ForeColor.RGB = (0, 0, 255)
        newShape2.BackColor.RGB = (0, 204, 255)
        newShape2.Patterned(msoPatternHorizontalBrick)
    //Sets formatting for default shapes
    newShape1.SetShapesDefaultProperties()
//New shape has default formatting
Shape.AddShape(msoShapeRectangle, 90, 90, 40, 30)
}
```

### Shape.Ungroup

取消组合指定形状中的所有组合形状。

**语法**

**express.Ungroup ()**

\*express是一个代表 Shape 对象的变量。

**返回值**

ShapeRange

**说明**

此方法分解指定形状中的图片和 OLE 对象并将取消组合后的形状作为单个 **ShapeRange** 对象返回。

因为将组合形状作为单个对象处理，所以对形状进行组合或取消组合操作时，将会更改 **Shapes** 集合中项的数目，并更改集合中受影响的项之后的项的索引号。

**示例**

``` JavaScript
/*以下示例取消组合 myDocument 中的所有组合形状，并分解所有的图片和 OLE 对象。*/
function test()
{
let myDocument = ActiveDocument
for(let i = 1; i <= myDocument.Shapes.Count; i++){
    myDocument.Shapes.Item(i).Ungroup()
}
}
```

``` JavaScript
/*以下示例将取消组合 myDocument 中的所有组合形状，但不分解文档中图片或 OLE 对象。*/
function test()
{
let myDocument = ActiveDocument
for(let i = 1; i <= myDocument.Shapes.Count; i++){
    if(myDocument.Shapes.Item(i).Type == msoGroup){
       myDocument.Shapes.Item(i).Ungroup()
    }
}
}
```

### Shape.ZOrder

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 Z 顺序中的位置）。

**语法**

**express.ZOrder ( *ZOrderCmd* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型     | 说明                                       |
|-------------|-----------|--------------|--------------------------------------------|
| *ZOrderCmd* | 必选      | MsoZOrderCmd | 指定相对于其他形状将指定形状移动到的位置。 |

**返回值**

无

**说明**

使用 **ZOrderPosition** 属性可确定形状在 Z 顺序中的当前位置。

**示例**

``` JavaScript
/*本示例在活动文档中添加一个椭圆形，并将该椭圆形置于 Z 顺序中倒数第二的位置（如果文档中至少还有另一个图形）。*/
function test()
{
let zorderposition = ActiveDocument.Shapes.AddShape(msoShapeOval,  100, 100, 100, 300)
while(zorderposition.ZOrderPosition >  2){
      zorderposition.ZOrder(msoSendBackward)
}
}
```

## 成员属性

### Shape.Adjustments

返回一个 **Adjustments** 对象，该对象包含所有对指定 **Shape** 对象（代表自选图形或艺术字）进行调整操作的调整值。只读。

**语法**

**express.Adjustments**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例将 myDocument 上图形 3 的调整值 1 设为 0.25。*/
ActiveDocument.Shapes.Item(3).Adjustments.Item(1, 0.25)
```

### Shape.AlternativeText

返回或设置与网页的图形相关联的 可选文字?（可选文字：由 Web 浏览器用于在图像下载过程中显示文本，其对象是那些关闭图形的用户，以及那些依赖屏幕阅读软件将屏幕上的图形转换为可读语言的用户。） 。 **String** 类型，可读写。

**语法**

**express.AlternativeText**

\*express是一个代表 Shape 对象的变量。

### Shape.Anchor

返回一个 **Range** 对象，该对象代表指定图形或图形区域的锁定范围。只读。

**语法**

**express.Anchor**

\*express是一个代表 Shape 对象的变量。

**说明**

所有 **Shape** 对象都锁定于某一个文字区域，但可将其置于锁定标记所在页的任意位置。如果创建图形时指定了锁定范围，则锁定标记位于包含该锁定范围的第一段落的段首。如果没有指定锁定范围，则将自动选择锁定范围，图形参照页面的左边框和上边框进行定位。

图形总是与锁定标记处在同一页上。如果图形的 **LockAnchor** 属性为 **True** ，则不能在页面上拖动锁定标记。

**示例**

``` JavaScript
/*本示例选定活动文档第一个图形锁定标记所在的段落。*/
ActiveDocument.Shapes.Item(1).Anchor.Paragraphs.Item(1).Range.Select()
```

### Shape.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Shape 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Shape.AutoShapeType

返回或设置指定的 **Shape** 对象的图形类型，该对象不是代表线条或任意多边形，而是代表自选图形。 **MsoAutoShapeType** 类型，可读写。

**语法**

**express.AutoShapeType**

\*express是一个代表 Shape 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

**示例**

``` JavaScript
/*本示例将活动文档中的所有 16 磅星形替换为 32 磅星形。*/
function ReplaceAutoShape(){
    let shpStar
    let docNew = ActiveDocument
    for(let i = 1; i <= docNew.Shapes.Count; i++){
        shpStar = docNew.Shapes.Item(i)
        if(shpStar.AutoShapeType == msoShape16pointStar){
        shpStar.AutoShapeType = msoShape32pointStar
        }
    }
}
```

### Shape.BackgroundStyle

设置或返回指定形状的背景样式。可读/写 MsoBackgroundStyleIndex 。

**语法**

**express.BackgroundStyle**

\*express是一个代表 Shape 对象的变量。

### Shape.Callout

返回 **CalloutFormat** 对象，该对象包含指定图形的标注格式属性。只读。

**语法**

**express.Callout**

\*express是一个代表 Shape 对象的变量。

**说明**

该属性应用于代表标注的 **Shape** 对象。

**示例**

``` JavaScript
/*该示例在 myDocument 中添加一个椭圆和一个指向椭圆的标注。标注文本不带有边框，但有一条强调线将文本和标注线隔开。*/
function test()
{
let Mydocument = ActiveDocument.Shapes
    Mydocument.AddShape(msoShapeOval, 180, 200, 280, 130)
let myDocument = Mydocument.AddCallout(msoCalloutTwo, 420, 170, 170, 40)
    myDocument.TextFrame.TextRange.Text = "My oval"
let newmyDocument = myDocument.Callout
    newmyDocument.Accent = true
    newmyDocument.Border = false
}
```

### Shape.CanvasItems

返回一个 **CanvasShapes** 对象，该对象代表绘图画布上图形的集合。

**语法**

**express.CanvasItems**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档中创建一个新的绘图画布，并在其上添加一个圆圈。*/
function NewCanvasShape(){
    let shpCanvas = ActiveDocument.Shapes.AddCanvas(100, 75, 150, 200)
    shpCanvas.CanvasItems.AddShape(msoShapeOval, 25, 25, 150, 150)
}
```

### Shape.Chart

返回一个 **Chart** 对象，该对象代表文档中形状集合内的图表。只读。

**语法**

**express.Chart**

\*express是一个代表 Shape 对象的变量。

### Shape.Child

如果图形是子图形或位于图形区域的所有图形都是同一父图形的子图形，则该属性值为 **True** 。 **MsoTriState** 类型，只读。

**语法**

**express.Child**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例选择绘图画布中的第一个图形，并且如果所选图形是一个子图形，用指定颜色填充图形。本示例假定活动文档中的第一个图形是包含多个图形的绘图画布。*/
function FillChildShape(){
    //Select the first shape in the drawing canvas
    let shpCanvasItem = ActiveDocument.Shapes.Item(1).CanvasItems.Item(1)

    //Fill selected shape if it is a child shape
    if(shpCanvasItem.Child == msoTrue){
        shpCanvasItem.Fill.ForeColor.RGB = 100, 0, 200
    }
    else {
        MsgBox("This shape is not a child shape.")
    }
}
```

### Shape.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Shape 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Shape.Fill

返回一个 **FillFormat** 对象，该对象包含指定图形的填充格式属性。只读。

**语法**

**express.Fill**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例首先向 myDocument 添加一个矩形，然后为其设置前景色、背景色和矩形填充的渐变类型。*/
function test()
{
let myDocument = Documents.Item(1)
myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
myDocument.ForeColor.RGB = 128, 0, 0
myDocument.BackColor.RGB = 170, 170, 170
myDocument.TwoColorGradient(msoGradientHorizontal, 1)
}
```

### Shape.Glow

返回一个 **GlowFormat** 对象，该对象代表形状的发光格式。只读。

**语法**

**express.Glow**

\*express是一个代表 Shape 对象的变量。

### Shape.GroupItems

返回一个 **GroupShapes** 对象，该对象代表指定图形组中的单个图形。只读。

**语法**

**express.GroupItems**

\*express是一个代表 Shape 对象的变量。

**说明**

该属性应用于代表组合图形的 **Shape** 对象。使用 **GroupShapes** 对象的 **Item** 方法可从图形组中返回单个图形。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加三个三角形，并组合这些三角形，为整个图形组设置颜色，然后又独立地更改第二个三角形的颜色。*/
function test()
{
let myShapes = ActiveDocument.Shapes

myShapes.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"
myShapes.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"
myShapes.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"

let gruopShapes = myShapes.Range(["shpOne", "shpTwo", "shpThree"]).Group()

gruopShapes.Fill.PresetTextured(msoTextureBlueTissuePaper)
gruopShapes.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble)

}
```

### Shape.HasChart

如果指定的形状具有图表，则为 **True** 。只读。

**语法**

**express.HasChart**

\*express是一个代表 Shape 对象的变量。

**说明**

对于 OLE 图表，此属性总是返回 False。此时请使用 ` InlineShape.OLEFormat.ProgID ` ，并且检查是否存在下列可能的值：“Excel.Chart.8”、“MSGraph.Chart.8”、“Excel.Sheet.8”、“Excel.Chart.5”、“MSGraph.Chart.5”或“Excel.Sheet.5”。

### Shape.HasSmartArt

如果形状上显示 SmartArt 图表则返回 **True** 。只读。

**语法**

**express.HasSmartArt**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示活动文档中的第一个形状是否包含 SmartArt。*/
function test()
{
let myShape = ActiveDocument.Shapes.Item(1)
if(myShape.HasSmartArt){
    MsgBox("The first shape contains SmartArt.")
}
else{
    MsgBox("The first shape contains no SmartArt.")
}
}
```

### Shape.Height

返回或设置指定图形的高度。 **Single** 类型，可读写。

**语法**

**express.Height**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例将一张图片作为嵌入式图形插入活动文档，并更改此图片的高度和宽度*/
function test()
{
    let aInLine = Application.ActiveDocument.InlineShapes.AddPicture("C:\\Windows\\Bubbles.bmp", Selection.Range)
    aInLine.Height = 100
    aInLine.Width = 200
}
```

### Shape.HeightRelative

返回或设置一个 **Single** 类型的值，该值代表形状高度的相对百分比。可读写。

**语法**

**express.HeightRelative**

\*express是一个代表 Shape 对象的变量。

**说明**

将此属性与 **RelativeVerticalSize** 属性一起使用。当设置为 **wdShapeSizeRelativeNone** (-999999)（参见 **WdShapeSizeRelative** 枚举）时，应忽略此属性，因为形状不能使用百分比大小。高度是由 **Height** 属性单独确定的。

**示例**

``` JavaScript
/*设置形状一的HeightRelative属性为50*/
Application.ActiveDocument.Shapes.Item(1).HeightRelative = 50
```

### Shape.HorizontalFlip

表示形状进行过水平翻转。 **MsoTriState** 类型，只读。

**语法**

**express.HorizontalFlip**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中所有进行过水平翻转或垂直翻转的形状还原至初始状态。*/
function FlipShape(){
    for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){
        if(ActiveDocument.Shapes.Item(i).HorizontalFlip){
            ActiveDocument.Shapes.Item(i).Flip(msoFlipHorizontal)
        }
        if(ActiveDocument.Shapes.Item(i).VerticalFlip){
            ActiveDocument.Shapes.Item(i).Flip(msoFlipVertical)
        }
    }
}
```

### Shape.Hyperlink

返回一个 **Hyperlink** 对象，该对象代表与 **Shape** 对象相关联的超链接。只读。

**语法**

**express.Hyperlink**

\*express是一个代表 Shape 对象的变量。

**说明**

``` JavaScript
/*如果不存在与指定形状相关联的超链接，则会发生错误。在此情况下，请将 Add 方法用于 Hyperlinks 集合来为指定形状添加超链接。以下示例说明了具体的操作方法。*/
ActiveDocument.Hyperlinks.Add(Selection.Shapes.Item(1), "http://www.microsoft.com")
```

**示例**

``` JavaScript
/*以下示例显示活动文档中第一个形状的超链接地址。*/
MsgBox(ActiveDocument.Shapes.Item(1).Hyperlink.Address)
```

### Shape.ID

返回指定形状的标识类型。只读 **Long** 类型。

**语法**

**express.ID**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*获取形状的ID*/
Application.ActiveDocument.Shapes.Item(1).ID
```

### Shape.LayoutInCell

返回一个 **Long** 类型的值，该值表示表格中的形状显示在表格内部还是表格外部。

**语法**

**express.LayoutInCell**

\*express是一个代表 Shape 对象的变量。

**说明**

如果该属性为 **True** ，则表示指定的图片显示在表格内部；如果该属性为 **False** ，则表示指定的图片显示在表格外部。

**LayoutInCell** 属性对应于图片格式的 **“高级版式”** 对话框中的 **“表格单元格中的版式”** 选项。

| 注释                                                                                                                                             |
|--------------------------------------------------------------------------------------------------------------------------------------------------|
| 仅当 **WrapFormat** 对象的 **Type** 属性设置为除 **wdWrapTypeInline** 或 **wdWrapTypeNone** 之外的值时，对 **LayoutInCell** 属性的设置才会生效。 |

**示例**

``` JavaScript
/*以下示例对活动文档中的第一个形状禁用“表格单元格中的版式”选项。本示例假设指定的形状在表格内并且不是嵌入形状。*/
ActiveDocument.Shapes.Item(1).LayoutInCell = false
```

### Shape.Left

返回或设置一个 **Single** 类型的值，该值代表指定形状或形状范围的水平位置（以磅为单位）。也可以是任何有效的 **WdShapePosition** 常量。可读写。

**语法**

**express.Left**

\*express是一个代表 Shape 对象的变量。

**说明**

形状的位置是根据形状边框的左上角到形状锁定标记的相对距离进行计算的。 **RelativeHorizontalPosition** 属性控制锁定标记是沿字符、文本栏、页边距还是页面边缘放置。

**示例**

``` JavaScript
/*本示例将活动文档中第一个图形的水平位置设置为距页面的左边缘 1 英寸*/

function test()
{
    let shapes = Application.ActiveDocument.Shapes.Item(1)
    shapes.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    shapes.Left = InchesToPoints(1)
}
```

### Shape.LeftRelative

返回或设置一个 **Single** 类型的值，该值代表形状左侧的相对位置。可读写。

**语法**

**express.LeftRelative**

\*express是一个代表 Shape 对象的变量。

**说明**

将此属性与 **RelativeHorizontalPosition** 属性一起使用。当设置为 **wdShapePositionRelativeNone** (-999999)（参见 **WdShapePositionRelative** 枚举）时，应忽略此属性，因为形状不能使用百分比放置。水平位置是由 **Left** 属性单独确定的。

### Shape.Line

返回一个 **LineFormat** 对象，该对象包含指定形状的线条格式属性。只读。

**语法**

**express.Line**

\*express是一个代表 Shape 对象的变量。

**说明**

对于线条来说， **LineFormat** 对象代表线条本身；而对于带有边框的形状来说， **LineFormat** 对象代表边框。

**示例**

``` JavaScript
/*以下示例向 myDocument 添加一条蓝色虚线。*/
function test()
{
let line = ActiveDocument.Shapes.AddLine(10, 10, 250, 250).Line
line.DashStyle = msoLineDashDotDot
line.ForeColor.RGB = (50, 0, 128)
}
```

``` JavaScript
/*以下示例向 myDocument 添加一个十字形，然后将其边框粗细设置为 8 磅、颜色为红色。*/
function test()
{
let line = ActiveDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line
line.Weight = 8
line.ForeColor.RGB = (255, 0, 0)
}
```

### Shape.LinkFormat

返回一个 **LinkFormat** 对象，该对象代表链接到文件的形状的链接选项。只读。

**语法**

**express.LinkFormat**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例将一个图形作为内嵌形状插入（使用 INCLUDEPICTURE 域），然后显示其源名称 (Tiles.bmp)。*/
function test()
{
let iShape = ActiveDocument.InlineShapes.AddPicture("C:\\windows\\Tiles.bmp", true, false, Selection.Range)
MsgBox(iShape.LinkFormat.SourceName)
}
```

### Shape.LockAnchor

如果 **Shape** 对象的锁定标记锁定到锁定范围，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.LockAnchor**

\*express是一个代表 Shape 对象的变量。

**说明**

如果形状有一个锁定的标记，则不能通过拖动来移动形状的锁定标记。锁定标记不会随形状移动而移动。虽然 **Shape** 对象已锁定到某一文字范围，但是您可以将该对象放在页面上的任何位置。形状锁定到包含锁定范围的第一段的开始。形状总是与其锁定标记位于同一页。

**示例**

``` JavaScript
/*以下示例新建一个文档，为其添加一个形状，然后锁定形状的锁定标记。*/
function test()
{
let myDoc = Documents.Add()
let myShape = myDoc.Shapes.AddShape(msoShapeBalloon, 100, 100, 140, 70)
myShape.LockAnchor = true
ActiveDocument.ActiveWindow.View.ShowObjectAnchors = true
}
```

``` JavaScript
/*以下示例返回一条消息，报告活动文档中每个形状的锁定状态。*/
function test()
{
for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){
    MsgBox("Shape "+ i + " is locked - " +ActiveDocument.Shapes.Item(i).LockAnchor)
}
}
```

### Shape.LockAspectRatio

如果在调整指定形状的大小时保留其最初比例，则该属性值为 **MsoTrue** ；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 **MsoFalse** 。可读写 **MsoTriState** 类型。

**语法**

**express.LockAspectRatio**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 添加一个立方体。可以移动此立方体并调整其大小，但不能重新设置其比例。*/
function test()
{
let myDocument = ActiveDocument
myDocument.Shapes.AddShape(msoShapeCube, 50, 50, 100, 200).LockAspectRatio = msoTrue
}
```

### Shape.Name

返回或设置指定对象的名称。 **String** 类型，可读写。

**语法**

**express.Name**

\*express是一个代表 Shape 对象的变量。

### Shape.Nodes

返回一个 **ShapeNodes** 集合，该集合代表指定形状的几何描述。

**语法**

**express.Nodes**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动文档的图形 3 的第四个顶点的后面添加一个带有一条曲线的平滑顶点。图形 3 必须为至少带有四个顶点的任意多边形。*/
function test()
{
let nodes = ActiveDocument.Shapes.Item(3).Nodes
nodes.Insert(4, msoSegmentCurve, msoEditingSmooth, 210, 100)
}
```

### Shape.OLEFormat

返回一个 **OLEFormat** 对象，该对象代表指定形状、内嵌形状或域的 OLE 特性（链接除外）。只读。

**语法**

**express.OLEFormat**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例遍历活动文档中的所有浮动形状，并将所有链接的 ET 工作表设置为自动更新。*/
function test()
{
for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){
    if(ActiveDocument.Shapes.Item(i).Type == msoLinkedOLEObject){
        if(ActiveDocument.Shapes.Item(i).OLEFormat.ProgID == "Excel.Sheet"){
            ActiveDocument.Shapes.Item(i).LinkFormat.AutoUpdate = ture
        }
    }
}
}
```

### Shape.Parent

返回一个 **Object** 类型值，该值代表指定 **Shape** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Shape 对象的变量。

### Shape.ParentGroup

返回一个 **Shape** 对象，该对象代表子形状或子形状范围的通用父形状。

**语法**

**express.ParentGroup**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动文档中创建两个形状并对其进行分组。然后利用一个形状访问该组的父组，并使用同样的填充颜色填充父组中的所有形状。本示例假定活动文档当前并不包括形状，如果包括形状，则产生错误。*/
function ParentGroupShape(){
    //Add two shapes to active document and group
    let shapes = ActiveDocument.Shapes
    shapes.AddShape(msoShapeOval, 72, 72, 100, 100)
    shapes.AddShape(msoShapeHeart, 110, 120, 100, 100)
    shapes.Range(Array(1, 2)).Group()
    let pgShape = ActiveDocument.Shapes.Item(1).GroupItems.Item(1).ParentGroup
    pgShape.Fill.ForeColor.RGB = (100, 0, 255)
}
```

### Shape.PictureFormat

返回一个 **PictureFormat** 对象，该对象包含指定对象的图片格式属性。

**语法**

**express.PictureFormat**

\*express是一个代表 Shape 对象的变量。

**说明**

该属性适用于代表图片或 OLE 对象的 **Shape** 对象。

**示例**

``` JavaScript
/*以下示例设置了 myDocument 中第一个形状的亮度和对比度。第一个形状必须是图片或 OLE 对象。*/
function test()
{
let myDocument = ActiveDocument.Shapes.Item(1).PictureFormat
myDocument.Brightness = 0.3
myDocument.Contrast = .75
}
```

### Shape.Reflection

返回一个 **ReflectionFormat** 对象，该对象代表形状的反射格式。只读。

**语法**

**express.Reflection**

\*express是一个代表 Shape 对象的变量。

### Shape.RelativeHorizontalPosition

指定形状的相对水平位置。可读写 **WdRelativeHorizontalPosition** 类型。

**语法**

**express.RelativeHorizontalPosition**

\*express是一个代表 Shape 对象的变量。

### Shape.RelativeHorizontalSize

返回或设置一个 **WdRelativeVerticalSize** 常量，该常量代表形状区域相对的对象。可读写。

**语法**

**express.RelativeHorizontalSize**

\*express是一个代表 Shape 对象的变量。

**说明**

此属性可与 **WidthRelative** 属性一起使用。

### Shape.RelativeVerticalPosition

指定形状的相对垂直位置。可读写 **WdRelativeVerticalPosition** 类型。

**语法**

**express.RelativeVerticalPosition**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例重新定位活动文档中第一个形状对象。*/
function test()
{
let shapes = ActiveDocument.Shapes.Item(1)
shapes.Left = InchesToPoints(0.6)
shapes.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
shapes.Top = InchesToPoints(1)
shapes.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
}
```

### Shape.RelativeVerticalSize

返回或设置一个 **WdRelativeVerticalSize** 常量，该常量代表形状的相对垂直大小。可读写。

**语法**

**express.RelativeVerticalSize**

\*express是一个代表 Shape 对象的变量。

**说明**

此属性可与 **HeightRelative** 属性一起使用。

### Shape.Rotation

返回或设置指定图形绕 Z 轴旋转的度数。正数表示按顺时针方向进行旋转，负值表示按逆时针方向旋转。 **Single** 类型，可读写。

**语法**

**express.Rotation**

\*express是一个代表 Shape 对象的变量。

**说明**

要设置三维图形绕 X 轴或 Y 轴的旋转，请使用 **ThreeDFormat** 对象的 **RotationX** 属性或 **RotationY** 属性。

**示例**

``` JavaScript
/*以下示例将 myDocument 中所有形状的旋转量设置为与形状 1 的旋转量一致。*/
function test()
{
let aShape = ActiveDocument.Shapes
let sh1Rotation = aShape.Item(1).Rotation
for(let i = 1; i <= aShape.Count; i++){
    aShape.Item(i).Rotation = sh1Rotation
}
}
```

### Shape.Script

返回一个 **Script** 对象，该对象代表网页中图像的一段脚本或代码。

**语法**

**express.Script**

\*express是一个代表 Shape 对象的变量。

**说明**

如果该网页中不包含脚本，则不返回任何内容。

**示例**

``` JavaScript
/*以下示例显示活动文档中第一个形状所用的脚本语言的类型。*/
function test()
{
let objScr = ActiveDocument.Shapes.Item(1).script
if(objScr){
    switch(objScr.Language){
        case msoScriptLanguageVisualBasic:
            MsgBox("VBScript")
            break
        case msoScriptLanguageVisualBasic:
            MsgBox("JavaScript")
            break
        case msoScriptLanguageVisualBasic:
            MsgBox("Active Server Pages")
            break
        default:
            Msgbox("Other scripting language")
    }
}
}
```

### Shape.Shadow

返回一个 **ShadowFormat** 对象，该对象代表指定形状的阴影格式。

**语法**

**express.Shadow**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例向活动文档中添加一个带阴影格式的箭头。*/
function test()
{
let myShape = ActiveDocument.Shapes.AddShape(msoShapeRightArrow, 90, 79, 64, 43)
myShape.Shadow.Type = msoShadow5
}
```

### Shape.ShapeStyle

返回或设置指定形状的形状样式。可读/写 MsoShapeStyleIndex。

**语法**

**express.ShapeStyle**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例更改活动文档中第一个形状的形状样式。*/
function test()
{
let myShape = ActiveDocument.Shapes.Item(1)
myShape.ShapeStyle = msoLineStylePreset12
}
```

### Shape.SmartArt

返回 SmartArt 对象，该对象提供一种处理与指定形状相关联的 SmartArt 的方式。只读。

**语法**

**express.SmartArt**

\*express是一个代表 Shape 对象的变量。

**说明**

**SmartArt** 属性为与形状相关联的 SmartArt 图形进行交互提供了一个入口点。

**示例**

``` JavaScript
/*下面的代码示例向活动文档中添加 SmartArt 图形。*/
function test()
{
let myShape = ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts.Item(1), 100, 100, 400, 400)
let mySmartArt = myShape.SmartArt
}
```

### Shape.SoftEdge

返回一个 **SoftEdgeFormat** 对象，该对象代表形状的软边缘格式。只读。

**语法**

**express.SoftEdge**

\*express是一个代表 Shape 对象的变量。

### Shape.TextEffect

返回一个 **TextEffectFormat** 对象，该对象包含指定形状的文本效果格式属性。只读。

**语法**

**express.TextEffect**

\*express是一个代表 Shape 对象的变量。

**说明**

该属性适用于代表艺术字的 **Shape** 对象。

**示例**

``` JavaScript
/*以下示例中，如果 myDocument 的第三个形状为艺术字，则将它的字形设为加粗。*/
function test()
{
let myShape = ActiveDocument.Shapes.Item(3)
if(myShape.Type == msoTextEffect){
    myShape.TextEffect.FontBold = true
}
}
```

### Shape.TextFrame

返回一个 **TextFrame** 对象，该对象包含指定形状的文字。

**语法**

**express.TextFrame**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例在 myDocument 中添加一个矩形，在该矩形中添加文本，并为文本框架设置边距。*/
function test()
{
let textframe = ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
textframe.TextRange.Text = "Here is some test text"
textframe.MarginBottom = 0
textframe.MarginLeft = 100
textframe.MarginRight = 0
textframe.MarginTop = 20
}
```

### Shape.TextFrame2

返回一个 **TextFrame2** 对象，该对象包含指定形状的文本。只读。

**语法**

**express.TextFrame2**

\*express是一个代表 Shape 对象的变量。

### Shape.ThreeD

返回一个 **ThreeDFormat** 对象，该对象包含指定形状的三维格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例设置应用于 myDocument 中第一个形状的三维效果的深度、延伸颜色、延伸方向以及照明方向。*/
function test()
{
let myDocument = ActiveDocument.Shapes.Item(1).ThreeD
myDocument.Visible = true
myDocument.Depth = 50
// RGB value for purple
myDocument.ExtrusionColor.RGB = (255, 100, 255)
myDocument.SetExtrusionDirection(msoExtrusionTop)
myDocument.PresetLightingDirection = msoLightingLeft
}
```

### Shape.Title

返回或设置包含指定形状的标题的 **String** 类型值。可读/写。

**语法**

**express.Title**

\*express是一个代表 Shape 对象的变量。

**说明**

使用 **Title** 属性可为形状提供可选文字标题。该属性将标题文本添加到 WPS 2015 中 **“设置形状格式”** 对话框的 **“可选文字”** 窗格上的 **“标题”** 文本框中。

| 注释                                                                                                                       |
|----------------------------------------------------------------------------------------------------------------------------|
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

``` JavaScript
/*下面的代码示例向活动文档中的第二个形状中添加可选文字标题。*/
ActiveDocument.Shapes.Item(2).Title = "Shape 2."
```

### Shape.Top

返回或设置指定形状或形状范围的垂直位置（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.Top**

\*express是一个代表 Shape 对象的变量。

**说明**

形状的位置是根据形状边框的左上角到形状锁定标记的相对距离进行计算的。 **RelativeVerticalPosition** 属性控制形状的锁定标记是沿行、段落、页边距还是页面边缘放置。

对于包含多个形状的 **ShapeRange** 对象， **Top** 属性设置每个形状的垂直位置。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个形状的垂直位置设置为距页面顶部 1 英寸*/

function test()
{
    let shapes = Application.ActiveDocument.Shapes.Item(1)
    shapes.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    shapes.Top = InchesToPoints(1)
}
```

### Shape.TopRelative

返回或设置一个 **Single** 类型的值，该值代表形状顶部的相对位置。可读写。

**语法**

**express.TopRelative**

\*express是一个代表 Shape 对象的变量。

**说明**

将此属性与 **RelativeHorizontalPosition** 属性一起使用。当设置为 **wdShapePositionRelativeNone** (-999999)（参见 **WdShapePositionRelative** 枚举）时，应忽略此属性，因为形状不能使用百分比放置。垂直位置是由 **Top** 属性单独确定的。

### Shape.Type

返回内嵌形状的类型。只读 **MsoShapeType** 类型。

**语法**

**express.Type**

\*express是一个代表 Shape 对象的变量。

### Shape.VerticalFlip

如果指定形状围绕垂直轴进行翻转，则该属性值为 **True** 。 **MsoTriState** 类型，只读。

**语法**

**express.VerticalFlip**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例将 myDocument 中所有进行过水平翻转或垂直翻转的形状还原至初始状态。*/
function test()
{
for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){
    if(ActiveDocument.Shapes.Item(i).HorizontalFlip == -1){
        ActiveDocument.Shapes.Item(i).Flip(msoFlipHorizontal)
    }
    if(ActiveDocument.Shapes.Item(i).VerticalFlip == -1){
        ActiveDocument.Shapes.Item(i).Flip(msoFlipVertical)
    }
}
}
```

### Shape.Vertices

该属性以一系列 坐标对 ?（坐标对：一对值，表示两维数组中存储的点的 x 和 y 坐标，该数组中包含许多点的坐标。） 的形式返回指定任意多边形图形顶点（和贝赛尔曲线的控点）的坐标。只读 **Variant** 类型。

**语法**

**express.Vertices**

\*express是一个代表 Shape 对象的变量。

**说明**

可将该属性返回的数组用作 **AddCurve** 或 **AddPolyLine** 方法的参数。

下表显示 **Vertices** 属性如何将数组 *vertArray()* 的值与三角形的顶点坐标相关联。

vertArray

``` JavaScript
/*第一个顶点至文档的左边界的水平距离。*/
vertArray(1, 1)


/*第一个顶点至文档顶端的垂直距离。*/
vertArray(1, 2)


/*第二个顶点至文档的左边界的水平距离。*/
vertArray(2, 1)


/*第二个顶点至文档顶端的垂直距离。*/
vertArray(2, 2)


/*第三个顶点至文档的左边界的水平距离。*/
vertArray(3, 1)


/*第三个顶点至文档顶端的垂直距离。*/
vertArray(3, 2)
```

**示例**

``` JavaScript
/*以下示例将活动文档中第一个形状的顶点坐标赋给一个数组变量，并显示第一个顶点的坐标。第一个形状必须是任意多边形图形。*/
function test()
{
let ishape = ActiveDocument.Shapes.Item(1)
vertArray = ishape.Vertices
x1 = vertArray[1][1]
y1 = vertArray[1][2]
MsgBox("First vertex coordinates: "+ x1 + "," + y1)
}
```

``` JavaScript
/*以下示例创建一条与活动文档中第一个形状具有相同几何特征的曲线。该示例假定第一个形状为包含 3n+1 个顶点的贝赛尔曲线，其中 n 为曲线段数量。*/
function test()
{
let shapes = ActiveDocument.Shapes
shapes.AddCurve(shapes.Item(1).Vertices, Selection.Range)
}
```

### Shape.Visible

如果指定对象或应用于该对象的格式是可见的，则该属性值为 **True** 。 **MsoTriState** 类型，可读写。

**语法**

**express.Visible**

\*express是一个代表 Shape 对象的变量。

**说明**

如果 **Visible** 属性为 **False** ，则某些方法和属性可能无法使用。

**示例**

``` JavaScript
/*本示例新建一篇文档，然后在文档中添加文本和一个矩形，同时将 WPS 设置为在打印文档时隐藏该矩形，然后在打印结束后再显示该图形。*/
function test()
{
let myDoc = Documents.Add(Selection.TypeText("This is some sample text."))
myDoc.Shapes.AddShape(msoShapeRectangle, 200, 70, 150, 60)
myDoc.Shapes.Item(1).Visible = false
myDoc.PrintOut()
myDoc.Shapes.Item(1).Visible = true
}
```

### Shape.Width

返回或设置指定形状的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

\*express是一个代表 Shape 对象的变量。

### Shape.WidthRelative

返回或设置一个代表形状的相对宽度的 **Single** 。可读写。

**语法**

**express.WidthRelative**

\*express是一个代表 Shape 对象的变量。

**说明**

将此属性与 **RelativeVerticalSize** 属性一起使用。当设置为 **wdShapeSizeRelativeNone** (-999999)（参见 **WdShapeSizeRelative** 枚举）时，应忽略此属性，因为形状不能使用百分比大小。宽度是由 **Width** 属性单独确定的。

### Shape.WrapFormat

返回一个 **WrapFormat** 对象，该对象包含在指定的形状四周文字环绕的属性。只读。

**语法**

**express.WrapFormat**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动文档中添加一个椭圆，并指定文档文字环绕该椭圆外切矩形的左右两侧。该示例将文档文字与该矩形四边之间的边距设置为 0.1 英寸。*/
function test()
{
let myOval = ActiveDocument.Shapes.AddShape(msoShapeOval, 36, 36, 90, 50).WrapFormat
myOval.Type = wdWrapSquare
myOval.Side = wdWrapBoth
myOval.DistanceTop = InchesToPoints(0.1)
myOval.DistanceBottom = InchesToPoints(0.1)
myOval.DistanceLeft = InchesToPoints(0.1)
myOval.DistanceRight = InchesToPoints(0.1)
}
```

### Shape.ZOrderPosition

返回一个 **Long** 类型的值，该值代表指定的形状在 Z 顺序中的位置。只读。

**语法**

**express.ZOrderPosition**

\*express是一个代表 Shape 对象的变量。

**说明**

` Shapes(1) ` 返回 Z 顺序中的最后一个形状，而 ` Shapes(Shapes.Count) ` 返回 Z 顺序中的第一个形状。该属性为只读。要设置形状在 Z 顺序中的位置，请使用 **ZOrder** 方法。

形状在 Z 顺序中的位置与 Shapes 集合中形状的索引号相对应。例如，如果在 myDocument 上有四个形状，则表达式 ` myDocument.Shapes(1) ` 返回 Z 顺序中的最后一个形状，而表达式 ` myDocument.Shapes(4) ` 返回 Z 顺序中的第一个形状。

无论何时将新形状添加到集合中，默认情况下该形状都将添加到 Z 顺序的最前端。

**示例**

``` JavaScript
/*此示例首先向 myDocument 中添加一个椭圆形。如果该文档另外至少还有一个形状，则按照 Z 顺序将此椭圆形放置在倒数第二的位置。*/
function test()
{
let myDocument = ActiveDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)
while(myDocument.ZOrderPosition > 2){
    myDocument.ZOrder(msoSendBackward)
}
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Shape](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Windows 对象

## 方法

| 序号 | 名称                                                          | 说明                                                                                      |
|:----:|:--------------------------------------------------------------|:------------------------------------------------------------------------------------------|
|  1   | [Add](#Windows.Add)                                           | 返回一个 Window 对象，该对象代表文档的新窗口。                                            |
|  2   | [Arrange](#Windows.Arrange)                                   | 在应用程序工作区排列所有打开的文档窗口。                                                  |
|  3   | [BreakSideBySide](#Windows.BreakSideBySide)                   | 如果两个窗口以并排方式排列，则终止该模式。返回一个代表该方法是否成功的 Boolean 类型的值。 |
|  4   | [CompareSideBySideWith](#Windows.CompareSideBySideWith)       | 打开两个窗口并以并排方式排列。返回一个 Boolean 类型的值。                                 |
|  5   | [Item](#Windows.Item)                                         | 返回集合中的单个 Window 对象。                                                            |
|  6   | [ResetPositionsSideBySide](#Windows.ResetPositionsSideBySide) | 重置 “并排比较” 查看模式中的两个文档窗口。                                                |

## 属性

| 序号 | 名称                                                        | 说明                                                                         |
|:----:|:------------------------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Windows.Application)                         | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Windows.Count)                                     | 返回一个 Long 类型的值，该值代表集合中窗口的数量。只读。                     |
|  3   | [Creator](#Windows.Creator)                                 | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#Windows.Parent)                                   | 返回一个 Object 类型值，该值代表指定 Windows 对象的父对象。                  |
|  5   | [SyncScrollingSideBySide](#Windows.SyncScrollingSideBySide) | 如果为 True ，则启用同时滚动窗口内容。 Boolean 类型，可读写。                |

## 成员方法

### Windows.Add

返回一个 **Window** 对象，该对象代表文档的新窗口。

**语法**

**express.Add ( *Window* )**

\*express是一个代表 Windows 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                           |
|----------|-----------|----------|--------------------------------------------------------------------------------|
| *Window* | 可选      | Variant  | 需要再打开一个窗口的 Window 对象。如果省略该参数，则为活动文档打开一个新窗口。 |

**说明**

当为某个文档打开多个窗口时，在窗口标题中会显示冒号 (：) 和数字。

**示例**

``` JavaScript
/*本示例为显示在活动窗口中的文档打开一个新窗口*/
Application.Windows.Add()
/*本示例为 MyDoc.doc 打开一个新窗口*/
Application.Windows.Add(Application.Documents.Item("MyDoc.doc").Windows.Item(1))
```

### Windows.Arrange

在应用程序工作区排列所有打开的文档窗口。

**语法**

**express.Arrange ( *ArrangeStyle* )**

\*express是一个代表 Windows 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                     |
|----------------|-----------|----------|--------------------------------------------------------------------------|
| *ArrangeStyle* | 可选      | Variant  | 窗口排列方式。可以是下列 WdArrangeStyle 常量之一：wdIcons 或者 wdTiled。 |

**说明**

WPS 使用的是单个文档界面（SDI），本方法不会产生任何效果。

**示例**

``` JavaScript
/*本示例排列所有打开的窗口，使其不相互重叠。*/
Windows.Arrange(wdTiled)
```

``` JavaScript
/*本示例将所有打开的窗口最小化，然后排列这些窗口。*/
function test()
{
for(let i = 1; i <= Windows.Count; i++ ){
    Windows.Item(i).Activate()
    Windows.Item(i).WindowState = wdWindowStateMinimize
}

Windows.Arrange(wdIcons)
}
```

### Windows.BreakSideBySide

如果两个窗口以并排方式排列，则终止该模式。返回一个代表该方法是否成功的 **Boolean** 类型的值。

**语法**

**express.BreakSideBySide ()**

\*express是一个代表 Windows 对象的变量。

**返回值**

Boolean

**示例**

``` JavaScript
/*以下示例结束并排模式。*/
ActiveDocument.Windows.BreakSideBySide()
```

### Windows.CompareSideBySideWith

打开两个窗口并以并排方式排列。返回一个 **Boolean** 类型的值。

**语法**

**express.CompareSideBySideWith ( *Document* )**

\*express是一个代表 Windows 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                       |
|------------|-----------|----------|----------------------------|
| *Document* | 必选      | Variant  | 要在并排窗口中查看的文档。 |

**返回值**

Boolean

**说明**

不能将 **Application** 对象或 **ActiveDocument** 属性与 **CompareSideBySideWith** 方法配合使用。

**示例**

``` JavaScript
/*以下示例将两个新文档排列在相邻的窗口中。*/
function test()
{
let objDoc1 = Documents.Add()
let objDoc2 = Documents.Add()

objDoc2.Activate()
objDoc2.Windows.CompareSideBySideWith (objDoc1)
Windows.ResetPositionsSideBySide()
}
```

### Windows.Item

返回集合中的单个 **Window** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Windows 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

### Windows.ResetPositionsSideBySide

重置 **“并排比较”** 查看模式中的两个文档窗口。

**语法**

**express.ResetPositionsSideBySide ()**

\*express是一个代表 Windows 对象的变量。

**说明**

对应于 **“并排比较”** 工具栏上的 **“重设窗口位置”** 按钮。使用 **ResetPositionsSideBySide** 方法可重置两个文档的显示。例如，如果用户最大化或最小化正在进行比较的两个文档窗口中的一个，则 **ResetPositionsSideBySide** 方法重置显示，从而再次并排显示两个窗口。

**示例**

``` JavaScript
/*以下示例将以前放置在并排窗口中的两个文档放在相邻的窗口中。*/
Windows.ResetPositionsSideBySide()
```

## 成员属性

### Windows.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Windows 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Windows.Count

返回一个 **Long** 类型的值，该值代表集合中窗口的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Windows 对象的变量。

### Windows.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Windows 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Windows.Parent

返回一个 **Object** 类型值，该值代表指定 **Windows** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Windows 对象的变量。

### Windows.SyncScrollingSideBySide

如果为 **True** ，则启用同时滚动窗口内容。 **Boolean** 类型，可读写。

**语法**

**express.SyncScrollingSideBySide**

\*express是一个代表 Windows 对象的变量。

**说明**

如果该值为 **False** ，则禁用同时滚动窗口。

**示例**

``` JavaScript
/*下列示例启用同时滚动相邻窗口的功能。*/
function test()
{
let objDoc1 = Documents.Add()
let objDoc2 = Documents.Add()

objDoc2.Activate()
objDoc2.Windows.CompareSideBySideWith(objDoc1)
Windows.ResetPositionsSideBySide()
Windows.SyncScrollingSideBySide = true
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Windows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Words 对象

## 简介

所选内容、范围或文档中的单词的集合。 Words 集合中的每一项均为代表一个单词的 Range 对象。不存在 WPS 对象。

## 方法

| 序号 | 名称                | 说明                          |
|:----:|:--------------------|:------------------------------|
|  1   | [Item](#Words.Item) | 返回集合中的单个 Range 对象。 |

## 属性

| 序号 | 名称                              | 说明                                                                         |
|:----:|:----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Words.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Words.Count)             | 返回一个 Long 类型的值，该值代表集合中的字数。只读。                         |
|  3   | [Creator](#Words.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [First](#Words.First)             | 返回一个 Range 对象，该对象代表单词集合中的第一个单词。                      |
|  5   | [Last](#Words.Last)               | 返回一个 Range 对象，该对象代表单词集合中的最后一个单词。                    |
|  6   | [Parent](#Words.Parent)           | 返回一个 Object 类型值，该值代表指定 Words 对象的父对象。                    |

## 成员方法

### Words.Item

返回集合中的单个 **Range** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Words 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

## 成员属性

### Words.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 Words 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Words.Count

返回一个 **Long** 类型的值，该值代表集合中的字数。只读。

**语法**

**express.Count**

\*express是一个代表 Words 对象的变量。

**示例**

``` JavaScript
/*本示例显示选定内容的字数*/
if(Application.Selection.Words.Count >= 1 && Application.Selection.Type != wdSelectionIP){
    alert("The selection contains " + Application.Selection.Words.Count + " words.")
}
```

### Words.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Words 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Words.First

返回一个 **Range** 对象，该对象代表单词集合中的第一个单词。

**语法**

**express.First**

\*express是一个代表 Words 对象的变量。

### Words.Last

返回一个 **Range** 对象，该对象代表单词集合中的最后一个单词。

**语法**

**express.Last**

\*express是一个代表 Words 对象的变量。

**示例**

``` JavaScript
/*以下示例对所选内容中的最后一个单词应用加粗格式。*/
function test()
{
if(Selection.Words.Count >= 2){
    Selection.Words.Last.Bold = true
}
}
```

### Words.Parent

返回一个 **Object** 类型值，该值代表指定 **Words** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Words 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Words](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XMLSchemaReference 对象

## 简介

代表附加到文档的单个架构。

## 说明

使用 XMLSchemaReference 属性可返回用于 XMLSchemaReference 对象的 ChildNodeSuggestion 对象。如果引用的 XML 架构是 SimpleSample 架构，则以下示例插入建议的 XML 子元素。

``` JavaScript
function test()
{
let objSuggestion = ActiveDocument.ChildNodeSuggestions
for(let i = 
 1;  i < =  objSuggestion.Count; i++) {    
    if(objSuggestion.XMLSchemaReference = "SimpleSample") {
        objSuggestion.Item(i).Insert()
    }
}
}
```

## 方法

| 序号 | 名称                                 | 说明                              |
|:----:|:-------------------------------------|:----------------------------------|
|  1   | [Delete](#XMLSchemaReference.Delete) | 删除指定的 XML 架构引用。         |
|  2   | [Reload](#XMLSchemaReference.Reload) | 重新加载在文档中引用的 XML 架构。 |

## 属性

| 序号 | 名称                                             | 说明                                                                         |
|:----:|:-------------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#XMLSchemaReference.Application)   | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Creator](#XMLSchemaReference.Creator)           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  3   | [Location](#XMLSchemaReference.Location)         | 返回一个 String 类型的值，该值代表 XML 架构的物理位置。只读。                |
|  4   | [NamespaceURI](#XMLSchemaReference.NamespaceURI) | 返回 String 值，该值表示指定对象的架构命名空间的统一资源标识符 (URI)。只读。 |
|  5   | [Parent](#XMLSchemaReference.Parent)             | 返回一个 Object 类型值，该值代表指定 XMLSchemaReference 对象的父对象。       |

## 成员方法

### XMLSchemaReference.Delete

删除指定的 XML 架构引用。

**语法**

**express.Delete ()**

\*express是一个代表 XMLSchemaReference 对象的变量。

### XMLSchemaReference.Reload

重新加载在文档中引用的 XML 架构。

**语法**

**express.Reload ()**

\*express是一个代表 XMLSchemaReference 对象的变量。

**返回值**

无

## 成员属性

### XMLSchemaReference.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 XMLSchemaReference 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XMLSchemaReference.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XMLSchemaReference 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XMLSchemaReference.Location

返回一个 **String** 类型的值，该值代表 XML 架构的物理位置。只读。

**语法**

**express.Location**

\*express是一个代表 XMLSchemaReference 对象的变量。

### XMLSchemaReference.NamespaceURI

返回 **String** 值，该值表示指定对象的架构命名空间的统一资源标识符 (URI)。只读。

**语法**

**express.NamespaceURI**

\*express是一个代表 XMLSchemaReference 对象的变量。

**说明**

如果创建的是用于 WPS 的 XML 架构，则强烈建议在架构中指定 targetNamespace 设置。

**示例**

以下示例重新加载 SimpleSample 架构，如果该架构没有附加到活动文档，则附加该架构。

``` JavaScript
/*SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中。有关详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document SDK。*/
function test()
{
if(ActiveDocument.XMLSchemaReferences.Item(1).NamespaceURI != "SimpleSample") {
    Application.XMLNamespaces.Item("SimpleSample").AttachToDocument(ActiveDocument)
}
}
```

### XMLSchemaReference.Parent

返回一个 **Object** 类型值，该值代表指定 **XMLSchemaReference** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XMLSchemaReference 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XMLSchemaReference](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ActionSetting 对象

## 简介

包含指定形状或文本范围在幻灯片放映中对鼠标动作的反应的有关信息。

## 说明

ActionSetting 对象是 ActionSettings 集合的成员。

ActionSettings 集合包含一个代表在幻灯片放映中用户单击指定对象时该对象的反应的 ActionSetting 对象，以及一个代表在幻灯片放映中用户将鼠标指针移动到指定对象上时该对象的反应的 ActionSetting 对象。

如果设置了 ActionSetting 对象的属性但没有效果，请确认已将 Action 属性设为适当的值。

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                                                                                                                                                 |
|:----:|:----------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Action](#ActionSetting.Action)               | 返回或设置在幻灯片放映过程中单击指定形状或将鼠标指针定位到形状上时将发生的动作的类型。                                                                                                                                                                                               |
|  2   | [ActionVerb](#ActionSetting.ActionVerb)       | 返回或设置包含 OLE 动作的字符串，该 OLE 动作将在用户在幻灯片放映过程中单击指定形状或将鼠标指针移过该形状时运行。要使该属性可影响幻灯片放映操作，必须先将 Action 属性设置为 ppActionOLEVerb。String 类型，可读/写。                                                                   |
|  3   | [AnimateAction](#ActionSetting.AnimateAction) | 发生指定的鼠标操作时，如果指定形状的颜色立即反转，则为 MsoTrue。MsoTriState 类型，可读/写 。                                                                                                                                                                                         |
|  4   | [Application](#ActionSetting.Application)     | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                              |
|  5   | [Hyperlink](#ActionSetting.Hyperlink)         | 返回一个 Hyperlink 对象，该对象代表指定形状的超链接。要在幻灯片放映过程中激活超链接，必须将 Action 属性设为 ppActionHyperlink。只读。                                                                                                                                                |
|  6   | [Parent](#ActionSetting.Parent)               | 返回指定对象的父对象。                                                                                                                                                                                                                                                               |
|  7   | [Run](#ActionSetting.Run)                     | 返回或设置演示文稿或宏的名称，当在幻灯片放映过程中单击指定形状或鼠标指针经过该形状时，将运行该演示文稿或宏。Action 属性必须设置为 ppActionRunMacro 或 ppActionRunProgram，此属性才能影响幻灯片的放映动作。可读/写。String 类型。                                                     |
|  8   | [ShowAndReturn](#ActionSetting.ShowAndReturn) | 决定 WPP是否返回初始的幻灯片放映以及返回的条件。可读/写。MsoTriState 类型。                                                                                                                                                                                                          |
|  9   | [SlideShowName](#ActionSetting.SlideShowName) | 返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（ActionSetting 对象）；返回或设置自定义幻灯片放映的名称以便打印（PrintOptions 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（PublishObject 对象）。String 类型，可读/写。 |
|  10  | [SoundEffect](#ActionSetting.SoundEffect)     | 返回一个 SoundEffect 对象，该对象代表切换到指定幻灯片时播放的声音。                                                                                                                                                                                                                  |

## 成员属性

### ActionSetting.Action

返回或设置在幻灯片放映过程中单击指定形状或将鼠标指针定位到形状上时将发生的动作的类型。

**语法**

**express.Action**

\*express是一个代表 ActionSetting 对象的变量。

**说明**

可以是下列 **PpActionType** 常量之一。

|                         |
|-------------------------|
| ppActionEndShow         |
| ppActionFirstSlide      |
| ppActionHyperlink       |
| ppActionLastSlide       |
| ppActionLastSlideViewed |
| ppActionMixed           |
| ppActionNamedSlideShow  |
| ppActionNextSlide       |
| ppActionNone            |
| ppActionOLEVerb         |
| ppActionPlay            |
| ppActionPreviousSlide   |
| ppActionRunMacro        |
| ppActionRunProgram      |

可将 **Action** 属性与 **ActionSetting** 对象的其他属性结合使用，如下表所示。

| 如果将Action 属性设置为此值 | 使用此属性    | 进行此操作                                                                 |
|-----------------------------|---------------|----------------------------------------------------------------------------|
| ppActionHyperlink           | Hyperlink     | 设置超链接的属性，在幻灯片放映过程中为响应形状上的鼠标动作将转至该超链接。 |
| ppActionRunProgram          | Run           | 返回或设置在幻灯片放映过程中为响应形状上的鼠标动作而运行的程序的名称。     |
| ppActionRunMacro            | Run           | 返回或设置在幻灯片放映过程中为响应形状上的鼠标动作而运行的宏的名称。       |
| ppActionOLEVerb             | ActionVerb    | 设置在幻灯片放映过程中为响应形状上的鼠标动作而调用的 OLE 动作。            |
| ppActionNamedSlideShow      | SlideShowName | 设置自定义放映名称以响应幻灯片放映过程中对形状的鼠标操作。                 |

**示例**

``` JavaScript
/*本示例设置活动演示文稿第一张幻灯片的第三个形状（OLE 对象），使其在幻灯片放映过程中当鼠标移过它时播放。可读/写 Long 类型。*/
function test(){
let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).ActionSettings.Item(ppMouseOver)
Act.ActionVerb = "Play"
Act.Action = ppActionOLEVerb
}
```

### ActionSetting.ActionVerb

返回或设置包含 OLE 动作的字符串，该 OLE 动作将在用户在幻灯片放映过程中单击指定形状或将鼠标指针移过该形状时运行。要使该属性可影响幻灯片放映操作，必须先将 Action 属性设置为 ppActionOLEVerb。String 类型，可读/写。

**语法**

**express.ActionVerb**

\*express是一个代表 ActionSetting 对象的变量。

**示例**

``` JavaScript
/*本示例将第一张幻灯片的第三个形状设置为：只要鼠标指针在幻灯片放映时移过该形状就会被播放。但第三个形状必须是支持“Play”动词的 OLE 对象。*/
function test(){
　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).ActionSettings.Item(ppMouseOver)
　　　　Act.ActionVerb = "Play"
　　　　Act.Action = ppActionOLEVerb}
```

### ActionSetting.AnimateAction

发生指定的鼠标操作时，如果指定形状的颜色立即反转，则为 MsoTrue。MsoTriState 类型，可读/写 。

**语法**

**express.AnimateAction**

\*express是一个代表 ActionSetting 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue                                       |

**示例**

``` JavaScript
/*本示例设置在幻灯片放映过程中，单击活动演示文稿第一张幻灯片的第三个形状时播放掌声，并立即反转该形状的颜色。
*/
function test(){
　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).ActionSettings.Item(ppMouseClick)
　　　　Act.SoundEffect.Name = "applause"
　　　　Act.AnimateAction = msoTrue
}
```

### ActionSetting.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 ActionSetting 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。
*/
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

``` JavaScript
/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test(){
　　　　for(let i= 1;i< = ActivePresentation.Slides.Item(1).Shapes.Count;i++) {
    　　　　if(ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {
        　　　　MsgBox(ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)
    　　　　}
　　　　}
}
```

### ActionSetting.Hyperlink

返回一个 Hyperlink 对象，该对象代表指定形状的超链接。要在幻灯片放映过程中激活超链接，必须将 Action 属性设为 ppActionHyperlink。只读。

**语法**

**express.Hyperlink**

\*express是一个代表 ActionSetting 对象的变量。

**示例**

``` JavaScript
/*本示例对当前演示文稿中第一张幻灯片上的第一个形状进行设置，使其在幻灯片演示过程中，如果被单击则跳转到 WPS 网站上。*/
function test(){
　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick)
　　　　Act.Action = ppActionHyperlink
　　　　Act.Hyperlink.Address = "http://www.wps.cn/"
}
```

### ActionSetting.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ActionSetting 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 */
function test(){
　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes
　　　　let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
　　　　addShp.TextRange.Text = "Test text"
　　　　addShp.Parent.Rotation = 45
}
```

### ActionSetting.Run

返回或设置演示文稿或宏的名称，当在幻灯片放映过程中单击指定形状或鼠标指针经过该形状时，将运行该演示文稿或宏。Action 属性必须设置为 ppActionRunMacro 或 ppActionRunProgram，此属性才能影响幻灯片的放映动作。可读/写。String 类型。

**语法**

**express.Run**

\*express是一个代表 ActionSetting 对象的变量。

**说明**

如果 **Action** 属性的值为 **ppActionRunMacro** ，则指定的字符串值应为一个当前已加载的全局宏的名称。如果 **Action** 属性的值为 **ppActionRunProgram** ，则指定的字符串应为一个程序的完整路径和文件名。

可以将 **Run** 属性设为无参数的或单个 *Shape* 或 *Object* 参数的宏。幻灯片放映时被单击的形状将被作为此参数传送。

**示例**

``` JavaScript
/*本示例指定在幻灯片放映过程中的任何时候，如果鼠标指针经过形状，将运行宏 CalculateTotal。*/
function test(){
　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).ActionSettings.Item(ppMouseOver)
　　　　Act.Action = ppActionRunMacro
　　　　Act.Run = "CalculateTotal"
　　　　Act.AnimateAction = true
}
```

### ActionSetting.ShowAndReturn

决定 WPP是否返回初始的幻灯片放映以及返回的条件。可读/写。MsoTriState 类型。

**语法**

**express.ShowAndReturn**

\*express是一个代表 ActionSetting 对象的变量。

**说明**

|                                                                                                                                        |
|----------------------------------------------------------------------------------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。                                                                                          |
| msoCTrue                                                                                                                               |
| msoFalse 默认值。WPP不会从未激活的自定义幻灯片放映返回初始幻灯片放映。                                                                 |
| msoTriStateMixed                                                                                                                       |
| msoTriStateToggle                                                                                                                      |
| msoTrue：WPP从取消激活的自定义幻灯片放映返回初始幻灯片放映，该自定义幻灯片放映通过初始演示文稿的 Hyperlink 或 ActionSetting 对象激活。 |

**示例**

``` JavaScript
/*以下示例设置当前演示文稿中第一张幻灯片上第五个形状的鼠标单击动作为：显示名为“techtalk”的自定义幻灯片放映。当该自定义幻灯片放映结束后，将自动返回演示文稿在鼠标单击前的最初状态。*/
function test(){
　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(5).ActionSettings.Item(ppMouseClick)
　　　　Act.Action = ppActionNamedSlideShow
　　　　Act.SlideShowName = "techtalk"
　　　　Act.ShowandReturn = msoTrue
}
```

### ActionSetting.SlideShowName

返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（ActionSetting 对象）；返回或设置自定义幻灯片放映的名称以便打印（PrintOptions 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（PublishObject 对象）。String 类型，可读/写。

**语法**

**express.SlideShowName**

\*express是一个代表 ActionSetting 对象的变量。

**说明**

必须将 **RangeType** 属性设置为 **ppPrintNamedSlideShow** 才能打印自定义幻灯片放映。

**示例**

``` JavaScript
/*本示例打印名为“tech talk”的现有自定义幻灯片放映。*/
function test(){
　　　　let Pri = Application.ActivePresentation.PrintOptions
　　　　Pri.RangeType = ppPrintNamedSlideShow
　　　　Pri.SlideShowName = "tech talk"
　　　　ActivePresentation.PrintOut()
}
```

``` JavaScript
/*下例将当前演示文稿保存为 HTML 4.0 文件格式，文件名为“mallard.htm”。然后显示一条消息，说明正将当前名称的演示文稿保存为WPP 和 HTML 两种格式。*/
function test(){
　　　　let Pub = ActivePresentation.PublishObjects.Item(1)
　　　　let PresName = Pub.SlideShowName
　　　　Pub.SourceType = ppPublishAll
　　　　Pub.FileName = "C:\\HTMLPres\\mallard.htm"
　　　　Pub.HTMLVersion = ppHTMLVersion4
　　　　MsgBox("Saving presentation " + "'" + PresName + "'" + " inWPP" + "\n" + "\r" + " format and HTML version 4.0 format")
　　　　Pub.Publish()
}
```

### ActionSetting.SoundEffect

返回一个 SoundEffect 对象，该对象代表切换到指定幻灯片时播放的声音。

**语法**

**express.SoundEffect**

\*express是一个代表 ActionSetting 对象的变量。

**示例**

``` JavaScript
/*以下示例设置在当前演示文稿第一张幻灯片的第一个形状演示动画时播放文件 Bass.wav。*/
function test(){
　　　　let Ani = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).AnimationSettings
　　　　Ani.Animate = true
　　　　Ani.TextLevelEffect = ppAnimateByAllLevels
　　　　Ani.SoundEffect.ImportFromFile("c:\\bass.wav")
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/ActionSetting](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AddIns 对象

## 简介

AddIn 对象的集合，这些对象代表与 WPP 相关的对WPP 可用的所有加载项，而无论这些加载项是否已加载。

## 说明

这里不包含 组件对象模型 (COM) 加载项 （COM 加载项：通过添加自定义命令和指定的功能来扩展 WPS Office程序的功能的补充程序。COM 加载项可在一个或多个 Office 程序中运行。COM 加载项使用文件扩展名 .dll 或 .exe。） 。

## 方法

| 序号 | 名称                     | 说明                                                                |
|:----:|:-------------------------|:--------------------------------------------------------------------|
|  1   | [Add](#AddIns.Add)       | 返回一个 AddIn 对象，该对象代表一个添加到加载宏列表中的加载宏文件。 |
|  2   | [Item](#AddIns.Item)     | 从指定集合中返回单个对象。                                          |
|  3   | [Remove](#AddIns.Remove) | 从加载宏集合中删除一个加载宏。                                      |

## 属性

| 序号 | 名称                               | 说明                                                    |
|:----:|:-----------------------------------|:--------------------------------------------------------|
|  1   | [Application](#AddIns.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#AddIns.Count)             | 返回指定集合中的对象数目。                              |
|  3   | [Parent](#AddIns.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### AddIns.Add

返回一个 AddIn 对象，该对象代表一个添加到加载宏列表中的加载宏文件。

**语法**

**express.Add ( *FileName* )**

\*express是一个代表 AddIns 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                       |
|------------|-----------|----------|----------------------------------------------------------------------------|
| *FileName* | 必选      | String   | 包含要添加到加载宏列表中的加载宏的文件的完整名称（包括路径和文件扩展名）。 |

**返回值**

AddIn

**说明**

此方法并不加载新的加载宏。必须设置 **Loaded** 属性才能加载该加载宏。

### AddIns.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 AddIns 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

AddIn

### AddIns.Remove

从加载宏集合中删除一个加载宏。

**语法**

**express.Remove ( *Index* )**

\*express是一个代表 AddIns 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                           |
|---------|-----------|----------|--------------------------------|
| *Index* | 必选      | Variant  | 要从集合中删除的加载宏的名称。 |

**示例**

本示例从可用加载宏的列表中删除名为“MyTools”的加载宏。

``` JavaScript
示例代码 
Appliaction.AddIns.Remove("mytools")
```

## 成员属性

### AddIns.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 AddIns 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

``` JavaScript
/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test(){
        for(let i=1;i <= Application.ActivePresentation.Slides.Item(1).Shapes.Count;i++) {
            if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {
             　　　alert(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)
          　}
        }
    }
```

### AddIns.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 AddIns 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除窗口1外的所有窗口。*/
function test(){
　　　　let Win = Application.Windows
       for(let i=2;i <= Win.Count;i++) {
　　　　    Win.Item(2).Close()
　　　　}
}
```

### AddIns.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 AddIns 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/AddIns](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Adjustments 对象

## 简介

包含指定自选图形、艺术字对象或连接符的调整值的集合。

## 说明

每个调整值代表了调整控点可以调整的方向。因为某些调整控点有两个调整方向（例如，某些调整控点可以在水平和垂直两个方向调整），所以形状的调整值数量可以大于调整控点的数量。一个形状最多可有八个调整值。

使用 Adjustments 属性返回 Adjustments 对象。使用 Adjustments ( *index* ) 返回单个调整值，其中 *index* 是该调整值的索引号。

不同的形状有不同数量的调整值，不同的调整值以不同的方式改变形状的几何外形，且不同的调整值有不同的有效值范围。例如，以下示例显示右箭头标注的四个调整值分别如何影响该标注几何外形的定义。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 因为每个可调整的形状都有一组不同的调整值，所以检验某一特定形状的调整行为的建议方法是：先手动创建该形状的一个实例，再打开宏记录器并调整该形状，然后查看所记录的代码。 |

下表概述了不同类型的调整所具有的有效的调整值范围。多数情况下，如果指定的值超过了有效值范围，将给调整分配最接近该值的有效值。

| 调整类型           | 有效值                                                                                                                                                                                                                                                                                             |
|--------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 线性（水平或垂直） | 通常 0.0 值代表形状的左边界或上边界，而 1.0 值代表形状的右边界或下边界。有效值对应于有效的手动调整。例如，如果只能将调整控点手动拖动形状的一半宽度，则相应的调整值最大为 0.5。对于象连接符和标注这样的形状，0.0 和 1.0 值代表由它们的起始和终止点定义的矩形界限，此时负值和大于 1.0 的值是有效的。 |
| 射线               | 调整值 1.0 对应于形状宽度。最大值为 0.5，或形状宽度的一半。                                                                                                                                                                                                                                        |
| 角                 | 值以度表示。如果指定的值超过了 -180 到 180 的范围，则将其折算为该范围内的值。                                                                                                                                                                                                                      |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                                                                                                         |
|:----:|:----------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Adjustments.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                                                      |
|  2   | [Count](#Adjustments.Count)             | 返回指定集合中的对象数目。                                                                                                                                                                                                                                                                                   |
|  3   | [Creator](#Adjustments.Creator)         | 返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。                                                                                                                                                |
|  4   | [Item](#Adjustments.Item)               | 返回或设置由 Index 参数指定的调整值。对于线性调整，0.0 调整值通常对应于形状的左边缘或上边缘，1.0 调整值通常对应于形状的右边缘或下边缘。但对某些形状，调整值可以超过形状边界。对于径向调整，1.0 调整值对应于形状宽度。对于角度调整，调整值用度数指定。Item 属性仅应用于含有调整的形状。可读/写。Single 类型。 |
|  5   | [Parent](#Adjustments.Parent)           | 返回指定对象的父对象。                                                                                                                                                                                                                                                                                       |

## 成员属性

### Adjustments.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Adjustments 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

``` JavaScript
/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test(){
    for(let i=1;i <= Application.ActivePresentation.Slides.Item(1).Shapes.Count;i++) {
        if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {
         　　　alert(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)
      　}
    }
}
```

### Adjustments.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 Adjustments 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭第2个窗口。*/
function test(){
　　let Win = Application.Windows
    for(let i=2;i <= Win.Count;i++) {
　　　　 Win.Item(2).Close()
　　}
}
```

### Adjustments.Creator

返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 Adjustments 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

``` JavaScript
/*以下示例显示关于 myObject 创建者的消息。*/
function test(){
　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
　　　　if(myObject.Creator == parseInt(50575054,16)) {
　　　　    alert("This is a WPP object")
　　　　}
　　　　else {
　　　　    alert("This is not a WPP object")
　　　　}}
```

### Adjustments.Item

返回或设置由 Index 参数指定的调整值。对于线性调整，0.0 调整值通常对应于形状的左边缘或上边缘，1.0 调整值通常对应于形状的右边缘或下边缘。但对某些形状，调整值可以超过形状边界。对于径向调整，1.0 调整值对应于形状宽度。对于角度调整，调整值用度数指定。Item 属性仅应用于含有调整的形状。可读/写。Single 类型。

**语法**

**express.Item**

\*express是一个代表 Adjustments 对象的变量。

**说明**

自选图形、连接符和艺术字对象最多可有八个调整。

**示例**

``` JavaScript
/*以下示例将两个十字形添加到 myDocument。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let myShape = myDocument.Shapes
　　　　myShape.AddShape(msoShapeCross, 10, 10, 100, 100).Adjustments.Item(1) 
　　　　myShape.AddShape(msoShapeCross, 150, 10, 100, 100).Adjustments.Item(1)
}
```

### Adjustments.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Adjustments 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Adjustments](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Application 对象

## 简介

代表整个 WPP 应用程序。

## 说明

Application 对象包括：

- 应用程序范围内的设置和选项（例如，当前打印机的名称）。
- 用于返回顶层对象的属性，例如 ActivePresentation 和 Windows 。

在WPS宏编辑器中编写要在WPP 中运行的代码时， Application 对象的下列属性可以在没有对象限定符的情况下使用： ActivePresentation 、 ActiveWindow 、 AddIns 、 Presentations 、 SlideShowWindows 和 Windows 。例如，可以编写 ActiveWindow.Height = 200 来代替 Application.ActiveWindow.Height = 200 。（仅WPS宏编辑器）

## 方法

| 序号 | 名称                              | 说明                                                |
|:----:|:----------------------------------|:----------------------------------------------------|
|  1   | [Activate](#Application.Activate) | 激活指定的对象。                                    |
|  2   | [Quit](#Application.Quit)         | 退出 WPP。该方法相当于单击“文件”菜单上的“退出WPP”。 |
|  3   | [Run](#Application.Run)           | 运行JS宏。                                          |

## 属性

| 序号 | 名称                                                                            | 说明                                                                                                                                                                                            |
|:----:|:--------------------------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Active](#Application.Active)                                                   | 返回指定的窗格或窗口是否处于激活的状态。只读。 MsoTriState 类型。                                                                                                                               |
|  2   | [ActiveEncryptionSession](#Application.ActiveEncryptionSession)                 | 此对象或成员加入对象模型中的时间太晚，以至于还没有完整的文档说明。 只读。                                                                                                                       |
|  3   | [ActivePresentation](#Application.ActivePresentation)                           | 返回一个 Presentation 对象，该对象代表当前窗口中打开的演示文稿。只读。请注意，如果嵌入的演示文稿处于活动状态，则 ActivePresentation 属性将返回嵌入的演示文稿。                                  |
|  4   | [ActivePrinter](#Application.ActivePrinter)                                     | 返回当前打印机的名称。只读。 String 类型。                                                                                                                                                      |
|  5   | [ActiveWindow](#Application.ActiveWindow)                                       | 返回一个代表活动文档窗口的 DocumentWindow 对象。只读。                                                                                                                                          |
|  6   | [Assistance](#Application.Assistance)                                           | 获取 WPS Office IAssistance 对象的引用，该对象为开发人员提供了一种可以为 WPS Office 用户创建自定义帮助体验的方法。只读。                                                                        |
|  7   | [AutoCorrect](#Application.AutoCorrect)                                         | 返回一个代表 WPP 中的“自动更正”功能的 AutoCorrect 对象。                                                                                                                                        |
|  8   | [Build](#Application.Build)                                                     | 返回WPP 内部版本号。只读 String 类型。                                                                                                                                                          |
|  9   | [Caption](#Application.Caption)                                                 | 返回在应用程序窗口的标题栏中显示的文本。 String 类型，可读/写。                                                                                                                                 |
|  10  | [COMAddIns](#Application.COMAddIns)                                             | 返回对 WPP 中当前加载的组件对象模型 (COM) 加载宏的引用。在 “WPP选项”对话框的 “加载宏” 选项卡中列出了这些加载宏。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。                   |
|  11  | [CommandBars](#Application.CommandBars)                                         | 返回一个代表WPP 中所有命令栏的 CommandBars 集合。只读。                                                                                                                                         |
|  12  | [Creator](#Application.Creator)                                                 | 返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。                                   |
|  13  | [DefaultWebOptions](#Application.DefaultWebOptions)                             | 返回一个 DefaultWebOptions 对象，该对象包含将部分或全部演示文稿作为网页发布或者打开一个网页时 WPP 使用的全局应用程序级属性。只读。                                                              |
|  14  | [DisplayAlerts](#Application.DisplayAlerts)                                     | 设置或返回 PpAlertLevel 常量，该常量代表 WPP 在运行宏时是否显示警告。可读/写。                                                                                                                  |
|  15  | [DisplayDocumentInformationPanel](#Application.DisplayDocumentInformationPanel) | 返回或设置是否在 WPP 用户界面中显示“文档属性”面板。可读/写。                                                                                                                                    |
|  16  | [DisplayGridLines](#Application.DisplayGridLines)                               | MsoTrue 属性值用于显示 WPP 中的网格线。可读/写 MsoTriState 类型。                                                                                                                               |
|  17  | [FeatureInstall](#Application.FeatureInstall)                                   | 返回或设置 WPP 如何处理对所需功能尚未安装的方法和属性的调用。可读/写 MsoFeatureInstall 类型。                                                                                                   |
|  18  | [FileDialog](#Application.FileDialog)                                           | 返回一个 FileDialog 对象，该对象表示文件对话框的单个实例。                                                                                                                                      |
|  19  | [Height](#Application.Height)                                                   | 以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读， Single 类型；用于所有其他对象时可读/写， Single 类型。                                                                             |
|  20  | [LanguageSettings](#Application.LanguageSettings)                               | 返回一个 LanguageSettings 对象，该对象包含 WPP 中有关语言设置的信息。只读。                                                                                                                     |
|  21  | [Left](#Application.Left)                                                       | 返回或设置一个 Single 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。 |
|  22  | [Name](#Application.Name)                                                       | 返回应用程序的名称字符串。 String 类型，只读 。                                                                                                                                                 |
|  23  | [NewPresentation](#Application.NewPresentation)                                 | 返回一个 NewFile 对象，该对象代表“新建演示文稿” 任务窗格中列出的演示文稿。只读。                                                                                                                |
|  24  | [OperatingSystem](#Application.OperatingSystem)                                 | 返回操作系统名称。只读 String 类型。                                                                                                                                                            |
|  25  | [Options](#Application.Options)                                                 | 返回一个 Options 对象，该对象代表 WPP 中的应用程序选项。                                                                                                                                        |
|  26  | [Path](#Application.Path)                                                       | 返回 String ，该值代表到指定的 Application 对象的路径， 只读。                                                                                                                                  |
|  27  | [Presentations](#Application.Presentations)                                     | 返回一个 Presentations 集合，该集合代表所有打开的演示文稿。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象 Presentation 。                                                          |
|  28  | [ProductCode](#Application.ProductCode)                                         | 返回 WPP 的全局唯一标识符 (GUID)。用户可能会用到 GUID，例如，当使用程序调用应用程序编程接口 (API) 时。只读 String 类型。                                                                        |
|  29  | [ShowStartupDialog](#Application.ShowStartupDialog)                             | 属性值为 MsoTrue 时，启动 WPP 后显示 “新建演示文稿”任务窗格。可读/写。                                                                                                                          |
|  30  | [ShowWindowsInTaskbar](#Application.ShowWindowsInTaskbar)                       | 决定是否每个打开的演示文稿都有单独的 Windows 任务栏按钮。可读/写 MsoTriState 类型。                                                                                                             |
|  31  | [SlideShowWindows](#Application.SlideShowWindows)                               | 返回一个 SlideShowWindows 集合，该集合代表所有打开的幻灯片放映窗口。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象 SlideShowWindow 。                                              |
|  32  | [Top](#Application.Top)                                                         | 返回或设置一个 Single 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。   |
|  33  | [Version](#Application.Version)                                                 | 返回WPP 版本号。只读 String 类型。                                                                                                                                                              |
|  34  | [Visible](#Application.Visible)                                                 | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 MsoTriState 类型。                                                                                                                  |
|  35  | [Width](#Application.Width)                                                     | 以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读， Single 类型；用于所有其他对象时可读/写， Single 类型。                                                                             |
|  36  | [Windows](#Application.Windows)                                                 | 返回一个代表所有打开的文档窗口的 DocumentWindow 集合。只读。                                                                                                                                    |
|  37  | [WindowState](#Application.WindowState)                                         | 返回或设置指定窗口的状态。可读/写。 PpWindowState 类型。                                                                                                                                        |

## 成员方法

### Application.Activate

激活指定的对象。

**语法**

**express.Activate ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例会激活WPP程序当前窗口*/
Application.Activate()
```

### Application.Quit

退出 WPP。该方法相当于单击“文件”菜单上的“退出WPP”。

**语法**

**express.Quit ()**

\*express是一个代表 Application 对象的变量。

**说明**

若要避免提示保存更改，可在调用 **Quit** 方法之前使用 **Save** 或 **SaveAs** 方法保存所有打开的演示文稿。

**示例**

``` JavaScript
/*本示例保存所有打开的演示文稿，然后退出WPP。*/
function test()
{
    let app = Application
    for(let i=1; i <= app.Presentations.Count; i++) 
    {
        app.Presentations.Item(i).Save()
    }
    app.Quit()
}
```

### Application.Run

运行JS宏。

**语法**

**express.Run ( *MacroName* , *safeArrayOfParams* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                     |
|---------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *MacroName*         | 必选      | String   | 要运行的宏的名称。该字符串可包含下列内容：后跟感叹号 (!) 的加载演示文稿或加载宏文件名、后跟句点 (.) 的有效模块名称以及宏名称。例如，“MyPres.ppt!Module1.Test”是一个有效的 MacroName 值。 |
| *safeArrayOfParams* | 必选      | Variant  | 要传递给该过程的参数。不能为此参数指定一个对象，也不能对此方法使用命名参数。参数必须按位置传递。                                                                                         |

**说明**

宏可能包含病毒，因此在运行宏时要格外小心。请采用下列预防措施：在计算机上运行最新的防病毒软件；将宏安全级别设置为“高”；清除“信任所有安装的加载项和模板” 复选框；使用数字签名；维护受信任的发布者的列表。

**示例**

``` JavaScript
/*在本示例中，Main 过程定义一个数组，然后运行宏 TestPass，将该数组作为参数传递。*/
function Main() {
    let x = []
    x[1] = "hi"
    x[2] = 7
    Application.Run("TestPass", x)
}

function TestPass(x) {
    alert(x[1])
    alert(x[2])
}
```

## 成员属性

### Application.Active

返回指定的窗格或窗口是否处于激活的状态。只读。 **MsoTriState** 类型。

**语法**

**express.Active**

\*express是一个代表 Application 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** ：指定的窗格或窗口处于活动状态。  |

**示例**

``` JavaScript
//以下示例检查演示文稿文件 "test.ppt" 是否位于当前窗口。如果不是，则将当前处于激活状态的演示文稿的名称保存在变量 oldWin 中，并激活 "test.ppt" 演示文稿。
function test(){
    let window = Application.Presentations.Item("test.ppt").Windows.Item(1)
    if(!window.Active) {
        let oldWin = Application.ActiveWindow
        window.Activate()
    }
}
```

### Application.ActiveEncryptionSession

此对象或成员加入对象模型中的时间太晚，以至于还没有完整的文档说明。 只读。

**语法**

**express.ActiveEncryptionSession**

\*express是一个代表 Application 对象的变量。

### Application.ActivePresentation

返回一个 Presentation 对象，该对象代表当前窗口中打开的演示文稿。只读。请注意，如果嵌入的演示文稿处于活动状态，则 **ActivePresentation** 属性将返回嵌入的演示文稿。

**语法**

**express.ActivePresentation**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将加载的演示文稿保存到D盘根目录下名为“TestFile”的文件中。*/
function test()
{
    let MyPath = "D:\TestFile"
    Application.ActivePresentation.SaveAs(MyPath)
}
```

### Application.ActivePrinter

返回当前打印机的名称。只读。 **String** 类型。

**语法**

**express.ActivePrinter**

\*express是一个代表 Application 对象的变量。

**说明**

**示例**

``` JavaScript
//以下示例显示当前打印机的名称。
alert("The name of the active printer is " + Application.ActivePrinter)
```

### Application.ActiveWindow

返回一个代表活动文档窗口的 **DocumentWindow** 对象。只读。

**语法**

**express.ActiveWindow**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口最小化。*/
Application.ActiveWindow.WindowState = ppWindowMinimized
```

### Application.Assistance

获取 WPS Office IAssistance 对象的引用，该对象为开发人员提供了一种可以为 WPS Office 用户创建自定义帮助体验的方法。只读。

**语法**

**express.Assistance**

\*express是一个代表 Application 对象的变量。

### Application.AutoCorrect

返回一个代表 WPP 中的“自动更正”功能的 **AutoCorrect** 对象。

**语法**

**express.AutoCorrect**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
//以下示例禁止显示“自动更正选项”和“自动版式选项”按钮。
function HideAutoCorrectOpButton() {
    let ac = Application.AutoCorrect
    ac.DisplayAutoCorrectOptions = msoFalse
    ac.DisplayAutoLayoutOptions = msoFalse
}
```

### Application.Build

返回WPP 内部版本号。只读 **String** 类型。

**语法**

**express.Build**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例显示WPP 内部版本号。*/
alert(Application.Build)
```

### Application.Caption

返回在应用程序窗口的标题栏中显示的文本。 **String** 类型，可读/写。

**语法**

**express.Caption**

\*express是一个代表 Application 对象的变量。

### Application.COMAddIns

返回对 WPP 中当前加载的组件对象模型 (COM) 加载宏的引用。在 “WPP选项”对话框的 **“加载宏”** 选项卡中列出了这些加载宏。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**语法**

**express.COMAddIns**

\*express是一个代表 Application 对象的变量。

### Application.CommandBars

返回一个代表WPP 中所有命令栏的 **CommandBars** 集合。只读。

**语法**

**express.CommandBars**

\*express是一个代表 Application 对象的变量。

### Application.Creator

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 Application 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

``` JavaScript
//以下示例显示关于 myObject 创建者的消息。
function test(){
    let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
    if(myObject.Creator == parseInt(50575054,16)) {
       alert("This is a WPP object")
    }
    else {
    alert("This is not a WPP object")
    }
}
```

### Application.DefaultWebOptions

返回一个 **DefaultWebOptions** 对象，该对象包含将部分或全部演示文稿作为网页发布或者打开一个网页时 WPP 使用的全局应用程序级属性。只读。

**语法**

**express.DefaultWebOptions**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
//本示例检查默认的文档编码是否为 Western。如果是 Western，则字符串 strDocEncoding 将根据它进行设置。
function test(){
    let objAppWebOptions = Application.DefaultWebOptions
    if(objAppWebOptions.Encoding == msoEncodingWestern) {
        let strDocEncoding = "Western"
    }
}
```

### Application.DisplayAlerts

设置或返回 **PpAlertLevel** 常量，该常量代表 WPP 在运行宏时是否显示警告。可读/写。

**语法**

**express.DisplayAlerts**

\*express是一个代表 Application 对象的变量。

**说明**

宏停止运行后， **DisplayAlerts** 属性的值不会重新设置；该值将在会话过程中保持不变。在会话过程之间不会保存该值，所以在WPP 开始运行时，该值重新设置为 **ppAlertsNone** 。

|                                                                                                     |
|-----------------------------------------------------------------------------------------------------|
| **PpAlertLevel** 可以是下列 **PpAlertLevel** 常量之一。                                             |
| **ppAlertsAll** 显示所有的消息框和警告；将错误返回宏。                                              |
| **ppAlertsNone** 默认值。不显示任何警告或消息框。如果运行宏时遇到消息框，则选中默认值并继续运行宏。 |

**示例**

``` JavaScript
//下面的代码行指示WPP 显示所有的消息框和警告，并将错误返回宏。
Application.DisplayAlerts = ppAlertsAll
```

### Application.DisplayDocumentInformationPanel

返回或设置是否在 WPP 用户界面中显示“文档属性”面板。可读/写。

**语法**

**express.DisplayDocumentInformationPanel**

\*express是一个代表 Application 对象的变量。

### Application.DisplayGridLines

**MsoTrue** 属性值用于显示 WPP 中的网格线。可读/写 **MsoTriState** 类型。

**语法**

**express.DisplayGridLines**

\*express是一个代表 Application 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**                                   |

**示例**

``` JavaScript
//本示例切换WPP 中网格线的显示状态。
function ToggleGridLines() {
    let a = Application
    if(a.DisplayGridLines == msoTrue) {
        a.DisplayGridLines = msoFalse
    }
    else {
        a.DisplayGridLines = msoTrue
    }
}
```

### Application.FeatureInstall

返回或设置 WPP 如何处理对所需功能尚未安装的方法和属性的调用。可读/写 **MsoFeatureInstall** 类型。

**语法**

**express.FeatureInstall**

\*express是一个代表 Application 对象的变量。

**说明**

可以使用 **msoFeatureInstallOnDemandWithUI** 常量防止用户误以为在安装某功能时系统没有反应。使用带有错误捕获例程的 **msoFeatureInstallNone** 常量可防止安装最终用户功能。

| 注释                                                                                                                                                                                                                                                                                                                   |
|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 如果在字符串中引用未安装的演示文稿设计模板，则将产生运行时错误。无论 **FeatureInstall** 属性设置是什么，该模板都不会自动安装。要对当前未安装的模板使用 **ApplyTemplate** 方法，必须首先安装附加设计模板。为此，请通过运行 WPS Office安装程序（单击 Windows 控制面板中的 “添加/删除程序”图标）安装WPP 的附加设计模板。  |

|                                                                                    |
|------------------------------------------------------------------------------------|
| MsoFeatureInstall 可以是下列 MsoFeatureInstall 常量之一。                          |
| **msoFeatureInstallNone** 默认值。调用未安装的功能时会产生可捕获的运行时自动错误。 |
| **msoFeatureInstallOnDemand** 显示对话框提示用户安装新功能。                       |
| **msoFeatureInstallOnDemandWithUI** 安装时显示进度条。不提示用户安装新功能。       |

**示例**

``` JavaScript
//以下示例检查 FeatureInstall 属性的值。如果该属性设置为 msoFeatureInstallNone，则代码显示一个消息框，以询问用户是否要更改属性设置。如果用户回答“是”，则此属性将设置为 msoFeatureInstallOnDemand。
function test(){
    let a = Application
    if(a.FeatureInstall == msoFeatureInstallNone) {
        let Reply = alert("Uninstalled features for this application " + "\r\n"
        + "may cause a run-time error when called." + "\r\n" + "\r\n" + "Would you like to change this setting" 
        + "\r\n"  + "to automatically install missing features when called?" , 52, "Feature Install Setting")
        if(Reply == 6) {
            a.FeatureInstall = msoFeatureInstallOnDemand
        }
    }
}
```

### Application.FileDialog

返回一个 **FileDialog** 对象，该对象表示文件对话框的单个实例。

**语法**

**express.FileDialog**

\*express是一个代表 Application 对象的变量。

**说明**

|                                                                   |
|-------------------------------------------------------------------|
| **MsoFileDialogType** 可以是下列 **MsoFileDialogType** 常量之一。 |
| **msoFileDialogFilePicker**                                       |
| **msoFileDialogFolderPicker**                                     |
| **msoFileDialogOpen**                                             |
| **msoFileDialogSaveAs**                                           |

**示例**

``` JavaScript
//以下示例显示“另存为”对话框。
function ShowSaveAsDialog() {
    let dlgSaveAs = Application.FileDialog( msoFileDialogSaveAs)
    dlgSaveAs.Show()
}
```

### Application.Height

以磅为单位返回或设置指定对象的高度。用于 **Master** 对象时只读， **Single** 类型；用于所有其他对象时可读/写， **Single** 类型。

**语法**

**express.Height**

\*express是一个代表 Application 对象的变量。

**说明**

一个 **Height** Shape 对象的 **Shape** Height 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

### Application.LanguageSettings

返回一个 **LanguageSettings** 对象，该对象包含 WPP 中有关语言设置的信息。只读。

**语法**

**express.LanguageSettings**

\*express是一个代表 Application 对象的变量。

### Application.Left

返回或设置一个 **Single** 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Left**

\*express是一个代表 Application 对象的变量。

### Application.Name

返回应用程序的名称字符串。 **String** 类型，只读 。

**语法**

**express.Name**

\*express是一个代表 Application 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则对形状集合对象Shapes调用 Item(" Rectangle 2 ") 将返回对该形状的引用。详情见示例。

**示例**

``` JavaScript
function test()
{
    let name = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Name;
    let shape = Application.ActivePresentation.Slides.Item(1).Shapes.Item(name)
    alert(shape.Name)
}
```

### Application.NewPresentation

返回一个 **NewFile** 对象，该对象代表“新建演示文稿” 任务窗格中列出的演示文稿。只读。

**语法**

**express.NewPresentation**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
//以下示例会在窗格中最后部分的底部列出“新建演示文稿”任务窗格上的演示文稿。 
function CreateNewPresentationListItem() {
    Application.NewPresentation.Add("C:\\Presentation.ppt")
    Application.CommandBars.Item(1).Visible = true
}
```

### Application.OperatingSystem

返回操作系统名称。只读 **String** 类型。

**语法**

**express.OperatingSystem**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
//本示例检测 OperatingSystem 属性以确定WPP 是否运行在 32 位版本的 Windows 中。
function test(){
    let os = Application.OperatingSystem
    if((os.indexOf("Windows (32-bit)") != -1) != 0) {
        alert("Running a 32-bit version of  Windows")
    }
}
```

### Application.Options

返回一个 **Options** 对象，该对象代表 WPP 中的应用程序选项。

**语法**

**express.Options**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
//使用 Options 属性可返回 Options 对象。以下示例设置WPP 的三个应用程序选项。
function TogglePasteOptionsButton() {
    let options = Application.Options
    if(options.DisplayPasteOptions == false) {
        options.DisplayPasteOptions = true
    }
}
```

### Application.Path

返回 **String** ，该值代表到指定的 **Application** 对象的路径， 只读。

**语法**

**express.Path**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例显示WPP程序的路径。*/
alert(Application.Path)
```

### Application.Presentations

返回一个 **Presentations** 集合，该集合代表所有打开的演示文稿。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象 **Presentation** 。

**语法**

**express.Presentations**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例打开名为“TEST”的演示文稿。*/
Application.Presentations.Open("C:\\Users\\wps\\Desktop\\TEST.pptm")

/*本示例将第一个演示文稿以“Year-End Report.ppt”名称保存在D盘。*/
Application.Presentations.Item(1).SaveAs("D:/Year-End Report.pptx")

/*本示例关闭演示文稿“Year-end report”。*/
Application.Presentations.Item("Year-End Report.ppt").Close()
```

### Application.ProductCode

返回 WPP 的全局唯一标识符 (GUID)。用户可能会用到 GUID，例如，当使用程序调用应用程序编程接口 (API) 时。只读 **String** 类型。

**语法**

**express.ProductCode**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将显示GUID。*/
alert(Application.ProductCode)
```

### Application.ShowStartupDialog

属性值为 **MsoTrue** 时，启动 WPP 后显示 “新建演示文稿”任务窗格。可读/写。

**语法**

**express.ShowStartupDialog**

\*express是一个代表 Application 对象的变量。

**说明**

|                                                       |
|-------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue** 不适用于此属性。                         |
| **msoFalse** 隐藏 “新建演示文稿”侧窗格。              |
| **msoTriStateMixed** 不适用于此属性。                 |
| **msoTriStateToggle** 不适用于此属性。                |
| **msoTrue** 默认值。显示 “新建演示文稿”侧窗格。       |

**示例**

``` JavaScript
//下面的代码行在WPP 启动时禁用“新建演示文稿”任务窗格。
Application.ShowStartupDialog = msoFalse
```

### Application.ShowWindowsInTaskbar

决定是否每个打开的演示文稿都有单独的 Windows 任务栏按钮。可读/写 **MsoTriState** 类型。

**语法**

**express.ShowWindowsInTaskbar**

\*express是一个代表 Application 对象的变量。

**说明**

设置为 **True** 时，此属性将模拟单文档界面 (SDI) 的外观，这样便于在打开的演示文稿之间切换。但是，如果在使用多个演示文稿的同时还打开了其他应用程序，则可能希望将此属性设置为 **False** ，以避免任务栏上布满多余的按钮。

仅当 WPS Office与 Windows Update 或 Windows 系统一同使用时此属性可用。

|                                                                       |
|-----------------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。                         |
| **msoCTrue**                                                          |
| **msoFalse**                                                          |
| **msoTriStateMixed**                                                  |
| **msoTriStateToggle**                                                 |
| **msoTrue** 默认值。每个打开的演示文稿都有单独的 Windows 任务栏按钮。 |

**示例**

``` JavaScript
//本示例指定每个打开的演示文稿都不具有单独的 Windows 任务栏按钮。
Application.ShowWindowsInTaskbar = msoFalse
```

### Application.SlideShowWindows

返回一个 **SlideShowWindows** 集合，该集合代表所有打开的幻灯片放映窗口。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象 **SlideShowWindow** 。

**语法**

**express.SlideShowWindows**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例在一个窗口中运行幻灯片放映，并设置该幻灯片放映窗口的高度和宽度。*/
function test()
{
    let a = Application
    a.Presentations.Item(1).SlideShowSettings.Run()
    let windows = a.SlideShowWindows.Item(1)
    windows.Height = 250
    windows.Width = 250
}
```

### Application.Top

返回或设置一个 **Single** 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Top**

\*express是一个代表 Application 对象的变量。

### Application.Version

返回WPP 版本号。只读 **String** 类型。

**语法**

**express.Version**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例在一个消息框中显示WPP 版本号和内部版本号以及操作系统名称。*/
function test()
{
    let a = Application
    alert("Welcome to WPP version " + a.Version + ", build " + a.Build + ", running on " + a.OperatingSystem + "!")
}
```

### Application.Visible

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 **MsoTriState** 类型。

**语法**

**express.Visible**

\*express是一个代表 Application 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** ：指定对象或对象格式是可见的。    |

**示例**

``` JavaScript
//以下示例为当前演示文稿中第一张幻灯片的第三个形状设置阴影的水平偏移量和垂直偏移量，使其向右偏移 5 磅并向上偏移 3 磅。如果该形状原来没有阴影，本例将会为其添上阴影。
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let shadow = myDocument.Shapes.Item(3).Shadow
    shadow.Visible = msoTrue
    shadow.OffsetX = 5
    shadow.OffsetY = -3
}
```

### Application.Width

以磅为单位返回或设置指定对象的宽度。用于 **Master** 对象时只读， **Single** 类型；用于所有其他对象时可读/写， **Single** 类型。

**语法**

**express.Width**

\*express是一个代表 Application 对象的变量。

### Application.Windows

返回一个代表所有打开的文档窗口的 DocumentWindow 集合。只读。

**语法**

**express.Windows**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口。*/
function test()
{
    let aw = Application.Windows
    for(let i = aw.Count; i >= 2; i--) 
    {
        aw.Item(i).Close()
    }
}
```

### Application.WindowState

返回或设置指定窗口的状态。可读/写。 **PpWindowState** 类型。

**语法**

**express.WindowState**

\*express是一个代表 Application 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| PpWindowState 可以是下列 PpWindowState 常量之一。 |
| **ppWindowMaximized**                             |
| **ppWindowMinimized**                             |
| **ppWindowNormal**                                |

当窗口状态为 **ppWindowNormal** 时，该窗口既未最大化也未最小化。

**示例**

``` JavaScript
/*以下示例将当前窗口最小化。*/
Application.ActiveWindow.WindowState = ppWindowMinimized
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FillFormat 对象

## 简介

代表形状的填充格式。形状可以有单色、渐变、纹理、图案、图片或半透明填充。

## 说明

FillFormat 对象的许多属性都是只读的。若要设置其中一个属性，必须应用相应的方法。

## 方法

| 序号 | 名称                                             | 说明                                                                                                     |
|:----:|:-------------------------------------------------|:---------------------------------------------------------------------------------------------------------|
|  1   | [Background](#FillFormat.Background)             | 指定形状的填充与幻灯片背景相匹配。如果将此方法应用于填充后又更改了幻灯片背景色，该填充会随之改变。       |
|  2   | [OneColorGradient](#FillFormat.OneColorGradient) | 将指定填充设置为单色渐变。                                                                               |
|  3   | [Patterned](#FillFormat.Patterned)               | 将指定填充设置为一种图案。                                                                               |
|  4   | [PresetGradient](#FillFormat.PresetGradient)     | 将指定填充设为一个预设的渐变。                                                                           |
|  5   | [PresetTextured](#FillFormat.PresetTextured)     | 将指定的填充方式设置为预设纹理。                                                                         |
|  6   | [Solid](#FillFormat.Solid)                       | 将指定的填充格式设置为均匀的颜色。可用本方法将带有渐变、纹理、图案或背景的填充格式转换为单色的填充格式。 |
|  7   | [TwoColorGradient](#FillFormat.TwoColorGradient) | 设置指定填充为双色渐变。                                                                                 |
|  8   | [UserPicture](#FillFormat.UserPicture)           | 用一幅大图像来填充指定形状。如果要通过平铺小图像来填充形状，请使用 UserTextured 方法。                   |
|  9   | [UserTextured](#FillFormat.UserTextured)         | 用平铺的小图像填充指定形状。如果要用一个大图像来填充该形状，请使用 UserPicture 方法。                    |

## 属性

| 序号 | 名称                                                 | 说明                                                                                                                                                                                                                                                                                                                                                                               |
|:----:|:-----------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FillFormat.Application)               | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                                                                                                                            |
|  2   | [BackColor](#FillFormat.BackColor)                   | 返回或设置 ColorFormat 对象，该对象代表指定填充对象或图案线条的背景色。可读/写。                                                                                                                                                                                                                                                                                                   |
|  3   | [Creator](#FillFormat.Creator)                       | 返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。                                                                                                                                                                                                                      |
|  4   | [ForeColor](#FillFormat.ForeColor)                   | 返回或设置 ColorFormat 对象，该对象表示填充、线条或阴影的前景色。可读/写。                                                                                                                                                                                                                                                                                                         |
|  5   | [GradientColorType](#FillFormat.GradientColorType)   | 返回指定填充的渐变颜色类型。此属性为只读。使用 OneColorGradient 、 PresetGradient 或 TwoColorGradient 方法设置填充的渐变类型。只读 MsoGradientColorType 类型。                                                                                                                                                                                                                     |
|  6   | [GradientDegree](#FillFormat.GradientDegree)         | 返回一个值，该值指示单色渐变填充的深浅程度。该值为 0（零）表示形状的前景色中混入黑色形成渐变；值为 1 表示混入白色；介于 0 和 1 之间的值表示混入前景色的偏深或偏浅的底纹。只读 Single 类型。此属性为只读。使用 OneColorGradient 方法设置填充的渐变程度。                                                                                                                            |
|  7   | [GradientStyle](#FillFormat.GradientStyle)           | 返回指定填充的渐变样式。使用 OneColorGradient 、 PresetGradient 或 TwoColorGradient 方法设置填充的渐变样式。尝试对未使用渐变的填充返回此属性将产生错误。使用 Type 属性决定填充是否使用渐变。只读 MsoGradientStyle 类型。                                                                                                                                                           |
|  8   | [GradientVariant](#FillFormat.GradientVariant)       | 返回指定填充的渐变变量，对于大多数渐变填充，其返回值为从 1 到 4 的整数。如果渐变样式为 msoGradientFromTitle 或 msoGradientFromCenter ，则该属性返回 1 或 2。该属性的值与 “形状填充” 选项卡中的 “渐变” 子选项卡上的渐变变量（从左向右，从上向下编号）相对应。 Long 类型，只读。此属性为只读。使用 OneColorGradient 、 PresetGradient 或 TwoColorGradient 方法可设置填充的渐变变量。 |
|  9   | [Parent](#FillFormat.Parent)                         | 返回指定对象的父对象。                                                                                                                                                                                                                                                                                                                                                             |
|  10  | [Pattern](#FillFormat.Pattern)                       | 设置或返回一个值，该值代表应用于指定填充的图案。使用 BackColor 和 ForeColor 属性可设置图案中使用的颜色。只读。 MsoPatternType 类型。                                                                                                                                                                                                                                               |
|  11  | [PresetGradientType](#FillFormat.PresetGradientType) | 返回指定填充的预设渐变类型。只读 MsoPresetGradientType 类型。使用 PresetGradient 方法可设置填充的预设渐变类型。                                                                                                                                                                                                                                                                    |
|  12  | [PresetTexture](#FillFormat.PresetTexture)           | 返回指定填充的预设纹理。只读 MsoPresetTexture 类型。                                                                                                                                                                                                                                                                                                                               |
|  13  | [TextureName](#FillFormat.TextureName)               | 返回指定填充的自定义纹理文件名称。只读 String 类型。                                                                                                                                                                                                                                                                                                                               |
|  14  | [TextureType](#FillFormat.TextureType)               | 返回指定填充的纹理类型。 MsoTextureType 类型，只读。使用 PresetTextured 或 UserTextured 方法可设置填充的纹理类型。                                                                                                                                                                                                                                                                 |
|  15  | [Transparency](#FillFormat.Transparency)             | 该属性返回或设置指定填充、阴影或线条的透明度，可取 0.0（不透明）至 1.0（透明）之间的数值。 Single 类型，可读/写。                                                                                                                                                                                                                                                                  |
|  16  | [Type](#FillFormat.Type)                             | 返回一个 MsoFillType 常量，该常量代表填充的类型。只读。                                                                                                                                                                                                                                                                                                                            |
|  17  | [Visible](#FillFormat.Visible)                       | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 MsoTriState 类型。                                                                                                                                                                                                                                                                                                     |

## 成员方法

### FillFormat.Background

指定形状的填充与幻灯片背景相匹配。如果将此方法应用于填充后又更改了幻灯片背景色，该填充会随之改变。

**语法**

**express.Background ()**

\*express是一个代表 FillFormat 对象的变量。

**说明**

请注意，将 **Background** 方法应用于形状的填充与设置该形状的透明填充不完全相同，也不完全等同于将形状填充设为与背景相同。请参见第二个示例。

**示例**

本示例将当前演示文稿中第一张幻灯片第一个形状的填充设为与幻灯片背景相匹配。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.Background()
```

本示例将当前演示文稿第一张幻灯片的背景设为一个预设的渐变填充，在幻灯片中添加一个矩形，然后在该矩形前面放置三个椭圆。第一个椭圆的填充与幻灯片背景相匹配，第二个的填充为透明，而第三个使用的填充方式与背景相同。请注意三个椭圆外观的不同。

``` JavaScript
function test(){
　　　　let aSlides = Application.ActivePresentation.Slides.Item(1)
　　　　aSlides.FollowMasterBackground = false
　　　　aSlides.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientDaybreak)
　　　　let dBackground = aSlides.Shapes
　　　　dBackground.AddShape(msoShapeRectangle, 50, 200, 600, 100)
　　　　dBackground.AddShape(msoShapeOval, 75, 150, 150, 100).Fill.Background()
　　　　dBackground.AddShape(msoShapeOval, 275, 150, 150, 100).Fill.Transparency = 1
　　　　dBackground.AddShape(msoShapeOval, 475, 150, 150, 100).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientDaybreak)
}
```

### FillFormat.OneColorGradient

将指定填充设置为单色渐变。

**语法**

**express.OneColorGradient ( *Style* , *Variant* , *Degree* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型         | 说明                                                                                                                                                                    |
|-----------|-----------|------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*   | 必选      | MsoGradientStyle | 渐变样式。                                                                                                                                                              |
| *Variant* | 必选      | Long             | 渐变变量。取值范围为 1 到 4，对应于“形状填充”选项卡中“渐变”选项卡上的四个渐变变量。如果 Style 为 msoGradientFromTitle 或 msoGradientFromCenter，则该参数可设为 1 或 2。 |
| *Degree*  | 必选      | Single           | 渐变程度。可为 0.0（深）到 1.0（浅）之间的值。                                                                                                                          |

**说明**

| MsoGradientStyle 可以是下列 MsoGradientStyle 常量之一。 |
|---------------------------------------------------------|
| msoGradientDiagonalDown                                 |
| msoGradientDiagonalUp                                   |
| msoGradientFromCenter                                   |
| msoGradientFromCorner                                   |
| msoGradientFromTitle                                    |
| msoGradientHorizontal                                   |
| msoGradientMixed                                        |
| msoGradientVertical                                     |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一个具有单色渐变填充效果的矩形。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill
　　　　dShapes.ForeColor.RGB = (0, 128, 128)
　　　　dShapes.OneColorGradient(msoGradientHorizontal, 1, 1)
}
```

### FillFormat.Patterned

将指定填充设置为一种图案。

**语法**

**express.Patterned ( *Pattern* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型       | 说明                     |
|-----------|-----------|----------------|--------------------------|
| *Pattern* | 必选      | MsoPatternType | 用于指定填充方式的图案。 |

**说明**

使用 **BackColor** 和 **ForeColor** 属性可设置图案中使用的颜色。

| MsoPatternType 可以是下列 MsoPatternType 常量之一。 |
|-----------------------------------------------------|
| msoPattern10Percent                                 |
| msoPattern20Percent                                 |
| msoPattern25Percent                                 |
| msoPattern30Percent                                 |
| msoPattern40Percent                                 |
| msoPattern50Percent                                 |
| msoPattern5Percent                                  |
| msoPattern60Percent                                 |
| msoPattern70Percent                                 |
| msoPattern75Percent                                 |
| msoPattern80Percent                                 |
| msoPattern90Percent                                 |
| msoPatternDarkDownwardDiagonal                      |
| msoPatternDarkHorizontal                            |
| msoPatternDarkUpwardDiagonal                        |
| msoPatternDashedDownwardDiagonal                    |
| msoPatternDashedHorizontal                          |
| msoPatternDashedUpwardDiagonal                      |
| msoPatternDashedVertical                            |
| msoPatternDiagonalBrick                             |
| msoPatternDivot                                     |
| msoPatternDottedDiamond                             |
| msoPatternDottedGrid                                |
| msoPatternHorizontalBrick                           |
| msoPatternLargeCheckerBoard                         |
| msoPatternLargeConfetti                             |
| msoPatternLargeGrid                                 |
| msoPatternLightDownwardDiagonal                     |
| msoPatternLightHorizontal                           |
| msoPatternLightUpwardDiagonal                       |
| msoPatternLightVertical                             |
| msoPatternMixed                                     |
| msoPatternNarrowHorizontal                          |
| msoPatternNarrowVertical                            |
| msoPatternOutlinedDiamond                           |
| msoPatternPlaid                                     |
| msoPatternShingle                                   |
| msoPatternSmallCheckerBoard                         |
| msoPatternSmallConfetti                             |
| msoPatternSmallGrid                                 |
| msoPatternSolidDiamond                              |
| msoPatternSphere                                    |
| msoPatternTrellis                                   |
| msoPatternWave                                      |
| msoPatternWeave                                     |
| msoPatternWideDownwardDiagonal                      |
| msoPatternWideUpwardDiagonal                        |
| msoPatternZigZag                                    |
| msoPatternDarkVertical                              |

**示例**

``` JavaScript
/*本示例将一个带图案填充的椭圆添加到 myDocument 中。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeOval, 60, 60, 80, 40).Fill
　　　　dShapes.ForeColor.RGB = (128, 0, 0)
　　　　dShapes.BackColor.RGB = (0, 0, 255)
　　　　dShapes.Patterned(msoPatternDarkVertical)
}
```

### FillFormat.PresetGradient

将指定填充设为一个预设的渐变。

**语法**

**express.PresetGradient ( *Style* , *Variant* , *PresetGradientType* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型              | 说明                                                                                                                                                                      |
|----------------------|-----------|-----------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*              | 必选      | MsoGradientStyle      | 渐变样式。                                                                                                                                                                |
| *Variant*            | 必选      | Integer               | 渐变变量。取值范围为 1 到 4，对应于“形状填充”选项卡的“渐变”子选项卡上的四个渐变变量。如果 Style 为 msoGradientFromTitle 或 msoGradientFromCenter，则该参数可设为 1 或 2。 |
| *PresetGradientType* | 必选      | MsoPresetGradientType | 渐变类型。                                                                                                                                                                |

**说明**

| MsoGradientStyle 可以是下列 MsoGradientStyle 常量之一。 |
|---------------------------------------------------------|
| msoGradientDiagonalDown                                 |
| msoGradientDiagonalUp                                   |
| msoGradientFromCenter                                   |
| msoGradientFromCorner                                   |
| msoGradientFromTitle                                    |
| msoGradientHorizontal                                   |
| msoGradientMixed                                        |
| msoGradientVertical                                     |

| MsoPresetGradientType 可以是下列 MsoPresetGradientType 常量之一。 |
|-------------------------------------------------------------------|
| msoGradientBrass                                                  |
| msoGradientCalmWater                                              |
| msoGradientChrome                                                 |
| msoGradientChromeII                                               |
| msoGradientDaybreak                                               |
| msoGradientDesert                                                 |
| msoGradientEarlySunset                                            |
| msoGradientFire                                                   |
| msoGradientFog                                                    |
| msoGradientGold                                                   |
| msoGradientGoldII                                                 |
| msoGradientHorizon                                                |
| msoGradientLateSunset                                             |
| msoGradientMahogany                                               |
| msoGradientMoss                                                   |
| msoGradientNightfall                                              |
| msoGradientOcean                                                  |
| msoGradientParchment                                              |
| msoGradientPeacock                                                |
| msoGradientRainbow                                                |
| msoGradientRainbowII                                              |
| msoGradientSapphire                                               |
| msoGradientSilver                                                 |
| msoGradientWheat                                                  |
| msoPresetGradientMixed                                            |

**示例**

本示例将一个带有预设渐变填充的矩形添加到 ` myDocument ` 。

``` JavaScript
function test() {
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 140, 80).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass)
}
```

### FillFormat.PresetTextured

将指定的填充方式设置为预设纹理。

**语法**

**express.PresetTextured ( *PresetTexture* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型         | 说明       |
|-----------------|-----------|------------------|------------|
| *PresetTexture* | 必选      | MsoPresetTexture | 预设纹理。 |

**说明**

| MsoPresetTexture 可以是下列 MsoPresetTexture 常量之一。 |
|---------------------------------------------------------|
| msoPresetTextureMixed                                   |
| msoTextureBlueTissuePaper                               |
| msoTextureBouquet                                       |
| msoTextureBrownMarble                                   |
| msoTextureCanvas                                        |
| msoTextureCork                                          |
| msoTextureDenim                                         |
| msoTextureFishFossil                                    |
| msoTextureGranite                                       |
| msoTextureGreenMarble                                   |
| msoTextureMediumWood                                    |
| msoTextureNewsprint                                     |
| msoTextureOak                                           |
| msoTexturePaperBag                                      |
| msoTexturePapyrus                                       |
| msoTextureParchment                                     |
| msoTexturePinkTissuePaper                               |
| msoTexturePurpleMesh                                    |
| msoTextureRecycledPaper                                 |
| msoTextureSand                                          |
| msoTextureStationery                                    |
| msoTextureWalnut                                        |
| msoTextureWaterDroplets                                 |
| msoTextureWhiteMarble                                   |
| msoTextureWovenMat                                      |

**示例**

``` JavaScript
/*本示例将一个带有灰褐色大理石纹理填充的圆柱形添加到 myDocument。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　myDocument.Shapes.AddShape(msoShapeCan, 90, 90, 40, 80).Fill.PresetTextured(msoTextureGreenMarble)
}
```

### FillFormat.Solid

将指定的填充格式设置为均匀的颜色。可用本方法将带有渐变、纹理、图案或背景的填充格式转换为单色的填充格式。

**语法**

**express.Solid ()**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*本示例将 myDocument 上的所有填充转换为统一的红色填充。*/
function test() {
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    for (let s = 1; s <= myDocument.Shapes.Count; s++) {
        myDocument.Shapes.Item(s).Fill.Solid()
        myDocument.Shapes.Item(s).Fill.ForeColor.RGB = (0, 255, 255)
    }
}　　
```

### FillFormat.TwoColorGradient

设置指定填充为双色渐变。

**语法**

**express.TwoColorGradient ( *Style* , *Variant* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型         | 说明                                                                                                                                                                      |
|-----------|-----------|------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*   | 必选      | MsoGradientStyle | 渐变样式。                                                                                                                                                                |
| *Variant* | 必选      | Long             | 渐变变量。取值范围为 1 到 4，对应于“形状填充”选项卡的“渐变”子选项卡上的四个渐变变量。如果 Style 为 msoGradientFromTitle 或 msoGradientFromCenter，则该参数可设为 1 或 2。 |

**说明**

| MsoGradientStyle 可以是下列 MsoGradientStyle 常量之一。 |
|---------------------------------------------------------|
| msoGradientDiagonalDown                                 |
| msoGradientDiagonalUp                                   |
| msoGradientFromCenter                                   |
| msoGradientFromCorner                                   |
| msoGradientFromTitle                                    |
| msoGradientHorizontal                                   |
| msoGradientMixed                                        |
| msoGradientVertical                                     |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一个具有双色渐变填充效果的矩形，并设置填充的前景色和背景色。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill
　　　　dShapes.ForeColor.RGB = (128, 0, 0)
　　　　dShapes.BackColor.RGB = (0, 170, 170)
　　　　dShapes.TwoColorGradient(msoGradientHorizontal, 1)
}
```

### FillFormat.UserPicture

用一幅大图像来填充指定形状。如果要通过平铺小图像来填充形状，请使用 **UserTextured** 方法。

**语法**

**express.UserPicture ( *PictureFile* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明         |
|---------------|-----------|----------|--------------|
| *PictureFile* | 必选      | String   | 图片文件名。 |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加了两个矩形。用 Tiles.bmp 文件中的一幅大图像填充左边的矩形；用 Tiles.bmp 文件中若干平铺的小图像填充右边的矩形*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes
　　　　dShapes.AddShape(msoShapeRectangle, 0, 0, 200, 100).Fill.UserPicture("c:\\windows\\tiles.bmp")
　　　　dShapes.AddShape(msoShapeRectangle, 300, 0, 200, 100).Fill.UserTextured("c:\\windows\\tiles.bmp")
}
```

### FillFormat.UserTextured

用平铺的小图像填充指定形状。如果要用一个大图像来填充该形状，请使用 **UserPicture** 方法。

**语法**

**express.UserTextured ()**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加了两个矩形。用 Tiles.bmp 文件中的一幅大图像填充左边的矩形；用 Tiles.bmp 文件中若干平铺的小图像填充右边的矩形*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes
　　　　dShapes.AddShape(msoShapeRectangle, 0, 0, 200, 100).Fill.UserPicture("c:\\windows\\tiles.bmp")
　　　　dShapes.AddShape(msoShapeRectangle, 300, 0, 200, 100).Fill.UserTextured("c:\\windows\\tiles.bmp")
}
```

## 成员属性

### FillFormat.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 FillFormat 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
funciton AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### FillFormat.BackColor

返回或设置 **ColorFormat** 对象，该对象代表指定填充对象或图案线条的背景色。可读/写。

**语法**

**express.BackColor**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
　　　　dShapes.ForeColor.RGB = (128, 0, 0)
　　　　dShapes.BackColor.RGB = (170, 170, 170)
　　　　dShapes.TwoColorGradient(msoGradientHorizontal, 1)
}
```

``` JavaScript
/*以下示例将一条带图案的线条添加至 myDocument。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dLine = myDocument.Shapes.AddLine(10, 100, 250, 0).Line
　　　　dLine.Weight = 6
　　　　dLine.ForeColor.RGB = (0, 0, 255)
　　　　dLine.BackColor.RGB = (128, 0, 0)
　　　　dLine.Pattern = msoPatternDarkDownwardDiagonal
}
```

### FillFormat.Creator

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 FillFormat 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

``` JavaScript
/*以下示例显示关于 myObject 创建者的消息。*/
function test(){
　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
　　　　if ( myObject.Creator == 0x50575054 ) {
　　　　    alert("This is a WPP object")
　　　　}
　　　　else {
　　　　    alert("This is not a WPP object")
　　　　}
}
```

### FillFormat.ForeColor

返回或设置 **ColorFormat** 对象，该对象表示填充、线条或阴影的前景色。可读/写。

**语法**

**express.ForeColor**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
　　　　dShapes.ForeColor.RGB = (128, 0, 0)
　　　　dShapes.BackColor.RGB = (170, 170, 170)
　　　　dShapes.TwoColorGradient(msoGradientHorizontal, 1)}
```

``` JavaScript
/*以下示例将一条带图案的线条添加至 myDocument。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dLine = myDocument.Shapes.AddLine(10, 100, 250, 0).Line
　　　　dLine.Weight = 6
　　　　dLine.ForeColor.RGB = (0, 0, 255)
　　　　dLine.BackColor.RGB = (128, 0, 0)
　　　　dLine.Pattern = msoPatternDarkDownwardDiagonal
}
```

### FillFormat.GradientColorType

返回指定填充的渐变颜色类型。此属性为只读。使用 **OneColorGradient** 、 **PresetGradient** 或 **TwoColorGradient** 方法设置填充的渐变类型。只读 **MsoGradientColorType** 类型。

**语法**

**express.GradientColorType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

| MsoGradientColorType 可为以下 MsoGradientColorType 常量之一。 |
|---------------------------------------------------------------|
| msoGradientColorMixed                                         |
| msoGradientOneColor                                           |
| msoGradientPresetColors                                       |
| msoGradientTwoColors                                          |

**示例**

``` JavaScript
/*本示例将 myDocument 中使用双色渐变填充的所有形状的填充更改为预设的渐变填充。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　for ( let s = 1; s <= myDocument.Shapes.Count; s++ ) {
　　　　    if ( myDocument.Shapes.Item(s).Fill.GradientColorType == msoGradientTwoColors ) {
　　　　        myDocument.Shapes.Item(s).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass)
　　　　    }
　　　　}}
```

### FillFormat.GradientDegree

返回一个值，该值指示单色渐变填充的深浅程度。该值为 0（零）表示形状的前景色中混入黑色形成渐变；值为 1 表示混入白色；介于 0 和 1 之间的值表示混入前景色的偏深或偏浅的底纹。只读 **Single** 类型。此属性为只读。使用 **OneColorGradient** 方法设置填充的渐变程度。

**语法**

**express.GradientDegree**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 添加矩形，并设置该矩形的填充渐变程度与名为“Rectangle 2”的形状中的渐变程度一致。如果“Rectangle 2”未使用单色渐变填充，则本示例失败。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes
　　　　let gradDegree1 = dShapes.Item("Rectangle 2").Fill.GradientDegree
　　　　let filldShapes = dShapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill
　　　　filldShapes.ForeColor.RGB = (128, 0, 0)
　　　　filldShapes.OneColorGradient(msoGradientHorizontal, 1, gradDegree1)
}
```

### FillFormat.GradientStyle

返回指定填充的渐变样式。使用 **OneColorGradient** 、 **PresetGradient** 或 **TwoColorGradient** 方法设置填充的渐变样式。尝试对未使用渐变的填充返回此属性将产生错误。使用 **Type** 属性决定填充是否使用渐变。只读 **MsoGradientStyle** 类型。

**语法**

**express.GradientStyle**

\*express是一个代表 FillFormat 对象的变量。

**说明**

| MsoGradientStyle 可以是下列 MsoGradientStyle 常量之一。 |
|---------------------------------------------------------|
| msoGradientDiagonalDown                                 |
| msoGradientDiagonalUp                                   |
| msoGradientFromCenter                                   |
| msoGradientFromCorner                                   |
| msoGradientFromTitle                                    |
| msoGradientHorizontal                                   |
| msoGradientMixed                                        |
| msoGradientVertical                                     |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加矩形，并设置该矩形的填充渐变样式与名为“rect1”的形状的渐变样式一致。“rect1”必须具有渐变填充，本示例才有效。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes
　　　　let gradStyle1 = dShapes.Item("rect1").Fill.GradientStyle
　　　　let filldShapes = dShapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill
　　　　filldShapes.ForeColor.RGB = (128, 0, 0)
　　　　filldShapes.OneColorGradient(gradStyle1, 1, 1)
}
```

### FillFormat.GradientVariant

返回指定填充的渐变变量，对于大多数渐变填充，其返回值为从 1 到 4 的整数。如果渐变样式为 **msoGradientFromTitle** 或 **msoGradientFromCenter** ，则该属性返回 1 或 2。该属性的值与 **“形状填充”** 选项卡中的 **“渐变”** 子选项卡上的渐变变量（从左向右，从上向下编号）相对应。 **Long** 类型，只读。此属性为只读。使用 **OneColorGradient** 、 **PresetGradient** 或 **TwoColorGradient** 方法可设置填充的渐变变量。

**语法**

**express.GradientVariant**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加矩形，并设置该矩形的填充渐变变量，以匹配名为“rect1”的形状的渐变变量。“rect1”必须具有渐变填充，本示例才有效。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes
　　　　let gradVar1 = dShapes.Item("rect1").Fill.GradientVariant
　　　　let filldShapes = dShapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill
　　　　filldShapes.ForeColor.RGB = (128, 0, 0)
}
```

### FillFormat.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

### FillFormat.Pattern

设置或返回一个值，该值代表应用于指定填充的图案。使用 **BackColor** 和 **ForeColor** 属性可设置图案中使用的颜色。只读。 **MsoPatternType** 类型。

**语法**

**express.Pattern**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一个矩形，并设置其填充图案与名为“rect1”的形状保持一致。新矩形使用和 rect1 相同的图案，但颜色未必相同。通过 BackColor 和 ForeColor 属性可设置该图案中使用的颜色。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes
　　　　let pattern1 = dShapes.Item("rect1").Fill.Pattern
　　　　let fillShapes = dShapes.AddShape(msoShapeRectangle, 100, 100, 120, 80).Fill
　　　　fillShapes.ForeColor.RGB = (128, 0, 0)
　　　　fillShapes.BackColor.RGB = (0, 0, 255)
　　　　fillShapes.Patterned(pattern1)
}
```

### FillFormat.PresetGradientType

返回指定填充的预设渐变类型。只读 **MsoPresetGradientType** 类型。使用 **PresetGradient** 方法可设置填充的预设渐变类型。

**语法**

**express.PresetGradientType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

| MsoPresetGradientType 可以是下列 MsoPresetGradientType 常量之一。 |
|-------------------------------------------------------------------|
| msoGradientBrass                                                  |
| msoGradientCalmWater                                              |
| msoGradientChrome                                                 |
| msoGradientChromeII                                               |
| msoGradientDaybreak                                               |
| msoGradientDesert                                                 |
| msoGradientEarlySunset                                            |
| msoGradientFire                                                   |
| msoGradientFog                                                    |
| msoGradientGold                                                   |
| msoGradientGoldII                                                 |
| msoGradientHorizon                                                |
| msoGradientLateSunset                                             |
| msoGradientMahogany                                               |
| msoGradientMoss                                                   |
| msoGradientNightfall                                              |
| msoGradientOcean                                                  |
| msoGradientParchment                                              |
| msoGradientPeacock                                                |
| msoGradientRainbow                                                |
| msoGradientRainbowII                                              |
| msoGradientSapphire                                               |
| msoGradientSilver                                                 |
| msoGradientWheat                                                  |
| msoPresetGradientMixed                                            |

**示例**

``` JavaScript
/*本示例将 myDocument 中所有具有 Moss 预设渐变填充的形状的填充更改为 Fog 预设渐变填充。*/
function test(){ 
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　for ( let s = 1; s <= myDocument.Shapes.Count; s++ ) {
　　　　    if ( myDocument.Shapes.Item(s).Fill.PresetGradientType == msoGradientMoss ) {
　　　　        myDocument.Shapes.Item(s).Fill.PresetGradient = msoGradientFog
　　　　    }
　　　　}
}
```

### FillFormat.PresetTexture

返回指定填充的预设纹理。只读 **MsoPresetTexture** 类型。

**语法**

**express.PresetTexture**

\*express是一个代表 FillFormat 对象的变量。

**说明**

| MsoPresetTexture 可以是下列 MsoPresetTexture 常量之一。 |
|---------------------------------------------------------|
| msoPresetTextureMixed                                   |
| msoTextureBlueTissuePaper                               |
| msoTextureBouquet                                       |
| msoTextureBrownMarble                                   |
| msoTextureCanvas                                        |
| msoTextureCork                                          |
| msoTextureDenim                                         |
| msoTextureFishFossil                                    |
| msoTextureGranite                                       |
| msoTextureGreenMarble                                   |
| msoTextureMediumWood                                    |
| msoTextureNewsprint                                     |
| msoTextureOak                                           |
| msoTexturePaperBag                                      |
| msoTexturePapyrus                                       |
| msoTextureParchment                                     |
| msoTexturePinkTissuePaper                               |
| msoTexturePurpleMesh                                    |
| msoTextureRecycledPaper                                 |
| msoTextureSand                                          |
| msoTextureStationery                                    |
| msoTextureWalnut                                        |
| msoTextureWaterDroplets                                 |
| msoTextureWhiteMarble                                   |
| msoTextureWovenMat                                      |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一个矩形，并设置该矩形的预设纹理与第二个形状的预设纹理一致。要使本示例正常运行，第二个形状必须使用预设纹理的填充。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let dShapes = myDocument.Shapes
　　　　let presetTexture2 = dShapes.Item(2).Fill.PresetTexture
　　　　dShapes.AddShape(msoShapeRectangle, 100, 0, 40, 80).Fill.PresetTextured(presetTexture2)}
```

### FillFormat.TextureName

返回指定填充的自定义纹理文件名称。只读 **String** 类型。

**语法**

**express.TextureName**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性为只读。使用 **UserTextured** 方法可设置填充的纹理文件。

**示例**

``` JavaScript
/*本示例给 myDocument 添加一个椭圆。如果 myDocument 的第一个形状使用用户定义的纹理填充，则新椭圆使用和第一个形状相同的填充。如果第一个形状使用其他任何类型的填充，则新椭圆使用绿色大理石填充*/
function test() {
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let newFill = myDocument.Shapes.AddShape(msoShapeOval, 0, 0, 200, 90).Fill
    let fill = myDocument.Shapes.Item(1).Fill
    if (fill.Type == msoFillTextured && fill.TextureType == msoTextureUserDefined) {
        newFill.UserTextured(fill.TextureName)
    }
    else {
        newFill.PresetTextured(msoTextureGreenMarble)
    }
}
```

### FillFormat.TextureType

返回指定填充的纹理类型。 **MsoTextureType** 类型，只读。使用 **PresetTextured** 或 **UserTextured** 方法可设置填充的纹理类型。

**语法**

**express.TextureType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

| MsoTextureType 可为以下 MsoTextureType 常量之一。 |
|---------------------------------------------------|
| msoTexturePreset                                  |
| msoTextureTypeMixed                               |
| msoTextureUserDefined                             |

**示例**

``` JavaScript
/*本示例会将 myDocument 中所有形状的填充更改为画布，这些形状全部使用自定义的纹理填充。 */
function test() {
    let shape = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shape.Count; i++) {
        if (shape.Item(i).Fill.TextureType == msoTextureUserDefined) {
            shape.Item(i).Fill.PresetTextured(msoTextureCanvas)
        }
    }
}
```

### FillFormat.Transparency

该属性返回或设置指定填充、阴影或线条的透明度，可取 0.0（不透明）至 1.0（透明）之间的数值。 **Single** 类型，可读/写。

**语法**

**express.Transparency**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

**示例**

``` JavaScript
/*本例将 myDocument 中第三个形状的阴影设置为半透明的红色。如果形状原来无阴影，以下示例为它添加阴影。*/
function test() {
    let sh = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).Shadow
    sh.Visible = true
    sh.ForeColor.RGB = (255, 0, 0)
    sh.Transparency = 0.5
}
```

### FillFormat.Type

返回一个 **MsoFillType** 常量，该常量代表填充的类型。只读。

**语法**

**express.Type**

\*express是一个代表 FillFormat 对象的变量。

**说明**

| MsoFillType 可为以下 MsoFillType 常量之一。 |
|---------------------------------------------|
| msoFillBackground                           |
| msoFillGradient                             |
| msoFillMixed                                |
| msoFillPatterned                            |
| msoFillPicture                              |
| msoFillSolid                                |
| msoFillTextured                             |

### FillFormat.Visible

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 **MsoTriState** 类型。

**语法**

**express.Visible**

\*express是一个代表 FillFormat 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue：指定对象或对象格式是可见的。         |

**示例**

``` JavaScript
/*以下示例为当前演示文稿中第一张幻灯片的第三个形状设置阴影的水平偏移量和垂直偏移量，使其向右偏移 5 磅并向上偏移 3 磅。如果该形状原来没有阴影，本例将会为其添上阴影。*/
function test() {
    let sh = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).Shadow
    sh.Visible = msoTrue
    sh.OffsetX = 5
    sh.OffsetY = -3
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/FillFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# OLEFormat 对象

## 简介

包含应用于 OLE 对象的属性和方法。

## 说明

LinkFormat 对象包含仅应用于链接 OLE 对象的属性和方法。 PictureFormat 对象包含应用于图片和 OLE 对象的属性和方法。

## 方法

| 序号 | 名称                            | 说明                                                                            |
|:----:|:--------------------------------|:--------------------------------------------------------------------------------|
|  1   | [Activate](#OLEFormat.Activate) | 激活指定的对象。                                                                |
|  2   | [DoVerb](#OLEFormat.DoVerb)     | 请求 OLE 对象执行其中一个动作。使用 ObjectVerbs 属性可确定 OLE 对象的可用动作。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                            |
|:----:|:----------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#OLEFormat.Application)   | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                         |
|  2   | [FollowColors](#OLEFormat.FollowColors) | 返回或设置指定对象中的颜色对幻灯片配色方案的继承程度。指定的对象必须是在 Graph 或 Organization Chart 中创建的图表。可读/写 PpFollowColors 类型。                                                |
|  3   | [Object](#OLEFormat.Object)             | 返回一个对象，该对象代表指定 OLE 对象的顶层界面。该属性使您可以访问某些应用程序中的属性和方法（OLE 对象在这些应用程序中创建）。只读。                                                           |
|  4   | [ObjectVerbs](#OLEFormat.ObjectVerbs)   | 返回一个 ObjectVerbs 集合，该集合包含指定的 OLE 对象的所有 OLE动作。只读。                                                                                                                      |
|  5   | [Parent](#OLEFormat.Parent)             | 返回指定对象的父对象。                                                                                                                                                                          |
|  6   | [ProgID](#OLEFormat.ProgID)             | 返回指定 OLE 对象的 程序标识符 (ProgID) （程序标识符：窗体 OLEServerName.ObjectName 中的标识符（例如，Excel.Sheet 或WPP.Slide），由 Windows 注册表用来唯一标识一个对象。） 。只读 String 类型。 |

## 成员方法

### OLEFormat.Activate

激活指定的对象。

**语法**

**express.Activate ()**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

本示例按顺序激活紧挨在活动窗口之后的文档窗口。

``` JavaScript
Application.Windows.Item(2).Activate()
```

### OLEFormat.DoVerb

请求 OLE 对象执行其中一个动作。使用 **ObjectVerbs** 属性可确定 OLE 对象的可用动作。

**语法**

**express.DoVerb ( *Index* )**

\*express是一个代表 OLEFormat 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                           |
|---------|-----------|----------|------------------------------------------------|
| *Index* | 必选      | Integer  | 要执行的动作。如果忽略该参数，则执行默认动作。 |

**示例**

``` JavaScript
/*本示例对当前演示文稿中第一张幻灯片的第三个形状执行默认动作，其中第三个形状是链接或嵌人的 OLE 对象。*/
function test(){
　　　　let sli = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
 　　　　   if(sli.Type == msoEmbeddedOLEObject || sli.Type == msoLinkedOLEObject){
     　　　　   sli.OLEFormat.DoVerb()
  　　　　  }
}
```

``` JavaScript
/*本示例对当前演示文稿中第一张幻灯片的第三个形状执行默认动作“Open”，其中第三个形状是支持动作“Open”的 OLE 对象。*
/
function test(){
　　　　let sli = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
  　　　　  if(sli.Type == msoEmbeddedOLEObject || sli.Type == msoLinkedOLEObject){
    　　　　    for(let i = 1; i <= sli.OLEFormat.ObjectVerbs.Count; i++){
      　　　　      nCount++
       　　　　     if(sli.OLEFormat.ObjectVerbs.Item(i) == "Open"){
        　　　　        sli.OLEFormat.DoVerb(nCount)
      　　　　      }
   　　　　     }
　　　　   }
}
```

## 成员属性

### OLEFormat.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres){
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### OLEFormat.FollowColors

返回或设置指定对象中的颜色对幻灯片配色方案的继承程度。指定的对象必须是在 Graph 或 Organization Chart 中创建的图表。可读/写 **PpFollowColors** 类型。

**语法**

**express.FollowColors**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

| PpFollowColors 可以是下列 PpFollowColors 常量之一。                |
|--------------------------------------------------------------------|
| ppFollowColorsNone 图表颜色不继承幻灯片配色方案。                  |
| ppFollowColorsMixed                                                |
| ppFollowColorsScheme 图表中所有的颜色都继承幻灯片配色方案。        |
| ppFollowColorsTextAndBackground 只有文本和背景继承幻灯片配色方案。 |

**示例**

本示例指定当前演示文稿第一张幻灯片第二个形状的文本和背景继承幻灯片配色方案。第二个形状必须是 Graph 或 Organization Chart 创建的图表。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).OLEFormat.FollowColors = ppFollowColorsTextAndBackground
```

### OLEFormat.Object

返回一个对象，该对象代表指定 OLE 对象的顶层界面。该属性使您可以访问某些应用程序中的属性和方法（OLE 对象在这些应用程序中创建）。只读。

**语法**

**express.Object**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

使用 **TypeName** 函数判断此属性为特定的 OLE 对象返回的对象的类型。

**示例**

本示例显示当前演示文稿第一张幻灯片第一个形状中包含的对象的类型。第一个形状必须包含 OLE 对象。

``` JavaScript
alert(TypeName(Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).OLEFormat.Object))
```

### OLEFormat.ObjectVerbs

返回一个 **ObjectVerbs** 集合，该集合包含指定的 OLE 对象的所有 OLE动作。只读。

**语法**

**express.ObjectVerbs**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*本示例显示了当前演示文稿中第二张幻灯片上第一个形状内的 OLE 对象具有的所有可用操作。只有当第一个形状是 OLE 对象时，这个例子才会有效。*/
function test(){
let ole = Application.ActivePresentation.Slides.Item(2).Shapes.Item(1).OLEFormat
    for(let i = 1; i <= ole.ObjectVerbs.Count; i++){
        alert(ole.ObjectVerbs)
    }
}
```

``` JavaScript
/*本示例表明：对于当前演示文稿上第二张幻灯片中第一个形状所代表的 OLE 对象而言，若“Open”为该对象的一个 OLE 操作，那么在幻灯片演示过程中如果单击该对象，则此对象会打开。只有当第一个形状是 OLE 对象时，这个例子才会有效。*/
function test() {
    let sli = Application.ActivePresentation.Slides.Item(2).Shapes.Item(1)
    for (let i = 1; i <= sli.OLEFormat.ObjectVerbs.Count; i++) {
        if (sli.OLEFormat.ObjectVerbs.Item(i) == "Open") {
            let set = sli.ActionSettings(ppMouseClick)
            set.Action = ppActionOLEVerb
            set.ActionVerb = sli.OLEFormat.ObjectVerbs.Item(i)
        }
    }
}
```

### OLEFormat.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let tex = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
　　　　tex.TextRange.Text = "Test text"
　　　　tex.Parent.Rotation = 45
}
```

### OLEFormat.ProgID

返回指定 OLE 对象的 程序标识符 (ProgID) （程序标识符：窗体 OLEServerName.ObjectName 中的标识符（例如，Excel.Sheet 或WPP.Slide），由 Windows 注册表用来唯一标识一个对象。） 。只读 **String** 类型。

**语法**

**express.ProgID**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*本示例搜索当前演示文稿所有幻灯片中的所有对象，将全部链接的 ET 工作表设为手动更新。*/
function test() {
    let te = Application.ActivePresentation.Slides
    for (let i = 1; i <= te.Count; i++) {
        for (let ii = 1; ii <= te.Item(i).Shapes;) {
            if (te.Item(i).Shapes.Item(ii).Type == msoLinkedOLEObject) {
                if (te.Item(i).Shapes.Item(ii).OLEFormat.ProgID == "Excel.Sheet") {
                    te.Item(i).Shapes.Item(ii).LinkFormat.AutoUpdate = ppUpdateOptionManual
                }
            }
        }
    }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/OLEFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Panes 对象

## 简介

Pane 对象的集合，这些对象代表普通视图内文档窗口中的幻灯片、大纲和备注窗格，或代表文档窗口中其他任意视图内的单个窗格。

## 说明

在普通视图中， Panes 集合包含三个成员。所有其他文档窗口视图则仅有一个窗格，因而 Panes 集合也就只有一个成员。

``` JavaScript
使用 Panes 属性可返回 Panes 集合。以下示例将测试活动窗口中窗格的数目。如果值为 1，则说明是除普通视图之外的其他视图，接着将激活普通视图，然后使用纵向窗格分隔线将文档窗口进行划分，使大纲窗格占 15% 位置，幻灯片窗格占 85% 位置。
function test() {
    let mypane = Application.ActiveWindow

    if(mypane.Panes.Count == 1){
       mypane.ViewType = ppViewNormal
       mypane.SplitHorizontal = 15
    }
}
```

## 方法

| 序号 | 名称                | 说明                                    |
|:----:|:--------------------|:----------------------------------------|
|  1   | [Item](#Panes.Item) | 从指定集合中返回单个对象。返回值为 Pane |

## 属性

| 序号 | 名称                              | 说明                                                    |
|:----:|:----------------------------------|:--------------------------------------------------------|
|  1   | [Application](#Panes.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#Panes.Count)             | 返回指定集合中的对象数目。                              |
|  3   | [Parent](#Panes.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### Panes.Item

从指定集合中返回单个对象。返回值为 Pane

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Panes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                             |
|---------|-----------|----------|----------------------------------|
| *Index* | 必选      | Long     | 集合中要返回的单个对象的索引号。 |

## 成员属性

### Panes.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Panes 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres){
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test() {
  let mypane = Application.ActivePresentation.Slides.Item(1).Shapes

  for(let i = 1; i <= mypane.Count; i++){
      if(mypane.Item(i).Type == msoLinkedOLEObject){
          alert(mypane.Item(i).OLEFormat.Application.Name)
      }
  }
}
```

### Panes.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 Panes 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口。*/
function test() {
  let mypane = Application.Windows  
  for(let i = 2; i <= mypane.Count; i++){
      mypane.Item(i).Close()
  }
}
```

### Panes.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Panes 对象的变量。

**示例**

``` JavaScript
//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Panes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Slides 对象

## 简介

指定的演示文稿中所有 Slide 对象的集合。

## 说明

以下示例说明如何执行下列操作：

- 创建一张幻灯片并将其添加到集合中
- 返回通过名称、索引号或幻灯片 ID 号指定的单张幻灯片
- 返回演示文稿中幻灯片的子集
- 同时将一种属性或方法应用到演示文稿中的所有幻灯片

## 示例

使用 Slides 属性返回 Slides 集合。使用 [Add](#Slides.Add) 方法新建幻灯片并将其添加到集合中。以下示例将新幻灯片添加到活动演示文稿中。

``` JavaScript
Application.ActivePresentation.Slides.Add(2, ppLayoutBlank)
```

使用 Slides ( *index* )（其中 *index* 是幻灯片名称或索引号）或使用 Slides.FindBySlideID ( *index* )（其中 *index* 是幻灯片 ID 号）返回单个 Slide 对象。以下示例设置活动演示文稿中第一张幻灯片的版式。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Layout = ppLayoutTitle
```

使用 Slides.Range ( *index* ) 返回代表 Slides 集合的一个子集的 SlideRange 对象，其中 *index* 是幻灯片索引号或名称，或者幻灯片索引号的数组或幻灯片名称的数组。

如果要同时对演示文稿中所有幻灯片进行某种操作（例如全部删除或设置它们的某些属性），可不带参数使用 Slides.Range 创建一个包含 Slides 集合中所有幻灯片的 SlideRange 集合，然后将适当的属性或方法应用于 SlideRange 集合。

## 方法

| 序号 | 名称                                     | 说明                                                                                                                                                                              |
|:----:|:-----------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Add](#Slides.Add)                       | 创建新的幻灯片，将其添加到 Slides 集合，然后返回幻灯片。                                                                                                                          |
|  2   | [AddSlide](#Slides.AddSlide)             | 创建新的幻灯片，将其添加到 Slides 集合，然后返回幻灯片。                                                                                                                          |
|  3   | [FindBySlideID](#Slides.FindBySlideID)   | 返回一个 Slide 对象，该对象代表具有指定幻灯片 ID 号的幻灯片。每张幻灯片在创建时都会自动分配唯一的幻灯片 ID 号。使用 SlideID 属性可返回幻灯片的 ID 号。                            |
|  4   | [InsertFromFile](#Slides.InsertFromFile) | 将来自某文件的幻灯片插入到演示文稿中的指定位置。返回一个代表所插入幻灯片数量的 integer 。                                                                                         |
|  5   | [Item](#Slides.Item)                     | 从指定集合中返回单个对象。                                                                                                                                                        |
|  6   | [Paste](#Slides.Paste)                   | 将剪贴板上的幻灯片粘贴到演示文稿的 Slides 集合中。用 Index 参数可以指定要插入幻灯片的位置。返回一个代表粘贴对象的 SlideRange 对象。每张粘贴幻灯片都成为指定的 Slides 集合的成员。 |
|  7   | [Range](#Slides.Range)                   | 返回一个代表 Slides 集合中的幻灯片子集的 SlideRange 对象。                                                                                                                        |

## 属性

| 序号 | 名称                               | 说明                                                    |
|:----:|:-----------------------------------|:--------------------------------------------------------|
|  1   | [Application](#Slides.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#Slides.Count)             | 返回指定集合中的对象数目。                              |

## 成员方法

### Slides.Add

创建新的幻灯片，将其添加到 Slides 集合，然后返回幻灯片。

**语法**

**express.Add ( *Index* , *PpSlideLayout* )**

\*express是一个代表 Slides 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型      | 说明                   |
|-----------------|-----------|---------------|------------------------|
| *Index*         | 必选      | Number        | 要添加的幻灯片的索引。 |
| *PpSlideLayout* | 必选      | PpSlideLayout | 幻灯片的版式。         |

**说明**

|                                                   |
|---------------------------------------------------|
| PpSlideLayout 可以是下列 PpSlideLayout 常量之一。 |
| **ppLayoutBlank**                                 |
| **ppLayoutChart**                                 |
| **ppLayoutChartAndText**                          |
| **ppLayoutClipartAndText**                        |
| **ppLayoutClipArtAndVerticalText**                |
| **ppLayoutFourObjects**                           |
| **ppLayoutLargeObject**                           |
| **ppLayoutMediaClipAndText**                      |
| **ppLayoutMixed**                                 |
| **ppLayoutObject**                                |
| **ppLayoutObjectAndText**                         |
| **ppLayoutObjectOverText**                        |
| **ppLayoutOrgchart**                              |
| **ppLayoutTable**                                 |
| **ppLayoutText**                                  |
| **ppLayoutTextAndChart**                          |
| **ppLayoutTextAndClipart**                        |
| **ppLayoutTextAndMediaClip**                      |
| **ppLayoutTextAndObject**                         |
| **ppLayoutTextAndTwoObjects**                     |
| **ppLayoutTextOverObject**                        |
| **ppLayoutTitle**                                 |
| **ppLayoutTitleOnly**                             |
| **ppLayoutTwoColumnText**                         |
| **ppLayoutTwoObjectsAndText**                     |
| **ppLayoutTwoObjectsOverText**                    |
| **ppLayoutVerticalText**                          |
| **ppLayoutVerticalTitleAndText**                  |
| **ppLayoutVerticalTitleAndTextOverChart**         |

### Slides.AddSlide

创建新的幻灯片，将其添加到 Slides 集合，然后返回幻灯片。

**语法**

**express.AddSlide ( *Index* , *pCustomLayout* )**

\*express是一个代表 Slides 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型     | 说明                   |
|-----------------|-----------|--------------|------------------------|
| *Index*         | 必选      | Integer      | 要添加的幻灯片的索引。 |
| *pCustomLayout* | 必选      | CustomLayout | 幻灯片的版式。         |

### Slides.FindBySlideID

返回一个 **Slide** 对象，该对象代表具有指定幻灯片 ID 号的幻灯片。每张幻灯片在创建时都会自动分配唯一的幻灯片 ID 号。使用 **SlideID** 属性可返回幻灯片的 ID 号。

**语法**

**express.FindBySlideID ( *SlideID* )**

\*express是一个代表 Slides 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                        |
|-----------|-----------|----------|-------------------------------------------------------------|
| *SlideID* | 必选      | Number   | 指定要返回的幻灯片的 ID 号。WPP在创建幻灯片时会分配该编号。 |

**说明**

与 **SlideIndex** 属性不同，在演示文稿中添加或重新排列幻灯片时， **Slide** 对象的 **SlideID** 属性不会改变。因此，与对 **Item** 方法使用幻灯片索引号相比，对 **FindBySlideID** 方法使用幻灯片 ID 号可以更可靠地从 **Slides** 集合返回特定的 **Slide** 对象。

**示例**

``` JavaScript
/*本示例示范如何检索 Slide 对象的唯一 ID 号，并使用此编号从 Slides 集合返回该 Slide 对象。*/
function test()
{
    let gslides = Application.ActivePresentation.Slides

    // Get slide ID
    let graphSlideID = gslides.Add(2, ppLayoutChart).SlideID
    
    // Use ID to return specific slide
    gslides.FindBySlideID(graphSlideID).SlideShowTransition.EntryEffect = ppEffectCoverLeft
}
```

### Slides.InsertFromFile

将来自某文件的幻灯片插入到演示文稿中的指定位置。返回一个代表所插入幻灯片数量的 **integer** 。

**语法**

**express.InsertFromFile ( *FileName* , *Index* , *SlideStart* , *SlideEnd* )**

\*express是一个代表 Slides 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                            |
|--------------|-----------|----------|---------------------------------------------------------------------------------|
| *FileName*   | 必选      | String   | 指示包含要插入的幻灯片的文件的文件名。                                          |
| *Index*      | 必选      | Number   | 要在其后插入新幻灯片的 Slide 对象在指定 Slides 集合中的索引号。                 |
| *SlideStart* | 可选      | Number   | 第一个 Slide 对象在 Slides 集合中的索引号，该集合在由 FileName 指定的文件中。   |
| *SlideEnd*   | 可选      | Number   | 最后一个 Slide 对象在 Slides 集合中的索引号，该集合在由 FileName 指定的文件中。 |

**示例**

``` JavaScript
/*本示例在活动演示文稿的第二张灯片后插入文件 C:\Ppt\Sales.ppt 的第三张到第六张幻灯片。*/
Application.ActivePresentation.Slides.InsertFromFile("C:\\ppt\\sales.ppt", 2, 3, 6)
```

### Slides.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Slides 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

### Slides.Paste

将剪贴板上的幻灯片粘贴到演示文稿的 **Slides** 集合中。用 **Index** 参数可以指定要插入幻灯片的位置。返回一个代表粘贴对象的 **SlideRange** 对象。每张粘贴幻灯片都成为指定的 **Slides** 集合的成员。

**语法**

**express.Paste ( *Index* )**

\*express是一个代表 Slides 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                           |
|---------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------|
| *Index* | 可选      | Number   | 表示要将剪贴板上的幻灯片粘贴在它前面的幻灯片的索引号。如果省略此参数，则剪贴板上的幻灯片将粘贴到演示文稿的最后一张幻灯片之后。 |

**说明**

将剪贴板内容粘贴到窗口中之前，请使用 **ViewType** 属性设置窗口视图。下表显示了可以粘贴到每个视图中的内容。

| 视图                   | 可以从剪贴板中粘贴以下内容                                                                                                                                                                                                                                                                                                |
|------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 幻灯片视图或备注页视图 | 形状、文本或整张幻灯片。如果从剪贴板粘贴一个幻灯片，该幻灯片的图像将作为嵌入对象被插入到幻灯片、母版或备注页中；如果选中形状，粘贴的文本将附加到形状文本之后；如果选中文本，粘贴的文本将替换选中的文本；如果选中任何其他对象，粘贴的文本将被放到它自己的文本框中。粘贴的形状将被放到 z 顺序的最上端且不会替换选中的形状。 |
| 大纲视图               | 文本或整张幻灯片。不能向大纲视图粘贴形状。粘贴的幻灯片将被插到光标所在的幻灯片之前。                                                                                                                                                                                                                                      |
| 幻灯片浏览视图         | 整张幻灯片。不能向幻灯片浏览视图粘贴形状或文本。粘贴的幻灯片将被插到光标处或演示文稿中最后选中的一张幻灯片之后。                                                                                                                                                                                                          |

**示例**

``` JavaScript
/*本示例从演示文稿“Old Sales”中剪切第三张和第五张幻灯片，然后将它们插入到活动演示文稿中的第四张幻灯片之前。*/
function test()
{
    Application.Presentations.Item("Old Sales").Slides.Range([3, 5]).Cut()
    Application.ActivePresentation.Slides.Paste(4)
}
```

### Slides.Range

返回一个代表 **Slides** 集合中的幻灯片子集的 **SlideRange** 对象。

**语法**

**express.Range ( *Index* )**

\*express是一个代表 Slides 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                    |
|---------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Index* | 可选      | Variant  | 要包含在范围内的各个幻灯片。可以是 Integer 类型的值（指定幻灯片的索引号）、String 类型的值（指定幻灯片的名称），或者是包含整数或字符串的数组。如果省略此参数，则 Range 方法将返回指定集合中的所有对象。 |

**说明**

虽然可以使用 **Range** 方法返回任意数目的形状或幻灯片，但如果只需要返回该集合的一个成员，则使用 **Item** 方法会更为简单。例如，Shapes(1) 比 Shapes.Range(1) 简单，Slides(2) 比 Slides.Range(2) 简单。

## 成员属性

### Slides.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Slides 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某函数。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres) 
{
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")  
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test()
{
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for(let i = 1; i <= shpOle.Count; i++) 
    {
        if(shpOle.Item(i).Type == msoLinkedOLEObject) 
        {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### Slides.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 Slides 对象的变量。

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Slides](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TextFrame 对象

## 简介

代表 Shape 对象中的文本框架。包含文本框架中的文本以及控制文本框架的对齐和定位的属性和方法。

## 说明

使用 TextFrame 属性可返回一个 TextFrame 对象。

``` JavaScript
/*下例向 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框的边距。*/
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250,140).TextFrame

    rng.Characters().Text = "Here is some test text"
    rng.MarginBottom = 10
    rng.MarginLeft = 10
    rng.MarginRight = 10
    rng.MarginTop = 10
}
```

## 方法

| 序号 | 名称                                | 说明 |
|:----:|:------------------------------------|:-----|
|  1   | [DeleteText](#TextFrame.DeleteText) |      |

## 属性

| 序号 | 名称                                            | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#TextFrame.Application)           | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AutoSize](#TextFrame.AutoSize)                 | 如果指定对象能自动调整大小，以适应其中所包含的文字，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                 |
|  3   | [HorizontalAnchor](#TextFrame.HorizontalAnchor) |                                                                                                                                                                                                                                 |
|  4   | [MarginBottom](#TextFrame.MarginBottom)         | 以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读/写。 Single 类型。                                                                                                                                   |
|  5   | [MarginLeft](#TextFrame.MarginLeft)             | 以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。 Single 类型。                                                                                                                                |
|  6   | [MarginRight](#TextFrame.MarginRight)           | 以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。 Single 类型。                                                                                                                                |
|  7   | [MarginTop](#TextFrame.MarginTop)               | 以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。 Single 类型。                                                                                                                                  |
|  8   | [Orientation](#TextFrame.Orientation)           | 返回或设置一个 Long 值，它代表文本框的方向。                                                                                                                                                                                    |
|  9   | [Ruler](#TextFrame.Ruler)                       |                                                                                                                                                                                                                                 |
|  10  | [TextRange](#TextFrame.TextRange)               |                                                                                                                                                                                                                                 |
|  11  | [VerticalAnchor](#TextFrame.VerticalAnchor)     |                                                                                                                                                                                                                                 |
|  12  | [WordWrap](#TextFrame.WordWrap)                 |                                                                                                                                                                                                                                 |

## 成员方法

### TextFrame.DeleteText

**语法**

**express.DeleteText ()**

\*express是一个代表 TextFrame 对象的变量。

## 成员属性

### TextFrame.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
  let myObject = Application.ActiveWorkbook
  if(myObject.Application.Value == "ET"){
      alert("This is an ET Application object.")
  }
  else{
      alert("This is not an ET Application object.")
  }
}
```

### TextFrame.AutoSize

如果指定对象能自动调整大小，以适应其中所包含的文字，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoSize**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例使第一个形状的文本框能自动调整大小，以适应其中所包含的文字。*/
Application.Worksheets.Item(1).Shapes.Item(1).TextFrame.AutoSize = true
```

### TextFrame.HorizontalAnchor

**语法**

**express.HorizontalAnchor**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.MarginBottom

以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读/写。 **Single** 类型。

**语法**

**express.MarginBottom**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
function test() {
  let myDocument = Application.Worksheets.Item(1)
  let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
  rng.Characters().Text = "Here is some test text"
  rng.MarginBottom = 0
  rng.MarginLeft = 100
  rng.MarginRight = 0
  rng.MarginTop = 20
}
```

### TextFrame.MarginLeft

以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。 **Single** 类型。

**语法**

**express.MarginLeft**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
function test() {
  let myDocument = Application.Worksheets.Item(1)
  let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
  rng.Characters().Text = "Here is some test text"
  rng.MarginBottom = 0
  rng.MarginLeft = 100
  rng.MarginRight = 0
  rng.MarginTop = 20
}
```

### TextFrame.MarginRight

以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。 **Single** 类型。

**语法**

**express.MarginRight**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
function test() {
  let myDocument = Application.Worksheets.Item(1)
  let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
  rng.Characters().Text = "Here is some test text"
  rng.MarginBottom = 0
  rng.MarginLeft = 100
  rng.MarginRight = 0
  rng.MarginTop = 20
}
```

### TextFrame.MarginTop

以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。 **Single** 类型。

**语法**

**express.MarginTop**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
function test() {
  let myDocument = Application.Worksheets.Item(1)
  let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
  rng.Characters().Text = "Here is some test text"
  rng.MarginBottom = 0
  rng.MarginLeft = 100
  rng.MarginRight = 0
  rng.MarginTop = 20
}
```

### TextFrame.Orientation

返回或设置一个 **Long** 值，它代表文本框的方向。

**语法**

**express.Orientation**

\*express是一个代表 TextFrame 对象的变量。

**说明**

此属性值可设为 -90 到 90 度之间的整数值或以下 **<span "="" r="" rlidawscontentredir?assetid="HV100730982052&amp;CTT=11&amp;Origin=HV103325142052&quot;" style="color: rgb(0, 0, 0); text-decoration-line: none;" target="_new">MsoTextOrientation</span>** 常量之一。

### TextFrame.Ruler

**语法**

**express.Ruler**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.TextRange

**语法**

**express.TextRange**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.VerticalAnchor

**语法**

**express.VerticalAnchor**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.WordWrap

**语法**

**express.WordWrap**

\*express是一个代表 TextFrame 对象的变量。

> WPS 开发文档： [WPS 基础接口/演示 API 参考/TextFrame](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# COMAddIn 对象

## 简介

代表 WPS Office宿主应用程序中的一个 COM 加载项。 COMAddIn 对象是 COMAddIns 集合的成员。

## 说明

可以使用 COMAddIns.Item(index)，其中 index 既可以是一个能够返回 COMAddIns 集合中相应位置的加载项的序数值，也可以是代表指定 COM 加载项的 ProgID 的 String 值。

## 属性

| 序号 | 名称                                 | 说明                                                                                                                            |
|:----:|:-------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#COMAddIn.Application) | 获取一个 Application 对象，代表 COMAddIn 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Connect](#COMAddIn.Connect)         | 获取或设置指定的 COMAddIn 对象的连接状态。可读/写。                                                                             |
|  3   | [Creator](#COMAddIn.Creator)         | 获取一个 32 位整数，指示创建 COMAddIn 对象时所使用的应用程序。只读。                                                            |
|  4   | [Description](#COMAddIn.Description) | 获取或设置指定的 COMAddin 对象的说明性 String 值。可读/写。                                                                     |
|  5   | [Guid](#COMAddIn.Guid)               | 获取指定 COMAddIn 对象的类标识符 (CLSID)。只读。                                                                                |
|  6   | [Object](#COMAddIn.Object)           | 获取或设置一个对象引用。可读/写。                                                                                               |
|  7   | [Parent](#COMAddIn.Parent)           | 获取 COMAddIn 对象的 Parent 对象。只读。                                                                                        |
|  8   | [ProgId](#COMAddIn.ProgId)           | 获取指定 COMAddIn 对象的编程标识符 (ProgID)。只读。                                                                             |

## 成员属性

### COMAddIn.Application

获取一个 **Application** 对象，代表 **COMAddIn** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 COMAddIn 对象的变量。

### COMAddIn.Connect

获取或设置指定的 **COMAddIn** 对象的连接状态。可读/写。

**语法**

**express.Connect**

\*express是一个代表 COMAddIn 对象的变量。

**说明**

如果该加载项是活动的，那么 **Connect** 属性返回 **True** ，反之返回 **False** 。活动加载项是经过注册并建立了连接的加载项；非活动加载项是经过注册但当前没有连接的加载项。

**示例**

``` JavaScript
//下面的示例显示一个消息框，并在其中指示第一个 COM 加载项是否已经注册以及当前是否已经连接。
function test()
{
    if(Application.COMAddIns.Item(1).Connect) 
    {
        alert("The add-in is connected.")
    }
    else 
    {
        alert("The add-in is not connected.")
    }
}
```

### COMAddIn.Creator

获取一个 32 位整数，指示创建 **COMAddIn** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 COMAddIn 对象的变量。

### COMAddIn.Description

获取或设置指定的 **COMAddin** 对象的说明性 **String** 值。可读/写。

**语法**

**express.Description**

\*express是一个代表 COMAddIn 对象的变量。

**示例**

``` JavaScript
//下面的示例显示用于绘图的 COM 加载项的说明文本。
alert("The description of this COMAddIn is \"" + Application.COMAddIns.Item("msodraa9.ShapeSelect").Description + "\"")
```

### COMAddIn.Guid

获取指定 **COMAddIn** 对象的类标识符 (CLSID)。只读。

**语法**

**express.Guid**

\*express是一个代表 COMAddIn 对象的变量。

**示例**

``` JavaScript
//下面的示例在消息框中显示 COMAddIns 集合中第一个 COM 加载项的 ProgID 和 CLSID。
alert("My ProgID is " + Application.COMAddIns.Item(1).ProgID + " and my CLSID is "+ Application.COMAddIns.Item(1).Guid)
```

### COMAddIn.Object

获取或设置一个对象引用。可读/写。

**语法**

**express.Object**

\*express是一个代表 COMAddIn 对象的变量。

**说明**

**Object** 属性是一个可读写属性，可存储任何对象引用。从这个角度来讲，它类似于某些 ActiveX 控件的通用 **Tag** 属性。

在某些情况下， **Object** 属性返回指定的 **COMAddIn** 对象所代表的对象。其他情况下，它默认返回 **Nothing** 。

**示例**

``` JavaScript
//下面的示例返回 COM 加载项 msodraa9.ShapeSelect 所代表的对象。
let objBaseObject = Application.COMAddIns.Item("msodraa9.ShapeSelect").Object
```

### COMAddIn.Parent

获取 **COMAddIn** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 COMAddIn 对象的变量。

### COMAddIn.ProgId

获取指定 **COMAddIn** 对象的编程标识符 (ProgID)。只读。

**语法**

**express.ProgId**

\*express是一个代表 COMAddIn 对象的变量。

**示例**

``` JavaScript
//下面的示例在消息框中显示 COMAddIns 集合中第一个 COM 加载项的 ProgID 和 CLSID。
alert("My ProgID is " + Application.COMAddIns.Item(1).ProgID + " and my CLSID is "+ Application.COMAddIns.Item(1).Guid)
```

> WPS 开发文档： [WPS 基础接口/通用 API 参考/COMAddIn](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# EffectParameters 对象

## 简介

代表 EffectParameter 对象的集合。

## 说明

图片效果被处理为由各个项构成的链，以便创建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。效果参数指定这些效果的属性。

下面的代码将设置 WPP 幻灯片中某个形状的几个图片效果填充属性。

``` JavaScript
function test(){
    // Setup a slide with one picture shape.
    let myEffects = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects

    // Insert a 150% Saturation effect.
    myEffects.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5

    // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.
    let brightnessContrast = myEffects.Insert(Application.Enum.msoEffectBrightnessContrast)

    brightnessContrast.EffectParameters.Item(1).Value = -0.5
    brightnessContrast.EffectParameters.Item(2).Value = 0.25

    // Remove all Picture effects.
    while(myEffects.Count > 0){
        myEffects.Delete(1)
    }
}
```

## 属性

| 序号 | 名称                                         | 说明                                                                             |
|:----:|:---------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [Application](#EffectParameters.Application) | 获取一个 Application 对象，该对象代表 EffectParameters 对象的容器应用程序。只读  |
|  2   | [Count](#EffectParameters.Count)             | 检索包含在 EffectParameters 集合中的 EffectParameter 对象数的计数。只读          |
|  3   | [Creator](#EffectParameters.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 EffectParameters 对象的应用程序。只读 |
|  4   | [Item](#EffectParameters.Item)               | 在指定的索引处或通过指定的唯一 ID 来检索 EffectParameter 对象。只读              |

## 成员属性

### EffectParameters.Application

获取一个 **Application** 对象，该对象代表 **EffectParameters** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 EffectParameters 对象的变量。

### EffectParameters.Count

检索包含在 **EffectParameters** 集合中的 **EffectParameter** 对象数的计数。只读

**语法**

**express.Count**

\*express是一个代表 EffectParameters 对象的变量。

### EffectParameters.Creator

获取一个 32 位整数，该整数指示在其中创建了 **EffectParameters** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 EffectParameters 对象的变量。

### EffectParameters.Item

在指定的索引处或通过指定的唯一 ID 来检索 **EffectParameter** 对象。只读

**语法**

**express.Item**

\*express是一个代表 EffectParameters 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/EffectParameters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartArt 对象

## 简介

用于与 SmartArt 图形交互的顶级类。

## 说明

``` JavaScript
/*下面的代码将顶级节点添加到带有项目符号的文本窗格中。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Add()
```

## 方法

| 序号 | 名称                     | 说明                               |
|:----:|:-------------------------|:-----------------------------------|
|  1   | [Reset](#SmartArt.Reset) | 将 SmartArt 图形重置为其原始状态。 |

## 属性

| 序号 | 名称                                 | 说明                                                                                                   |
|:----:|:-------------------------------------|:-------------------------------------------------------------------------------------------------------|
|  1   | [AllNodes](#SmartArt.AllNodes)       | 检索包含 SmartArt 图表内部所有节点的 SmartArtNodes 对象。只读。                                        |
|  2   | [Application](#SmartArt.Application) | 获取一个 Application 对象，该对象代表 SmartArt 对象的容器应用程序。只读。                              |
|  3   | [Color](#SmartArt.Color)             | 检索或设置应用到 Smart Art 图形的 Smart Art 颜色样式。可读/写。                                        |
|  4   | [Creator](#SmartArt.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArt 对象的应用程序。只读。                             |
|  5   | [Layout](#SmartArt.Layout)           | 检索或设置与 Smart Art 图形关联的 Smart Art 布局。可读/写。                                            |
|  6   | [Nodes](#SmartArt.Nodes)             | 检索 SmartArt 图表中根节点的子级。只读。                                                               |
|  7   | [Parent](#SmartArt.Parent)           | 返回调用对象。只读。                                                                                   |
|  8   | [QuickStyle](#SmartArt.QuickStyle)   | 检索或设置应用到 SmartArt 图形的 SmartArt 快速样式。可读/写。                                          |
|  9   | [Reverse](#SmartArt.Reverse)         | 获取或设置 SmartArt 图表与（从左到右）LTR 或（从右到左）RTL（如果该图表支持反向）有关的状态。可读/写。 |

## 成员方法

### SmartArt.Reset

将 SmartArt 图形重置为其原始状态。

**语法**

**express.Reset ()**

\*express是一个代表 SmartArt 对象的变量。

**说明**

此操作对应于 SmartArt 的“设计”选项卡下的“重设图形”按钮提供的功能。

## 成员属性

### SmartArt.AllNodes

检索包含 SmartArt 图表内部所有节点的 **SmartArtNodes** 对象。只读。

**语法**

**express.AllNodes**

\*express是一个代表 SmartArt 对象的变量。

**说明**

节点是按照顺序检索的，与数据模型无关。例如，下面的数据模型将按照 A、B、C、D、E、F 的顺序检索节点。

- A

- - B

  - - C

  - D

- - - E

- F

**示例**

``` JavaScript
/*下面的示例设置第一个节点内部的文本。*/Application.ActiveWorkbook.ActiveChart.Shapes.Item(1).SmartArt.AllNodes.Item(1).TextFrame2.TextRange.Text="Node 1"
```

### SmartArt.Application

获取一个 **Application** 对象，该对象代表 **SmartArt** 对象的容器应用程序。只读。

**语法**

**express.Application**

\*express是一个代表 SmartArt 对象的变量。

### SmartArt.Color

检索或设置应用到 Smart Art 图形的 Smart Art 颜色样式。可读/写。

**语法**

**express.Color**

\*express是一个代表 SmartArt 对象的变量。

**示例**

``` JavaScript
/*下面的代码设置 Smart Art 图表的配色方案。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Color = Application.SmartArtColors.Item(1)
```

### SmartArt.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArt** 对象的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SmartArt 对象的变量。

### SmartArt.Layout

检索或设置与 Smart Art 图形关联的 Smart Art 布局。可读/写。

**语法**

**express.Layout**

\*express是一个代表 SmartArt 对象的变量。

**示例**

``` JavaScript
/*下面的代码设置 Smart Art 布局。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Layout = Application.SmartArtLayouts.Item(1)
```

### SmartArt.Nodes

检索 SmartArt 图表中根节点的子级。只读。

**语法**

**express.Nodes**

\*express是一个代表 SmartArt 对象的变量。

**说明**

根节点没有父节点并且仅包含子级（如果 SmartArt 图形的数据模型内有子级存在）。在下面的示例中，将返回节点 A 和 F。

- A

- - B

  - - C

  - D

- - - E

- F

**示例**

``` JavaScript
/*下面的代码在 WPP 中添加一个顶级节点。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Add()
```

### SmartArt.Parent

返回调用对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SmartArt 对象的变量。

### SmartArt.QuickStyle

检索或设置应用到 SmartArt 图形的 SmartArt 快速样式。可读/写。

**语法**

**express.QuickStyle**

\*express是一个代表 SmartArt 对象的变量。

**示例**

------------------------------------------------------------------------

``` JavaScript
/*下面的代码更改 WPP 中 Smart Art 的快速样式。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(1)
```

### SmartArt.Reverse

获取或设置 SmartArt 图表与（从左到右）LTR 或（从右到左）RTL（如果该图表支持反向）有关的状态。可读/写。

**语法**

**express.Reverse**

\*express是一个代表 SmartArt 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArt](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Sync 对象

## 简介

WPS 中的 Document 对象、ET 中的 Workbook 对象以及 WPP 中的 Presentation 对象的 Sync 属性均可返回一个 Sync 对象。

| 注释                                                        |
|-------------------------------------------------------------|
| 自 WPS Office 2015 开始，此对象或成员已弃用，不应再行使用。 |

## 说明

使用 Sync 对象可管理存储在 SharePoint网站中的共享文档的本地副本和服务器副本之间的同步。 Status 属性返回有关同步当前状态的重要信息。使用 GetUpdate 方法可刷新同步状态。使用 LastSyncTime 、 ErrorType 和 WorkspaceLastChangedBy 属性可返回其他信息。

请参阅 Status 属性以获取有关共享文档的本地副本和服务器副本之间的差异和冲突的其他信息。

使用 PutUpdate 方法可将本地更改保存到服务器中。在没有进行本地更改时，关闭并重新打开文档可检索来自服务器的最新版本。使用 ResolveConflict 方法可解决本地副本和服务器副本之间的差异，或者使用 OpenVersion 方法可在当前打开的文档的本地版本旁打开另一版本。

由于 Sync 对象的 GetUpdate 、 PutUpdate 和 ResolveConflict 方法以异步方式完成任务，因此它们不返回状态代码。 Sync 对象通过单个事件提供重要的状态信息，开发人员可以通过以下应用程序特定事件来访问这些信息：

- 在 WPS 中，通过 Document 对象的 Sync 事件或 Application 对象的 DocumentSync 事件；
- 在 ET 中，通过 Workbook 对象的 Sync 事件或 Application 对象的 WorkbookSync 事件；
- 在 WPP 中，通过 Application 对象的 PresentationSync 事件。

上述 Sync 事件返回一个 msoSyncEventType 值。

无论活动文档是启用还是禁用了共享和同步， Sync 对象模型都可用。当活动文档未共享或没有启用同步时， Document 、 Workbook 和 Presentation 对象的 Sync 属性不返回 Nothing 。使用 Status 属性可确定文档是否共享以及是否启用了同步。

并不是所有文档同步问题都会引发可捕获的运行时错误。在使用了 Sync 对象的方法后，最好对 Status 属性进行检查。如果 Status 属性是 msoSyncStatusError ，则应检查 ErrorType 属性，以获取有关已发生的错误类型的其他信息。

在许多情况下，解决错误条件的最好方法是调用 GetUpdate 方法。例如，如果对 PutUpdate 的调用导致错误条件的产生，则对 GetUpdate 的调用会将状态重置为 msoSyncStatusLocalChanges 。

``` JavaScript
/*下面的示例将基于活动文档的状态说明 Sync 对象的多种方法。*/
function test(){
    var objSync = Application.ActiveWorkbook.Sync;
    if(objSync.Status > msoSyncEventDownloadFailed)
    {
        switch (objSync.Status)
        {
            case msoSyncStatusConflict:
                objSync.ResolveConflict(msoSyncConflictMerge);
                ActiveDocument.Save();
                objSync.ResolveConflict(msoSyncConflictClientWins);
                strStatus = "Conflict resolved by merging changes.";
                break;
            case msoSyncStatusError:
                strStatus = "Last error type: " + objSync.ErrorType;
                break;
            case msoSyncStatusLatest:
                strStatus = "Document copies already in sync.";
                break;
            case msoSyncStatusLocalChanges:
                objSync.PutUpdate();
                strStatus = "Local changes saved to server.";
                break;
            case msoSyncStatusNewerAvailable:
                objSync.GetUpdate();
                strStatus = "Local copy updated from server.";
                break;
            case msoSyncStatusSuspended:
                objSync.Unsuspend();
                strStatus = "Synchronization resumed.";
                break;
        }
    }
    else
        strStatus = "Not a shared workspace document.";
    alert(strStatus + " -- Sync Information");
}
```

## 方法

| 序号 | 名称                                     | 说明                                               |
|:----:|:-----------------------------------------|:---------------------------------------------------|
|  1   | [OpenVersion](#Sync.OpenVersion)         | 在当前打开的本地版本旁打开共享文档的另一个版本。   |
|  2   | [PutUpdate](#Sync.PutUpdate)             | 用共享文档的本地副本更新其服务器副本。             |
|  3   | [ResolveConflict](#Sync.ResolveConflict) | 解决共享文档的本地副本和服务器副本之间的冲突。     |
|  4   | [Unsuspend](#Sync.Unsuspend)             | 继续执行共享文档的本地副本和服务器副本之间的同步。 |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                        |
|:----:|:-------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Sync.Application)                       | 获取一个 Application 对象，代表 Sync 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Creator](#Sync.Creator)                               | 获取一个 32 位整数，指示创建 Sync 对象时所使用的应用程序。只读。                                                            |
|  3   | [ErrorType](#Sync.ErrorType)                           | 获取指示最新文档同步错误类型的 MsoSyncErrorType 常量。只读。                                                                |
|  4   | [LastSyncTime](#Sync.LastSyncTime)                     | 获取上一次将活动文档的本地副本与服务器副本同步的日期和时间。只读。                                                          |
|  5   | [Parent](#Sync.Parent)                                 | 获取 Sync 对象的 Parent 对象。只读。                                                                                        |
|  6   | [Status](#Sync.Status)                                 | 获取活动文档的本地副本与服务器副本的同步状态。只读。                                                                        |
|  7   | [WorkspaceLastChangedBy](#Sync.WorkspaceLastChangedBy) | 显示最后一个在共享文档服务器副本中保存更改的用户的显示名称。只读。                                                          |

## 成员方法

### Sync.OpenVersion

在当前打开的本地版本旁打开共享文档的另一个版本。

**语法**

**express.OpenVersion ( *SyncVersionType* )**

\*express是一个代表 Sync 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型           | 说明             |
|-------------------|-----------|--------------------|------------------|
| *SyncVersionType* | 必选      | MsoSyncVersionType | 代表版本的类型。 |

**说明**

使用 **OpenVersion** 方法可在当前打开的本地版本旁打开上次查看的共享文档版本 ( **msoSyncVersionLastViewed** ) 或其服务器副本 ( **msoSyncVersionServer** )。

**The msoSyncVersionLastViewed** 选项显示用户用服务器副本覆盖本地副本时创建的文档副本。例如，如果使用 **msoSyncConflictServerWins** 选项调用 **ResolveConflict** 方法，则将保存您的本地更改，并可以通过调用 **OpenVersion(msoSyncVersionLastViewed)** 进行查看。

并非所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一个操作后，最好对 **Status** 属性进行检查；如果 **Status** 属性是 **msoSyncStatusError** ，则应检查 **ErrorType** 属性，以获取有关已发生的错误类型的其他信息。

**示例**

``` JavaScript
/*下面的示例将在当前打开的共享文档本地版本旁打开该文档的服务器副本。*/
function test(){
    var objSync = Application.ActiveWorkbook.Sync; 
    objSync.OpenVersion(Application.Enum.msoSyncVersionServer);
}
```

### Sync.PutUpdate

用共享文档的本地副本更新其服务器副本。

**语法**

**express.PutUpdate ()**

\*express是一个代表 Sync 对象的变量。

**说明**

如果客户端未识别出最近对共享文档服务器副本所做的更改， **PutUpdate** 方法可能会遇到冲突。在调用 **PutUpdate** 之前先调用 **GetUpdate** 方法可刷新服务器副本的状态并检测可能的冲突。

如果本地文档有未保存的更改， **PutUpdate** 方法将发生运行时错误。

并非所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一个操作后，最好对 **Status** 属性进行检查；如果 **Status** 属性是 **msoSyncStatusError** ，则应检查 **ErrorType** 属性，以获取有关已发生的错误类型的其他信息。

在许多情况下，解决错误条件的最好方法是调用 **GetUpdate** 方法。例如，如果对 **PutUpdate** 的调用导致错误条件的产生，则对 **GetUpdate** 的调用会将状态重置为 **msoSyncStatusLocalChanges** 。

**示例**

``` JavaScript
/*如果文档的本地副本已被编辑，则下面的示例会使用 PutUpdate 方法根据本地副本更新服务器副本。*/
function test(){
    var objSync = Application.ActiveWorkbook.Sync;
    if(objSync.Status == Application.Enum.msoSyncStatusLocalChanges) {
        objSync.PutUpdate();
        strStatus = "Local changes saved to server.";
        alert(strStatus + "Sync Information");
    }
}
```

### Sync.ResolveConflict

解决共享文档的本地副本和服务器副本之间的冲突。

**语法**

**express.ResolveConflict ( *SyncConflictResolution* )**

\*express是一个代表 Sync 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型                      | 说明                 |
|--------------------------|-----------|-------------------------------|----------------------|
| *SyncConflictResolution* | 必选      | MsoSyncConflictResolutionType | 指定解决冲突的方式。 |

**说明**

使用 **ResolveConflict** 方法可解决活动文档的本地副本和服务器副本之间的差异。使用 **msoSyncConflictMerge** 选项（不适用于 ET 工作簿）可以合并各个文档的更改。通过 **msoSyncConflictClientWins** 选项使用本地更改替换服务器副本，或者通过 **msoSyncConflictServerWins** 选项使用更改的服务器副本替换本地副本。

**msoSyncConflictMerge** 选项将对服务器副本所做的更改合并到本地副本中，但实际上并没有解决冲突。要解决冲突并成功地合并更改，必须在合并更改后保存活动文档，然后再次使用 **msoSyncConflictClientWins** 选项调用 **ResolveConflict** 方法。

如果客户端未检测出对共享文档服务器副本的最新更改， **ResolveConflict** 方法可能会遇到冲突条件。在调用 **ResolveConflict** 之前先调用 **GetUpdate** 方法可刷新服务器副本的状态并检测可能的冲突。

如果本地文档有未保存的更改或者在文档的两个副本之间不存在冲突， **ResolveConflict** 方法就会发生运行时错误。

并不是所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一个操作后，最好对 **ErrorType** 属性进行检查。如果 **Status** 属性是 **msoSyncStatusError** ，则应检查 **Status** 属性，以进一步获取有关所发生错误类型的信息。

**示例**

``` JavaScript
/*下面的示例试图通过合并更改来解决活动文档的本地副本和服务器副本之间的冲突。*/
function test(){
    var objSync = Application.ActiveWorkbook.Sync;
    if(objSync.Status == Application.Enum.msoSyncStatusConflict) {
        objSync.ResolveConflict(Application.Enum.msoSyncConflictMerge);
        Application.ActiveWorkbook.Save();
        objSync.ResolveConflict(Application.Enum.msoSyncConflictClientWins);
        strStatus = "Conflict resolved by merging changes.";
        alert(strStatus + "Sync Information");
    }
}
```

### Sync.Unsuspend

继续执行共享文档的本地副本和服务器副本之间的同步。

**语法**

**express.Unsuspend ()**

\*express是一个代表 Sync 对象的变量。

**说明**

使用 **Unsuspend** 方法可在 **Status** 属性返回 **msoSyncStatusSuspended** 时继续执行文档同步。

并非所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一个操作后，最好对 **Status** 属性进行检查；如果 **Status** 属性是 **msoSyncStatusError** ，则应检查 **ErrorType** 属性，以获取有关已发生的错误类型的其他信息。

**示例**

``` JavaScript
/*下面的示例将在文档同步暂停后继续执行文档同步。*/
function test(){
    var objSync = Application.ActiveWorkbook.Sync;
    if(objSync.Status == Application.Enum.msoSyncStatusSuspended) {
        objSync.Unsuspend();
        alert(strStatus + "Sync Information");
    }
}
```

## 成员属性

### Sync.Application

获取一个 **Application** 对象，代表 **Sync** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 Sync 对象的变量。

### Sync.Creator

获取一个 32 位整数，指示创建 **Sync** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 Sync 对象的变量。

### Sync.ErrorType

获取指示最新文档同步错误类型的 **MsoSyncErrorType** 常量。只读。

**语法**

**express.ErrorType**

\*express是一个代表 Sync 对象的变量。

**说明**

使用 **ErrorType** 属性确定最新文档同步错误的类型。并不是所有文档同步问题都会引发可捕获的运行时错误。在使用 **Sync** 对象执行完一项操作后，最好对 **Status** 属性进行检查；如果 **Status** 属性是 **msoSyncStatusError** ，则应检查 **ErrorType** 属性，以获取所发生错误类型的其他信息。

### Sync.LastSyncTime

获取上一次将活动文档的本地副本与服务器副本同步的日期和时间。只读。

**语法**

**express.LastSyncTime**

\*express是一个代表 Sync 对象的变量。

**说明**

使用 **LastSyncTime** 属性可确定距活动文档的本地副本上一次与服务器副本同步已有多长时间。检查 **Status** 属性可确定本地副本与服务器副本是否同步。

如果未将活动文档配置为在本地副本与服务器副本之间进行同步， **LastSyncTime** 属性就会产生运行时错误。

**示例**

### Sync.Parent

获取 **Sync** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Sync 对象的变量。

### Sync.Status

获取活动文档的本地副本与服务器副本的同步状态。只读。

**语法**

**express.Status**

\*express是一个代表 Sync 对象的变量。

**说明**

使用 **Status** 属性可确定活动文档的本地副本是否与共享服务器副本同步。使用 **GetUpdate** 方法可刷新状态。对于不同的状态条件，请分别使用下列方法和属性：

- **msoSyncStatusConflict** - 当本地副本和服务器副本都发生更改时为 **True** 。使用 **ResolveConflict** 方法可解决这些差异。
- **msoSyncStatusError** - 检查 **ErrorType** 属性。
- **msoSyncStatusLocalChanges** - 当只有本地副本发生更改时为 **True** 。使用 **PutUpdate** 方法可将本地更改保存到服务器副本中。
- **msoSyncStatusNewerAvailable** - 当只有服务器副本发生更改时为 **True** 。关闭并重新打开文档可使用来自服务器中的最新副本。
- **msoSyncStatusSuspended** - 使用 **Unsuspend** 方法可恢复同步。

**Status** 属性按照下面的优先级顺序从列表中返回单个常量：

1.  **msoSyncStatusNoSharedWorkspace**
2.  **msoSyncStatusError**
3.  **msoSyncStatusSuspended**
4.  **msoSyncStatusConflict**
5.  **msoSyncStatusNewerAvailable**
6.  **msoSyncStatusLocalChanges**
7.  **msoSyncStatusLatest**

**示例**

``` JavaScript
/*下面的示例将检查 Status 属性，并在必要时进行适当的操作以使文档的本地副本与服务器副本同步。*/
function test(){
    var objSync = Application.ActiveWorkbook.Sync;
    if(objSync.Status > Application.Enum.msoSyncEventDownloadFailed)
    {
        switch (objSync.Status)
        {
            case Application.Enum.msoSyncStatusConflict:
                objSync.ResolveConflict(Application.Enum.msoSyncConflictMerge);
                　　　　　　　　 Application.ActiveDocument.Save();
                　　　　　　　　 objSync.ResolveConflict(Application.Enum.msoSyncConflictClientWins);
                　　　　　　　　 strStatus = "Conflict resolved by merging changes.";
                break;
            case Application.Enum.msoSyncStatusError:
                strStatus = "Last error type: " + objSync.ErrorType;
               　　　　　　  break;
            case Application.Enum.msoSyncStatusLatest:
                strStatus = "Document copies already in sync.";
                　　　　　　　　 break;
            case Application.Enum.msoSyncStatusLocalChanges:
                objSync.PutUpdate();
                　　　　　　　　 strStatus = "Local changes saved to server.";
                　　　　　　　　 break;
            case Application.Enum.msoSyncStatusNewerAvailable:
                objSync.GetUpdate();
                　　　　　　　　 strStatus = "Local copy updated from server.";
                　　　　　　　　 break;
            case Application.Enum.msoSyncStatusSuspended:
                objSync.Unsuspend();
                　　　　　　　　 strStatus = "Synchronization resumed.";
                　　　　　　　　 break;
        }
    }
    else
    　　　　strStatus = "Not a shared workspace document.";
    　　　　alert(strStatus + "Sync Information");
}
```

### Sync.WorkspaceLastChangedBy

显示最后一个在共享文档服务器副本中保存更改的用户的显示名称。只读。

**语法**

**express.WorkspaceLastChangedBy**

\*express是一个代表 Sync 对象的变量。

**说明**

如果未将活动文档配置为在本地副本和服务器副本之间进行同步， **WorkspaceLastChangedBy** 属性就会产生运行时错误。

**示例**

``` JavaScript
/*下面的示例将检查共享文档的本地副本和服务器副本之间的冲突，并报告最后一个在服务器副本中保存更改的用户的名称。*/
function test(){
    var objSync = Application.ActiveWorkbook.Sync;
    if(objSync.Status == Application.Enum.msoSyncStatusConflict) {
        strStatus = "The server copy has been changed by" + objSync.WorkspaceLastChangedBy;
        alert(strStatus + "Server Copy Changed");
    }
}
```

> WPS 开发文档： [WPS 基础接口/通用 API 参考/Sync](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# WorkflowTask 对象

## 简介

代表 WorkflowTasks 集合中的单个工作流任务。

## 说明

``` JavaScript
/*以下示例显示当前文档中每个工作流任务的名称，然后显示特定任务的工作流任务编辑用户界面。应该注意到，调用 GetWorkflowTasks 方法涉及到服务器的往返行程。*/
function test(){
    const objWorkflowTasks = Application.ActiveWorkbook.GetWorkflowTasks();
    for(let i = 1; i <= objWorkflowTasks.Count; i++)
        Debug.Print(objWorkflowTasks.Item(i).Name);
        
    var objWorkflowTask = objWorkflowTasks.Item(1);
    objWorkflowTask.Show();
}
```

## 方法

| 序号 | 名称                       | 说明                                                 |
|:----:|:---------------------------|:-----------------------------------------------------|
|  1   | [Show](#WorkflowTask.Show) | 显示指定 WorkflowTask 对象的工作流任务编辑用户界面。 |

## 属性

| 序号 | 名称                                     | 说明                                                                         |
|:----:|:-----------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#WorkflowTask.Application) | 获取一个代表 WorkflowTemplate 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [AssignedTo](#WorkflowTask.AssignedTo)   | 获取向其分配了工作流任务的人员的姓名。只读。                                 |
|  3   | [CreatedBy](#WorkflowTask.CreatedBy)     | 获取创建工作流任务的人员的姓名。只读                                         |
|  4   | [CreatedDate](#WorkflowTask.CreatedDate) | 获取工作流任务的创建日期。只读。                                             |
|  5   | [Creator](#WorkflowTask.Creator)         | 获取一个 32 位整数，指示创建 WorkflowTemplate 对象时所使用的应用程序。只读。 |
|  6   | [Description](#WorkflowTask.Description) | 获取工作流模板的说明。只读。                                                 |
|  7   | [DueDate](#WorkflowTask.DueDate)         | 获取工作流任务的截止日期。只读                                               |
|  8   | [Id](#WorkflowTask.Id)                   | 获取用于创建工作流实例的模板的 ID。只读。                                    |
|  9   | [ListID](#WorkflowTask.ListID)           | 获取包含工作流任务的列表的 ID。只读。                                        |
|  10  | [Name](#WorkflowTask.Name)               | 获取 WorkflowTemplate 对象的名称。只读。                                     |
|  11  | [WorkflowID](#WorkflowTask.WorkflowID)   | 获取与工作流任务关联的工作流的 ID。只读。                                    |

## 成员方法

### WorkflowTask.Show

显示指定 **WorkflowTask** 对象的工作流任务编辑用户界面。

**语法**

**express.Show ()**

\*express是一个代表 WorkflowTask 对象的变量。

**返回值**

Integer

**示例**

``` JavaScript
/*以下示例显示当前文档中每个工作流任务的名称，然后显示特定任务的工作流任务编辑用户界面。*/
function test(){
    const objWorkflowTasks = Application.ActiveWorkbook.GetWorkflowTasks();
    for(let i = 1; i <= objWorkflowTasks.Count; i++)
        Debug.Print(objWorkflowTasks.Item(i).Name);

    var objWorkflowTask = objWorkflowTasks.Item(1);
    objWorkflowTask.Show();
}
```

## 成员属性

### WorkflowTask.Application

获取一个代表 **WorkflowTemplate** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 WorkflowTask 对象的变量。

**说明**

### WorkflowTask.AssignedTo

获取向其分配了工作流任务的人员的姓名。只读。

**语法**

**express.AssignedTo**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.CreatedBy

获取创建工作流任务的人员的姓名。只读

**语法**

**express.CreatedBy**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.CreatedDate

获取工作流任务的创建日期。只读。

**语法**

**express.CreatedDate**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.Creator

获取一个 32 位整数，指示创建 **WorkflowTemplate** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.Description

获取工作流模板的说明。只读。

**语法**

**express.Description**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.DueDate

获取工作流任务的截止日期。只读

**语法**

**express.DueDate**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.Id

获取用于创建工作流实例的模板的 ID。只读。

**语法**

**express.Id**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.ListID

获取包含工作流任务的列表的 ID。只读。

**语法**

**express.ListID**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.Name

获取 **WorkflowTemplate** 对象的名称。只读。

**语法**

**express.Name**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.WorkflowID

获取与工作流任务关联的工作流的 ID。只读。

**语法**

**express.WorkflowID**

\*express是一个代表 WorkflowTask 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/WorkflowTask](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
