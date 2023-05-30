# PlotArea 对象

## 简介

代表图表的绘图区。

## 说明

该区域为绘制图表数据的区域。二维图表中的绘图区包含数据标志、网格线、数据标签、趋势线和可选的置于图表区内的图表项。三维图表的绘图区中除包含上述各项外，还在图表中包含背景墙、基底、坐标轴、坐标轴标题和刻度线标签。绘图区被图表区所包围。二维图表的图表区包含坐标轴、图表标题、坐标轴标题和图例。三维图表的图表区包含图表标题和图例。有关设置图表区格式的详细信息，请参阅 ChartArea 对象。

## 方法

| 序号 | 名称                                   | 说明                 |
|:----:|:---------------------------------------|:---------------------|
|  1   | [ClearFormats](#PlotArea.ClearFormats) | 清除对象的格式设置。 |
|  2   | [Select](#PlotArea.Select)             | 选择对象。           |

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
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Format">Format</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 ChartFormat 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Height">Height</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.InsideHeight">InsideHeight</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回绘图区内部高度。可读写 Double 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.InsideLeft">InsideLeft</a></td>
<td style="text-align: left;" data-valian="middle"><p>以 <span>磅</span> 为单位返回从图表边界到绘图区内部左边界的距离。可读写 Double 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.InsideTop">InsideTop</a></td>
<td style="text-align: left;" data-valian="middle"><p>以 <span>磅</span> 为单位返回从图表边界到绘图区内部上边界的距离。可读写 Double 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.InsideWidth">InsideWidth</a></td>
<td style="text-align: left;" data-valian="middle"><p>以 <span>磅</span> 为单位返回绘图区内部宽度。可读写 Double 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Left">Left</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Name">Name</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 String 值，它代表对象的名称。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Position">Position</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置图表上绘制区域的位置。可读/写 <span>XlChartElementPosition</span> 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Top">Top</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Width">Width</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PlotArea.ClearFormats

清除对象的格式设置。

**语法**

**express.ClearFormats ()**

\*express是一个代表 PlotArea 对象的变量。

**返回值**

Variant

### PlotArea.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 PlotArea 对象的变量。

**返回值**

Varint

## 成员属性

### PlotArea.Application

如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PlotArea 对象的变量。

**示例**

``` JavaScript
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

### PlotArea.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PlotArea 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。Creator 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PlotArea.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Height

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.InsideHeight

以磅为单位返回绘图区内部高度。可读写 **Double** 类型。

**语法**

**express.InsideHeight**

\*express是一个代表 PlotArea 对象的变量。

**说明**

这种度量方式的图形区域不包含坐标轴标签。图形区域的 Height 属性使用包含坐标轴标签的封闭矩形。

**示例**

``` JavaScript
//本示例在 Chart1 中的绘图区内绘制带点线的矩形。
function test()
{
    let pa = Application.Charts.Item("chart1").PlotArea
    pa.Shapes.AddShape(msoShapeRectangle,pa.InsideLeft,pa.InsideTop,pa.InsideWidth,pa.InsideHeight)
    pa.Fill.Transparency = 1
    pa.Line.DashStyle = msoLineDashDot
}
```

### PlotArea.InsideLeft

以 磅 为单位返回从图表边界到绘图区内部左边界的距离。可读写 **Double** 类型。

**语法**

**express.InsideLeft**

\*express是一个代表 PlotArea 对象的变量。

**说明**

这种度量方式的图形区域不包含坐标轴标签。图形区域的 **Left** 属性使用包含坐标轴标签的封闭矩形。

**示例**

``` JavaScript
//本示例在 Chart1 中的绘图区内绘制带点线的矩形。
function test()
{
    let pa = Application.Charts.Item("chart1").PlotArea
    pa.Shapes.AddShape(msoShapeRectangle,pa.InsideLeft,pa.InsideTop,pa.InsideWidth,pa.InsideHeight)
    pa.Fill.Transparency = 1
    pa.Line.DashStyle = msoLineDashDot
}
```

### PlotArea.InsideTop

以 磅 为单位返回从图表边界到绘图区内部上边界的距离。可读写 **Double** 类型。

**语法**

**express.InsideTop**

\*express是一个代表 PlotArea 对象的变量。

**说明**

这种度量方式的图形区域不包含坐标轴标签。图形区域的 **Top** 属性使用包含坐标轴标签的封闭矩形。

**示例**

``` JavaScript
//本示例在 Chart1 中的绘图区内绘制带点线的矩形。
function test()
{
    let pa = Application.Charts.Item("chart1").PlotArea
    pa.Shapes.AddShape(msoShapeRectangle,pa.InsideLeft,pa.InsideTop,pa.InsideWidth,pa.InsideHeight)
    pa.Fill.Transparency = 1
    pa.Line.DashStyle = msoLineDashDot
}
```

### PlotArea.InsideWidth

以 磅 为单位返回绘图区内部宽度。可读写 **Double** 类型。

**语法**

**express.InsideWidth**

\*express是一个代表 PlotArea 对象的变量。

**说明**

这种度量方式的图形区域不包含坐标轴标签。图形区域的 **Width** 属性使用包含坐标轴标签的封闭矩形。

**示例**

``` JavaScript
//本示例在 Chart1 中的绘图区内绘制带点线的矩形。
function test()
{
    let pa = Application.Charts.Item("chart1").PlotArea
    pa.Shapes.AddShape(msoShapeRectangle,pa.InsideLeft,pa.InsideTop,pa.InsideWidth,pa.InsideHeight)
    pa.Fill.Transparency = 1
    pa.Line.DashStyle = msoLineDashDot
}
```

### PlotArea.Left

table { word-break:break-all; }

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Position

返回或设置图表上绘制区域的位置。可读/写 **XlChartElementPosition** 类型。

**语法**

**express.Position**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Width

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 PlotArea 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PlotArea](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
