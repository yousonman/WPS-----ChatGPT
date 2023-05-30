# TextFrame 对象

## 简介

代表 Shape 对象中的文本框架。包含文本框架中的文本以及控制文本框架的对齐和定位的属性和方法。

## 说明

使用 TextFrame 属性可返回一个 TextFrame 对象。

下例向 *myDocument* 中添加一个矩形，向矩形中添加文本，然后设置文本框的边距。

``` JavaScript
let myDocument = Application.Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 10
rng.MarginLeft = 10
rng.MarginRight = 10
rng.MarginTop = 10
```

## 方法

| 序号 | 名称                                | 说明                                                                                                                                |
|:----:|:------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Characters](#TextFrame.Characters) | 返回一个 Ch aracters 对象，该对象表示某个形状的文本框架中的字符区域。可以使用 Characters 对象向文本框架中添加字符和设置字符的格式。 |

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
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.AutoMargins">AutoMargins</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 ET 是否自动计算文本框边距。可读/写</p>
<p>返回Boolean值</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.AutoSize">AutoSize</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定对象能自动调整大小，以适应其中所包含的文字，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.HorizontalAlignment">HorizontalAlignment</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 XlHAlign 值，它代表指定对象的水平对齐方式。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.HorizontalOverflow">HorizontalOverflow</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定对象的水平溢出设置。可读/写</p>
<p>返回 XlOartHorizontalOverflow 值</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.MarginBottom">MarginBottom</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读/写。 Single 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.MarginLeft">MarginLeft</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。 Single 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.MarginRight">MarginRight</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。 Single 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.MarginTop">MarginTop</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。 Single 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.Orientation">Orientation</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Long 值，它代表文本框的方向。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.ReadingOrder">ReadingOrder</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定对象的阅读次序。可为以下常量之一： xlRTL （从右到左）、 xlLTR （从左到右）或 xlContext 。 Long 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.VerticalAlignment">VerticalAlignment</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 XlVAlign 值，它代表指定对象的垂直对齐方式。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.VerticalOverflow">VerticalOverflow</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定对象的垂直溢出设置。可读/写</p>
<p>返回 XlOartVerticalOverflow 值</p></td>
</tr>
</tbody>
</table>

## 成员方法

### TextFrame.Characters

返回一个 **Ch aracters** 对象，该对象表示某个形状的文本框架中的字符区域。可以使用 **Characters** 对象向文本框架中添加字符和设置字符的格式。

**语法**

**express.Characters ( *Start* , *Length* )**

\*express是一个代表 TextFrame 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                    |
|----------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------|
| *Start*  | 可选      | Variant  | 要返回的第一个字符。如果此参数设置为 1 或被省略，则 Characters 方法将返回以第一个字符为起始字符的字符区域。             |
| *Length* | 可选      | Variant  | 要返回的字符个数。如果省略此参数，则 Characters 方法将返回该字符串的剩余部分（由 Start 参数设置的字符以后的所有字符）。 |

**说明**

**Characters** 对象不是集合。

**示例**

``` JavaScript
/*本示例将活动工作表中第一个形状的文本框架中第三个字符的格式设置为加粗。*/
let rng = Application.ActiveSheet.Shapes.Item(1).TextFrame
rng.Characters().Text = "abcdefg"
rng.Characters(3, 1).Font.Bold = true
```

## 成员属性

### TextFrame.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
let myObject = ActiveWorkbook
if(myObject.Application.Value == "ET"){
    MsgBox("This is an ET Application object.")
}
else{
    MsgBox("This is not an ET Application object.")
}
```

### TextFrame.AutoMargins

返回或设置 ET 是否自动计算文本框边距。可读/写

返回Boolean值

**语法**

**express.AutoMargins**

\*express是一个代表 TextFrame 对象的变量。

**说明**

如果 ET 自动计算文本框边距，则为 **True** ；否则为 **False** 。当此属性为 **True** 时， **MarginLeft** , **MarginRight** 、 **MarginTop** 和 **MarginBottom** 属性将被忽略。

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

### TextFrame.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TextFrame 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### TextFrame.HorizontalAlignment

返回或设置一个 **XlHAlign** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 TextFrame 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### TextFrame.HorizontalOverflow

返回或设置指定对象的水平溢出设置。可读/写

返回 **XlOartHorizontalOverflow 值**

**语法**

**express.HorizontalOverflow**

\*express是一个代表 TextFrame 对象的变量。

**说明**

此属性只在 **Word Wrap** 属性为 **msoFalse** (0) 时有效。

### TextFrame.MarginBottom

以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读/写。 **Single** 类型。

**语法**

**express.MarginBottom**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myDocument = Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 0
rng.MarginLeft = 100
rng.MarginRight = 0
rng.MarginTop = 20
```

### TextFrame.MarginLeft

以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。 **Single** 类型。

**语法**

**express.MarginLeft**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myDocument = Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 0
rng.MarginLeft = 100
rng.MarginRight = 0
rng.MarginTop = 20
```

### TextFrame.MarginRight

以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。 **Single** 类型。

**语法**

**express.MarginRight**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myDocument = Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 0
rng.MarginLeft = 100
rng.MarginRight = 0
rng.MarginTop = 20
```

### TextFrame.MarginTop

以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。 **Single** 类型。

**语法**

**express.MarginTop**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myDocument = Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 0
rng.MarginLeft = 100
rng.MarginRight = 0
rng.MarginTop = 20
```

### TextFrame.Orientation

返回或设置一个 **Long** 值，它代表文本框的方向。

**语法**

**express.Orientation**

\*express是一个代表 TextFrame 对象的变量。

**说明**

此属性值可设为 -90 到 90 度之间的整数值或以下 **<span "="" r="" rlidawscontentredir?assetid="HV100730982052&amp;CTT=11&amp;Origin=HV103325142052&quot;" style="color: rgb(0, 0, 0); text-decoration-line: none;" target="_new">MsoTextOrientation</span>** 常量之一。

### TextFrame.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.ReadingOrder

返回或设置指定对象的阅读次序。可为以下常量之一： **xlRTL** （从右到左）、 **xlLTR** （从左到右）或 **xlContext** 。 **Long** 类型，可读写。

**语法**

**express.ReadingOrder**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.VerticalAlignment

返回或设置一个 **XlVAlign** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 TextFrame 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### TextFrame.VerticalOverflow

返回或设置指定对象的垂直溢出设置。可读/写

返回 **XlOartVerticalOverflow** 值

**语法**

**express.VerticalOverflow**

\*express是一个代表 TextFrame 对象的变量。

**说明**

此属性只在 **AutoS ize** 属性为 **False** 时有效。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/TextFrame](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
