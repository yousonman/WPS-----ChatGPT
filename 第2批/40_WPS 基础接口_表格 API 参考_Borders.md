# Borders 对象

## 简介

由四个 Border 对象组成的集合，它们分别代表 Range 或 Style 对象的四个边框。

## 方法

| 序号 | 名称                  | 说明                                                     |
|:----:|:----------------------|:---------------------------------------------------------|
|  1   | [Item](#Borders.Item) | 返回一个 Border 对象，它代表单元格区域或样式的边框之一。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Borders.Application)   | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Color](#Borders.Color)               | 返回或设置对象的主要颜色，如注释部分中的表格所示。 Variant 型，可读写。                                                                                                                                                         |
|  3   | [ColorIndex](#Borders.ColorIndex)     | 返回或设置一个 Variant 值，它代表全部四条边框的颜色。                                                                                                                                                                           |
|  4   | [Count](#Borders.Count)               | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  5   | [Creator](#Borders.Creator)           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  6   | [LineStyle](#Borders.LineStyle)       | 返回或设置边框的线型。 XlLineStyle 、 xlGray25 、 xlGray50 、 xlGray75 或 xlAutomatic 类型，可读写。                                                                                                                            |
|  7   | [Parent](#Borders.Parent)             | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  8   | [ThemeColor](#Borders.ThemeColor)     | 返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 Variant 类型。                                                                                                                                      |
|  9   | [TintAndShade](#Borders.TintAndShade) | 返回或设置一个 Single ，使颜色变深或变浅。                                                                                                                                                                                      |
|  10  | [Value](#Borders.Value)               | 等价于 Borders.LineStyle                                                                                                                                                                                                        |
|  11  | [Weight](#Borders.Weight)             | 返回或设置一个 XlBorderWeight 值，它代表边框的粗细。                                                                                                                                                                            |

## 成员方法

### Borders.Item

返回一个 **Border** 对象，它代表单元格区域或样式的边框之一。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Borders 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型       | 说明                        |
|---------|-----------|----------------|-----------------------------|
| *Index* | 必选      | XlBordersIndex | XlBordersIndex 的常量之一。 |

**说明**

|                                                           |
|-----------------------------------------------------------|
| **XlBordersIndex** 可为下列 **XlBordersIndex** 常量之一。 |
| **xlDiagonalDown**                                        |
| **xlDiagonalUp**                                          |
| **xlEdgeBottom**                                          |
| **xlEdgeLeft**                                            |
| **xlEdgeRight**                                           |
| **xlEdgeTop**                                             |
| **xlInsideHorizontal**                                    |
| **xlInsideVertical**                                      |

**示例**

``` JavaScript
/* 下例设置单元格区域 A1:G1 的底部边界的颜色*/
Application.Worksheets.Item("Sheet1").Range("a1:g1").Borders.Item(xlEdgeBottom).Color = (255, 0, 0)
```

## 成员属性

### Borders.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
function test() {
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### Borders.Color

返回或设置对象的主要颜色，如注释部分中的表格所示。 **Variant** 型，可读写。

**语法**

**express.Color**

\*express是一个代表 Borders 对象的变量。

**说明**

| 对象         | 对应颜色                                                                            |
|--------------|-------------------------------------------------------------------------------------|
| **边框**     | 边框的颜色。                                                                        |
| **Borders**  | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font**     | 字体的颜色。                                                                        |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。                                              |
| **Tab**      | 选项卡的颜色。                                                                      |

**示例**

``` JavaScript
/* 此示例对 Chart1 中数值坐标轴的刻度线标志颜色进行设置*/
Application.Charts.Item("Chart1").Axes(xlValue).TickLabels.Font.Color = (0, 255, 0)
```

### Borders.ColorIndex

返回或设置一个 **Variant** 值，它代表全部四条边框的颜色。

**语法**

**express.ColorIndex**

\*express是一个代表 Borders 对象的变量。

**说明**

如果全部四条边框不是同一种颜色，此属性返回 **Null** 。

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 **XlColorIndex** 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

### Borders.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Borders 对象的变量。

### Borders.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Borders 对象的变量。

### Borders.LineStyle

返回或设置边框的线型。 **XlLineStyle** 、 **xlGray25** 、 **xlGray50** 、 **xlGray75** 或 **xlAutomatic** 类型，可读写。

**语法**

**express.LineStyle**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/* 本示例为 Chart1 的图表区和绘图区域设置边框*/
function test() {
    let charts = Application.Charts.Item("Chart1")
    charts.ChartArea.Border.LineStyle = xlDashDot
    let border = charts.PlotArea.Border
    border.LineStyle = xlDashDotDot
    border.Weight = xlThick
}
```

### Borders.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Borders 对象的变量。

### Borders.ThemeColor

返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。

**语法**

**express.ThemeColor**

\*express是一个代表 Borders 对象的变量。

**说明**

如果对象当前未应用主题颜色，试图访问其主题颜色则会引起无效请求运行时错误。

### Borders.TintAndShade

返回或设置一个 **Single** ，使颜色变深或变浅。

**语法**

**express.TintAndShade**

\*express是一个代表 Borders 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

### Borders.Value

等价于 Borders.LineStyle

**语法**

**express.Value**

\*express是一个代表 Borders 对象的变量。

### Borders.Weight

返回或设置一个 **XlBorderWeight** 值，它代表边框的粗细。

**语法**

**express.Weight**

\*express是一个代表 Borders 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Borders](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
