# LineFormat 对象

## 简介

代表线条和箭头的格式设置。对于线条， LineFormat 对象包含该线条自身的格式信息；对于具有边框的形状，该对象包含形状的边框的格式信息。

## 说明

可以使用 Line 属性返回一个 LineFormat 对象。以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形。

``` JavaScript
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

## 方法

| 序号 | 名称                               | 说明                                                |
|:----:|:-----------------------------------|:----------------------------------------------------|
|  1   | [BreakLink](#LineFormat.BreakLink) | 断开源文件与指定 OLE 对象、图片或链接域之间的链接。 |
|  2   | [Update](#LineFormat.Update)       | 更新指定的链接格式。                                |

## 属性

| 序号 | 名称                                                     | 说明                                                                      |
|:----:|:---------------------------------------------------------|:--------------------------------------------------------------------------|
|  1   | [Application](#LineFormat.Application)                   | 返回一个代表 WPS 应用程序的 Application 对象。                            |
|  2   | [BackColor](#LineFormat.BackColor)                       | 返回或设置一个 [ColorFormat]() 对象，该对象代表图案线条的背景色。可读写。 |
|  3   | [BeginArrowheadLength](#LineFormat.BeginArrowheadLength) | 返回或设置指定线条开头的箭头长度。 MsoArrowheadLength 类型，可读写。      |
|  4   | [BeginArrowheadStyle](#LineFormat.BeginArrowheadStyle)   | 返回或设置指定线条开头的箭头样式。 MsoArrowheadStyle 类型，可读写。       |
|  5   | [BeginArrowheadWidth](#LineFormat.BeginArrowheadWidth)   | 返回或设置指定线条开头的箭头宽度。 MsoArrowheadWidth 类型，可读写。       |
|  6   | [DashStyle](#LineFormat.DashStyle)                       | 返回或设置指定线条的虚线样式。 MsoLineDashStyle 类型，可读写。            |
|  7   | [EndArrowheadLength](#LineFormat.EndArrowheadLength)     | 返回或设置指定线条末尾箭头的长度。 MsoArrowheadLength 类型，可读写。      |
|  8   | [EndArrowheadStyle](#LineFormat.EndArrowheadStyle)       | 返回或设置指定线条末尾箭头的样式。 MsoArrowheadStyle 类型，可读写。       |
|  9   | [EndArrowheadWidth](#LineFormat.EndArrowheadWidth)       | 返回或设置指定线条末尾箭头的宽度。 MsoArrowheadWidth 类型，可读写。       |
|  10  | [ForeColor](#LineFormat.ForeColor)                       | 返回或设置一个 [ColorFormat]() 对象，该对象代表线条的前景色。可读写。     |
|  11  | [Pattern](#LineFormat.Pattern)                           | 设置或返回代表应用于指定线条的图案的值。可读写 MsoPatternType 类型。      |
|  12  | [Style](#LineFormat.Style)                               | 返回或设置线条格式样式。可读写 MsoLineStyle 类型。                        |
|  13  | [Transparency](#LineFormat.Transparency)                 | 返回或设置线条的透明度。可读写 Number 类型。                              |
|  14  | [Weight](#LineFormat.Weight)                             | 返回或设置指定线条的粗细（以磅为单位）。 Number 类型，可读写。            |

## 成员方法

### LineFormat.BreakLink

断开源文件与指定 OLE 对象、图片或链接域之间的链接。

**语法**

**express.BreakLink ()**

\*express是一个代表 LineFormat 对象的变量。

**说明**

使用该方法后，在源文件发生变化的情况下，链接结果不会自动更新。

**示例**

``` JavaScript
/*本示例更新活动文档中所有的 OLE 对象，然后断开链接。*/
function test()
{
let myShapes = ActiveDocument.Shapes

for(let i = 1; i <= myShapes.Count; i++){
    if(myShapes.Item(i).Type == msoLinkedOLEObject){
        myShapes.Item(i).Update()
        myShapes.Item(i).LinkFormat.BreakLink()
    }
}
}
```

### LineFormat.Update

更新指定的链接格式。

**语法**

**express.Update ()**

\*express是一个代表 LineFormat 对象的变量。

## 成员属性

### LineFormat.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.BackColor

返回或设置一个 **[ColorFormat]()** 对象，该对象代表图案线条的背景色。可读写。

**语法**

**express.BackColor**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例为活动文档添加一条带图案线条 */
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.BackColor.RGB = 0x00ff00;
    lnFormat.Pattern = msoPatternDarkDownwardDiagonal;
}
```

### LineFormat.BeginArrowheadLength

返回或设置指定线条开头的箭头长度。 **MsoArrowheadLength** 类型，可读写。

**语法**

**express.BeginArrowheadLength**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.BeginArrowheadStyle

返回或设置指定线条开头的箭头样式。 **MsoArrowheadStyle** 类型，可读写。

**语法**

**express.BeginArrowheadStyle**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.BeginArrowheadWidth

返回或设置指定线条开头的箭头宽度。 **MsoArrowheadWidth** 类型，可读写。

**语法**

**express.BeginArrowheadWidth**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.DashStyle

返回或设置指定线条的虚线样式。 **MsoLineDashStyle** 类型，可读写。

**语法**

**express.DashStyle**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.EndArrowheadLength

返回或设置指定线条末尾箭头的长度。 **MsoArrowheadLength** 类型，可读写。

**语法**

**express.EndArrowheadLength**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.EndArrowheadStyle

返回或设置指定线条末尾箭头的样式。 **MsoArrowheadStyle** 类型，可读写。

**语法**

**express.EndArrowheadStyle**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.EndArrowheadWidth

返回或设置指定线条末尾箭头的宽度。 **MsoArrowheadWidth** 类型，可读写。

**语法**

**express.EndArrowheadWidth**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.ForeColor

返回或设置一个 **[ColorFormat]()** 对象，该对象代表线条的前景色。可读写。

**语法**

**express.ForeColor**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Pattern

设置或返回代表应用于指定线条的图案的值。可读写 **MsoPatternType** 类型。

**语法**

**express.Pattern**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Style

返回或设置线条格式样式。可读写 **MsoLineStyle** 类型。

**语法**

**express.Style**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Transparency

返回或设置线条的透明度。可读写 **Number** 类型。

**语法**

**express.Transparency**

\*express是一个代表 LineFormat 对象的变量。

**说明**

此属性的值可以取 0.0（不透明）至 1.0（透明）之间的值。该属性只影响单色线条的外观；而不影响带图案的线条的外观。

**示例**

``` JavaScript
/* 添加一条直线，并且设置透明度为0.5 */
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.BackColor.RGB = 0x00ff00;
    lnFormat.Transparency = 0.5
}
```

### LineFormat.Weight

返回或设置指定线条的粗细（以磅为单位）。 **Number** 类型，可读写。

**语法**

**express.Weight**

\*express是一个代表 LineFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/LineFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
