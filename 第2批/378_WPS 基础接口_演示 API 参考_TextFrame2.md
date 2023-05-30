# TextFrame2 对象

## 简介

代表 Shape 、 ShapeRange 或 ChartFormat 对象的文本框。

## 说明

该对象包含文本框中的文本，还包含控制文本框对齐方式和位置的属性和方法。使用 TextFrame2 属性可返回 TextFrame2 对象。

``` JavaScript
/*下例向 myDocument 添加一个矩形，向矩形添加文本，然后设置文本框的边距。*/
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2
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
|  2   | [HorizontalAnchor](#TextFrame2.HorizontalAnchor) | 返回或设置指定文本的水平定位类型。可读/写 MsoHorizontalAnchor 类型。                                |
|  3   | [MarginBottom](#TextFrame2.MarginBottom)         | 返回或设置从文本框底端到包含文本的形状的内接矩形底端的距离（以磅为单位）。可读/写 Single 类型。     |
|  4   | [MarginLeft](#TextFrame2.MarginLeft)             | 以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形的左边界的距离。可读/写 Single 类型。   |
|  5   | [MarginRight](#TextFrame2.MarginRight)           | 返回或设置从文本框右边缘到包含文本的形状的内接矩形右边缘的距离（以磅为单位）。可读/写 Single 类型。 |
|  6   | [MarginTop](#TextFrame2.MarginTop)               | 返回或设置从文本框顶端到包含文本的形状的内接矩形顶端的距离（以磅为单位）。可读/写 Single 类型。     |
|  7   | [Orientation](#TextFrame2.Orientation)           | 返回或设置一个值，该值代表文本框方向。可读/写 MsoTextOrientation 类型。                             |
|  8   | [Ruler](#TextFrame2.Ruler)                       | 返回一个 Ruler 对象，该对象代表指定文本的标尺。只读。                                               |
|  9   | [TextRange](#TextFrame2.TextRange)               | 返回 TextRange2 对象，该对象代表对象中的文本。只读。                                                |
|  10  | [ThreeD](#TextFrame2.ThreeD)                     | 返回一个 ThreeDFormat 对象，该对象包含指定文本的三维效果格式属性。只读。                            |
|  11  | [VerticalAnchor](#TextFrame2.VerticalAnchor)     | 返回或设置指定文本的垂直定位类型。可读/写 MsoVerticalAnchor 类型。                                  |
|  12  | [WarpFormat](#TextFrame2.WarpFormat)             | 返回或设置指定文本框的弯曲类型。可读/写 MsoWarpFormat 。                                            |
|  13  | [WordWrap](#TextFrame2.WordWrap)                 | 返回或设置形状边界内外的文本换行。可读/写 MsoTriState 类型。                                        |

## 成员方法

### TextFrame2.DeleteText

删除文本框中的文字以及所有关联的文字属性。

**语法**

**express.DeleteText ()**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

关联的文字属性包括 **Font** 属性，例如加粗、下划线等等。

**示例**

``` JavaScript
/*如果文本框包含文字，则此示例将删除文本框中的文字。*/
function test() {
  let deleteText = Application.ActiveSheet.Shapes.Item(1).TextFrame2
  if ( deleteText.HasText ) {
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

### TextFrame2.Orientation

返回或设置一个值，该值代表文本框方向。可读/写 **MsoTextOrientation** 类型。

**语法**

**express.Orientation**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

此属性值可设置为从 -90 到 90 度的整数值，或者以下 **<span "="" r="" rlidawscontentredir?assetid="HV100730982052&amp;CTT=11&amp;Origin=HV103325302052&quot;" style="color: rgb(0, 0, 0); text-decoration-line: none;" target="_new">MsoTextOrientation</span>** 常量之一。

### TextFrame2.Ruler

返回一个 **Ruler** 对象，该对象代表指定文本的标尺。只读。

**语法**

**express.Ruler**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.TextRange

返回 **TextRange2** 对象，该对象代表对象中的文本。只读。

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

### TextFrame2.WordWrap

返回或设置形状边界内外的文本换行。可读/写 **MsoTriState** 类型。

**语法**

**express.WordWrap**

\*express是一个代表 TextFrame2 对象的变量。

> WPS 开发文档： [WPS 基础接口/演示 API 参考/TextFrame2](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
