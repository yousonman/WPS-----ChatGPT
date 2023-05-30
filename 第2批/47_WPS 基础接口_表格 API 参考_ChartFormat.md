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
