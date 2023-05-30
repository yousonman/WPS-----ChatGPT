# TableStyleElement 对象

## 简介

代表单个表格样式元素。

## 说明

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，标题行是表格的元素。表格样式可以规定标题行的填充色为红色。

表格中每个表格样式元素的格式设置可在适用于该元素的表格样式中指定。

## 方法

| 序号 | 名称                              | 说明               |
|:----:|:----------------------------------|:-------------------|
|  1   | [Clear](#TableStyleElement.Clear) | 清除此元素的格式。 |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                               |
|:----:|:----------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#TableStyleElement.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Borders](#TableStyleElement.Borders)         | 返回一个代表表样式元素的边框的 Borders 集合。只读。                                                                                                                |
|  3   | [Creator](#TableStyleElement.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Font](#TableStyleElement.Font)               | 返回一个 Font 对象，它代表指定对象的字体。 只读。                                                                                                                  |
|  5   | [HasFormat](#TableStyleElement.HasFormat)     | 返回表样式元素是否具有应用到指定元素的格式设置。只读 Boolean 类型。                                                                                                |
|  6   | [Interior](#TableStyleElement.Interior)       | 返回一个 Interior 对象，该对象代表指定对象的内部。 只读。                                                                                                          |
|  7   | [Parent](#TableStyleElement.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  8   | [StripeSize](#TableStyleElement.StripeSize)   | 返回或设置条带的大小。可读/写 Long 类型。                                                                                                                          |

## 成员方法

### TableStyleElement.Clear

清除此元素的格式。

**语法**

**express.Clear ()**

\*express是一个代表 TableStyleElement 对象的变量。

## 成员属性

### TableStyleElement.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TableStyleElement 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### TableStyleElement.Borders

返回一个代表表样式元素的边框的 **Borders** 集合。只读。

**语法**

**express.Borders**

\*express是一个代表 TableStyleElement 对象的变量。

**示例**

本示例将表格的上边框的颜色设置为红色。

``` JavaScript
function test() {
    let sty = Application.ActiveWorkbook.TableStyles.Item("Table Style 4").TableStyleElements.Item(Application.Enum.xlWholeTable).Borders.Item(Application.Enum.xlEdgeTop)
    sty.Color = 255
    sty.TintAndShade = 0
    sty.Weight = 2
    sty.LineStyle = 1
}
```

### TableStyleElement.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TableStyleElement 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL

### TableStyleElement.Font

返回一个 **Font** 对象，它代表指定对象的字体。 只读。

**语法**

**express.Font**

\*express是一个代表 TableStyleElement 对象的变量。

### TableStyleElement.HasFormat

返回表样式元素是否具有应用到指定元素的格式设置。只读 **Boolean** 类型。

**语法**

**express.HasFormat**

\*express是一个代表 TableStyleElement 对象的变量。

### TableStyleElement.Interior

返回一个 **Interior** 对象，该对象代表指定对象的内部。 只读。

**语法**

**express.Interior**

\*express是一个代表 TableStyleElement 对象的变量。

### TableStyleElement.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 TableStyleElement 对象的变量。

### TableStyleElement.StripeSize

返回或设置条带的大小。可读/写 **Long** 类型。

**语法**

**express.StripeSize**

\*express是一个代表 TableStyleElement 对象的变量。

**说明**

此属性不会应用于所有 **TableStyleElement** 对象。它只应用于 **xlColumnStripe1** 、 **xlColumnStripe2** 、 **xlRowStripe1** 和 **xlRowStripe2** 类型。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/TableStyleElement](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
