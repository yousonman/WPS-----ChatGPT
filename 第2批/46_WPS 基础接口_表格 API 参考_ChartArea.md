# ChartArea 对象

## 简介

代表图表的图表区。

## 说明

图表区包含绘图区在内的一切内容。但是，绘图区具有它自己的填充方式，因此填充绘图区并不会填充图表区。

有关设置绘图区格式的信息，请参阅 PlotArea 对象 。

使用 ChartArea 属性可返回 ChartArea 对象。

示例

以下示例关闭名为“Sheet1”的工作表上嵌入图表 1 的图表区边框。

``` JavaScript
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format.Line.Visible = false
```

## 方法

| 序号 | 名称                                      | 说明                             |
|:----:|:------------------------------------------|:---------------------------------|
|  1   | [Clear](#ChartArea.Clear)                 | 清除整个对象。                   |
|  2   | [ClearContents](#ChartArea.ClearContents) | 清除图表中的数据但保留格式设置。 |
|  3   | [ClearFormats](#ChartArea.ClearFormats)   | 清除对象的格式设置。             |
|  4   | [Copy](#ChartArea.Copy)                   | 将对象复制到剪贴板。             |
|  5   | [Select](#ChartArea.Select)               | 选择对象。                       |

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartArea.Application)       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#ChartArea.Creator)               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Format](#ChartArea.Format)                 | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  4   | [Height](#ChartArea.Height)                 | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  5   | [Left](#ChartArea.Left)                     | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  6   | [Name](#ChartArea.Name)                     | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  7   | [Parent](#ChartArea.Parent)                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  8   | [RoundedCorners](#ChartArea.RoundedCorners) | 如果图表的图表区域使用圆角，则为 True 。可读写 Boolean 类型。                                                                                                                                                                   |
|  9   | [Shadow](#ChartArea.Shadow)                 | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  10  | [Top](#ChartArea.Top)                       | 返回一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                            |
|  11  | [Width](#ChartArea.Width)                   | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |

## 成员方法

### ChartArea.Clear

清除整个对象。

**语法**

**express.Clear ()**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*此示例清除 Chart1 中的图表区域（图表数据和格式设置）。*/
Application.Charts.Item("Chart1").ChartArea.Clear()
```

### ChartArea.ClearContents

清除图表中的数据但保留格式设置。

**语法**

**express.ClearContents ()**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*此示例清除 Chart1 的图表数据，但保留其格式设置。*/
Application.Charts.Item("Chart1").ChartArea.ClearContents()
```

### ChartArea.ClearFormats

清除对象的格式设置。

**语法**

**express.ClearFormats ()**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*此示例清除工作表 Sheet1 上第一个嵌入式图表的格式设置。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats()
```

### ChartArea.Copy

将对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 ChartArea 对象的变量。

## 成员属性

### ChartArea.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    if(Application.ActiveWorkbook.Application.Value == "ET"){
        alert("This is an ET Application object.")
    }
    else{
        alert("This is not an ET Application object.")
    }
}
```

### ChartArea.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartArea 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ChartArea.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 ChartArea 对象的变量。

**说明**

**ChartFormat** 对象包含图表区的线条、填充和效果格式。

### ChartArea.Height

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.RoundedCorners

如果图表的图表区域使用圆角，则为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.RoundedCorners**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*此示例为 Sheet1 上的第一个图表添加圆角。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.RoundedCorners = true
```

### ChartArea.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Top

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Width

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 ChartArea 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartArea](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
