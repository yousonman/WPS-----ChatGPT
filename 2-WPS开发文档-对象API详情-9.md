# HSpacer 对象

## 简介

填充水平方向上无用的空隙。

## 说明

填充水平方向上无用的空隙。

## 属性

| 序号 | 名称                                  | 说明                                                 |
|:----:|:--------------------------------------|:-----------------------------------------------------|
|  1   | [Height](#HSpacer.Height)             | 获取或设置指定对象的高度                             |
|  2   | [HeightPolicy](#HSpacer.HeightPolicy) | 获取或设置指定对像的高度策略                         |
|  3   | [Left](#HSpacer.Left)                 | 您可以通过该属性指定控件在对象或报表中的位置         |
|  4   | [Name](#HSpacer.Name)                 | 您可以通过该属性指定或确定表示对象名称的字符串，只读 |
|  5   | [Top](#HSpacer.Top)                   | 您可以通过该属性指定控件在对象或报表中的位置         |
|  6   | [Width](#HSpacer.Width)               | 获取或设置指定对象的宽度                             |
|  7   | [WidthPolicy](#HSpacer.WidthPolicy)   | 获取或设置指定对象的宽度策略                         |

## 成员属性

### HSpacer.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 HSpacer 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将HSpacer1控件的高度设置为100*/
function func1()
{
    UserForm1.HSpacer1.Height = 100
}
```

### HSpacer.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 HSpacer 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将HSpacer1的高度设置为固定*/
function func1()
{
    UserForm1.HSpacer1.HeightPolicy = 2
}
```

### HSpacer.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 HSpacer 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将HSpacer1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.HSpacer1.Left = 100
}
```

### HSpacer.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 HSpacer 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### HSpacer.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 HSpacer 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将HSpacer1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.HSpacer1.Top = 100
}
```

### HSpacer.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 HSpacer 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将HSpacer1的宽度设置为100*/
function func1()
{
    UserForm1.HSpacer1.Width = 100
}
```

### HSpacer.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 HSpacer 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将HSpacer1的宽度设置为固定*/
function func1()
{
    UserForm1.HSpacer1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/HSpacer](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Adjustments 对象

## 简介

它包含指定的自选图形、艺术字对象或连接符的调整值的集合。

## 说明

每个调整值代表一种调整控点的调整方法。由于某些调整控点可以按两种方法调整（例如，某些控点既可以水平调整也可以垂直调整），所以形状的调整值数量可以大于调整控点数量。一个形状最多可以有八个调整值。

使用 Adjustments 属性可返回 Adjustments 对象。使用 Adjustments ( *index* )（其中 *index* 是调整值的索引号）可返回单个调整值。

不同的形状具有不同数目的调整值，不同类型的调整值在不同的方向上调整形状的几何性质，不同类型的调整值有不同的取值范围。例如，下面的图示显示了右箭头标注的四个调整值各对该标注的几何形状起什么作用。

![](服务器端图像/adjlabel_ZA06051188.gif)

| 注释                                                                                                                                 |
|--------------------------------------------------------------------------------------------------------------------------------------|
| 由于每个形状有不同的调整值集，校验指定形状的调整行为的最好方法是手动创建一个图例，在打开宏记录器的情况下作调整，然后检查记录的代码。 |

下表概括了不同类型的调整值的有效取值范围。多数情况下，如果指定的调整值超出了有效取值范围，就将用最接近的有效值来代替。

| 调整类型           | 有效值                                                                                                                                                                                                                                                                                             |
|--------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 线性（水平或垂直） | 通常 0.0 值代表形状的左边界或上边界，而 1.0 值代表形状的右边界或下边界。有效值对应于有效的手动调整。例如，如果只能将调整控点手动拖动形状的一半宽度，则相应的调整值最大为 0.5。对于象连接符和标注这样的形状，0.0 和 1.0 值代表由它们的起始和终止点定义的矩形界限，此时负值和大于 1.0 的值是有效的。 |
| 射线图             | 调整值 1.0 对应于形状宽度。最大值为 0.5，或形状宽度的一半。                                                                                                                                                                                                                                        |
| 角                 | 值以度表示。如果指定的值超过了 -180 到 180 的范围，则将其折算为该范围内的值。                                                                                                                                                                                                                      |

示例

``` JavaScript
/*
 * 本示例向 myDocument 中添加右箭头标注，并且设置该标注的调整值。
 * 请注意，尽管形状只有三个调整控点，但是它有四个调整值。第三和第四个调整值都和箭头头部和颈部间的调整句柄相对应。
*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    let rac = myDocument.Shapes.AddShape(msoShapeRightArrowCallout,10, 10, 250, 190)
    let adj = rac.Adjustments
    adj.Item(1) = 0.5    //adjusts width of text box
    adj.Item(2) = 0.15   //adjusts width of arrow head
    adj.Item(3) = 0.8    //adjusts length of arrow head
    adj.Item(4) = 0.4    //adjusts width of arrow neck
}
```

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Adjustments.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Adjustments.Count)             | 返回一个 Integer 值，它代表集合中对象的数量。                                                                                                                                                                                   |
|  3   | [Creator](#Adjustments.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Item](#Adjustments.Item)               | 返回或设置由 Index 参数指定的调整值。 Single 型，可读写。                                                                                                                                                                       |
|  5   | [Parent](#Adjustments.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员属性

### Adjustments.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Adjustments 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
  let myObject = Application.ActiveWorkbook
  if(myObject.Application.Value == "ET") {
      alert("This is an ET Application object.")
  }
  else {
      alert("This is not an ET Application object.")
  }
}
```

### Adjustments.Count

返回一个 **Integer** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Adjustments 对象的变量。

### Adjustments.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Adjustments 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Adjustments.Item

返回或设置由 **Index** 参数指定的调整值。 **Single** 型，可读写。

**语法**

**express.Item**

\*express是一个代表 Adjustments 对象的变量。

**说明**

自选图形、连接符和艺术字对象最多可有八个调整。

对于线性调整，调整值 0.0 通常对应于形状的左边缘或上边缘，值 1.0 通常对应于形状的右边缘或下边缘。不过，对于某些形状，调整可以超越边界。对于放射状调整，调整值 1.0 对应于形状的宽度。对于角度的调整，调整值为指定的角度。 **Item** 属性只应用于 具有调整的形状。

参数

| **名称** | **必选/可选** | **数据类型** | **说明**                    |
|----------|---------------|--------------|-----------------------------|
| *Index*  | 必选          | **\[INT\]**  | **Long** 型。调整的索引号。 |

### Adjustments.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Adjustments 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Adjustments](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Axis 对象

## 简介

代表图表中的单个坐标轴。

## 说明

Axis 对象是 Axes 集合的成员。

使用 Axes ( *type* , *group* )（其中 *type* 为坐标轴类型，而 *group* 为坐标轴组）可返回单个 Axis 对象。 *Type* 可为以下 XlAxisType 常量之一： xlCategory 、 xlSeries 或 xlValue 。 *Group* 可为以下 XlAxisGroup 常量之一： xlPrimary 或 xlSecondary 。有关详细信息，请参阅 Axes 方法。

示例

``` JavaScript
/*下例在名为“Chart1”的图表工作表中设置分类轴的标题文本。*/
function test()
{
    let myChart = Application.Charts.Item("Chart1").Axes.Item(xlCategory)
    myChart.HasTitle = true
    myChart.AxisTitle.Caption = "1994"
}
```

## 方法

| 序号 | 名称                   | 说明       |
|:----:|:-----------------------|:-----------|
|  1   | [Delete](#Axis.Delete) | 删除对象。 |
|  2   | [Select](#Axis.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                                                                                                                             |
|:----:|:-------------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Axis.Application)                       | 如果不使用对象识别符，则该属性返回一个 App lication 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AxisBetweenCategories](#Axis.AxisBetweenCategories)   | 如果数值轴与分类坐标轴相交于分类之间，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                |
|  3   | [AxisGroup](#Axis.AxisGroup)                           | 返回指定轴的组。只读                                                                                                                                                                                                             |
|  4   | [AxisTitle](#Axis.AxisTitle)                           | 返回一个 AxisTitle 对象，该对象表示指定坐标轴的标题。只读。                                                                                                                                                                      |
|  5   | [BaseUnit](#Axis.BaseUnit)                             | 返回或设置指定分类轴的基本单位。 XlTim eUnit 类型，可读写。                                                                                                                                                                      |
|  6   | [BaseUnitIsAuto](#Axis.BaseUnitIsAuto)                 | 如果由 ET 为指定的分类轴选取适当的基本单位，则该值为 True 。默认值为 True 。 Boolean 类型，可读写。                                                                                                                              |
|  7   | [Border](#Axis.Border)                                 | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                         |
|  8   | [CategoryNames](#Axis.CategoryNames)                   | 以文本数组形式返回或设置指定坐标轴中所有分类的名称。该属性既可设为一个数组，也可设为一个包含所有分类名称的 Range 对象。 Variant 类型，可读写。                                                                                   |
|  9   | [CategoryType](#Axis.CategoryType)                     | 返回或设置分类轴类型。 XlCategoryType 类型，可读写。                                                                                                                                                                             |
|  10  | [Creator](#Axis.Creator)                               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                       |
|  11  | [Crosses](#Axis.Crosses)                               | 返回或设置指定坐标轴与其他坐标轴相交的点。 Long 类型，可读写。                                                                                                                                                                   |
|  12  | [CrossesAt](#Axis.CrossesAt)                           | 返回或设置数值轴中与分类坐标轴的交点。仅应用于数值轴。 Double 类型，可读写。                                                                                                                                                     |
|  13  | [DisplayUnit](#Axis.DisplayUnit)                       | 返回或设置数值轴的单位标签。 XlDisplayUnit 、 xlCustom 或 xlNone 类型，可读写。                                                                                                                                                  |
|  14  | [DisplayUnitCustom](#Axis.DisplayUnitCustom)           | 如果 DisplayUnit 属性的值是 xlCustom ，则 DisplayUnitCustom 返回或设置显示的单位的值。该值范围必须是从 0 到 10E307。 Double 类型，可读写。                                                                                       |
|  15  | [DisplayUnitLabel](#Axis.DisplayUnitLabel)             | 返回指定坐标轴的 DisplayUnitLabel 对象。如果 HasDisplayUnitLabel 属性设置为 False ，则返回 null 。只读。                                                                                                                         |
|  16  | [Format](#Axis.Format)                                 | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                    |
|  17  | [HasDisplayUnitLabel](#Axis.HasDisplayUnitLabel)       | 如果由 DisplayUnit 或 DisplayUnitCustom 属性指定的标签显示在指定坐标轴上，则该属性值为 True 。默认值为 True 。 Boolean 类型，可读写。                                                                                            |
|  18  | [HasMajorGridlines](#Axis.HasMajorGridlines)           | 如果坐标轴有主要网格线，则该值为 True 。只有主要坐标轴组中的坐标轴才能有网格线。 Boolean 类型，可读写。                                                                                                                          |
|  19  | [HasMinorGridlines](#Axis.HasMinorGridlines)           | 如果坐标轴有次要网格线，则该值为 True 。只有主要坐标轴组中的坐标轴才能有网格线。 Boolean 类型，可读写。                                                                                                                          |
|  20  | [HasTitle](#Axis.HasTitle)                             | 如果坐标轴或图表有可见标题，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                  |
|  21  | [Height](#Axis.Height)                                 | 返回一个 Double 值，它代表对象的高度（以磅为单位）。                                                                                                                                                                             |
|  22  | [Left](#Axis.Left)                                     | 返回一个 Double 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。                                                                                                                                                       |
|  23  | [LogBase](#Axis.LogBase)                               | 在使用对数刻度时返回或设置对数的底。可读/写 Double 类型。                                                                                                                                                                        |
|  24  | [MajorGridlines](#Axis.MajorGridlines)                 | 返回一个 Gridlines 对象，该对象表示指定坐标轴的主要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。                                                                                                                        |
|  25  | [MajorTickMark](#Axis.MajorTickMark)                   | 返回或设置指定坐标轴的主刻度线类型。 XlTickMark 类型，可读写。                                                                                                                                                                   |
|  26  | [MajorUnit](#Axis.MajorUnit)                           | 返回或设置数值轴的主要单位。 Double 类型，可读写。                                                                                                                                                                               |
|  27  | [MajorUnitIsAuto](#Axis.MajorUnitIsAuto)               | 如果 ET 计算数值轴的主要单位，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                        |
|  28  | [MajorUnitScale](#Axis.MajorUnitScale)                 | 当 CategoryType 属性设置为 xlTimeScale 时，返回或设置分类轴主要单位的刻度值。 XlTimeUnit 类型，可读写。                                                                                                                          |
|  29  | [MaximumScale](#Axis.MaximumScale)                     | 返回或设置数值轴上的最大值。 Double 类型，可读写。                                                                                                                                                                               |
|  30  | [MaximumScaleIsAuto](#Axis.MaximumScaleIsAuto)         | 如果 ET 计算该数值轴的最大值，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                        |
|  31  | [MinimumScale](#Axis.MinimumScale)                     | 返回或设置数值轴上的最小值。 Double 类型，可读写。                                                                                                                                                                               |
|  32  | [MinimumScaleIsAuto](#Axis.MinimumScaleIsAuto)         | 如果 ET 为数值轴计算最小值，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                          |
|  33  | [MinorGridlines](#Axis.MinorGridlines)                 | 返回 Gridlines 对象，该对象表示指定坐标轴的次要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。                                                                                                                            |
|  34  | [MinorTickMark](#Axis.MinorTickMark)                   | 返回或设置指定坐标轴的次要刻度线类型。 XlTickMark 类型，可读写。                                                                                                                                                                 |
|  35  | [MinorUnit](#Axis.MinorUnit)                           | 返回或设置数值轴上的次要单位。 Double 类型，可读写。                                                                                                                                                                             |
|  36  | [MinorUnitIsAuto](#Axis.MinorUnitIsAuto)               | 如果 ET 为数值轴计算次要单位，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                        |
|  37  | [MinorUnitScale](#Axis.MinorUnitScale)                 | 返回或设置当 CategoryType 属性设置为 xlTimeScale 时分类轴次要单位的刻度值。 XlTimeUnit 类型，可读写。                                                                                                                            |
|  38  | [Parent](#Axis.Parent)                                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                     |
|  39  | [ReversePlotOrder](#Axis.ReversePlotOrder)             | 如果 ET 的绘图区数据点的顺序为从后往前，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                              |
|  40  | [ScaleType](#Axis.ScaleType)                           | 返回或设置数值轴的刻度类型。 XlScaleType 类型，可读写。                                                                                                                                                                          |
|  41  | [TickLabelPosition](#Axis.TickLabelPosition)           | 说明指定坐标轴上刻度线标志的位置。 XlTickLabelPosition 类型，可读写。                                                                                                                                                            |
|  42  | [TickLabelSpacing](#Axis.TickLabelSpacing)             | 返回或设置刻度线标签之间的分类数或数据系列数。仅应用于分类轴和系列轴。可以是 1 到 31999 之间的一个数值。 Long 类型，可读写。                                                                                                     |
|  43  | [TickLabelSpacingIsAuto](#Axis.TickLabelSpacingIsAuto) | 返回或设置刻度标签间距是否是自动设置的。可读/写 Boolean 类型。                                                                                                                                                                   |
|  44  | [TickLabels](#Axis.TickLabels)                         | 返回一个 TickLabels 对象，该对象表示指定坐标轴的刻度线标签。只读。                                                                                                                                                               |
|  45  | [TickMarkSpacing](#Axis.TickMarkSpacing)               | 返回或设置刻度线之间的分类数或数据系列数。仅应用于分类轴和系列轴。可以是 1 到 31999 之间的一个数值。 Long 类型，可读写。                                                                                                         |
|  46  | [Top](#Axis.Top)                                       | 返回一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                             |
|  47  | [Type](#Axis.Type)                                     | 返回一个 XlAxisType 值，它代表“坐标轴”类型。                                                                                                                                                                                     |
|  48  | [Width](#Axis.Width)                                   | 返回一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                             |

## 成员方法

### Axis.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Axis 对象的变量。

### Axis.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 Axis 对象的变量。

## 成员属性

### Axis.Application

如果不使用对象识别符，则该属性返回一个 **App** **lication** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Axis 对象的变量。

### Axis.AxisBetweenCategories

如果数值轴与分类坐标轴相交于分类之间，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AxisBetweenCategories**

\*express是一个代表 Axis 对象的变量。

**说明**

该属性仅适用于分类轴，而不适用于三维图表。

**示例**

``` JavaScript
/*本示例设置 Chart1 的数值轴与分类坐标轴相交于两个分类之间。*/
Application.Charts.Item("Chart1").Axes(xlCategory).AxisBetweenCategories = true
```

### Axis.AxisGroup

返回指定轴的组。只读

**语法**

**express.AxisGroup**

\*express是一个代表 Axis 对象的变量。

### Axis.AxisTitle

返回一个 **AxisTitle** 对象，该对象表示指定坐标轴的标题。只读。

**语法**

**express.AxisTitle**

\*express是一个代表 Axis 对象的变量。

**说明**

本示例为 Chart1 的分类轴添加轴标签。

**示例**

``` JavaScript
function test()
{
  let myChart = Application.Charts.Item("Chart1").Axes(xlCategory)
  myChart.HasTitle = true
  myChart.AxisTitle.Text = "July Sales"
}
```

### Axis.BaseUnit

返回或设置指定分类轴的基本单位。 **XlTim** **eUnit** 类型，可读写。

**语法**

**express.BaseUnit**

\*express是一个代表 Axis 对象的变量。

**说明**

如果指定坐标轴的 **CategoryType** 属性设置为 **xlCategoryScale** ，则设置该属性不会看到任何效果。不过，将保留设置值，并在 **CategoryType** 属性设置为 **xlTimeScale** 时起作用。 不能对数值轴设置本属性。

**示例**

``` JavaScript
/*本示例使第一张工作表上嵌入的第一个图表中的分类轴使用时间刻度，并以月为基本单位。*/
function test()
{
  let axes = Application.Worksheets.Item(1).ChartObjects(1).Chart.Axes(xlCategory)
  
  axes.CategoryType = xlTimeScale
  axes.BaseUnit = xlMonths
}
```

### Axis.BaseUnitIsAuto

如果由 ET 为指定的分类轴选取适当的基本单位，则该值为 **True** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.BaseUnitIsAuto**

\*express是一个代表 Axis 对象的变量。

**说明**

不能对数值轴设置本属性。

**示例**

``` JavaScript
/*本示例使第一张工作表上嵌入的第一个图表中的分类轴使用时间刻度，并使用自动选取的基本单位。*/
function test()
{
  let axes = Application.Worksheets.Item(1).ChartObjects(1).Chart.Axes(xlCategory)
  
  axes.CategoryType = xlTimeScale
  axes.BaseUnitIsAuto = true
}
```

### Axis.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 Axis 对象的变量。

### Axis.CategoryNames

以文本数组形式返回或设置指定坐标轴中所有分类的名称。该属性既可设为一个数组，也可设为一个包含所有分类名称的 **Range** 对象。 **Variant** 类型，可读写。

**语法**

**express.CategoryNames**

\*express是一个代表 Axis 对象的变量。

**说明**

该属性仅适用于分类轴。

**示例**

``` JavaScript
/*本示例将 Chart1 的分类名设为 Sheet1 中 B1:B5 单元格的值。*/
Application.Charts("Chart1").Axes(xlCategory).CategoryNames = Application.Worksheets("Sheet1").Range("B1:B5")

/*本示例使用数组对 Chart1 的个别分类名进行设置。*/
Charts("Chart1").Axes(xlCategory).CategoryNames = Array("1985", "1986", "1987", "1988", "1989")
```

### Axis.CategoryType

返回或设置分类轴类型。 **XlCategoryType** 类型，可读写。

**语法**

**express.CategoryType**

\*express是一个代表 Axis 对象的变量。

**说明**

不能对数值轴设置本属性。

**示例**

``` JavaScript
/*本示例使第一张工作表上嵌入的第一个图表中的分类轴使用时间刻度，并以月为基本单位。*/
function test()
{
  let chart = Application.Worksheets(1).CharObjects(1).Chart
  let axes = chart.Axes(xlCategory)
  axes.CategoryType = xlTimeScale
  axes.BaseUnit = xlMonths
}
```

### Axis.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Axis 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Axis.Crosses

返回或设置指定坐标轴与其他坐标轴相交的点。 **Long** 类型，可读写。

**语法**

**express.Crosses**

\*express是一个代表 Axis 对象的变量。

**说明**

可以是下表所列的 **XlAxisCrosses** 常量之一。

| 常量                       | 含义                               |
|----------------------------|------------------------------------|
| **xlAxisCrossesAutomatic** | 由 ET 设置坐标轴交点。             |
| **xlMinimum**              | 坐标轴在最小值处相交。             |
| **xlMaximum**              | 坐标轴在最大值处相交。             |
| **xlAxisCrossesCustom**    | **CrossesAt** 属性指定坐标轴交点。 |

该属性对雷达图无效。对于三维图表，该属性只能应用于数值轴并指示由分类轴所定义的平面与数值轴相交的位置。 本属性既可用于分类轴，也可用于数值轴。在分类轴上，xlMinimum 使数值轴在第一个分类处与之相交，而 xlMaximum 使数值轴在最后一个分类处与之相交。 请注意，根据坐标轴的不同，xlMinimum 和 xlMaximum 可能有不同的含义。

**示例**

``` JavaScript
/*本示例设置 Chart1 的数值轴与分类坐标轴的交点在最大的 x 值处。*/
Application.Charts.Item("Chart1").Axes(xlCategory).Crosses = xlMaximum
```

### Axis.CrossesAt

返回或设置数值轴中与分类坐标轴的交点。仅应用于数值轴。 **Double** 类型，可读写。

**语法**

**express.CrossesAt**

\*express是一个代表 Axis 对象的变量。

**说明**

设置该属性将使 **Crosses** 属性更改为 **xlAxisCrossesCustom** 。

该属性不能用于雷达图。对于三维图表，本属性指示由分类轴和数值轴交叉所定义的平面的位置。

**示例**

``` JavaScript
/*本示例设置活动图表中的分类坐标轴与数值轴的交点位于数值坐标轴上数值 3 的位置。*/
function Chart(){
    // Create a sample source of data.
    Application.Range("A1").Value2 = "2"
    Application.Range("A2").Value2 = "4"
    Application.Range("A3").Value2 = "6"
    Application.Range("A4").Value2 = "3"

    // Create a chart based on the sample source of data.
    Application.Charts.Add()

    Application.ActiveChart.ChartType = xlLineMarkersStacked
    Application.ActiveChart.SetSourceData(Sheets.Item("Sheet1").Range("A1:A4"), xlColumns)
    Application.ActiveChart.Location(xlLocationAsObject, "Sheet1")

    // Set the category axis to cross the value axis at value 3.
    Application.ActiveChart.Axes(xlValue).Select()
    Application.Selection.CrossesAt = 3
}
```

### Axis.DisplayUnit

返回或设置数值轴的单位标签。 **XlDisplayUnit** 、 **xlCustom** 或 **xlNone** 类型，可读写。

**语法**

**express.DisplayUnit**

\*express是一个代表 Axis 对象的变量。

**说明**

|                                                         |
|---------------------------------------------------------|
| **XlDisplayUnit** 可为以下 **XlDisplayUnit** 常量之一。 |
| **xlHundredMillions**                                   |
| **xlHundredThousands**                                  |
| **xlMillions**                                          |
| **xlTenThousands**                                      |
| **xlThousands**                                         |
| **xlHundreds**                                          |
| **xlMillionMillions**                                   |
| **xlTenMillions**                                       |
| **xlThousandMillions**                                  |
| 此单位标签还可为以下常量之一                            |
| **xlCustom**                                            |
| **xlNone**                                              |

在绘制大数值图表时，使用单位标志可使刻度线标志易于阅读。例如，如果将数值轴的单位设置为百、千或百万，则可以在坐标轴的刻度标志上使用较小的数字值。

**示例**

``` JavaScript
/*本示例将 Chart1 中数值轴的显示单位设置为百。*/
function test()
{
  let axes = Application.Charts.Item("Chart1").Axes(xlValue)
  axes.DisplayUnit = xlHundreds
  axes.HasTitle = true
  axes.AxisTitle.Caption = "Rebate Amounts"
}
```

### Axis.DisplayUnitCustom

如果 **DisplayUnit** 属性的值是 **xlCustom** ，则 **DisplayUnitCustom** 返回或设置显示的单位的值。该值范围必须是从 0 到 10E307。 **Double** 类型，可读写。

**语法**

**express.DisplayUnitCustom**

\*express是一个代表 Axis 对象的变量。

**说明**

在绘制大数值图表时，使用单位标志可使刻度线标志易于阅读。例如，如果将数值轴的单位设置为百、千或百万，则可以在坐标轴的刻度标志上使用较小的数字值。

**示例**

``` JavaScript
/*本示例将 Chart1 中数值轴上显示的单位设置为 500。*/
function test()
{
  let axes = Application.Charts.Item("Chart1").Axes(xlValue)
  axes.DisplayUnit = xlCustom
  axes.DisplayUnitCustom = 500
  axes.HasTitle = true
  axes.AxisTitle.Caption = "Rebate Amounts"
}
```

### Axis.DisplayUnitLabel

返回指定坐标轴的 **DisplayUnitLabel** 对象。如果 **HasDisplayUnitLabel** 属性设置为 **False** ，则返回 **null** 。只读。

**语法**

**express.DisplayUnitLabel**

\*express是一个代表 Axis 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 中数值轴的标签标题设置为“Millions”，然后关闭自动字体缩放。*/
function test()
{
  let myLabel = Application.Charts.Item("Chart1").Axes(xlValue).DisplayUnitLabel
  myLabel.Caption = "Millions"
  myLabel.AutoScaleFont = false
}
```

### Axis.Format

返回 ChartFormat 对象。只读。

**语法**

**express.Format**

\*express是一个代表 Axis 对象的变量。

### Axis.HasDisplayUnitLabel

如果由 **DisplayUnit** 或 **DisplayUnitCustom** 属性指定的标签显示在指定坐标轴上，则该属性值为 **True** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasDisplayUnitLabel**

\*express是一个代表 Axis 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 中数值轴上的单位设置为 500，但隐藏单位标志。*/
function test()
{
  let axes = Application.Charts.Item("Chart1").Axes(xlValue)
  axes.DisplayUnit = xlCustom
  axes.DisplayUnitCustom = 500
  axes.AxisTitle.Caption = "Rebate Amounts"
  axes.HasDisplayUnitLabel = false
}
```

### Axis.HasMajorGridlines

如果坐标轴有主要网格线，则该值为 **True** 。只有主要坐标轴组中的坐标轴才能有网格线。 **Boolean** 类型，可读写。

**语法**

**express.HasMajorGridlines**

\*express是一个代表 Axis 对象的变量。

**示例**

``` JavaScript
/*本示例为 Chart1 中数值轴的主要网格线设置了颜色。*/
function test()
{
  let axes = Application.Charts.Item("Chart1").Axes(xlValue)
  
  if(axes.HasMajorGridlines){
      // Set color to red
      axes.MajorGridlines.Border.ColorIndex = 3
  }
}
```

### Axis.HasMinorGridlines

如果坐标轴有次要网格线，则该值为 **True** 。只有主要坐标轴组中的坐标轴才能有网格线。 **Boolean** 类型，可读写。

**语法**

**express.HasMinorGridlines**

\*express是一个代表 Axis 对象的变量。

**示例**

``` JavaScript
/*本示例对 Chart1 数值轴的次要网格线的颜色进行设置。*/
function test()
{
  let axes = Application.Charts.Item("Chart1").Axes(xlValue)
  
  if(axes.HasMinorGridlines){
      // Set color to green
      axes.MinorGridlines.Border.ColorIndex = 4
  }
}
```

### Axis.HasTitle

如果坐标轴或图表有可见标题，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasTitle**

\*express是一个代表 Axis 对象的变量。

**说明**

坐标轴标题由 **AxisTitle** 对象代表。

**示例**

``` JavaScript
/*此示例为 Chart1 的分类坐标轴添加坐标轴标志。*/
function test()
{
  Application.Charts.Item("Chart1").Axes().Item(xlCategory).HasTitle = true
  Application.Charts.Item("Chart1").Axes().Item(xlCategory).AxisTitle.Text = "July Sales"
}
```

### Axis.Height

返回一个 **Double** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 Axis 对象的变量。

### Axis.Left

返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 Axis 对象的变量。

### Axis.LogBase

在使用对数刻度时返回或设置对数的底。可读/写 **Double** 类型。

**语法**

**express.LogBase**

\*express是一个代表 Axis 对象的变量。

**说明**

试图将此属性设置为小于或等于零 (0) 的值将导致错误。默认值是 10。

### Axis.MajorGridlines

返回一个 **Gridlines** 对象，该对象表示指定坐标轴的主要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。

**语法**

**express.MajorGridlines**

\*express是一个代表 Axis 对象的变量。

**示例**

``` JavaScript
/*本示例为 Chart1 中数值轴的主要网格线设置了颜色。*/
function test()
{
  let axes = Application.Charts.Item("Chart1").Axes().Item(xlValue)
  if(axes.HasMajorGridlines){
      axes.MajorGridlines.Border.ColorIndex = 5    //set color to blue
  }
}
```

### Axis.MajorTickMark

返回或设置指定坐标轴的主刻度线类型。 **XlTickMark** 类型，可读写。

**语法**

**express.MajorTickMark**

\*express是一个代表 Axis 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| **XlTickMark** 可为以下 **XlTickMark** 常量之一。 |
| **xlTickMarkInside**                              |
| **xlTickMarkOutside**                             |
| **xlTickMarkCross**                               |
| **xlTickMarkNone**                                |

**示例**

``` JavaScript
/*本示例将 Chart1 中的数值轴主要刻度线设置为在轴外侧。*/
Application.Charts.Item("Chart1").Axes().Item(xlValue).MajorTickMark = xlTickMarkOutside
```

### Axis.MajorUnit

返回或设置数值轴的主要单位。 **Double** 类型，可读写。

**语法**

**express.MajorUnit**

\*express是一个代表 Axis 对象的变量。

**说明**

设置该属性会将 **MajorUnitIsAuto** 属性设置为 False。 使用 **TickMarkSpacing** 属性可设置分类轴上的刻度线间距。

**示例**

``` JavaScript
/*本示例为 Chart1 中的数值轴设置主要单位和次要单位。*/
function test()
{
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MajorUnit = 100
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MinorUnit = 20
}
```

### Axis.MajorUnitIsAuto

如果 ET 计算数值轴的主要单位，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MajorUnitIsAuto**

\*express是一个代表 Axis 对象的变量。

**说明**

设置 **MajorUnit** 属性会将该属性设置为 False。

**示例**

``` JavaScript
/*本示例自动设置 Chart1 中的数值轴的主要单位和次要单位。*/
function test()
{
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MajorUnitIsAuto = true
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MinorUnitIsAuto = true
}
```

### Axis.MajorUnitScale

当 **CategoryType** 属性设置为 **xlTimeScale** 时，返回或设置分类轴主要单位的刻度值。 **XlTimeUnit** 类型，可读写。

**语法**

**express.MajorUnitScale**

\*express是一个代表 Axis 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| **XlTimeUnit** 可为以下 **XlTimeUnit** 常量之一。 |
| **xlMonths**                                      |
| **xlDays**                                        |
| **xlYears**                                       |

**示例**

``` JavaScript
/*本示例将分类轴设置为使用时间刻度，并设置主要刻度单位和次要刻度单位。*/
function test()
{
  let axes = Application.Charts.Item(1).Axes().Item(xlCategory)
  axes.CategoryType = xlTimeScale
  axes.MajorUnit = 5
  axes.MajorUnitScale = xlDays
  axes.MinorUnit = 1
  axes.MinorUnitScale = xlDays
}
```

### Axis.MaximumScale

返回或设置数值轴上的最大值。 **Double** 类型，可读写。

**语法**

**express.MaximumScale**

\*express是一个代表 Axis 对象的变量。

**说明**

设置该属性会将 **MaximumScaleIsAuto** 属性设置为 **False** 。

**示例**

``` JavaScript
/*本示例对 Chart1 中的数值轴的最小值和最大值进行设置。*/
function test()
{
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MinimumScale = 10
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MaximumScale = 120
}
```

### Axis.MaximumScaleIsAuto

如果 ET 计算该数值轴的最大值，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MaximumScaleIsAuto**

\*express是一个代表 Axis 对象的变量。

**说明**

设置 **MaximumScale** 属性会将该属性设置为 False。

**示例**

``` JavaScript
/*本示例自动计算 Chart1 中的数值轴的最小刻度和最大刻度。*/
function test()
{
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MinimumScaleIsAuto = true
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MaximumScaleIsAuto = true
}
```

### Axis.MinimumScale

返回或设置数值轴上的最小值。 **Double** 类型，可读写。

**语法**

**express.MinimumScale**

\*express是一个代表 Axis 对象的变量。

**说明**

设置该属性会将 **MinimumScaleIsAuto** 属性设置为 **False** 。

**示例**

``` JavaScript
/*本示例对 Chart1 中的数值轴的最小值和最大值进行设置。*/
function test()
{
Application.Charts.Item("Chart1").Axes().Item(xlValue).MinimumScale = 10
Application.Charts.Item("Chart1").Axes().Item(xlValue).MaximumScale = 120
}
```

### Axis.MinimumScaleIsAuto

如果 ET 为数值轴计算最小值，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MinimumScaleIsAuto**

\*express是一个代表 Axis 对象的变量。

**说明**

设置 **MinimumScale** 属性会将该属性设置为 False。

**示例**

``` JavaScript
/*本示例自动计算 Chart1 中的数值轴的最小刻度和最大刻度。*/
function test()
{
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MinimumScaleIsAuto = true
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MaximumScaleIsAuto = true
}
```

### Axis.MinorGridlines

返回 **Gridlines** 对象，该对象表示指定坐标轴的次要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。

**语法**

**express.MinorGridlines**

\*express是一个代表 Axis 对象的变量。

**示例**

``` JavaScript
/*本示例对 Chart1 数值轴的次要网格线的颜色进行设置。*/
function test()
{
  let minorGridlines  = Application.Charts.Item("Chart1").Axes().Item(xlValue)
  if(minorGridlines.HasMinorGridlines){
      minorGridlines.MinorGridlines.Border.ColorIndex = 5    //set color to blue
  }
}
```

### Axis.MinorTickMark

返回或设置指定坐标轴的次要刻度线类型。 **XlTickMark** 类型，可读写。

**语法**

**express.MinorTickMark**

\*express是一个代表 Axis 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| **XlTickMark** 可为以下 **XlTickMark** 常量之一。 |
| **xlTickMarkInside**                              |
| **xlTickMarkOutside**                             |
| **xlTickMarkCross**                               |
| **xlTickMarkNone**                                |

**示例**

``` JavaScript
/*本示例为 Chart1 数值轴内侧设置次要刻度线。*/
Application.Charts.Item("Chart1").Axes().Item(xlValue).MinorTickMark = xlTickMarkInside
```

### Axis.MinorUnit

返回或设置数值轴上的次要单位。 **Double** 类型，可读写。

**语法**

**express.MinorUnit**

\*express是一个代表 Axis 对象的变量。

**说明**

设置该属性会将 **MinorUnitIsAuto** 属性设置为 **False** 。 使用 **TickMarkSpacing** 属性可设置分类轴上的刻度线间距。

**示例**

``` JavaScript
/*本示例为 Chart1 中的数值轴设置主要单位和次要单位。*/
function test()
{
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MajorUnit = 100
  Application.Charts.Item("Chart1").Axes().Item(xlValue).MinorUnit = 20
}
```

### Axis.MinorUnitIsAuto

如果 ET 为数值轴计算次要单位，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MinorUnitIsAuto**

\*express是一个代表 Axis 对象的变量。

**说明**

设置 **MinorUnit** 属性会将该属性设置为 False。

**示例**

``` JavaScript
/*本示例自动计算 Chart1 数值轴的主要单位和次要单位。*/
function test()
{
  Charts.Item("Chart1").Axes().Item(xlValue).MajorUnitIsAuto = true
  Charts.Item("Chart1").Axes().Item(xlValue).MinorUnitIsAuto = true
}
```

### Axis.MinorUnitScale

返回或设置当 **CategoryType** 属性设置为 **xlTimeScale** 时分类轴次要单位的刻度值。 **XlTimeUnit** 类型，可读写。

**语法**

**express.MinorUnitScale**

\*express是一个代表 Axis 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| **XlTimeUnit** 可为以下 **XlTimeUnit** 常量之一。 |
| **xlMonths**                                      |
| **xlDays**                                        |
| **xlYears**                                       |

**示例**

``` JavaScript
/*本示例将分类轴设置为使用时间刻度，并设置主要刻度单位和次要刻度单位。*/
function test()
{
  let axes = Application.Charts.Item(1).Axes().Item(xlCategory)
  axes.CategoryType = xlTimeScale
  axes.MajorUnit = 5
  axes.MajorUnitScale = xlDays
  axes.MinorUnit = 1
  axes.MinorUnitScale = xlDays
}
```

### Axis.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Axis 对象的变量。

### Axis.ReversePlotOrder

如果 ET 的绘图区数据点的顺序为从后往前，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ReversePlotOrder**

\*express是一个代表 Axis 对象的变量。

**说明**

该属性不能用于雷达图。

**示例**

``` JavaScript
/*本示例将 Chart1 数值轴的绘图区数据点的顺序设置为从后往前。*/
Application.Charts.Item("Chart1").Axes().Item(xlValue).ReversePlotOrder = true
```

### Axis.ScaleType

返回或设置数值轴的刻度类型。 **XlScaleType** 类型，可读写。

**语法**

**express.ScaleType**

\*express是一个代表 Axis 对象的变量。

**说明**

|                                                     |
|-----------------------------------------------------|
| **XlScaleType** 可为以下 **XlScaleType** 常量之一。 |
| **xlScaleLinear**                                   |
| **xlScaleLogarithmic**                              |

对数刻度使用以 10 为底的对数。

**示例**

``` JavaScript
/*本示例将 Chart1 中的数值轴设置为使用对数刻度。*/
Application.Charts.Item("Chart1").Axes().Item(xlValue).ScaleType = xlScaleLogarithmic
```

### Axis.TickLabelPosition

说明指定坐标轴上刻度线标志的位置。 **XlTickLabelPosition** 类型，可读写。

**语法**

**express.TickLabelPosition**

\*express是一个代表 Axis 对象的变量。

**说明**

|                                                                     |
|---------------------------------------------------------------------|
| **XlTickLabelPosition** 可为以下 **XlTickLabelPosition** 常量之一。 |
| **xlTickLabelPositionLow**                                          |
| **xlTickLabelPositionNone**                                         |
| **xlTickLabelPositionHigh**                                         |
| **xlTickLabelPositionNextToAxis**                                   |

**示例**

``` JavaScript
/*本示例将 Chart1 中分类轴的刻度线标签设置为高位（在图表之上）。*/
Application.Charts.Item("Chart1").Axes().Item(xlCategory).TickLabelPosition = xlTickLabelPositionHigh
```

### Axis.TickLabelSpacing

返回或设置刻度线标签之间的分类数或数据系列数。仅应用于分类轴和系列轴。可以是 1 到 31999 之间的一个数值。 **Long** 类型，可读写。

**语法**

**express.TickLabelSpacing**

\*express是一个代表 Axis 对象的变量。

**说明**

数值轴上刻度线标签间距总是由 ET 进行计算。

**示例**

``` JavaScript
/*本示例设置 Chart1 中分类轴上刻度线标签之间的分类数。*/
Application.Charts.Item("Chart1").Axes().Item(xlCategory).TickLabelSpacing = 10
```

### Axis.TickLabelSpacingIsAuto

返回或设置刻度标签间距是否是自动设置的。可读/写 **Boolean** 类型。

**语法**

**express.TickLabelSpacingIsAuto**

\*express是一个代表 Axis 对象的变量。

**说明**

**TickLabelSpacing** 属性返回当前刻度标签间距是否是自动设置的。

### Axis.TickLabels

返回一个 **TickLabels** 对象，该对象表示指定坐标轴的刻度线标签。只读。

**语法**

**express.TickLabels**

\*express是一个代表 Axis 对象的变量。

**示例**

``` JavaScript
/*本示例为 Chart1 中数值轴刻度线标签的字体设置颜色。*/
Application.Charts.Item("Chart1").Axes().Item(xlValue).TickLabels.Font.ColorIndex = 3
```

### Axis.TickMarkSpacing

返回或设置刻度线之间的分类数或数据系列数。仅应用于分类轴和系列轴。可以是 1 到 31999 之间的一个数值。 **Long** 类型，可读写。

**语法**

**express.TickMarkSpacing**

\*express是一个代表 Axis 对象的变量。

**说明**

可用 **MajorUnit** 属性和 **MinorUnit** 属性设置数值轴上的刻度线间距。

**示例**

``` JavaScript
/*本示例设置 Chart1 中分类轴上刻度线之间的分类数。*/
Application.Charts.Item("Chart1").Axes().Item(xlCategory).TickMarkSpacing = 10
```

### Axis.Top

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 Axis 对象的变量。

### Axis.Type

返回一个 **XlAxisType** 值，它代表“坐标轴”类型。

**语法**

**express.Type**

\*express是一个代表 Axis 对象的变量。

**说明**

在对散点图使用此属性时， **xlCategory** 将被返回为“坐标轴”类型。

### Axis.Width

返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 Axis 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Axis](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CalculatedItems 对象

## 简介

由 PivotItem 对象组成的集合，这些对象代表指定数据透视表中的所有计算项。

## 方法

| 序号 | 名称                          | 说明                              |
|:----:|:------------------------------|:----------------------------------|
|  1   | [Add](#CalculatedItems.Add)   | 新建计算项。返回 PivotItem 对象。 |
|  2   | [Item](#CalculatedItems.Item) | 从集合中返回一个对象。            |

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CalculatedItems.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#CalculatedItems.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#CalculatedItems.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#CalculatedItems.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### CalculatedItems.Add

新建计算项。返回 **PivotItem** 对象。

**语法**

**express.Add ( *Name* , *Formula* , *UseStandardFormula* )**

\*express是一个代表 CalculatedItems 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                                                                                  |
|----------------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Name*               | 必选      | String   | 项的名称。                                                                                                                                            |
| *Formula*            | 必选      | String   | 项的公式。                                                                                                                                            |
| *UseStandardFormula* | 可选      | Variant  | 为实现向上兼容，该值默认为 False。对于作为项名称的任何参数中所含的字符串来说，该值为 True，并将被解释为已按标准美国英语进行了格式设置，而非本地设置。 |

**说明**

返回值 一个代表新计算项的 **PivotItem** 对象。

### CalculatedItems.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 CalculatedItems 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**说明**

返回值

包含在集合中的一个 **PivotItem** 对象。

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

``` JavaScript
/*此示例隐藏第一个计算项。*/
Application.Worksheets.Item(1).PivotTables(1).PivotFields("year").CalculatedItems().Item(1).Visible = false
```

## 成员属性

### CalculatedItems.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CalculatedItems 对象的变量。

### CalculatedItems.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 CalculatedItems 对象的变量。

### CalculatedItems.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CalculatedItems 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### CalculatedItems.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CalculatedItems 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/CalculatedItems](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Comments 对象

## 简介

由单元格批注组成的集合。

## 说明

每个批注都由一个 Comment 对象代表。

使用 Comments 属性可返回 Comments 集合。下例隐藏第一张工作表中的所有批注。

``` JavaScript
function test() {
    let cmt = Application.Worksheets.Item(1).Comments
    for(let c = 1;c <= cmt.Count;c++) {
        cmt.Item(c).Visible = false
    }
} 
```

使用 AddCom ment 方法可在区域内添加批注。下例在第一张工作表的单元格 E5 中添加批注。

``` JavaScript
function test() {
    let myComment = Application.Worksheets.Item(1).Range("E5").AddComment()
    myComment.Visible = false
    myComment.Text("reviewed on " + Date())
} 
```

使用 Comments ( *index* )（其中 *index* 为批注号）可返回 Comments 集合中的单条批注。下例隐藏第一张工作表中的第二条批注。

``` JavaScript
Application.Worksheets.Item(1).Comments.Item(2).Visible = false
```

## 方法

| 序号 | 名称                   | 说明                                                     |
|:----:|:-----------------------|:---------------------------------------------------------|
|  1   | [Item](#Comments.Item) | 从集合中返回一个对象， 包含在集合中的一个 Comment 对象。 |

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Comments.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Comments.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Comments.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Comments.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Comments.Item

从集合中返回一个对象， 包含在集合中的一个 **Comment** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Comments 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 必选      | Long     | 对象的索引号。 |

**示例**

``` JavaScript
/* 本示例隐藏第二条批注*/
Application.Worksheets.Item(1).Comments.Item(2).Visible = false
```

## 成员属性

### Comments.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Comments 对象的变量。

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

### Comments.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Comments 对象的变量。

### Comments.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Comments 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Comments.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Comments 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Comments](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Dialog 对象

## 简介

代表内置的 ET 对话框。

## 说明

Dialog 对象是 Dialogs 集合的成员。 Dialogs 集合包含 ET 中的所有内置对话框。不能在集合中新建或添加内置对话框。 Dialog 对象只能在 Show 方法中用来显示相应的对话框。

ET Visual Basic 对象库包含许多内置对话框的内置常量。每个常量的前缀均为“xlDialog”，随后即为对话框的名称。例如， “应用名称” 对话框常量为 xlDialogApplyNames ， “查找文件” 对话框常量为 xlDialogFindFile 。这些常量是 XlBuiltinDialog 枚举类型的成员。

使用 Di alogs ( *index* )（其中 *index* 是用于标识对话框的内置常量）可返回单个 Dialog 对象。下例运行 “文件” 菜单中的内置 “打开” 对话框。如果 ET 成功地打开了文件，则 Show 方法返回 True ；如果用户取消了对话框，则返回 False 。

``` JavaScript
let dlgAnswer = Application.Dialogs.Item(xlDialogOpen).Show()
```

## 方法

| 序号 | 名称                 | 说明                                                                                                                                              |
|:----:|:---------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Show](#Dialog.Show) | 显示内置的对话框，等待用户输入数据，然后返回一个代表用户响应的 Boolean 值 。如果用户单击“确定”，则返回 True ；如果用户单击“取消”，则返回 False 。 |

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Dialog.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#Dialog.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Parent](#Dialog.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Dialog.Show

显示内置的对话框，等待用户输入数据，然后返回一个代表用户响应的 **Boolean** 值 。如果用户单击“确定”，则返回 **True** ；如果用户单击“取消”，则返回 **False** 。

**语法**

**express.Show ( *Arg1* , *...* , *Arg30* )**

\*express是一个代表 Dialog 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                   |
|---------|-----------|----------|------------------------------------------------------------------------|
| *Arg1*  | 可选      | Variant  | 仅应用于内置对话框，是命令的初始参数。有关详细信息，请参阅“注解”部分。 |
| *...*   | 可选      | Variant  | 仅应用于内置对话框，是命令的初始参数。有关详细信息，请参阅“注解”部分。 |
| *Arg30* | 可选      | Variant  | 仅应用于内置对话框，是命令的初始参数。有关详细信息，请参阅“注解”部分。 |

**说明**

对于内置对话框，如果用户单击 **“确定”** ，则本方法返回 **True** ；如果用户单击 **“取消”** ，则返回 **False** 。

可以使用单个对话框同时更改许多属性。例如，可以使用“设置单元格格式”对话框更改 Font 对象的所有属性。

对于一些内置对话框（如 **“打开”** 对话框），可以使用 *arg1* , *arg2* , ..., *arg30* 设置初始值。要查找想设置的参数，请在 **内置对话框参数列表** 中查找相应的对话框常量。例如，搜索 **xlDialogOpen** 常量以查找 **“打开”** 对话框的参数。有关内置对话框的详细信息，请参阅 Dialogs 集合。

**示例**

``` JavaScript
/* 本示例显示“打开”对话框*/
Application.Dialogs.Item(xlDialogOpen).Show()
```

## 成员属性

### Dialog.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Dialog 对象的变量。

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

### Dialog.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Dialog 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Dialog.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Dialog 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Dialog](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FormatConditions 对象

## 简介

代表一个区域内所有条件格式的集合。

## 说明

FormatConditions 集合可以包含多个条件格式。每个格式由一个 FormatCondition 对象代表。

有关条件格式的详细信息，请参阅 FormatCondition 对象。

使用 FormatConditions 属性可返回 FormatConditions 对象。使用 Add 方法可新建条件格式，使用 Mo dify 方法可更改现有的条件格式。

下例为 E1:E10 单元格添加条件格式。

``` JavaScript
function test() {
    let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Add(xlCellValue, xlGreater, "=$a$1")
    let range3 = range2.Borders
    range3.LineStyle = xlContinuous
    range3.Weight = xlThin
    range3.ColorIndex = 6
    let range4 = range2.Font
    range4.Bold = true
    range4.ColorIndex = 3
} 
```

## 方法

| 序号 | 名称                                                         | 说明                                                                                                                    |
|:----:|:-------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------|
|  1   | [Add](#FormatConditions.Add)                                 | 添加新的条件格式。返回 一个 FormatCondition 对象，它代表新的条件格式。                                                  |
|  2   | [AddAboveAverage](#FormatConditions.AddAboveAverage)         | 返回一个代表指定区域条件格式规则的新 AboveAverage 对象。返回一个 AboveAverage 对象                                      |
|  3   | [AddColorScale](#FormatConditions.AddColorScale)             | 返回一个代表条件格式规则的新 ColorSc ale 对象，该条件格式规则使用单元格颜色渐变来表示选定区域中所含单元格值的相对差异。 |
|  4   | [AddDatabar](#FormatConditions.AddDatabar)                   | 返回一个代表指定区域数据条条件格式规则的 Dat abar 对象。                                                                |
|  5   | [AddIconSetCondition](#FormatConditions.AddIconSetCondition) | 返回一个代表指定区域图标集条件格式规则的新 IconSetCon dition 对象。                                                     |
|  6   | [AddTop10](#FormatConditions.AddTop10)                       | 返回一个代表指定区域条件格式规则的 Top 10 对象。                                                                        |
|  7   | [AddUniqueValues](#FormatConditions.AddUniqueValues)         | 返回一个代表指定区域条件格式规则的新 Unique Values 对象。                                                               |
|  8   | [Delete](#FormatConditions.Delete)                           | 删除对象。                                                                                                              |
|  9   | [Item](#FormatConditions.Item)                               | 从集合中返回一个对象。                                                                                                  |

## 属性

| 序号 | 名称                                         | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FormatConditions.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#FormatConditions.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#FormatConditions.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#FormatConditions.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### FormatConditions.Add

添加新的条件格式。返回 一个 **FormatCondition** 对象，它代表新的条件格式。

**语法**

**express.Add ( *Type* , *Operator* , *Formula1* , *Formula2* )**

\*express是一个代表 FormatConditions 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型              | 说明                                                                                                                                                                                                           |
|------------|-----------|-----------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Type*     | 必选      | XlFormatConditionType | 指定条件格式是基于单元格值还是基于表达式。                                                                                                                                                                     |
| *Operator* | 可选      | Variant               | 条件格式运算符。可为以下 XlFormatConditionOperator 常量之一：xlBetween、xlEqual、xlGreater、xlGreaterEqual、xlLess、xlLessEqual、xlNotBetween 或 xlNotEqual。如果 Type 为 xlExpression，则忽略 Operator 参数。 |
| *Formula1* | 可选      | Variant               | 与条件格式相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。                                                                                                                                         |
| *Formula2* | 可选      | Variant               | 当 Operator 为 xlBetween 或 xlNotBetween 时，它是与条件格式第二部分相关联的值或表达式（否则忽略该参数）。可为常量值、字符串值、单元格引用或公式。                                                              |

**说明**

对单个区域定义的条件格式不能超过三个。使用 **Mo dify** 方法可修改现有的条件格式，使用 **D elete** 方法可在添加新条件格式前删除现有的格式。

**示例**

``` JavaScript
/* 本示例向单元格区域 E1:E10 中添加条件格式*/
function test() {
    let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Add(xlCellValue, xlGreater, "=$a$1")
    let range3 = range2.Borders
    range3.LineStyle = xlContinuous
    range3.Weight = xlThin
    range3.ColorIndex = 6
    let range4 = range2.Font
    range4.Bold = true
    range4.ColorIndex = 3
}
```

### FormatConditions.AddAboveAverage

返回一个代表指定区域条件格式规则的新 **AboveAverage** 对象。返回一个 **AboveAverage** 对象

**语法**

**express.AddAboveAverage ()**

\*express是一个代表 FormatConditions 对象的变量。

**说明**

**AboveAverage** 对象用于查找单元格区域中高于或低于平均或标准偏差的值。例如，可在年度业绩评估中查找高于平均业绩的人员，或者在质量评估中查找低于两个标准偏差的加工原料。

### FormatConditions.AddColorScale

返回一个代表条件格式规则的新 **ColorSc ale** 对象，该条件格式规则使用单元格颜色渐变来表示选定区域中所含单元格值的相对差异。

**语法**

**express.AddColorScale ( *ColorScaleType* )**

\*express是一个代表 FormatConditions 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明         |
|------------------|-----------|----------|--------------|
| *ColorScaleType* | 必选      | Long     | 色阶的类型。 |

**说明**

色阶是直观的参照，可以帮助您了解数据的分布和变化。色阶可帮助您利用颜色的变化来标出给定区域中单元格值的相对差异。不同的颜色和颜色之间的渐变代表单元格值的差异。例如，在三色色阶中，您可以指定包含最高相对数据值的单元格为绿色，包含中间值的单元格为黄色，包含最低值的单元格为红色。

### FormatConditions.AddDatabar

返回一个代表指定区域数据条条件格式规则的 **Dat abar** 对象。

**语法**

**express.AddDatabar ()**

\*express是一个代表 FormatConditions 对象的变量。

**说明**

数据条可帮助您查看某个单元格相对于其他单元格的值。数据条的长度代表单元格中的值。较长的数据条代表较高的值，较短的数据条代表较低的值。数据条在辨认较高和较低数值，特别是有大量数据存在的情况下很有帮助，比如，辨认假日销售报告中的最高玩具销量和最低玩具销量。

### FormatConditions.AddIconSetCondition

返回一个代表指定区域图标集条件格式规则的新 **IconSetCon dition** 对象。

**语法**

**express.AddIconSetCondition ()**

\*express是一个代表 FormatConditions 对象的变量。

**说明**

使用图标集为数据添加注释并将数据分为按阈值隔开的三到五类数据。每种图标均代表某一值范围。例如，在 3 箭头图标集中，红色向上箭头代表较高值，黄色侧箭头代表中间值，绿色向下箭头代表较低值。

### FormatConditions.AddTop10

返回一个代表指定区域条件格式规则的 **Top 10** 对象。

**语法**

**express.AddTop10 ()**

\*express是一个代表 FormatConditions 对象的变量。

**说明**

使用 **Top10** 对象，可以基于指定的截止值查找单元格区域中的最高值和最低值。例如，可以查找区域报告中位居前五位的销售产品、客户调查中位居最后百分之十五的产品，或者部门人员分析中位居前 25 位的薪金。

### FormatConditions.AddUniqueValues

返回一个代表指定区域条件格式规则的新 **Unique Values** 对象。

**语法**

**express.AddUniqueValues ()**

\*express是一个代表 FormatConditions 对象的变量。

**说明**

可以使用 **UniqueValues** 对象快速显现包含唯一值或重复值的单元格。

### FormatConditions.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 FormatConditions 对象的变量。

### FormatConditions.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 FormatConditions 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
/* 此示例为 E1:E10 单元格区域的现有条件格式设置格式属性*/
function test() {
    let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)
    let range3 = range2.Borders
    range3.LineStyle = xlContinuous
    range3.Weight = xlThin
    range3.ColorIndex = 6
}
```

## 成员属性

### FormatConditions.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 FormatConditions 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
```

### FormatConditions.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 FormatConditions 对象的变量。

### FormatConditions.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FormatConditions 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### FormatConditions.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FormatConditions 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/FormatConditions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Hyperlink 对象

## 简介

代表一个超链接。

## 说明

Hyperlink 对象是 Hyperlinks 集合的成员。

使用 Hyperlink 属性可返回某个形状的超链接（一个形状只能有一个超链接）。下例激活形状一的超链接。

``` JavaScript
Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink.Follow(true)
```

区域或工作表可以有多个超链接。使用 Hyperlinks ( *index* ) 可返回单个 Hyperlink 对象，其中 *index* 是超链接编号。下例激活 A1:B2 区域中的第二个超链接。

``` JavaScript
Application.Worksheets.Item(1).Range("A1:B2").Hyperlinks.Item(2).Follow()
```

## 方法

| 序号 | 名称                                              | 说明                                                                                                                                   |
|:----:|:--------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddToFavorites](#Hyperlink.AddToFavorites)       | 将工作簿或超链接的快捷方式添加到“收藏夹”文件夹。                                                                                       |
|  2   | [CreateNewDocument](#Hyperlink.CreateNewDocument) | 新建一篇链接到指定超链接的文档。                                                                                                       |
|  3   | [Delete](#Hyperlink.Delete)                       | 删除对象。                                                                                                                             |
|  4   | [Follow](#Hyperlink.Follow)                       | 如果已经下载指定文档，则显示缓冲区中的该文档。否则，本方法对指定超链接进行处理以下载目标文档，然后将该文档在适当的应用程序中显示出来。 |

## 属性

| 序号 | 名称                                      | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Address](#Hyperlink.Address)             | 返回或设置一个 String 型，它代表目标文档的地址。                                                                                                                                                                                |
|  2   | [Application](#Hyperlink.Application)     | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  3   | [Creator](#Hyperlink.Creator)             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [EmailSubject](#Hyperlink.EmailSubject)   | 返回或设置指定超链接的电子邮件主题行的文本字符串。主题行是添加到超链接地址上的。 String 类型，可读写。                                                                                                                          |
|  5   | [Name](#Hyperlink.Name)                   | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  6   | [Parent](#Hyperlink.Parent)               | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  7   | [Range](#Hyperlink.Range)                 | 返回一个 Range 对象，它代表附加的指定超链接的区域。                                                                                                                                                                             |
|  8   | [ScreenTip](#Hyperlink.ScreenTip)         | 返回或设置指定超链接的屏幕提示文字。 String 类型，可读写。                                                                                                                                                                      |
|  9   | [Shape](#Hyperlink.Shape)                 | 返回一个 Shape 对象，它代表附加到指定超链接的形状。                                                                                                                                                                             |
|  10  | [SubAddress](#Hyperlink.SubAddress)       | 返回或设置文档中与指定超链接相关联的位置。 String 类型，可读写。                                                                                                                                                                |
|  11  | [TextToDisplay](#Hyperlink.TextToDisplay) | 返回或设置要为指定超链接显示的文本。默认值为超链接的地址。 String 类型，可读写。                                                                                                                                                |
|  12  | [Type](#Hyperlink.Type)                   | 返回一个包含 MsoHyperlinkType 常量的 Long 值，它代表 HTML 框架的位置。                                                                                                                                                          |

## 成员方法

### Hyperlink.AddToFavorites

将工作簿或超链接的快捷方式添加到“收藏夹”文件夹。

**语法**

**express.AddToFavorites ()**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/* 此示例将活动工作簿的快捷方式添加到“收藏夹”文件夹*/
Application.ActiveWorkbook.AddToFavorites()
```

### Hyperlink.CreateNewDocument

新建一篇链接到指定超链接的文档。

**语法**

**express.CreateNewDocument ( *Filename* , *EditNow* , *Overwrite* )**

\*express是一个代表 Hyperlink 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                                        |
|-------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Filename*  | 必选      | String   | 指定文档的文件名。                                                                                                                                          |
| *EditNow*   | 必选      | Boolean  | 如果该值为 True，则立即在关联的编辑环境中打开指定文档。默认值为 True。                                                                                      |
| *Overwrite* | 必选      | Boolean  | 如果该值为 True，则覆盖相同文件夹中任何现有的同名文件。如果该值为 False，则保留任何现有的同名文件，并且参数 Filename 将指定一个新的文件名。默认值为 False。 |

**示例**

``` JavaScript
/* 本示例在第一张工作表中新建一个基于新的超链接的文档，然后将该文档加载到 ET 中进行编辑。该文档名为“Report.xls”，它将覆盖“\\Server1\Annual”文件夹中的所有同名文件*/
function test() {
  let sheet2 = Application.Worksheets.Item(1)
  let objHyper = sheet2.Hyperlinks.Add(sheet2.Range("A10"), "\\\Server1\\Annual\\Report.xls")
  objHyper.CreateNewDocument("\\\Server1\\Annual\\Report.xls", true, true)
}
```

### Hyperlink.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.Follow

如果已经下载指定文档，则显示缓冲区中的该文档。否则，本方法对指定超链接进行处理以下载目标文档，然后将该文档在适当的应用程序中显示出来。

**语法**

**express.Follow ( *NewWindow* , *AddHistory* , *ExtraInfo* , *Method* , *HeaderInfo* )**

\*express是一个代表 Hyperlink 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                           |
|--------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------|
| *NewWindow*  | 可选      | Variant  | 如果为 True，则在新窗口中显示目标应用程序。默认值为 False。                                                                    |
| *AddHistory* | 可选      | Variant  | 未使用。保留供将来使用。                                                                                                       |
| *ExtraInfo*  | 可选      | Variant  | 指定解析超链接时要使用的 HTTP 附加信息的 String 或字节数组。例如，可以使用 ExtraInfo 指定图像的坐标、窗体的内容或 FAT 文件名。 |
| *Method*     | 可选      | Variant  | 指定附加 ExtraInfo 的方法。可以是下列 MsoExtraInfoMethod 常量之一。                                                            |
| *HeaderInfo* | 可选      | Variant  | 指定 HTTP 请求的标题信息的 String。默认值为空字符串。                                                                          |

**示例**

``` JavaScript
/* 本示例加载附于第一个形状（第一张工作表中）的超链接的文档*/
Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink.Follow(true)
```

## 成员属性

### Hyperlink.Address

返回或设置一个 **String** 型，它代表目标文档的地址。

**语法**

**express.Address**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
```

### Hyperlink.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Hyperlink.EmailSubject

返回或设置指定超链接的电子邮件主题行的文本字符串。主题行是添加到超链接地址上的。 **String** 类型，可读写。

**语法**

**express.EmailSubject**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

该属性常与电子邮件超链接一起使用。

该属性的值优先于使用同一 **Hyperlink** 对象的 **Address** 属性指定的任何电子邮件主题行。

**示例**

``` JavaScript
/* 本示例设置第一张工作表中第一个超链接的电子邮件主题行*/
Application.Worksheets.Item(1).Hyperlinks.Item(1).EmailSubject = "Quote Request"
```

### Hyperlink.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.Range

返回一个 **Range** 对象，它代表附加的指定超链接的区域。

**语法**

**express.Range**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/* 下例将应用于“Crew”工作表的“自动筛选”的地址保存在变量中*/
/* 示例代码*/
let rAddress = Application.Worksheets.Item("Crew").AutoFilter.Range.Addres
/* 此示例滚动工作簿窗口，直至超链接区域位于活动窗口的左上角*/
/* 示例代码*/
Application.Workbooks.Item(1).Activate()
let hr = Application.ActiveSheet.Hyperlinks.Item(1).Range
Application.ActiveWindow.ScrollRow = hr.Row
Application.ActiveWindow.ScrollColumn = hr.Column
```

### Hyperlink.ScreenTip

返回或设置指定超链接的屏幕提示文字。 **String** 类型，可读写。

**语法**

**express.ScreenTip**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

在将文档保存到网页中之后，如果在 Web 浏览器中查看该文档，则当鼠标移动到超链接上时，“屏幕提示”文字就会显示。某些浏览器可能不支持“屏幕提示”。

**示例**

``` JavaScript
/* 本示例设置活动工作表上第一个超链接的屏幕提示*/
Application.ActiveSheet.Hyperlinks.Item(1).ScreenTip = "Return to the home page"
```

### Hyperlink.Shape

返回一个 **Shape** 对象，它代表附加到指定超链接的形状。

**语法**

**express.Shape**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.SubAddress

返回或设置文档中与指定超链接相关联的位置。 **String** 类型，可读写。

**语法**

**express.SubAddress**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/* 本示例将一个区域位置添加到第一张形状的超链接中*/
Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink.SubAddress = "A1:B10"
```

### Hyperlink.TextToDisplay

返回或设置要为指定超链接显示的文本。默认值为超链接的地址。 **String** 类型，可读写。

**语法**

**express.TextToDisplay**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/* 本示例为活动工作表上第一个超链接设置要显示的文本*/
Application.ActiveSheet.Hyperlinks.Item(1).TextToDisplay = "Company Home Page"
```

### Hyperlink.Type

返回一个包含 **MsoHyperlinkType** 常量的 **Long** 值，它代表 HTML 框架的位置。

**语法**

**express.Type**

\*express是一个代表 Hyperlink 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Hyperlink](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Interior 对象

## 简介

代表一个对象的内部。

## 说明

使用 Interior 属性可返回 Interior 对象。下例将单元格 A1 的内部设置为红色。

``` JavaScript
Application.Worksheets.Item("Sheet1").Range("A1").Interior.ColorIndex = 3
```

## 属性

| 序号 | 名称                                                 | 说明                                                                                                                                                                                                                        |
|:----:|:-----------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Interior.Application)                 | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 Application 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 Application 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
|  2   | [Color](#Interior.Color)                             | 返回或设置对象的主要颜色，如注释部分中的表格所示。使用 RGB 函数可创建颜色值。 Variant 型，可读写。                                                                                                                          |
|  3   | [ColorIndex](#Interior.ColorIndex)                   | 返回或设置一个 Variant 值，它代表内部颜色。                                                                                                                                                                                 |
|  4   | [Creator](#Interior.Creator)                         | 返回一个 32 位整数，该整数指示创建对象的应用程序。只读 Long 类型。                                                                                                                                                          |
|  5   | [Gradient](#Interior.Gradient)                       | 返回或设置所选内容的 Interior 对象的 Gradient 属性。只读。                                                                                                                                                                  |
|  6   | [InvertIfNegative](#Interior.InvertIfNegative)       | 如果 指定项与一个负数相对应时 ET 就将其反色，则为 True 。 Variant 类型，可读写。                                                                                                                                            |
|  7   | [Parent](#Interior.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                |
|  8   | [Pattern](#Interior.Pattern)                         | 返回或设置一个包含 xlPattern 常量的 Variant 值，它代表内部图案。                                                                                                                                                            |
|  9   | [PatternColor](#Interior.PatternColor)               | 以 RGB 值返回或设置内部图案的颜色。可读/写 Variant 类型。                                                                                                                                                                   |
|  10  | [PatternColorIndex](#Interior.PatternColorIndex)     | 返回或设置内部图案的颜色，表示为当前调色板中的颜色编号或下列 XlColorIndex 常量之一： xlColorIndexAutomatic 或 xlColorIndexNone 。 Long 类型，可读写。                                                                       |
|  11  | [PatternThemeColor](#Interior.PatternThemeColor)     | 返回或设置 Interior 对象的主题颜色图案。可读/写 Variant 类型。                                                                                                                                                              |
|  12  | [PatternTintAndShade](#Interior.PatternTintAndShade) | 返回或设置 Interior 对象的淡色和底纹图案。可读/写 Variant 类型。                                                                                                                                                            |
|  13  | [ThemeColor](#Interior.ThemeColor)                   | 返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 Variant 类型。                                                                                                                                  |
|  14  | [TintAndShade](#Interior.TintAndShade)               | 返回或设置一个 Single ，使颜色变深或变浅。                                                                                                                                                                                  |

## 成员属性

### Interior.Application

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 Interior 对象的变量。

**示例**

``` JavaScript
function test{
  /*本示例显示一条有关创建 myObject 的应用程序的消息。*/
  let myObject = Application.ActiveWorkbook
  if(myObject.Application.Value == "ET") {
      alert("This is an ET Application object.")
  }
  else {
      alert("This is not an ET Application object.")
  }
}
```

### Interior.Color

返回或设置对象的主要颜色，如注释部分中的表格所示。使用 **RGB** 函数可创建颜色值。 **Variant** 型，可读写。

**语法**

**express.Color**

\*express是一个代表 Interior 对象的变量。

**说明**

|              |                                                                                     |
|--------------|-------------------------------------------------------------------------------------|
| **Border**   | 边框的颜色。                                                                        |
| **Borders**  | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font**     | 字体的颜色。                                                                        |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。                                              |
| **Tab**      | 选项卡的颜色。                                                                      |

**示例**

``` JavaScript
/*此示例对 Chart1 中数值坐标轴的刻度线标志颜色进行设置。*/
Application.Charts.Item("Chart1").Axes(xlValue).TickLabels.Font.Color = (0, 255, 0)
```

### Interior.ColorIndex

返回或设置一个 **Variant** 值，它代表内部颜色。

**语法**

**express.ColorIndex**

\*express是一个代表 Interior 对象的变量。

**说明**

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 **XlColorIndex** 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

### Interior.Creator

返回一个 32 位整数，该整数指示创建对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Interior 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Interior.Gradient

返回或设置所选内容的 **Interior** 对象的 **Gradient** 属性。只读。

**语法**

**express.Gradient**

\*express是一个代表 Interior 对象的变量。

**说明**

使用 LinearGradient 或 RectangularGradient。

### Interior.InvertIfNegative

如果 指定项与一个负数相对应时 ET 就将其反色，则为 **True** 。 **Variant** 类型，可读写。

**语法**

**express.InvertIfNegative**

\*express是一个代表 Interior 对象的变量。

### Interior.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Interior 对象的变量。

### Interior.Pattern

返回或设置一个包含 **xlPattern** 常量的 **Variant** 值，它代表内部图案。

**语法**

**express.Pattern**

\*express是一个代表 Interior 对象的变量。

**示例**

``` JavaScript
/*本示例为 Sheet1 中 A1 单元格的内部添加十字图案。*/
Application.Worksheets.Item("Sheet1").Range("A1").Interior.Pattern = xlPatternCrissCross
```

### Interior.PatternColor

以 RGB 值返回或设置内部图案的颜色。可读/写 **Variant** 类型。

**语法**

**express.PatternColor**

\*express是一个代表 Interior 对象的变量。

**示例**

``` JavaScript
function test{
  /*本示例对 Sheet1 中矩形一的内部图案的颜色进行设置。*/
  let rectangles2 = Application.Worksheets.Item("Sheet1").Rectangles(1).Interior
  rectangles2.Pattern = xlGrid
  rectangles2.PatternColor = (255,0,0)
}
```

### Interior.PatternColorIndex

返回或设置内部图案的颜色，表示为当前调色板中的颜色编号或下列 **XlColorIndex** 常量之一： **xlColorIndexAutomatic** 或 **xlColorIndexNone** 。 **Long** 类型，可读写。

**语法**

**express.PatternColorIndex**

\*express是一个代表 Interior 对象的变量。

**说明**

将本属性设置为 **xlColorIndexAutomatic** ，表示单元格使用自动图案或图形对象使用自动填充样式。将本属性设置为 **xlColorIndexNone** ，则表示不希望使用图案（等价于将 **Interior** 对象的 **Pattern** 属性设置为 **xlPatternNone** ）。

**示例**

``` JavaScript
function test{
  /*本示例对 Sheet1 中矩形一的内部图案的颜色进行设置。*/
  let rectangles2 = Worksheets.Item("Sheet1").Rectangles(1).Interior
  rectangles2.Pattern = xlChecker
  rectangles2.PatternColorIndex = 5
}
```

### Interior.PatternThemeColor

返回或设置 **Interior** 对象的主题颜色图案。可读/写 **Variant** 类型。

**语法**

**express.PatternThemeColor**

\*express是一个代表 Interior 对象的变量。

### Interior.PatternTintAndShade

返回或设置 **Interior** 对象的淡色和底纹图案。可读/写 **Variant** 类型。

**语法**

**express.PatternTintAndShade**

\*express是一个代表 Interior 对象的变量。

### Interior.ThemeColor

返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。

**语法**

**express.ThemeColor**

\*express是一个代表 Interior 对象的变量。

### Interior.TintAndShade

返回或设置一个 **Single** ，使颜色变深或变浅。

**语法**

**express.TintAndShade**

\*express是一个代表 Interior 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Interior](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LineFormat 对象

## 简介

代表线条和箭头格式。

## 说明

对于线条， LineFormat 对象包含该线条自身的格式信息；对于有边界的形状，该对象包含形状的边界的格式信息。

使用 Line 属性可返回一个 LineFormat 对象。

``` JavaScript
/*下例给 myDocument 添加一条蓝色虚线。在该线的起点有一个短而窄的椭圆，在该线的终点有一个长而宽的三角形。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
line.DashStyle = msoLineDashDotDot
line.ForeColor.RGB = (50, 0, 128)
line.BeginArrowheadLength = msoArrowheadShort
line.BeginArrowheadStyle = msoArrowheadOval
line.BeginArrowheadWidth = msoArrowheadNarrow
line.EndArrowheadLength = msoArrowheadLong
line.EndArrowheadStyle = msoArrowheadTriangle
line.EndArrowheadWidth = msoArrowheadWide
}
```

## 属性

| 序号 | 名称                                                     | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#LineFormat.Application)                   | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [BackColor](#LineFormat.BackColor)                       | 返回或设置一个 ColorFormat 对象，它代表指定的填充背景色。                                                                                                                                                                       |
|  3   | [BeginArrowheadLength](#LineFormat.BeginArrowheadLength) | 返回或设置指定线条起点的箭头的长度。 MsoArrowheadLength 类型，可读写。                                                                                                                                                          |
|  4   | [BeginArrowheadStyle](#LineFormat.BeginArrowheadStyle)   | 返回或设置指定线条起点的箭头样式。 MsoArrowheadStyle 类型，可读写。                                                                                                                                                             |
|  5   | [BeginArrowheadWidth](#LineFormat.BeginArrowheadWidth)   | 返回或设置指定线条起点的箭头宽度。 MsoArrowheadWidth 类型，可读写。                                                                                                                                                             |
|  6   | [Creator](#LineFormat.Creator)                           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  7   | [DashStyle](#LineFormat.DashStyle)                       | 返回或设置指定直线的虚线样式。可以是 MsoLineDashStyle 常量之一。可读/写 Long 类型。                                                                                                                                             |
|  8   | [EndArrowheadLength](#LineFormat.EndArrowheadLength)     | 返回或设置指定线条末尾的箭头的长度。 MsoArrowheadLength 类型，可读写。                                                                                                                                                          |
|  9   | [EndArrowheadStyle](#LineFormat.EndArrowheadStyle)       | 返回或设置指定线条端点的箭头样式。 MsoArrowheadStyle 类型，可读写。                                                                                                                                                             |
|  10  | [EndArrowheadWidth](#LineFormat.EndArrowheadWidth)       | 返回或设置指定线条端点的箭头的宽度。 MsoArrowheadWidth 类型，可读写。                                                                                                                                                           |
|  11  | [ForeColor](#LineFormat.ForeColor)                       | 返回或设置一个 ColorFormat 对象，它代表指定的前景填充色或纯色。                                                                                                                                                                 |
|  12  | [InsetPen](#LineFormat.InsetPen)                         | 返回或设置线条是否绘制在指定形状的边界以内。可读/写。                                                                                                                                                                           |
|  13  | [Parent](#LineFormat.Parent)                             | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  14  | [Pattern](#LineFormat.Pattern)                           | 返回或设置一个代表填充图案的 MsoPatternType 值。                                                                                                                                                                                |
|  15  | [Style](#LineFormat.Style)                               | 返回或设置一个 MsoLineStyle 值，它代表线条的样式。                                                                                                                                                                              |
|  16  | [Transparency](#LineFormat.Transparency)                 | 返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。 Double 型，可读写。                                                                                                                                    |
|  17  | [Visible](#LineFormat.Visible)                           | 返回或设置一个 MsoTriState 值，它确定对象是否可见。可读写。                                                                                                                                                                     |
|  18  | [Weight](#LineFormat.Weight)                             | 返回或设置一个 Single 值，它代表线条的粗细。                                                                                                                                                                                    |

## 成员属性

### LineFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 LineFormat 对象的变量。

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

### LineFormat.BackColor

返回或设置一个 **ColorFormat** 对象，它代表指定的填充背景色。

**语法**

**express.BackColor**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.BeginArrowheadLength

返回或设置指定线条起点的箭头的长度。 **MsoArrowheadLength** 类型，可读写。

**语法**

**express.BeginArrowheadLength**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
line.BeginArrowheadLength = msoArrowheadShort
line.BeginArrowheadStyle = msoArrowheadOval
line.BeginArrowheadWidth = msoArrowheadNarrow
line.EndArrowheadLength = msoArrowheadLong
line.EndArrowheadStyle = msoArrowheadTriangle
line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.BeginArrowheadStyle

返回或设置指定线条起点的箭头样式。 **MsoArrowheadStyle** 类型，可读写。

**语法**

**express.BeginArrowheadStyle**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
line.BeginArrowheadLength = msoArrowheadShort
line.BeginArrowheadStyle = msoArrowheadOval
line.BeginArrowheadWidth = msoArrowheadNarrow
line.EndArrowheadLength = msoArrowheadLong
line.EndArrowheadStyle = msoArrowheadTriangle
line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.BeginArrowheadWidth

返回或设置指定线条起点的箭头宽度。 **MsoArrowheadWidth** 类型，可读写。

**语法**

**express.BeginArrowheadWidth**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
line.BeginArrowheadLength = msoArrowheadShort
line.BeginArrowheadStyle = msoArrowheadOval
line.BeginArrowheadWidth = msoArrowheadNarrow
line.EndArrowheadLength = msoArrowheadLong
line.EndArrowheadStyle = msoArrowheadTriangle
line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 LineFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### LineFormat.DashStyle

返回或设置指定直线的虚线样式。可以是 **MsoLineDashStyle** 常量之一。可读/写 **Long** 类型。

**语法**

**express.DashStyle**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加蓝色的虚线。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let line = myDocument.Shapes.AddLine(10, 10, 250, 250).Line
line.DashStyle = msoLineDashDotDot
line.ForeColor.RGB = RGB(50, 0, 128)
}
```

### LineFormat.EndArrowheadLength

返回或设置指定线条末尾的箭头的长度。 **MsoArrowheadLength** 类型，可读写。

**语法**

**express.EndArrowheadLength**

\*express是一个代表 LineFormat 对象的变量。

**说明**

|                                                                     |
|---------------------------------------------------------------------|
| **MsoArrowheadLength** 可以是下列 **MsoArrowheadLength** 常量之一。 |
| **msoArrowheadLengthMixed**                                         |
| **msoArrowheadShort**                                               |
| **msoArrowheadLengthMedium**                                        |
| **msoArrowheadLong**                                                |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
line.BeginArrowheadLength = msoArrowheadShort
line.BeginArrowheadStyle = msoArrowheadOval
line.BeginArrowheadWidth = msoArrowheadNarrow
line.EndArrowheadLength = msoArrowheadLong
line.EndArrowheadStyle = msoArrowheadTriangle
line.EndArrowheadWidth = msoArrowheadWideh
}
```

### LineFormat.EndArrowheadStyle

返回或设置指定线条端点的箭头样式。 **MsoArrowheadStyle** 类型，可读写。

**语法**

**express.EndArrowheadStyle**

\*express是一个代表 LineFormat 对象的变量。

**说明**

|                                                                   |
|-------------------------------------------------------------------|
| **MsoArrowheadStyle** 可以是下列 **MsoArrowheadStyle** 常量之一。 |
| **msoArrowheadNone**                                              |
| **msoArrowheadOval**                                              |
| **msoArrowheadStyleMixed**                                        |
| **msoArrowheadDiamond**                                           |
| **msoArrowheadOpen**                                              |
| **msoArrowheadStealth**                                           |
| **msoArrowheadTriangle**                                          |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
line.BeginArrowheadLength = msoArrowheadShort
line.BeginArrowheadStyle = msoArrowheadOval
line.BeginArrowheadWidth = msoArrowheadNarrow
line.EndArrowheadLength = msoArrowheadLong
line.EndArrowheadStyle = msoArrowheadTriangle
line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.EndArrowheadWidth

返回或设置指定线条端点的箭头的宽度。 **MsoArrowheadWidth** 类型，可读写。

**语法**

**express.EndArrowheadWidth**

\*express是一个代表 LineFormat 对象的变量。

**说明**

|                                                                   |
|-------------------------------------------------------------------|
| **MsoArrowheadWidth** 可以是下列 **MsoArrowheadWidth** 常量之一。 |
| **msoArrowheadNarrow**                                            |
| **msoArrowheadWidthMedium**                                       |
| **msoArrowheadWide**                                              |
| **msoArrowheadWidthMixed**                                        |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/
function test(){
let myDocument = Worksheets.Item(1)
let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
line.BeginArrowheadLength = msoArrowheadShort
line.BeginArrowheadStyle = msoArrowheadOval
line.BeginArrowheadWidth = msoArrowheadNarrow
line.EndArrowheadLength = msoArrowheadLong
line.EndArrowheadStyle = msoArrowheadTriangle
line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.ForeColor

返回或设置一个 **ColorFormat** 对象，它代表指定的前景填充色或纯色。

**语法**

**express.ForeColor**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.InsetPen

返回或设置线条是否绘制在指定形状的边界以内。可读/写。

**语法**

**express.InsetPen**

\*express是一个代表 LineFormat 对象的变量。

**说明**

如果线条绘制在形状的边界以内，则为 **msoTrue** (-1)；否则为 **msoFalse** (0)。

**示例**

``` JavaScript
/*下面的代码示例将两个矩形添加到活动工作表，第一个矩形的线条绘制在其边界以内，第二个矩形的线条绘制在其边界上。*/
function test(){
let shpNew = Application.ActiveSheet.Shapes.AddShape(msoShapeRectangle, 200, 150, 150, 100)
shpNew.Line.Weight = 24
shpNew.Line.InsetPen = msoTrue

let shpNew2 = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 200, 150, 150, 100)
shpNew2.Line.Weight = 24
shpNew2.Line.InsetPen = msoFalse
}
```

### LineFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Pattern

返回或设置一个代表填充图案的 **MsoPatternType** 值。

**语法**

**express.Pattern**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Style

返回或设置一个 **MsoLineStyle** 值，它代表线条的样式。

**语法**

**express.Style**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Transparency

返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。 **Double** 型，可读写。

**语法**

**express.Transparency**

\*express是一个代表 LineFormat 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

### LineFormat.Visible

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Weight

返回或设置一个 **Single** 值，它代表线条的粗细。

**语法**

**express.Weight**

\*express是一个代表 LineFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/LineFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ModuleView 对象

## 简介

代表工作簿中的现有 模块 视图。

## 说明

模块区包含图表的绘图区在内的一切内容。但是，绘图区具有它自己的填充方式，因此填充绘图区并不会填充模块区。

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                               |
|:----:|:---------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ModuleView.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#ModuleView.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Parent](#ModuleView.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  4   | [Sheet](#ModuleView.Sheet)             | 返回指定的 ModuleView 对象的工作表名称。只读。                                                                                                                     |

## 成员属性

### ModuleView.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ModuleView 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### ModuleView.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ModuleView 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ModuleView.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ModuleView 对象的变量。

### ModuleView.Sheet

返回指定的 **ModuleView** 对象的工作表名称。只读。

**语法**

**express.Sheet**

\*express是一个代表 ModuleView 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ModuleView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Names 对象

## 简介

应用程序或工作簿中所有 Name 对象的集合。

## 说明

每一个 Name 对象都代表一个单元格区域的定义名称。名称可以是内置名称（如“Database”、“Print_Area”和“Auto_Open”）或自定义名称。

*RefersTo* 参数必须以 A1 样式表示法指定，包括必要时使用的美元符 (\$)。例如，如果在 Sheet1 上选定了单元格 A10，并且通过将 *RefersTo* 参数“=Sheet1!A1:B1”而定义了一个名称，那么该新名称实际上指向单元格区域 A10:B10（因为指定的是相对引用）。若要指定绝对引用，请使用“=Sheet1!\$A\$1:\$B\$1”。

## 示例

使用 Names 属性可返回 Names 集合。下例创建活动工作簿中所有名称及其引用地址的列表。

``` JavaScript
function test() {
    let nms = Application.ActiveWorkbook.Names
    let wks = Application.Worksheets.Item(1)
    for(let r = 1; r <= nms.Count; r++){
        wks.Cells.Item(r, 2).Value2 = nms.Item(r).Name
        wks.Cells.Item(r, 3).Value2 = nms.Item(r).RefersToRange.Address
    }
} 
```

使用 Add 方法可创建一个名称并将它添加到集合。下例创建一个新名称，该名称引用名为“Sheet1”的工作表上的单元格 A1:C20。

``` JavaScript
function test() {
    Application.Names.Add ("test1", "=sheet1!$a$1:$c$20")
} 
```

使用 Names ( *index* )（其中 *index* 是名称索引号或定义名称）可返回一个 Name 对象。下例从活动工作簿中删除名称“mySortRange”。

``` JavaScript
function test() {
    Application.ActiveWorkbook.Names.Item("mySortRange").Delete()
} 
```

## 方法

| 序号 | 名称                | 说明                              |
|:----:|:--------------------|:----------------------------------|
|  1   | [Add](#Names.Add)   | 为单元格区域定义新名称。          |
|  2   | [Item](#Names.Item) | 从 Names 集合返回一个 Name 对象。 |

## 属性

| 序号 | 名称                              | 说明                                                                                                                                                                                                                        |
|:----:|:----------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Names.Application) | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 Application 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 Application 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
|  2   | [Count](#Names.Count)             | 返回一个 Number 值，它代表集合中对象的数量。                                                                                                                                                                                |
|  3   | [Parent](#Names.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                |

## 成员方法

### Names.Add

为单元格区域定义新名称。

**语法**

**express.Add ( *Name* , *RefersTo* , *Visible* , *MacroType* , *ShortcutKey* , *Category* , *NameLocal* , *RefersToLocal* , *CategoryLocal* , *RefersToR1C1* , *RefersToR1C1Local* )**

\*express是一个代表 Names 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                   |
|---------------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Name*              | 可选      | Variant  | 如果未指定 NameLocal 参数，则指定要用作名称的英文文本。名称不能包括空格，并且不能设置为单元格引用的格式。                                                                              |
| *RefersTo*          | 可选      | Variant  | 如果未指定 RefersToLocal、RefersToR1C1 和 RefersToR1C1Local 参数，则说明名称引用的内容（使用 A1 格式表示法以英文表示）。 注释 如果引用不存在，则返回 Nothing。                         |
| *Visible*           | 可选      | Variant  | True 指定将名称定义为可见。False 指定将名称定义为隐藏。已隐藏的名称不会在“定义名称”、“粘贴名称”或“转到”对话框中显示。默认值为 True。                                                   |
| *MacroType*         | 可选      | Variant  | 由以下值之一确定的宏类型： 1 - 用户定义函数（Function 过程） 2 - 宏（Sub 过程） 3 或省略 - 无（该名称不引用用户定义函数或宏）                                                          |
| *ShortcutKey*       | 可选      | Variant  | 指定宏的快捷键。必须是单个字母，例如“z”或“Z”。仅适用于命令宏。                                                                                                                         |
| *Category*          | 可选      | Variant  | 如果 MacroType 参数等于 1 或 2，则此参数为宏或函数的分类。该分类在“函数向导”中使用。可以用数字（从 1 开始）或名称（以英文指定）引用现有的分类。如果指定的分类不存在，ET 将创建新分类。 |
| *NameLocal*         | 可选      | Variant  | 如果未指定 Name 参数，则指定要用作名称的本地化的文本。名称不能包括空格，并且不能设置为单元格引用的格式。                                                                               |
| *RefersToLocal*     | 可选      | Variant  | 如果未指定 RefersTo、RefersToR1C1 和 RefersToR1C1Local 参数，则说明名称引用的内容（使用 A1 格式表示法以本地化的文本表示）。                                                            |
| *CategoryLocal*     | 可选      | Variant  | 如果未指定 Category 参数，则指定标识自定义函数分类的本地化的文本。                                                                                                                     |
| *RefersToR1C1*      | 可选      | Variant  | 如果未指定 RefersTo、RefersToLocal 和 RefersToR1C1Local 参数，则说明名称引用的内容（使用 R1C1 格式表示法以英文表示）。                                                                 |
| *RefersToR1C1Local* | 可选      | Variant  | 如果未指定 RefersTo、RefersToLocal 和 RefersToR1C1 参数，则说明名称引用的内容（使用 R1C1 格式表示法以本地化的文本表示）。                                                              |

**示例**

``` JavaScript
/* 此示例为活动工作簿中 Sheet1 上的区域 A1:D3 定义一个新名称。 */
/* 注释:如果工作表 Sheet1 不存在，则返回 Nothing。 */
function test() {
    Application.ActiveWorkbook.Names.Add("tempRange", "=Sheet1!$A$1:$D$3")
}
```

### Names.Item

从 **Names** 集合返回一个 **Name** 对象。

**语法**

**express.Item ( *Index* , *IndexLocal* , *RefersTo* )**

\*express是一个代表 Names 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                           |
|--------------|-----------|----------|----------------------------------------------------------------|
| *Index*      | 可选      | Variant  | 要返回的定义名称的名称或编号。                                 |
| *IndexLocal* | 可选      | Variant  | 以用户语言表示的定义名称的名称。如果使用该参数，将不转换名称。 |
| *RefersTo*   | 可选      | Variant  | 名称引用的内容。使用该参数以引用的内容标识名称。               |

**说明**

必须只指定这三个参数中的一个。

**示例**

``` JavaScript
/* 此示例将活动工作簿中的 mySortRange 名称删除。 */
function test() {
    Application.ActiveWorkbook.Names.Item("mySortRange").Delete()
}
```

## 成员属性

### Names.Application

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 Names 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息。 */
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "Microsoft Excel") {
        alert("This is a Microsoft Excel Application object.")
    }
    else {
        alert("This is not a Microsoft Excel Application object.")
    }
}
```

### Names.Count

返回一个 **Number** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Names 对象的变量。

### Names.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Names 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Names](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotFilters 对象

## 简介

PivotFilters 对象是 PivotFilter 对象的集合。

## 说明

PivotFilters 集合包含用于添加新筛选器、计算集合中现有筛选器个数以及引用特定 PivotFilter 对象的属性和方法。

在以下示例中，将向位于当前活动单元格位置的 PivotField 添加新的 PivotFilter。

``` JavaScript
ActiveCell.PivotField.PivotFilters.Add(xlThisWeek)
```

## 方法

| 序号 | 名称                     | 说明                               |
|:----:|:-------------------------|:-----------------------------------|
|  1   | [Add](#PivotFilters.Add) | 将新筛选添加到 PivotFilters 集合。 |

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
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilters.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 <span>Application</span> 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilters.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回 PivotFilters 集合中的项数。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilters.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilters.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>根据 PivotFilters 集合对象特定元素在集合中的位置返回该特定元素。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilters.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定 PivotFilters 对象的父对象。只读。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotFilters.Add

将新筛选添加到 **PivotFilters** 集合。

**语法**

**express.Add ( *Type* , *DataField* , *Value1* , *Value2* , *Order* , *Name* , *Description* , *MemberPropertyField* )**

\*express是一个代表 PivotFilters 对象的变量。

**参数**

| 名称                  | 必选/可选 | 数据类型          | 说明                                |
|-----------------------|-----------|-------------------|-------------------------------------|
| *Type*                | 必选      | XlPivotFilterType | 要求 XlPivotFilterType 类型的筛选。 |
| *DataField*           | 可选      | Variant           | 筛选附加到的字段。                  |
| *Value1*              | 可选      | Variant           | 筛选值 1。                          |
| *Value2*              | 可选      | Variant           | 筛选值 1。                          |
| *Order*               | 可选      | Variant           | 筛选数据的顺序。                    |
| *Name*                | 可选      | Variant           | 筛选的名称。                        |
| *Description*         | 可选      | Variant           | 筛选的简要说明                      |
| *MemberPropertyField* | 可选      | Variant           | 指定标签筛选所基于的成员属性字段。  |

**返回值**

PivotFilter

**示例**

下面是一些关于如何正确使用 **Add** 函数的示例。

``` JavaScript
ActiveCell.PivotField.PivotFilters.Add(xlThisWeek)

ActiveCell.PivotField.PivotFilters.Add(xlTopCount,MyPivotField2,10)

ActiveCell.PivotField.PivotFilters.Add(xlCaptionIsNotBetween,null,"A","G")

ActiveCell.PivotField.PivotFilters.Add(xlValueIsGreaterThanOrEqualTo,MyPivotField2,10000)
```

由于 *Value1* 的数据类型无效，本示例将返回一个运行时错误。

``` JavaScript
ActiveCell.PivotField.PivotFilters.Add(xlValueIsGreaterThanOrEqualTo,MyPivotField2,“Allan”)
```

## 成员属性

### PivotFilters.Application

table { word-break:break-all; }

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotFilters 对象的变量。

**说明**

table { word-break:break-all; }

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### PivotFilters.Count

table { word-break:break-all; }

返回 **PivotFilters** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 PivotFilters 对象的变量。

### PivotFilters.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotFilters 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotFilters.Item

table { word-break:break-all; }

根据 **PivotFilters** 集合对象特定元素在集合中的位置返回该特定元素。只读。

**语法**

**express.Item**

\*express是一个代表 PivotFilters 对象的变量。

### PivotFilters.Parent

table { word-break:break-all; }

返回指定 **PivotFilters** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotFilters 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotFilters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotItemList 对象

## 简介

指定的数据透视表中所有 PivotItem 对象的集合。

## 说明

每个 PivotItem 代表数据透视表字段中的一个项。

使用 PivotCell 对象的 RowItems 或 ColumnItems 属性可返回一个 PivotItemList 集合。

一旦返回了 PivotItemList 集合，您就可以使用 Item 方法来标识某个特定的 PivotItem 列表。下例向用户显示与单元格 B5 相关联的 PivotItem 列表。本示例假定数据透视表位于活动工作表上。

``` JavaScript
function CheckPivotItemList() {
    // Identify contents associated with PivotItemList.
    MsgBox("Contents associated with cell B5: " + Application.Range("B5").PivotCell.RowItems.Item(1))
}
```

## 方法

| 序号 | 名称                        | 说明                   |
|:----:|:----------------------------|:-----------------------|
|  1   | [Item](#PivotItemList.Item) | 从集合中返回一个对象。 |

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
<td style="text-align: left;" data-valian="middle"><a href="#PivotItemList.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItemList.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 Long 值，它代表集合中对象的数量。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItemList.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItemList.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定对象的父对象。只读。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotItemList.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 PivotItemList 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 PivotItem 对象。

**说明**

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

本示例隐藏第一个计算项。

``` JavaScript
Worksheets.Item(1).PivotTables(1).PivotFields("year").CalculatedItems().Item(1).Visible = false
```

## 成员属性

### PivotItemList.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotItemList 对象的变量。

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

### PivotItemList.Count

table { word-break:break-all; }

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 PivotItemList 对象的变量。

### PivotItemList.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotItemList 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotItemList.Parent

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotItemList 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotItemList](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotItems 对象

## 简介

数据透视表字段中所有 PivotItem 对象的集合。

## 说明

这些项目是某个字段类型中的各个数据项。

使用 PivotItems 方法可返回 PivotItems 集合。下例为 Sheet4 上第一张数据透视表创建那些字段中包含的字段名和项目的枚举列表。

``` JavaScript
Worksheets.Item("sheet4").Activate()
let pTab = Worksheets.Item("sheet3").PivotTables(1)
let c = 1
for(let i=1;i <= pTab.PivotFields().Count;i++) {
    let r = 1
    Cells.Item(r, c).Value2 = pTab.PivotFields(i).Name
    r++
    for(let x=1;x <= pTab.PivotFields(i).PivotItems.Count;x++) {
        Cells.Item(r, c).Value2 = pTab.PivotFields(i).PivotItems(x).Name
        r++
    }
    c++
}
```

使用 PivotItems ( *index* )（其中 *index* 是项目索引号或名称）可返回一个 PivotItem 对象。下例隐藏“Sheet3”上第一张数据透视表中“Year”字段中包含“1998”的所有数据项。

``` JavaScript
Worksheets.Item("sheet3").PivotTables(1).PivotFields("year").PivotItems("1998").Visible = false
```

## 方法

| 序号 | 名称                     | 说明                   |
|:----:|:-------------------------|:-----------------------|
|  1   | [Add](#PivotItems.Add)   | 新建数据透视表项。     |
|  2   | [Item](#PivotItems.Item) | 从集合中返回一个对象。 |

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
<td style="text-align: left;" data-valian="middle"><a href="#PivotItems.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItems.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 Long 值，它代表集合中对象的数量。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItems.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItems.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定对象的父对象。只读。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotItems.Add

新建数据透视表项。

**语法**

**express.Add ( *Name* )**

\*express是一个代表 PivotItems 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| *Name* | 必选      | String   | 新数据透视表项的名称。 |

**示例**

此示例在第一张工作表上为第一个数据透视表创建一个新数据透视表项。

``` JavaScript
Worksheets.Item(1).PivotTables(1).PivotItems("Year").Add("1998")
```

### PivotItems.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 PivotItems 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

一个 Object 值，它代表集合包含的某个对象。

**说明**

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

此示例隐藏工作表 Sheet3 中第一个数据透视表的数据项 1998。

``` JavaScript
Worksheets.Item("sheet3").PivotTables(1).PivotFields("year").PivotItems().Item("1998").Visible = false
```

## 成员属性

### PivotItems.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotItems 对象的变量。

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

### PivotItems.Count

table { word-break:break-all; }

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 PivotItems 对象的变量。

### PivotItems.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotItems 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotItems.Parent

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotItems 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotItems](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# VPageBreak 对象

## 简介

代表一个垂直分页符。

## 说明

VPageBreak 对象是 VPageBreaks 集合的成员。

使用 VPageBreaks ( *index* )（其中 *index* 是该分页符的分页符索引号）可返回一个 VPageBreak 对象。下例更改第一个垂直分页符的位置。

``` JavaScript
Application.Worksheets.Item(1).VPageBreaks.Item(1).Location = Worksheets.Item(1).Range("e5")
```

## 方法

| 序号 | 名称                           | 说明                       |
|:----:|:-------------------------------|:---------------------------|
|  1   | [Delete](#VPageBreak.Delete)   | 删除对象。                 |
|  2   | [DragOff](#VPageBreak.DragOff) | 将一个分页符拖出打印区域。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#VPageBreak.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#VPageBreak.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Extent](#VPageBreak.Extent)           | 返回指定分页符的类型：全屏或仅在打印区域内。可为以下任一 XlPageBreakExtent 常量： xlPageBreakFull 或 xlPageBreakPartial 。 Long 类型，只读。                                                                                    |
|  4   | [Location](#VPageBreak.Location)       | 返回或设置定义分页符位置的单元格（ Range 对象）。水平分页符与定位单元格的上边缘对齐；垂直分页符与定位单元格的左边缘对齐。 Range 类型，可读写。                                                                                  |
|  5   | [Parent](#VPageBreak.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [Type](#VPageBreak.Type)               | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### VPageBreak.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 VPageBreak 对象的变量。

### VPageBreak.DragOff

将一个分页符拖出打印区域。

**语法**

**express.DragOff ( *Direction* , *RegionIndex* )**

\*express是一个代表 VPageBreak 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型    | 说明                                                                                                                                                           |
|---------------|-----------|-------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Direction*   | 必选      | XlDirection | 分页符拖动方向。                                                                                                                                               |
| *RegionIndex* | 必选      | Long        | 分页符的打印区域索引（当用户按下鼠标按钮拖动分页符时鼠标指针所在的位置）。如果打印区域是连续的，则只有一个打印区域。如果打印区域不是连续的，则有多个打印区域。 |

**说明**

该方法主要用于宏记录器。使用 **De lete** 方法可在 JS 中删除分页符。

**示例**

``` JavaScript
/* 本示例将活动工作表的第一个垂直分页符拖出第一个打印区域的右边界，以删除该分页符*/
​Application.ActiveSheet.VPageBreaks.Item(1).DragOff(xlToRight, 1)
```

## 成员属性

### VPageBreak.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 VPageBreak 对象的变量。

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

### VPageBreak.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 VPageBreak 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### VPageBreak.Extent

返回指定分页符的类型：全屏或仅在打印区域内。可为以下任一 **XlPageBreakExtent** 常量： **xlPageBreakFull** 或 **xlPageBreakPartial** 。 **Long** 类型，只读。

**语法**

**express.Extent**

\*express是一个代表 VPageBreak 对象的变量。

**示例**

``` JavaScript
/* 本示例显示全屏水平分页符和打印区水平分页符的总数*/
function test() {
    let cFull = 0
    let cPartial = 0
    for(let i = 1; i <= Application.Worksheets.Item(1).HPageBreaks.Count; i++){
        if(Application.Worksheets.Item(1).HPageBreaks.Item(i).Extent == xlPageBreakFull){
            cFull = cFull + 1
        }
        else{
            cPartial = cPartial + 1
        }
    }
    alert(cFull + " full-screen page breaks, " + cPartial + " print-area page breaks")
}
```

### VPageBreak.Location

返回或设置定义分页符位置的单元格（ **Range** 对象）。水平分页符与定位单元格的上边缘对齐；垂直分页符与定位单元格的左边缘对齐。 **Range** 类型，可读写。

**语法**

**express.Location**

\*express是一个代表 VPageBreak 对象的变量。

**示例**

``` JavaScript
/* 本示例移动垂直分页符的位置*/
Application.Worksheets.Item(1).VPageBreaks.Item(1).Location = Worksheets.Item(1).Range("e5")
```

### VPageBreak.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 VPageBreak 对象的变量。

### VPageBreak.Type

返回指定对象的父对象。只读。

**语法**

**express.Type**

\*express是一个代表 VPageBreak 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/VPageBreak](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Border 对象

## 简介

代表一个对象的一个边框。 Border 对象是 Borders 集合的成员。

## 说明

使用 Borders ( *index* ) 来返回单个 Border 对象（其中 *index* 标识边框）。Index 可以是某个 WdBorderType 常量。由于选择或安装的语言支持（如美国英语）不同，有些 WdBorderType 常量可能无法使用。

使用 LineStyle 属性将边框线应用于 Border 对象。以下示例在活动文档的第一段下添加双线边框。

``` JavaScript
function test() {
    let borders = Application.ActiveDocument.Paragraphs.Item(1).Borders.Item(wdBorderBottom)
    borders.LineStyle = wdLineStyleDouble
    borders.LineWidth = wdLineWidth025pt
}
```

以下示例将单线边框加在选定内容的首字符四周。

``` JavaScript
function test() {
    let characters = Application.Selection.Characters.Item(1)
    characters.Font.Size = 36
    characters.Borders.Enable = true 
}
```

以下示例在第一节中每一页的四周添加艺术型边框。

``` JavaScript
function test() {
    let borders = Application.ActiveDocument.Sections.Item(1).Borders
    for(let i=1;i<=borders.Count;++i){
        borders.Item(i).ArtStyle = wdArtSeattle
        borders.Item(i).ArtWidth = 20
    } 
}
```

不能将 Border 对象添加到 Borders 集合。 Borders 集合中的成员数是有限的，具体数目因对象类型而异。例如，表格在 Borders 集合中有六个元素，而段落有四个。

## 属性

| 序号 | 名称                               | 说明                                                                         |
|:----:|:-----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Border.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                         |
|  2   | [ArtStyle](#Border.ArtStyle)       | 返回或设置文档的艺术型页面边框。 WdPageBorderArt 类型，可读写。              |
|  3   | [ArtWidth](#Border.ArtWidth)       | 返回或设置指定艺术型边框的宽度（以磅为单位）。 Number 类型，可读写。         |
|  4   | [Color](#Border.Color)             | 返回或设置指定 Border 对象的 24 位颜色。                                     |
|  5   | [ColorIndex](#Border.ColorIndex)   | 返回或设置指定边框或字体对象的颜色。 WdColorIndex 类型，可读写。             |
|  6   | [Creator](#Border.Creator)         | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。 |
|  7   | [Inside](#Border.Inside)           | 如果指定对象可以使用内部边框，则该属性值为 True 。 Boolean 类型，只读。      |
|  8   | [LineStyle](#Border.LineStyle)     | 返回或设置指定对象的边框样式。 WdLineStyle 类型，可读写。                    |
|  9   | [LineWidth](#Border.LineWidth)     | 返回或设置对象边框的线条宽度。可读写。                                       |
|  10  | [Parent](#Border.Parent)           | 返回一个 Object 类型的值，该值代表指定 Border 对象的父对象。                 |
|  11  | [Visible](#Border.Visible)         | 如果指定对象可见，则该属性值为 True 。 Boolean 类型，可读写。                |

## 成员属性

### Border.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Border 对象的变量。

### Border.ArtStyle

返回或设置文档的艺术型页面边框。 **WdPageBorderArt** 类型，可读写。

**语法**

**express.ArtStyle**

\*express是一个代表 Border 对象的变量。

**示例**

``` JavaScript
/* 本示例为选定内容的第一节的每个页面添加由黑点构成的边框 */
function test() {
    for(let i=1;i <= Application.Selection.Sections.Item(1).Borders.Count;i++) {
        Application.Selection.Sections.Item(1).Borders.Item(i).ArtStyle = wdArtBasicBlackDots
        Application.Selection.Sections.Item(1).Borders.Item(i).ArtWidth = 6
    }
}

/* 本示例为活动文档中的第一节的每个页面添加由特定图片所构成的边框 */
function test() {
    let Sec = Application.ActiveDocument.Sections.Item(1)
    Sec.Borders.AlwaysInFront = true
    for(let i=1;i <= Application.ActiveDocument.Sections.Item(1).Borders.Count;i++) {
        Application.ActiveDocument.Sections.Item(1).Borders.Item(i).ArtStyle = wdArtPeople
        Application.ActiveDocument.Sections.Item(1).Borders.Item(i).ArtWidth = 15
    }
}
```

### Border.ArtWidth

返回或设置指定艺术型边框的宽度（以磅为单位）。 **Number** 类型，可读写。

**语法**

**express.ArtWidth**

\*express是一个代表 Border 对象的变量。

**示例**

``` JavaScript
/* 本示例为选定内容的第一节的每个页面添加由黑点构成的宽为 6 磅的边框 */
function test() {
    for(let i=1;i <= Application.Selection.Sections.Item(1).Borders.Count;i++) {
        let Bor = Application.Selection.Sections.Item(1).Borders.Item(i)
        Bor.ArtStyle = wdArtBasicBlackDots
        Bor.ArtWidth = 6
    }
}
```

### Border.Color

返回或设置指定 **Border** 对象的 24 位颜色。

**语法**

**express.Color**

\*express是一个代表 Border 对象的变量。

**说明**

可以是任何有效的 **WdColor** 常量

**示例**

``` JavaScript
/* 本示例在第一个表格中每个单元格周围添加靛蓝色的点划线边框 */
function test() {
    if(ActiveDocument.Tables.Count >= 1) {
        for(let i=1;i <= ActiveDocument.Tables.Item(1).Borders.Count;i++) {
            ActiveDocument.Tables.Item(1).Borders.Item(i).Color = wdColorIndigo
            ActiveDocument.Tables.Item(1).Borders.Item(i).LineStyle = wdLineStyleDashDot
            ActiveDocument.Tables.Item(1).Borders.Item(i).LineWidth = wdLineWidth075pt
        }
    }
}
```

### Border.ColorIndex

返回或设置指定边框或字体对象的颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.ColorIndex**

\*express是一个代表 Border 对象的变量。

**说明**

**wdByAuthor** 常量对边框和字体对象无效。

**示例**

``` JavaScript
/* 本示例为第一个表格的每个单元格添加红虚线边框 */
function test() {
    if(Application.ActiveDocument.Tables.Count >= 1) {
        for(let i=1;i <= Application.ActiveDocument.Tables.Item(1).Borders.Count;i++) {
            let Bor = Application.ActiveDocument.Tables.Item(1).Borders.Item(i)
            Bor.ColorIndex = wdRed
            Bor.LineStyle = wdLineStyleDashDot
            Bor.LineWidth = wdLineWidth075pt
        }
    }
}
```

### Border.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Border 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

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

### Border.Inside

如果指定对象可以使用内部边框，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.Inside**

\*express是一个代表 Border 对象的变量。

**示例**

``` JavaScript
/* 如果当前选定内容支持内部边框（即选定了多个段落或单元格），则本示例为其设置单线型内部边框 */
function test() {
    for(let i=1;i <= Application.Selection.Borders.Count;i++) {
        if(Application.Selection.Borders.Item(i).Inside == true) {
            Application.Selection.Borders.Item(i).LineStyle = wdLineStyleSingle
        }
    }
}
```

### Border.LineStyle

返回或设置指定对象的边框样式。 **WdLineStyle** 类型，可读写。

**语法**

**express.LineStyle**

\*express是一个代表 Border 对象的变量。

**说明**

如果是为单个字符或单词区域设置 **LineStyle** 属性，应选用字符边框。

如果是为一个或几个段落区域设置 **LineStyle** 属性，应选用段落边框。连续段落之间的边框可应用 **InsideLineStyle** 属性。

如果是为节设置 **LineStyle** 属性，应选用节中的页面边框。

**示例**

``` JavaScript
/* 如果所选内容是一个段落，或折叠起来的选定内容，本示例将在当前段落之前添加 0.75 磅宽的单线段落边框。如果所选内容中不包括段落，则为所选文字添加边框 */
function test() {
    let Bor = Application.Selection.Borders.Item(wdBorderTop)
    Bor.LineStyle = wdLineStyleSingle
    Bor.LineWidth = wdLineWidth075pt
}
```

### Border.LineWidth

返回或设置对象边框的线条宽度。可读写。

**语法**

**express.LineWidth**

\*express是一个代表 Border 对象的变量。

**示例**

``` JavaScript
/* 本示例为活动文档的第一个表格的第一行添加下边框 */
function test() {
    if(Application.ActiveDocument.Tables.Count >= 1) {
        let Bor = Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Borders.Item(wdBorderBottom)
        Bor.LineStyle = wdLineStyleDouble
        Bor.LineWidth = wdLineWidth050pt
    }
}
```

### Border.Parent

返回一个 **Object** 类型的值，该值代表指定 **Border** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Border 对象的变量。

### Border.Visible

如果指定对象可见，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Visible**

\*express是一个代表 Border 对象的变量。

**说明**

对于任何对象，如果其 **Visible** 属性值为 **False** ，则某些方法和属性可能无法使用。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Border](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Documents 对象

## 简介

当前在 WPS 中打开的所有 Document 对象的集合。

## 方法

| 序号 | 名称                                                | 说明                                                                              |
|:----:|:----------------------------------------------------|:----------------------------------------------------------------------------------|
|  1   | [Add](#Documents.Add)                               | 返回一个 Document 对象，该对象代表添加到打开文档集合中的新建空文档。              |
|  2   | [AddBlogDocument](#Documents.AddBlogDocument)       | 返回一个 Document 对象，该对象表示 WPS 向前三个参数所描述的帐户发布的新博客文档。 |
|  3   | [CanCheckOut](#Documents.CanCheckOut)               | 如果该属性值为 True ，则 WPS 可将指定的文档从服务器签出。可读/写 Boolean 类型。   |
|  4   | [Close](#Documents.Close)                           | 关闭指定的文档。                                                                  |
|  5   | [Item](#Documents.Item)                             | 返回集合中的单个 Document 对象。                                                  |
|  6   | [Open](#Documents.Open)                             | 打开指定的文档并将其添加到 Documents 集合。返回一个 Document 对象。               |
|  7   | [OpenNoRepairDialog](#Documents.OpenNoRepairDialog) | 打开指定的文档并将其添加到 Documents 集合。                                       |
|  8   | [Parent](#Documents.Parent)                         | 返回一个 Object 类型值，该值代表指定 Documents 对象的父对象。                     |
|  9   | [Save](#Documents.Save)                             | 保存 Documents 集合中的所有文档。                                                 |

## 属性

| 序号 | 名称                                  | 说明                                                                         |
|:----:|:--------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Add](#Documents.Add)                 | 返回一个 Document 对象，该对象代表添加到打开文档集合中的新建空文档。         |
|  2   | [Application](#Documents.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  3   | [Count](#Documents.Count)             | 返回一个 Number 类型的值，该值代表集合中的文档数。只读。                     |
|  4   | [Creator](#Documents.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |

## 成员方法

### Documents.Add

返回一个 **Document** 对象，该对象代表添加到打开文档集合中的新建空文档。

**语法**

**express.Add ( *Template* , *NewTemplate* , *DocumentType* , *Visible* )**

\*express是一个代表 Documents 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                                                                           |
|----------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------|
| *Template*     | 可选      | Variant  | 要用于新文档的模板名。如果省略该参数，则使用 Normal 模板。                                                                                     |
| *NewTemplate*  | 可选      | Variant  | 如果该属性值为 True，则将文档作为模板打开。默认值为 False。                                                                                    |
| *DocumentType* | 可选      | Variant  | 可以是下列 WdNewDocumentType 常量之一：wdNewBlankDocument、wdNewEmailMessage、wdNewFrameset 或 wdNewWebPage。默认常量是 wdNewBlankDocument。   |
| *Visible*      | 可选      | Variant  | 如果该属性值为 True，将在可见窗口中打开文档。如果该属性值为 False，则 WPS 打开此文档，但将文档窗口的 Visible 属性设置为 False。默认值为 True。 |

**示例**

``` JavaScript
/* 本示例基于 Normal 模板新建一篇文档*/
Application.Documents.Add()
```

### Documents.AddBlogDocument

返回一个 **Document** 对象，该对象表示 WPS 向前三个参数所描述的帐户发布的新博客文档。

**语法**

**express.AddBlogDocument ( *ProviderID* , *PostURL* , *BlogName* , *PostID* )**

\*express是一个代表 Documents 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                             |
|--------------|-----------|----------|------------------------------------------------------------------|
| *ProviderID* | 必选      | String   | GUID，即提供商在向 WPS 进行注册时所用的唯一值。                  |
| *PostURL*    | 必选      | String   | 用于向博客添加帖子的 URL。                                       |
| *BlogName*   | 必选      | String   | 将用于 WPS 中的博客的显示名称。                                  |
| *PostID*     | 可选      | String   | 用来填充使用 AddBlogDocument 方法创建的文档的现有张贴内容的 ID。 |

**说明**

此方法不仅创建一个新文档，还会向 WPS 注册指定的博客帐户（如果未注册）。此外，如果指定了 *PostID* 参数，则从提供商网站中用 *PostID* 参数的值所指定的张贴内容来填充新文档。

### Documents.CanCheckOut

如果该属性值为 **True** ，则 WPS 可将指定的文档从服务器签出。可读/写 **Boolean** 类型。

**语法**

**express.CanCheckOut ( *FileName* )**

\*express是一个代表 Documents 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                       |
|------------|-----------|----------|----------------------------|
| *FileName* | 必选      | String   | 文档的服务器路径和文件名。 |

**返回值**

Boolean

**说明**

要利用 WPS 中内置的协作功能，文档必须存储在 Microsoft SharePoint Portal Server 上。

**示例**

``` JavaScript
/*本示例验证没有其他用户编辑文档，并可将其签出。如果可以签出文档，则将文档复制到本地计算机进行编辑。*/
function CheckDocInOut(docCheckOut) {
if(Documents.CanCheckOut(docCheckOut) == true) {
        Documents.CheckOut(docCheckOut)
    }
    else {
        MsgBox("You are unable to check out this document at this time.")
    }
}
```

``` JavaScript
/*若要调用 CheckInOut 子程序，请使用下列子程序并将文件名“http://servername/workspace/report.doc”替换为以上“说明”一节中提到的服务器上实际的文件。 */
function CheckDocInOut() {
    Call CheckInOut("http://servername/workspace/report.doc")
}
```

### Documents.Close

关闭指定的文档。

**语法**

**express.Close ( *SaveChanges* , *OriginalFormat* , *RouteDocument* )**

\*express是一个代表 Documents 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                 |
|------------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------|
| *SaveChanges*    | 可选      | Variant  | 指定保存文档的操作。可以是下列 WdSaveOptions 常量之一：wdDoNotSaveChanges、wdPromptToSaveChanges 或 wdSaveChanges。  |
| *OriginalFormat* | 可选      | Variant  | 指定保存文档的格式。可以是下列 WdOriginalFormat 常量之一：wdOriginalDocumentFormat、wdPromptUser 或 wdWordDocument。 |
| *RouteDocument*  | 可选      | Variant  | 如果为 True，则将文档传送给下一个收件人。如果文档没有附加的传送名单，则忽略该参数。                                  |

### Documents.Item

返回集合中的单个 **Document** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Documents 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                           |
|---------|-----------|----------|----------------------------------------------------------------|
| *Index* | 可选      | Number   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

**示例**

``` JavaScript
/* 本示例显示 Documents 集合的第一篇文档的名称 */
Application.Documents.Item(1).Name
```

### Documents.Open

打开指定的文档并将其添加到 **Documents** 集合。返回一个 **Document** 对象。

**语法**

**express.Open ( *FileName* , *ConfirmConversions* , *ReadOnly* , *AddToRecentFiles* , *PasswordDocument* , *PasswordTemplate* , *Revert* , *WritePasswordDocument* , *WritePasswordTemplate* , *Format* , *Encoding* , *Visible* , *OpenConflictDocument* , *OpenAndRepair* , *DocumentDirection* , *NoEncodingDialog* )**

\*express是一个代表 Documents 对象的变量。

**参数**

| 名称                    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                 |
|-------------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName*              | 必选      | Variant  | 文档名（可包含路径）。                                                                                                                                                                               |
| *ConfirmConversions*    | 可选      | Variant  | 如果该属性为 True，则当文件不是 WPS 格式时，将显示“文件转换”对话框。                                                                                                                                 |
| *ReadOnly*              | 可选      | Variant  | 如果该属性值为 True，则以只读方式打开文档。该参数不会覆盖保存的文档的只读建议设置。例如，如果文档在只读建议启用的情况下保存，则将 ReadOnly 参数设置为 False 不会导致文件以可读写方式打开。           |
| *AddToRecentFiles*      | 可选      | Variant  | 如果该属性值为 True，则将文件名添加到“文件”菜单底部最近使用过的文件列表中。                                                                                                                          |
| *PasswordDocument*      | 可选      | Variant  | 打开文档时所需的密码。                                                                                                                                                                               |
| *PasswordTemplate*      | 可选      | Variant  | 打开模板时所需的密码。                                                                                                                                                                               |
| *Revert*                | 可选      | Variant  | 控制当 FileName 是一篇打开的文档的名称时应采取的操作。如果该属性值为 True，则放弃对打开文档进行的所有尚未保存的更改，并将重新打开该文件。如果该属性值为 False，则激活打开的文档。                    |
| *WritePasswordDocument* | 可选      | Variant  | 用于保存文档更改的密码。                                                                                                                                                                             |
| *WritePasswordTemplate* | 可选      | Variant  | 用于保存模板更改的密码。                                                                                                                                                                             |
| *Format*                | 可选      | Variant  | 用于打开文档的文件转换器。可以是 WdOpenFormat 常量之一。默认值为 wdOpenFormatAuto。                                                                                                                  |
| *Encoding*              | 可选      | Variant  | 当您查看保存的文档时 WPS 所使用的文档编码（代码页或字符集）。可以是任何有效的 MsoEncoding 常量。要查看有效 MsoEncoding 常量的列表，请参阅“Visual Basic 编辑器”中的“对象浏览器”。默认值为系统代码页。 |
| *Visible*               | 可选      | Variant  | 如果该属性值为 True，则可在可见窗口中打开文档。默认值为 True。                                                                                                                                       |
| *OpenConflictDocument*  | 可选      | Variant  | 指定是否打开具有脱机冲突的文档的冲突文件。                                                                                                                                                           |
| *OpenAndRepair*         | 可选      | Variant  | 如果该属性为 True，则修复文档，以防止文档毁坏。                                                                                                                                                      |
| *DocumentDirection*     | 可选      | Variant  | 表示文档中的横排文字。默认值为 wdLeftToRight。                                                                                                                                                       |
| *NoEncodingDialog*      | 可选      | Variant  | 如果该属性值为 True，则跳过显示当文字编码无法识别时 WPS 所显示的“编码”对话框。默认值为 False。                                                                                                       |

**说明**

请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

``` JavaScript
/* 本示例以只读方式打开文档 MyDoc.doc*/
Application.Documents.Open("C:\\MyDoc.doc", null, true)
```

### Documents.OpenNoRepairDialog

打开指定的文档并将其添加到 Documents 集合。

**语法**

**express.OpenNoRepairDialog ( *FileName* , *ConfirmConversions* , *ReadOnly* , *AddToRecentFiles* , *PasswordDocument* , *PasswordTemplate* , *Revert* , *WritePasswordDocument* , *WritePasswordTemplate* , *Format* , *Encoding* , *Visible* , *OpenAndRepair* , *DocumentDirection* , *NoEncodingDialog* , *XMLTransform* )**

\*express是一个代表 Documents 对象的变量。

**参数**

| 名称                    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                             |
|-------------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName*              | 必选      | Variant  | 文档名（可包含路径）。                                                                                                                                                                           |
| *ConfirmConversions*    | 可选      | Variant  | 如果该属性值为 True，则当文件不是 WPS 格式时，将显示“转换文件”对话框。                                                                                                                           |
| *ReadOnly*              | 可选      | Variant  | 如果值为 True，则以只读方式打开文档。该参数不会覆盖已保存文档的只读建议设置。例如，如果在只读建议设置打开的情况下保存了文档，则将 ReadOnly 参数设置为 False 不会导致文件以可读/写方式打开。      |
| *AddToRecentFiles*      | 可选      | Variant  | 如果该属性值为 True，则将文件名添加到“文件”菜单底部最近使用过的文件列表中。                                                                                                                      |
| *PasswordDocument*      | 可选      | Variant  | 打开文档时所需的密码。                                                                                                                                                                           |
| *PasswordTemplate*      | 可选      | Variant  | 打开模板时所需的密码。                                                                                                                                                                           |
| *Revert*                | 可选      | Variant  | 控制当 FileName 是一篇打开的文档的名称时应采取的操作。如果该属性值为 True，则放弃对打开文档进行的所有尚未保存的更改，并将重新打开该文件。如果该属性值为 False，则激活打开的文档。                |
| *WritePasswordDocument* | 可选      | Variant  | 用于保存文档更改的密码。                                                                                                                                                                         |
| *WritePasswordTemplate* | 可选      | Variant  | 用于保存模板更改的密码。                                                                                                                                                                         |
| *Format*                | 可选      | Variant  | 用于打开文档的文件转换器。可以是 WdOpenFormat 常量之一。默认值为 wdOpenFormatAuto。                                                                                                              |
| *Encoding*              | 可选      | Variant  | 当您查看已保存文档时 WPS 所使用的文档编码（代码页或字符集）。可以是任何有效的 MsoEncoding 常量。要查看有效 MsoEncoding 常量的列表，请参阅“Visual Basic 编辑器中的对象浏览器”。默认为系统代码页。 |
| *Visible*               | 可选      | Variant  | 如果值为 True，则在可见窗口中打开文档。默认值为 True。                                                                                                                                           |
| *OpenAndRepair*         | 可选      | Variant  | 如果值为 True，则修复文档，以防止文档毁坏。                                                                                                                                                      |
| *DocumentDirection*     | 可选      | Variant  | 表示文档中文字的水平排列方式。可以是任何有效的 WdDocumentDirection 常量。默认值为 wdLeftToRight。                                                                                                |
| *NoEncodingDialog*      | 可选      | Variant  | 如果值为 True，则跳过显示当无法识别文字编码时 WPS 所显示的“编码”对话框。默认值为 False。                                                                                                         |
| *XMLTransform*          | 可选      | Variant  | 指定要使用的转换。                                                                                                                                                                               |

**返回值**

一个代表指定文档的 Document 对象

**说明**

## 安全性

请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

``` JavaScript
/*以下示例以只读方式打开文档 MyDoc.doc。*/
function OpenDoc() {
    Documents.OpenNoRepairDialog("C:\\MyFiles\\MyDoc.doc", null, true)
}
```

``` JavaScript
/*以下示例利用 WordPerfect 6.x 文件转换器打开 Test.wp。*/
function OpenDoc2() {
    let fmt = Application.FileConverters.Item("WordPerfect6x").OpenFormat
    Documents.OpenNoRepairDialog("C:\\MyFiles\\Test.wp", null, null, null, null, null, null, null, null, fmt)
}
```

### Documents.Parent

返回一个 **Object** 类型值，该值代表指定 **Documents** 对象的父对象。

**语法**

**express.Parent ()**

\*express是一个代表 Documents 对象的变量。

### Documents.Save

保存 **Documents** 集合中的所有文档。

**语法**

**express.Save ( *NoPrompt* , *OriginalFormat* )**

\*express是一个代表 Documents 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                      |
|------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------|
| *NoPrompt*       | 可选      | Variant  | 如果该属性值为 True，则 WPS 将自动保存所有文档。如果该属性值为 False， WPS 将提示用户保存自上次保存后已修改过的每篇文档。 |
| *OriginalFormat* | 可选      | Variant  | 指定文档的保存方式。可以是 WdOriginalFormat 常量之一。                                                                    |

**说明**

如果此文档以前未保存过，则 **“另存为”** 对话框将提示用户键入文件名。

## 成员属性

### Documents.Add

返回一个 **Document** 对象，该对象代表添加到打开文档集合中的新建空文档。

**语法**

**express.Add**

\*express是一个代表 Documents 对象的变量。

**示例**

``` JavaScript
/*本示例基于 Normal 模板新建一篇文档。*/
Documents.Add()
```

``` JavaScript
/*本示例基于“专业型备忘录”模板新建一篇文档。*/
Documents.Add("C:\\Program Files\\Kingsoft\\WPS Office Professional\\office6\\mui\\default\\templates\\Normal.wpt")
```

``` JavaScript
/*本示例以活动文档的模板为样例，创建并打开一个新模板。*/
function test()
{
let tmpName = ActiveDocument.AttachedTemplate.FullName
Documents.Add(tmpName, true)
}
```

### Documents.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Documents 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Documents.Count

返回一个 **Number** 类型的值，该值代表集合中的文档数。只读。

**语法**

**express.Count**

\*express是一个代表 Documents 对象的变量。

### Documents.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Documents 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Documents](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Endnotes 对象

## 简介

Endnote 对象的集合，该集合中的对象代表所选内容、范围或文档中的所有尾注。

## 方法

| 序号 | 名称                                                               | 说明                                                  |
|:----:|:-------------------------------------------------------------------|:------------------------------------------------------|
|  1   | [Add](#Endnotes.Add)                                               | 返回一个 Endnote 对象，该对象代表添加到区域中的尾注。 |
|  2   | [Convert](#Endnotes.Convert)                                       | 将尾注转换为脚注。                                    |
|  3   | [Item](#Endnotes.Item)                                             | 返回集合中的单个 Endnote 对象。                       |
|  4   | [ResetContinuationNotice](#Endnotes.ResetContinuationNotice)       | 将尾注延续标记重新设置为默认标记。                    |
|  5   | [ResetContinuationSeparator](#Endnotes.ResetContinuationSeparator) | 将尾注延续分隔符重新设置为默认分隔符。                |
|  6   | [ResetSeparator](#Endnotes.ResetSeparator)                         | 将尾注分隔符重新设置为默认分隔符。                    |
|  7   | [SwapWithFootnotes](#Endnotes.SwapWithFootnotes)                   | 将一篇文档中的所有尾注转换成脚注或将脚注转换为尾注。  |

## 属性

| 序号 | 名称                                                     | 说明                                                                            |
|:----:|:---------------------------------------------------------|:--------------------------------------------------------------------------------|
|  1   | [Application](#Endnotes.Application)                     | 返回一个代表 WPS 应用程序的 Application 对象。                                  |
|  2   | [ContinuationNotice](#Endnotes.ContinuationNotice)       | 返回一个 Range 对象，该对象代表脚注的接续注明。只读。                           |
|  3   | [ContinuationSeparator](#Endnotes.ContinuationSeparator) | 返回一个 Range 对象，该对象代表尾注的接续分隔符。只读。                         |
|  4   | [Count](#Endnotes.Count)                                 | 返回一个 Long 类型的值，该值代表集合中尾注的数量。只读。                        |
|  5   | [Creator](#Endnotes.Creator)                             | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。    |
|  6   | [Location](#Endnotes.Location)                           | 返回或设置所有尾注的位置。 WdEndnoteLocation 类型，可读写。                     |
|  7   | [NumberStyle](#Endnotes.NumberStyle)                     | 返回或设置编号样式。可读/写 WdNoteNumberStyle 类型。                            |
|  8   | [NumberingRule](#Endnotes.NumberingRule)                 | 返回或设置在分页符或分节符后对尾注进行编号的方式。可读写 WdNumberingRule 类型。 |
|  9   | [Parent](#Endnotes.Parent)                               | 返回一个 Object 类型值，该值代表指定 Endnotes 对象的父对象。                    |
|  10  | [Separator](#Endnotes.Separator)                         | 返回一个 Range 对象，该对象代表尾注分隔符。                                     |
|  11  | [StartingNumber](#Endnotes.StartingNumber)               | 返回或设置注释编号、行号或页码的起始值。可读/写 Long 类型。                     |

## 成员方法

### Endnotes.Add

返回一个 **Endnote** 对象，该对象代表添加到区域中的尾注。

**语法**

**express.Add ( *Range* , *Reference* , *Text* )**

\*express是一个代表 Endnotes 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型     | 说明                                                                 |
|-------------|-----------|--------------|----------------------------------------------------------------------|
| *Range*     | 必选      | Range object | 标记用于尾注或脚注的区域。可以是折叠区域。                           |
| *Reference* | 可选      | Variant      | 自定义引用标记的文字。如果省略该参数，WPS 将插入自动编号的引用标记。 |
| *Text*      | 可选      | Variant      | 尾注或脚注文本。                                                     |

**说明**

要为 *Reference* 参数指定一个符号，可用 ` {FontName CharNum} ` 语法。FontName 是包含该符号的字体名称。修饰字体的名称显示在 **“插入”** 菜单 **“符号”** 对话框的 **“字体”** 框中。CharNum 是要插入符号的相应位置序数与 31 之和，在符号表中从左向右计数。例如，如果指定的符号是“Symbol”字体的 omega (ω)，它在符号表格中的位置为 56，该参数就应为“{Symbol 87}”。

**示例**

``` JavaScript
/*本示例为活动文档的第三段添加一条尾注。*/
function test() {
  let myRange = Application.ActiveDocument.Paragraphs.Item(3).Range
  Application.ActiveDocument.Endnotes.Add(myRange, undefined, "Ibid., 314.")
}
```

### Endnotes.Convert

将尾注转换为脚注。

**语法**

**express.Convert ()**

\*express是一个代表 Endnotes 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中所有的尾注转换为脚注。*/
function test() {
  let endDocEndnotes = Application.ActiveDocument.Endnotes
  if(endDocEndnotes.Count > 0) {
      myEndnotes.Convert()
  }
}
```

### Endnotes.Item

返回集合中的单个 **Endnote** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Endnotes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

### Endnotes.ResetContinuationNotice

将尾注延续标记重新设置为默认标记。

**语法**

**express.ResetContinuationNotice ()**

\*express是一个代表 Endnotes 对象的变量。

**说明**

默认标记为空（无文本）。

**示例**

``` JavaScript
/*以下示例重新设置活动文档的尾注延续标记。*/
Application.ActiveDocument.Endnotes.ResetContinuationNotice()
```

### Endnotes.ResetContinuationSeparator

将尾注延续分隔符重新设置为默认分隔符。

**语法**

**express.ResetContinuationSeparator ()**

\*express是一个代表 Endnotes 对象的变量。

**说明**

默认分隔符是一条长横线，该横线分隔文档文字与上接前一页的注释。

**示例**

``` JavaScript
/*以下示例重新设置每个打开的文档中第一节的尾注延续分隔符。*/
function test() {
  for(let i = 1; i <= Application.Documents.Count; i++) {
      Application.Documents.Item(i).Sections.Item(1).Range.Endnotes.ResetContinuationSeparator()
  }
}
```

### Endnotes.ResetSeparator

将尾注分隔符重新设置为默认分隔符。

**语法**

**express.ResetSeparator ()**

\*express是一个代表 Endnotes 对象的变量。

**说明**

/\* 默认分隔符为一条短横线，该横线分隔文档文字和注释。 \*/

**示例**

``` JavaScript
/*以下示例为所选内容所在的文档中的注释重新设置尾注分隔符。*/
Application.Selection.Endnotes.ResetSeparator()
```

### Endnotes.SwapWithFootnotes

将一篇文档中的所有尾注转换成脚注或将脚注转换为尾注。

**语法**

**express.SwapWithFootnotes ()**

\*express是一个代表 Endnotes 对象的变量。

**说明**

要将某一区域的尾注转换为脚注，请使用 **Convert** 方法。

**示例**

``` JavaScript
/*本示例将活动文档中的尾注转换为脚注，并将脚注转换为尾注。*/
Application.ActiveDocument.Endnotes.SwapWithFootnotes()
```

## 成员属性

### Endnotes.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 Endnotes 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Endnotes.ContinuationNotice

返回一个 **Range** 对象，该对象代表脚注的接续注明。只读。

**语法**

**express.ContinuationNotice**

\*express是一个代表 Endnotes 对象的变量。

**示例**

``` JavaScript
/*本示例用文字“Continued...”来代替当前的脚注接续注明。*/
function test() {
  let notice = Application.ActiveDocument.Footnotes.ContinuationNotice
  notice.Delete()
  notice.InsertBefore("Continued...")
}
```

### Endnotes.ContinuationSeparator

返回一个 **Range** 对象，该对象代表尾注的接续分隔符。只读。

**语法**

**express.ContinuationSeparator**

\*express是一个代表 Endnotes 对象的变量。

**示例**

``` JavaScript
/*本示例用一组下划线字符来代替当前的尾注接续分隔符。*/
function test() {
  let separator = Application.ActiveDocument.Endnotes.ContinuationSeparator
  separator.Delete()
  separator.InsertBefore("____")
}
```

### Endnotes.Count

返回一个 **Long** 类型的值，该值代表集合中尾注的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Endnotes 对象的变量。

### Endnotes.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Endnotes 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Endnotes.Location

返回或设置所有尾注的位置。 WdEndnoteLocation 类型，可读写。

**语法**

**express.Location**

\*express是一个代表 Endnotes 对象的变量。

**示例**

``` JavaScript
/*本示例将所有的尾注置于节的结尾。*/
Application.ActiveDocument.Endnotes.Location = wdEndOfSection
```

### Endnotes.NumberStyle

返回或设置编号样式。可读/写 **WdNoteNumberStyle** 类型。

**语法**

**express.NumberStyle**

\*express是一个代表 Endnotes 对象的变量。

**说明**

根据所选择或安装的语言支持（例如，英语（美国））的不同，以上某些常量可能不可用。

**示例**

``` JavaScript
/*本示例设置活动文档中脚注和尾注的格式。*/
function test() {
  Application.ActiveDocument.Footnotes.NumberStyle = wdNoteNumberStyleLowercaseRoman
  Application.ActiveDocument.Endnotes.NumberStyle = wdNoteNumberStyleUppercaseRoman
}
```

### Endnotes.NumberingRule

返回或设置在分页符或分节符后对尾注进行编号的方式。可读写 WdNumberingRule 类型。

**语法**

**express.NumberingRule**

\*express是一个代表 Endnotes 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动文档中的每个分节符后对尾注重新开始编号。*/
Application.ActiveDocument.Endnotes.NumberingRule = wdRestartSection
```

### Endnotes.Parent

返回一个 **Object** 类型值，该值代表指定 **Endnotes** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Endnotes 对象的变量。

### Endnotes.Separator

返回一个 **Range** 对象，该对象代表尾注分隔符。

**语法**

**express.Separator**

\*express是一个代表 Endnotes 对象的变量。

### Endnotes.StartingNumber

返回或设置注释编号、行号或页码的起始值。可读/写 **Long** 类型。

**语法**

**express.StartingNumber**

\*express是一个代表 Endnotes 对象的变量。

**说明**

只有在页面视图中才能看到行号。

当应用于页码时，该属性返回或设置指定 **HeaderFooter** 对象的起始页码。该页码是否在第一页显示取决于 **ShowFirstPageNumber** 属性的设置。如果 **RestartNumberingAtSection** 属性设置为 **False** ，它将覆盖 **StartingNumber** 属性，以便可以从前一节继续编排页码。

**示例**

``` JavaScript
/*本示例创建一篇新文档，将脚注的起始编号设置为 10，然后在插入点添加一个脚注。*/
function test() {
  let myDoc = Application.Documents.Add()
  myDoc.Footnotes.StartingNumber=10
  myDoc.Footnotes.Add(Selection.Range,"Text for a footnote")
}

/*本示例将为活动文档设置行号。起始编号设置为 5，编号间隔设置为 5，并且文档每节的开头独立编号。*/
function test() {
  let numbering= Application.ActiveDocument.PageSetup.LineNumbering
  numbering.Active = true
  numbering.StartingNumber = 5
  numbering.CountBy = 5
  numbering.RestartMode = wdRestartSection
}

/*本示例设置页码属性，然后为活动文档的页眉添加页码。*/
function test() {
  let Starting= Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).PageNumbers
  Starting.NumberStyle = wdPageNumberStyleArabic
  Starting.IncludeChapterNumber = false
  Starting.RestartNumberingAtSection = true
  Starting.StartingNumber = 5
  Starting.Add(wdAlignPageNumberCenter,true)
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Endnotes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListGallery 对象

## 简介

代表单个列表格式库。 ListGallery 对象是 ListGalleries 集合的一个成员。

## 说明

每个 ListGallery 对象代表 “项目符号和编号” 对话框中的三个选项卡之一。

使用 ListGalleries ( *Index* ) 可返回一个 ListGallery 对象，其中 *Index* 为 wdBulletGallery 、 wdNumberGallery 或 wdOutlineNumberGallery 。

以下示例返回 “项目符号和编号” 对话框中 “项目符号” 选项卡上的第三种列表格式（不包括 “无” ），并将其应用于所选内容。

``` JavaScript
function test() {
    let temp3 = Application.ListGalleries.Item(wdBulletGallery).ListTemplates.Item(3)
    Application.Selection.Range.ListFormat.ApplyListTemplate(temp3)
}
```

要查看指定列表模板是否包含 WPS 内置格式，请对 ListGallery 对象使用 Modified 属性。要将格式重新设置为原始列表格式，请对 ListGallery 对象使用 Reset 方法。

## 属性

| 序号 | 名称                                        | 说明                                                                    |
|:----:|:--------------------------------------------|:------------------------------------------------------------------------|
|  1   | [ListTemplates](#ListGallery.ListTemplates) | 返回一个 ListTemplates 集合，该集合代表指定列表库的所有列表格式。只读。 |

## 成员属性

### ListGallery.ListTemplates

返回一个 **ListTemplates** 集合，该集合代表指定列表库的所有列表格式。只读。

**语法**

**express.ListTemplates**

\*express是一个代表 ListGallery 对象的变量。

**示例**

``` JavaScript
/* 以下代码将选区设置为项目符号的第3钟符号模板 */
function test() {
    let temp3 = Application.ListGalleries.Item(wdBulletGallery).ListTemplates.Item(3)
    Application.Selection.Range.ListFormat.ApplyListTemplate(temp3)
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ListGallery](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Pages 对象

## 简介

文档中页面的集合。使用 Pages 集合及相关对象和属性可通过编程方式定义文档的页面版式。

## 方法

| 序号 | 名称                | 说明                                   |
|:----:|:--------------------|:---------------------------------------|
|  1   | [Item](#Pages.Item) | 返回集合中的单个 Page 对象。返回Page值 |

## 属性

| 序号 | 名称                              | 说明                                                                         |
|:----:|:----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Pages.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Pages.Count)             | 返回一个 Long 类型的值，该值代表集合中的页数。只读。                         |
|  3   | [Creator](#Pages.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#Pages.Parent)           | 返回一个 Object 类型值，该值代表指定 Pages 对象的父对象。                    |

## 成员方法

### Pages.Item

返回集合中的单个 **Page** 对象。返回Page值

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Pages 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

## 成员属性

### Pages.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Pages 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

**示例**

``` JavaScript
返回一个代表 WPS 应用程序的 Application 对象。
```

### Pages.Count

返回一个 **Long** 类型的值，该值代表集合中的页数。只读。

**语法**

**express.Count**

\*express是一个代表 Pages 对象的变量。

### Pages.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Pages 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Pages.Parent

返回一个 **Object** 类型值，该值代表指定 **Pages** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Pages 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Pages](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PageSetup 对象

## 简介

代表页面设置说明。 PageSetup 对象包含文档的所有页面设置属性（如左边距、下边距和纸张大小）。

## 说明

使用 PageSetup 属性可返回 PageSetup 对象。以下示例将活动文档的第一节设为横向并打印该文档。

``` JavaScript
function test() {
    Application.ActiveDocument.Sections.Item(1).PageSetup.Orientation = wdOrientLandscape
    Application.ActiveDocument.PrintOut()
}
```

以下示例设置命名为“Sales.doc”的文档的所有页边距。

``` JavaScript
function test() {
    let pages = Application.Documents.Item("Sales.doc").PageSetup
    pages.LeftMargin = Application.InchesToPoints(0.75)
    pages.RightMargin = Application.InchesToPoints(0.75)
    pages.TopMargin = Application.InchesToPoints(1.5)
    pages.BottomMargin = Application.InchesToPoints(1)
}
```

## 方法

| 序号 | 名称                                                    | 说明                                                                   |
|:----:|:--------------------------------------------------------|:-----------------------------------------------------------------------|
|  1   | [SetAsTemplateDefault](#PageSetup.SetAsTemplateDefault) | 将指定的页面设置格式设为活动文档和所有基于活动模板的新文档的默认格式。 |
|  2   | [TogglePortrait](#PageSetup.TogglePortrait)             | 在文档或节的纵向和横向页面方向之间切换。                               |

## 属性

| 序号 | 名称                                                                        | 说明                                                                                                                                                        |
|:----:|:----------------------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#PageSetup.Application)                                       | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                              |
|  2   | [BookFoldPrinting](#PageSetup.BookFoldPrinting)                             | 如果为 True ，则 WPS 将文档打印为一系列小册子，这样可以折叠打印的页面并以书籍的形式阅读。 Boolean 类型，可读写。                                            |
|  3   | [BookFoldPrintingSheets](#PageSetup.BookFoldPrintingSheets)                 | 返回或设置一个 Long 类型的值，该值代表每一本小册子的页数。 Boolean 类型，可读写。                                                                           |
|  4   | [BookFoldRevPrinting](#PageSetup.BookFoldRevPrinting)                       | 如果为 True ，则 WPS 反转用于双向或亚洲语言文档的 书籍折页打印 的打印顺序。 Boolean 类型，可读写。                                                          |
|  5   | [BottomMargin](#PageSetup.BottomMargin)                                     | 返回或设置页面底边与正文文本边界之间的距离（以磅为单位）。 Single 类型，可读写。                                                                            |
|  6   | [CharsLine](#PageSetup.CharsLine)                                           | 返回或设置文档网格中每行的字符数。 Single 类型，可读写。                                                                                                    |
|  7   | [Creator](#PageSetup.Creator)                                               | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                |
|  8   | [DifferentFirstPageHeaderFooter](#PageSetup.DifferentFirstPageHeaderFooter) | 如果第一页使用了不同的页眉或页脚，则为 True 。该属性值可以是 True 、 False 或 wdUndefined 。 Long 类型，可读写。                                            |
|  9   | [FirstPageTray](#PageSetup.FirstPageTray)                                   | 返回或设置用于文档或章节首页的纸盒。 WdPaperTray 类型，可读写。                                                                                             |
|  10  | [FooterDistance](#PageSetup.FooterDistance)                                 | 返回或设置页脚与页面底边之间的距离（以磅为单位）。 Single 类型，可读写。                                                                                    |
|  11  | [Gutter](#PageSetup.Gutter)                                                 | 返回或设置文档或节的每一页的额外页边距的大小（以磅为单位），以备装订需要。 Single 类型，可读写。                                                            |
|  12  | [GutterPos](#PageSetup.GutterPos)                                           | 返回或设置文档中的装订线位于哪一侧。 WdGutterStyle 类型，可读写。                                                                                           |
|  13  | [GutterStyle](#PageSetup.GutterStyle)                                       | 返回或设置 WPS 将装订线应用于当前文档时，基于从右向左的语言还是从左向右的语言。 WdGutterStyleOld 类型，可读写。                                             |
|  14  | [HeaderDistance](#PageSetup.HeaderDistance)                                 | 返回或设置页眉与页面顶端之间的距离（以磅为单位）。 Single 类型，可读写。                                                                                    |
|  15  | [LayoutMode](#PageSetup.LayoutMode)                                         | 返回或设置当前文档的版式模式。 WdLayoutMode 类型，可读写。                                                                                                  |
|  16  | [LeftMargin](#PageSetup.LeftMargin)                                         | 返回或设置页面左边缘与正文左边界之间的距离（以磅为单位）。 Single 类型，可读写。                                                                            |
|  17  | [LineNumbering](#PageSetup.LineNumbering)                                   | 返回或设置一个 LineNumbering 对象，该对象代表指定 PageSetup 对象的行号。                                                                                    |
|  18  | [LinesPage](#PageSetup.LinesPage)                                           | 返回或设置文档网格中每页的行数。 Single 类型，可读写。                                                                                                      |
|  19  | [MirrorMargins](#PageSetup.MirrorMargins)                                   | 如果对开页的内外边距宽度相等，则该属性值为 True 。 Long 类型，可读写。                                                                                      |
|  20  | [OddAndEvenPagesHeaderFooter](#PageSetup.OddAndEvenPagesHeaderFooter)       | 如果指定的 PageSetup 对象的奇数页和偶数页具有不同的页眉和页脚，则该属性值为 True 。 Long 类型，可读写。                                                     |
|  21  | [Orientation](#PageSetup.Orientation)                                       | 返回或设置页面的方向。可读写 WdOrientation 类型。                                                                                                           |
|  22  | [OtherPagesTray](#PageSetup.OtherPagesTray)                                 | 返回或设置用于文档或节中除第一页以外其他所有页的纸盒。 WdPaperTray 类型，可读写。                                                                           |
|  23  | [PageHeight](#PageSetup.PageHeight)                                         | 返回或设置页面高度（以磅为单位）。 Single 类型，可读写。                                                                                                    |
|  24  | [PageWidth](#PageSetup.PageWidth)                                           | 返回或设置页面宽度（以磅为单位）。 Single 类型，可读写。                                                                                                    |
|  25  | [PaperSize](#PageSetup.PaperSize)                                           | 返回或设置纸张大小。 WdPaperSize 类型，可读写。                                                                                                             |
|  26  | [Parent](#PageSetup.Parent)                                                 | 返回一个 Object 类型值，该值代表指定 PageSetup 对象的父对象。                                                                                               |
|  27  | [RightMargin](#PageSetup.RightMargin)                                       | 返回或设置页面右边距与正文右边界之间的距离（以磅为单位）。 Single 类型，可读写。                                                                            |
|  28  | [SectionDirection](#PageSetup.SectionDirection)                             | 返回或设置指定节的阅读顺序和对齐方式。 WdSectionDirection 类型，可读写。                                                                                    |
|  29  | [SectionStart](#PageSetup.SectionStart)                                     | 返回或设置指定对象的分节符类型。 WdSectionStart 类型，可读写。                                                                                              |
|  30  | [ShowGrid](#PageSetup.ShowGrid)                                             | 您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 PageSetup 对象的 ShowGrid 属性的信息，请查阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。 |
|  31  | [SuppressEndnotes](#PageSetup.SuppressEndnotes)                             | 如果在下一个没有隐藏尾注的节的末尾打印尾注，则该属性值为 True 。 Long 类型，可读写。                                                                        |
|  32  | [TextColumns](#PageSetup.TextColumns)                                       | 返回一个 TextColumns 集合，该集合代表指定 PageSetup 对象的文本栏集合。                                                                                      |
|  33  | [TopMargin](#PageSetup.TopMargin)                                           | 返回或设置页面的上边缘与正文文本的上边界之间的距离（以磅为单位）。可读写 Single 类型。                                                                      |
|  34  | [TwoPagesOnOne](#PageSetup.TwoPagesOnOne)                                   | 如果 WPS 将指定文档的每两页打印在同一张纸上，则该属性值为 True 。 Boolean 类型，可读写。                                                                    |
|  35  | [VerticalAlignment](#PageSetup.VerticalAlignment)                           | 返回或设置文档或节中每页文本的垂直对齐方式。可读写 WdVerticalAlignment 类型。                                                                               |

## 成员方法

### PageSetup.SetAsTemplateDefault

将指定的页面设置格式设为活动文档和所有基于活动模板的新文档的默认格式。

**语法**

**express.SetAsTemplateDefault ()**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*以下示例更改活动文档的左右页边距设置，然后将该页面设置格式设为默认格式。*/
function test()
{
let pagesetup = ActiveDocument.PageSetup
pagesetup.LeftMargin = InchesToPoints(1)
pagesetup.RightMargin = InchesToPoints(1)
pagesetup.SetAsTemplateDefault()
}
```

### PageSetup.TogglePortrait

在文档或节的纵向和横向页面方向之间切换。

**语法**

**express.TogglePortrait ()**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例改变活动文档的页面方向。*/
ActiveDocument.PageSetup.TogglePortrait()
```

``` JavaScript
/*本示例将改变文档选定部分的所有节的页面方向。如果每节原来的方向与其他节不同，将会导致出错。*/
Selection.PageSetup.TogglePortrait()
```

## 成员属性

### PageSetup.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 PageSetup 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### PageSetup.BookFoldPrinting

如果为 **True** ，则 WPS 将文档打印为一系列小册子，这样可以折叠打印的页面并以书籍的形式阅读。 **Boolean** 类型，可读写。

**语法**

**express.BookFoldPrinting**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档转换为以一页四开的形式打印的小册子。*/
function Booklet() { 
    let pages = ActiveDocument.PageSetup
    pages.BookFoldPrinting = true
    pages.BookFoldPrintingSheets = 4
}
```

### PageSetup.BookFoldPrintingSheets

返回或设置一个 **Long** 类型的值，该值代表每一本小册子的页数。 **Boolean** 类型，可读写。

**语法**

**express.BookFoldPrintingSheets**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档转换为以 16 页形式打印的小册子。*/
function Booklet() { 
    let pages = ActiveDocument.PageSetup
    pages.BookFoldPrinting = true
    pages.BookFoldPrintingSheets = 16
}
```

### PageSetup.BookFoldRevPrinting

如果为 **True** ，则 WPS 反转用于双向或亚洲语言文档的 书籍折页打印 的打印顺序。 **Boolean** 类型，可读写。

**语法**

**express.BookFoldRevPrinting**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将从左向右的书籍打印切换为从右向左，并用于以一页 16 开打印的双向或亚洲语言文档。*/
function BookletRev() { 
    let pages = ActiveDocument.PageSetup
    pages.BookFoldRevPrinting = true
    pages.BookFoldPrintingSheets = 16
}
```

### PageSetup.BottomMargin

返回或设置页面底边与正文文本边界之间的距离（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.BottomMargin**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档的下边距设置为 72 磅（即 1 英寸），上边距设置为 2 英寸。InchesToPoints 方法用来将英寸转化为磅值。*/
function test() {
  let pageSetup = Application.ActiveDocument.PageSetup
  pageSetup.BottomMargin = 72
  pageSetup.TopMargin = InchesToPoints(2)
}

/*本示例将当前选定内容所有节的下边距设置为 2.5 英寸。*/
Application.Selection.PageSetup.BottomMargin = InchesToPoints(2.5)

/*本示例返回所选部分第一节的下边距。PointsToInches 方法用来将结果转换为英寸。*/
function test() {
  let sngMargin = Application.Selection.Sections.Item(1).PageSetup.BottomMargin
  alert(PointsToInches(sngMargin) + " inches")
}
```

### PageSetup.CharsLine

返回或设置文档网格中每行的字符数。 **Single** 类型，可读写。

**语法**

**express.CharsLine**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档每行字符数设置为 42。*/
ActiveDocument.PageSetup.CharsLine = 42
```

### PageSetup.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### PageSetup.DifferentFirstPageHeaderFooter

如果第一页使用了不同的页眉或页脚，则为 **True** 。该属性值可以是 **True** 、 **False** 或 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.DifferentFirstPageHeaderFooter**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例检查活动文档的每一节，查找第一页是否有不同的页眉和页脚，如果有，则显示一条消息。*/
function test()
{
let sec = ActiveDocument.Sections
for(let i=1;i<=sec.Count;i++) {
    if(sec.Item(i).PageSetup.DifferentFirstPageHeaderFooter==true) {
        MsgBox("Section" + i +"has different first page headers&footers.")
    }
}
}
```

### PageSetup.FirstPageTray

返回或设置用于文档或章节首页的纸盒。 **WdPaperTray** 类型，可读写。

**语法**

**express.FirstPageTray**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例设置送纸盒，用来打印活动文档每节的首页。*/
ActiveDocument.PageSetup.FirstPageTray = wdPrinterLowerBin
```

``` JavaScript
/*本示例设置送纸盒，用来打印所选部分每节的首页。*/
Selection.PageSetup.FirstPageTray = wdPrinterUpperBin
```

### PageSetup.FooterDistance

返回或设置页脚与页面底边之间的距离（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.FooterDistance**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将页脚与页底边之间的距离设置为 0.5 英寸。InchesToPoints 方法用来将英寸转换为磅值。*/
ActiveDocument.PageSetup.FooterDistance = InchesToPoints(0.5)
```

``` JavaScript
/*本示例将选定的所有节的页脚与页底边的距离设置为 1 英寸。*/
Selection.Range.PageSetup.FooterDistance = 72
```

### PageSetup.Gutter

返回或设置文档或节的每一页的额外页边距的大小（以磅为单位），以备装订需要。 **Single** 类型，可读写。

**语法**

**express.Gutter**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果将 **MirrorMargins** 属性设置为 **True** ，则 **Gutter** 属性会将额外页边距添至内边距。否则，将额外页边距添至左边距。

**示例**

``` JavaScript
/*本示例将活动文档的内边距增加 1 英寸（72 磅）。*/
function test()
{
ActiveDocument.PageSetup.MirrorMargins = true
ActiveDocument.PageSetup.Gutter = 72
}
```

### PageSetup.GutterPos

返回或设置文档中的装订线位于哪一侧。 **WdGutterStyle** 类型，可读写。

**语法**

**express.GutterPos**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例设置装订线位于文档的右侧。*/
ActiveDocument.PageSetup.GutterPos = wdGutterPosRight
```

### PageSetup.GutterStyle

返回或设置 WPS 将装订线应用于当前文档时，基于从右向左的语言还是从左向右的语言。 **WdGutterStyleOld** 类型，可读写。

**语法**

**express.GutterStyle**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例设置当前文档使用从右向左语言文档的装订线样式。*/
ActiveDocument.PageSetup.GutterStyle = wdGutterStyleBidi
```

### PageSetup.HeaderDistance

返回或设置页眉与页面顶端之间的距离（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.HeaderDistance**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例显示页眉与页面顶边之间的距离。PointsToInches 方法用来将磅值转换为英寸。*/
function test()
{
let sngDistance = ActiveDocument.PageSetup.HeaderDistance
MsgBox(PointsToInches(sngDistance) + " inches")
}
```

### PageSetup.LayoutMode

返回或设置当前文档的版式模式。 **WdLayoutMode** 类型，可读写。

**语法**

**express.LayoutMode**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动文档的版式模式，以便 WPS 可以自动将键入的文本与网格线对齐。*/
ActiveDocument.PageSetup.LayoutMode = wdLayoutModeGenko
```

### PageSetup.LeftMargin

返回或设置页面左边缘与正文左边界之间的距离（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.LeftMargin**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果将 **MirrorMargins** 属性设置为 **True** ，则 LeftMargin 属性将控制内边距的设置，而 **MirrorMargins** 属性将控制外边距的设置。

**示例**

``` JavaScript
/*本示例将活动文档第二节的左边距设置为 1 英寸（72 磅）。*/
Application.ActiveDocument.Sections.Item(2).PageSetup.LeftMargin = 72
```

### PageSetup.LineNumbering

返回或设置一个 **LineNumbering** 对象，该对象代表指定 **PageSetup** 对象的行号。

**语法**

**express.LineNumbering**

\*express是一个代表 PageSetup 对象的变量。

**说明**

只有在页面视图中才能看到行号。

**示例**

``` JavaScript
/*本示例将为活动文档设置行号。*/
ActiveDocument.PageSetup.LineNumbering.Active = true
```

``` JavaScript
/*本示例为名为“MyDocument.doc”的文档设置行号。起始行号设置为 1，每 5 行显示一次行号，行号文档的各节中连续编号。*/
function test()
{
let myDoc = Documents.Item("MyDocument.doc")
myDoc.PageSetup.LineNumbering.Active = true
myDoc.PageSetup.LineNumbering.StartingNumber = 1
myDoc.PageSetup.LineNumbering.CountBy = 5
myDoc.PageSetup.LineNumbering.RestartMode = wdRestartContinuous
}
```

``` JavaScript
/*本示例将活动文档的行号设置为与 MyDocument.doc 相同。*/
ActiveDocument.PageSetup.LineNumbering = Documents.Item("MyDocument.doc").PageSetup.LineNumbering
 
```

### PageSetup.LinesPage

返回或设置文档网格中每页的行数。 **Single** 类型，可读写。

**语法**

**express.LinesPage**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例设定活动文档的每页具有 15 行。*/
ActiveDocument.PageSetup.LinesPage = 15
```

### PageSetup.MirrorMargins

如果对开页的内外边距宽度相等，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.MirrorMargins**

\*express是一个代表 PageSetup 对象的变量。

**说明**

**MirrorMargins** 属性可为 **True** 、 **False** 或 **wdUndefined** 。如果将 **MirrorMargins** 属性设为 **True** ，则 **LeftMargin** 属性将控制内边距的设置，而 **RightMargin** 属性将控制外边距的设置。

**示例**

``` JavaScript
/*本示例将活动文档的内边距设置为 1 英寸（72 磅），将外边距设置为 0.5 英寸。可以使用 InchesToPoints 方法将英寸转换为磅值。*/
function test()
{
let pagesetup = ActiveDocument.PageSetup
pagesetup.MirrorMargins = true
pagesetup.LeftMargin = 72
pagesetup.RightMargin = InchesToPoints(0.5)
}
```

### PageSetup.OddAndEvenPagesHeaderFooter

如果指定的 **PageSetup** 对象的奇数页和偶数页具有不同的页眉和页脚，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.OddAndEvenPagesHeaderFooter**

\*express是一个代表 PageSetup 对象的变量。

**说明**

**OddAndEvenPagesHeaderFooter** 属性值可以为 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*本示例为 Document1 的奇偶页创建不同的页眉和页脚。*/
function test()
{
let myDoc = Documents.Item("Document1")
myDoc.PageSetup.OddAndEvenPagesHeaderFooter = true
let primary = myDoc.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary)
primary.Range.InsertAfter("Odd Header")
let evenPages = myDoc.Sections.Item(1).Headers.Item(wdHeaderFooterEvenPages)
evenPages.Range.InsertAfter("Even Header")
}
```

### PageSetup.Orientation

返回或设置页面的方向。可读写 **WdOrientation** 类型。

**语法**

**express.Orientation**

\*express是一个代表 PageSetup 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdOrientation** 常量可能不可用。

**示例**

``` JavaScript
/*以下示例先改变文档“MyDocument.doc”的方向，再打印该文档。然后本示例将文档的方向恢复为纵向。*/
function test()
{
let myDoc = Documents.Item("MyDocument.doc")
myDoc.PageSetup.Orientation = wdOrientLandscape
myDoc.PrintOut()
myDoc.PageSetup.Orientation = wdOrientPortrait
}
```

### PageSetup.OtherPagesTray

返回或设置用于文档或节中除第一页以外其他所有页的纸盒。 **WdPaperTray** 类型，可读写。

**语法**

**express.OtherPagesTray**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例设置纸盒用于打印活动文档各节除第一页以外的所有页。*/
ActiveDocument.PageSetup.OtherPagesTray = wdPrinterUpperBin
```

``` JavaScript
/*本示例设置纸盒用于打印选定内容各节中除第一页以外的所有页。*/
Selection.PageSetup.OtherPagesTray = wdPrinterLowerBin
```

### PageSetup.PageHeight

返回或设置页面高度（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.PageHeight**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果设置了 **PageHeight** 属性，则 **PaperSize** 属性会改为 wdPaperCustom 。使用 **PaperSize** 属性可设置预定义大小的纸张的页面高度和宽度，例如 Letter 或 A4 纸。

**示例**

``` JavaScript
/*本示例将活动文档的页面高度设置为 9 英寸。*/
function test() {
  Application.ActiveDocument.PageSetup.PageHeight = InchesToPoints(9)
  Application.ActiveDocument.PageSetup.PageWidth = InchesToPoints(7)
}
```

### PageSetup.PageWidth

返回或设置页面宽度（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.PageWidth**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果设置了 **PageWidth** 属性，则 **PaperSize** 属性将改为 wdPaperCustom 。使用 **PaperSize** 属性可设置预定义大小的纸张的页面高度和宽度，例如 Letter 或 A4 纸。

**示例**

``` JavaScript
/*本示例返回 Document1 的页面宽度。可用 PointsToInches 方法将磅值转换为英寸。*/
function test() {
  let doc1set = Application.Documents.Item("Document1.docx").PageSetup
  alert("The page width is "+PointsToInches(doc1set.PageWidth)+" inches.")
}
```

### PageSetup.PaperSize

返回或设置纸张大小。 **WdPaperSize** 类型，可读写。

**语法**

**express.PaperSize**

\*express是一个代表 PageSetup 对象的变量。

**说明**

设置 **PageHeight** 或 **PageWidth** 属性会将 **PaperSize** 属性更改为 **wdPaperCustom** 。

**示例**

``` JavaScript
/*本示例将第一个文档的纸张大小设置为 legal 型。*/
Application.Documents.Item(1).PageSetup.PaperSize = wdPaperLegal
```

### PageSetup.Parent

返回一个 **Object** 类型值，该值代表指定 **PageSetup** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.RightMargin

返回或设置页面右边距与正文右边界之间的距离（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.RightMargin**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果将 **MirrorMargins** 属性设置为 **True** ，则 **RightMargin** 属性将控制外边距的设置，而 **LeftMargin** 属性控制内边距的设置。

**示例**

``` JavaScript
/*本示例显示活动文档的右边距设置。可以使用 PointsToInches 方法将结果转换为英寸。*/
alert("The right margin is set to " + PointsToInches(Application.ActiveDocument.PageSetup.RightMargin) + " inches.")

/*本示例设置选定区域第二节的右边距。可以使用 InchesToPoints 方法将英寸转换为磅值。*/
Application.Selection.Sections.Item(2).PageSetup.RightMargin = InchesToPoints(1)
```

### PageSetup.SectionDirection

返回或设置指定节的阅读顺序和对齐方式。 **WdSectionDirection** 类型，可读写。

**语法**

**express.SectionDirection**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档的第一节的方向设置为从右向左。*/
ActiveDocument.Sections.Item(1).PageSetup.SectionDirection = wdSectionDirectionRtl
```

### PageSetup.SectionStart

返回或设置指定对象的分节符类型。 **WdSectionStart** 类型，可读写。

**语法**

**express.SectionStart**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中所有节的分节符都改为连续的。*/
ActiveDocument.PageSetup.SectionStart = wdSectionContinuous
```

``` JavaScript
/*本示例返回 MyDoc.doc 中第二节开始部分使用的分节符类型，并将其应用于活动文档中的所有节。*/
function test()
{
mytype = Documents.Item("MyDoc.doc").Sections.Item(2).PageSetup.SectionStart
ActiveDocument.PageSetup.SectionStart = mytype
}
```

### PageSetup.ShowGrid

您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 **PageSetup** 对象的 **ShowGrid** 属性的信息，请查阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

**语法**

**express.ShowGrid**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.SuppressEndnotes

如果在下一个没有隐藏尾注的节的末尾打印尾注，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.SuppressEndnotes**

\*express是一个代表 PageSetup 对象的变量。

**说明**

在该节的尾注之前打印隐藏的尾注。仅当 **Location** 属性设为 **wdEndOfSection** 时，该属性才有效。

**示例**

``` JavaScript
/*本示例隐藏活动文档中第一节的尾注。*/
function test()
{
if(ActiveDocument.Endnotes.Location == wdEndOfSection) {
    ActiveDocument.Sections.Item(1).PageSetup.SuppressEndnotes = true
}
}
```

### PageSetup.TextColumns

返回一个 **TextColumns** 集合，该集合代表指定 **PageSetup** 对象的文本栏集合。

**语法**

**express.TextColumns**

\*express是一个代表 PageSetup 对象的变量。

**说明**

集合中应始终至少包含一个文本列。在创建新的文本列时，会向集合添加一列。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例为活动文档中的第二节创建四个等宽的文本栏。*/
function test()
{
let textcolumns = ActiveDocument.Sections.Item(2).PageSetup.TextColumns
textcolumns.SetCount(3)
textcolumns.Add(null,null,true)
}
```

``` JavaScript
/*本示例创建一个具有两个文本栏的文档。第一个文本栏宽 1.5 英寸，第二个文本栏宽 3 英寸。*/
function test()
{
let myDoc = Documents.Add()
let textcolumns = myDoc.PageSetup.TextColumns
textcolumns.SetCount(1)
textcolumns.Add(InchesToPoints(3))
let item = myDoc.PageSetup.TextColumns.Item(1)
item.Width = InchesToPoints(1.5)
item.SpaceAfter = InchesToPoints(0.5)
}
```

### PageSetup.TopMargin

返回或设置页面的上边缘与正文文本的上边界之间的距离（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.TopMargin**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中第一节的上边距设置为 72 磅（1 英寸）。*/
Application.ActiveDocument.Sections.Item(1).PageSetup.TopMargin = 72
```

### PageSetup.TwoPagesOnOne

如果 WPS 将指定文档的每两页打印在同一张纸上，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.TwoPagesOnOne**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为将活动文档的每两页打印在同一张纸上。*/
ActiveDocument.PageSetup.TwoPagesOnOne = true
```

### PageSetup.VerticalAlignment

返回或设置文档或节中每页文本的垂直对齐方式。可读写 **WdVerticalAlignment** 类型。

**语法**

**express.VerticalAlignment**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*以下示例改变第一个文档的垂直对齐方式，使文本在上边距和下边距之间居中。*/
Documents.Item(1).PageSetup.VerticalAlignment = wdAlignVerticalCenter
```

``` JavaScript
/*以下示例新建一个文档，并将同一段落插入 10 次。然后设置新文档的垂直对齐方式，使 10 个段落在上边距和下边距之间等距排列（两端对齐）。*/
function test()
{
let myDoc = Documents.Add()
for(let i= 1;i<=9;i++) {
    myDoc.Content.InsertAfter("This is a sentence.")
    myDoc.Content.InsertParagraphAfter()
}
myDoc.Content.InsertAfter("This is a sentence.")
myDoc.PageSetup.VerticalAlignment = wdAlignVerticalJustify
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/PageSetup](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ParagraphFormat 对象

## 简介

代表段落的所有格式。

## 说明

使用 Format 属性可返回一个或多个段落的 ParagraphFormat 对象。 ParagraphFormat 属性返回所选内容、范围、样式、 Find 对象或 Replacement 对象的 ParagraphFormat 对象。以下示例将活动文档中的第三段居中。

``` JavaScript
Application.ActiveDocument.Paragraphs.Item(3).Format.Alignment = wdAlignParagraphCenter
```

可以使用 JS 的 n ew 关键词新建一个单独的 ParagraphFormat 对象。以下示例创建一个 ParagraphFormat 对象，为其设置部分格式属性，然后将其所有属性应用于活动文档中的第一段。

``` JavaScript
function test() {
    let myParaF = Application.ActiveDocument.Paragraphs.Item(1).Format
    myParaF.Alignment = wdAlignParagraphCenter
    myParaF.Borders.Enable = true
} 
```

## 方法

| 序号 | 名称                                                                  | 说明                                       |
|:----:|:----------------------------------------------------------------------|:-------------------------------------------|
|  1   | [CloseUp](#ParagraphFormat.CloseUp)                                   | 删除指定段落格式中段落前的任何间距。       |
|  2   | [IndentCharWidth](#ParagraphFormat.IndentCharWidth)                   | 将一个或多个段落缩进指定的字符数。         |
|  3   | [IndentFirstLineCharWidth](#ParagraphFormat.IndentFirstLineCharWidth) | 将一个或多个段落的首行缩进指定的字符数。   |
|  4   | [OpenOrCloseUp](#ParagraphFormat.OpenOrCloseUp)                       | 切换指定段落的段前间距。                   |
|  5   | [OpenUp](#ParagraphFormat.OpenUp)                                     | 为指定段落设置 12 磅的段前间距。           |
|  6   | [Reset](#ParagraphFormat.Reset)                                       | 删除手动段落格式（不使用样式应用的格式）。 |
|  7   | [Space1](#ParagraphFormat.Space1)                                     | 为指定段落设置单倍行距。                   |
|  8   | [Space15](#ParagraphFormat.Space15)                                   | 为指定段落设置 1.5 倍行距。                |
|  9   | [Space2](#ParagraphFormat.Space2)                                     | 为指定段落设置 2 倍行距。                  |
|  10  | [TabHangingIndent](#ParagraphFormat.TabHangingIndent)                 | 将悬挂缩进量设置为指定的制表位数。         |
|  11  | [TabIndent](#ParagraphFormat.TabIndent)                               | 将指定段落的左缩进量设置为指定的制表位数。 |

## 属性

| 序号 | 名称                                                                                | 说明                                                                                                                                                                                            |
|:----:|:------------------------------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddSpaceBetweenFarEastAndAlpha](#ParagraphFormat.AddSpaceBetweenFarEastAndAlpha)   | 如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 True 。如果仅对于某些指定段落将该属性设置为 True ，则该属性会返回 wdUndefined 。 Long 类型，可读写。                      |
|  2   | [AddSpaceBetweenFarEastAndDigit](#ParagraphFormat.AddSpaceBetweenFarEastAndDigit)   | 如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 True 。如果仅对于某些指定段落将该属性设置为 True ，则该属性会返回 wdUndefined 。 Long 类型，可读写。                        |
|  3   | [Alignment](#ParagraphFormat.Alignment)                                             | 返回或设置一个 WdParagraphAlignment 常量，该常量代表指定段落的对齐方式，可读写。                                                                                                                |
|  4   | [Application](#ParagraphFormat.Application)                                         | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                                                  |
|  5   | [AutoAdjustRightIndent](#ParagraphFormat.AutoAdjustRightIndent)                     | 如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 True 。如果只将某些指定段落的 AutoAdjustRightIndent 属性设置为 True ，则该属性会返回 wdUndefined 。 Long 类型，可读写。 |
|  6   | [BaseLineAlignment](#ParagraphFormat.BaseLineAlignment)                             | 返回或设置一个 WdBaselineAlignment 常量，该常量代表行中字体的垂直位置，可读写。                                                                                                                 |
|  7   | [Borders](#ParagraphFormat.Borders)                                                 | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                                                                                                           |
|  8   | [CharacterUnitFirstLineIndent](#ParagraphFormat.CharacterUnitFirstLineIndent)       | 返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。 Single 类型，可读写。                                                                                    |
|  9   | [CharacterUnitLeftIndent](#ParagraphFormat.CharacterUnitLeftIndent)                 | 该属性返回或设置指定段落的左缩进量（以字符为单位）。 Single 类                                                                                                                                  |
|  10  | [CharacterUnitRightIndent](#ParagraphFormat.CharacterUnitRightIndent)               | 该属性返回或设置指定段落的右缩进量（以字符为单位）。 Single 类型，可读写。                                                                                                                      |
|  11  | [Creator](#ParagraphFormat.Creator)                                                 | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                                                    |
|  12  | [DisableLineHeightGrid](#ParagraphFormat.DisableLineHeightGrid)                     | 如果该属性的值为 True ，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 DisableLineHeightGrid 属性设置为 True ，则返回 wdUndefined 。 Long 类型，可读写。    |
|  13  | [Duplicate](#ParagraphFormat.Duplicate)                                             | 返回一个只读 ParagraphFormat 对象，该对象代表指定段落的段落格式。                                                                                                                               |
|  14  | [FarEastLineBreakControl](#ParagraphFormat.FarEastLineBreakControl)                 | 如果为 True ，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 FarEastLineBreakControl 属性设定为 True ，则返回 wdUndefined 。 Long 类型，可读写。                     |
|  15  | [FirstLineIndent](#ParagraphFormat.FirstLineIndent)                                 | 返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。 Single 类型，可读写。                                                                    |
|  16  | [HalfWidthPunctuationOnTopOfLine](#ParagraphFormat.HalfWidthPunctuationOnTopOfLine) | 如果为 True ，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 True ，则此属性将返回 wdUndefined 。 Long 类型，可读写。                                        |
|  17  | [HangingPunctuation](#ParagraphFormat.HangingPunctuation)                           | 如果为 True ，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 True ，则返回 wdUndefined 。 Long 类型，可读写。                                                             |
|  18  | [Hyphenation](#ParagraphFormat.Hyphenation)                                         | 如果指定的段落进行自动断字，则该属性值为 True 。如果指定的段落不进行自动断字，则该属性值为 False 。可读写 Long 类型。                                                                           |
|  19  | [KeepTogether](#ParagraphFormat.KeepTogether)                                       | 在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 True 。可读写 Long 类型。                                                                                           |
|  20  | [KeepWithNext](#ParagraphFormat.KeepWithNext)                                       | 在 WPS 对文档重新分页时，如果指定段落与它的下一段位于同一页上，则该属性值为 True 。可读写 Long 类型。                                                                                           |
|  21  | [LeftIndent](#ParagraphFormat.LeftIndent)                                           | 返回或设置一个 Single 类型的值，该值代表指定段落格式的左缩进值（以磅为单位）。可读写。                                                                                                          |
|  22  | [LineSpacing](#ParagraphFormat.LineSpacing)                                         | 返回或设置指定段落的行距（以磅为单位）。 Single 类型，可读写。                                                                                                                                  |
|  23  | [LineSpacingRule](#ParagraphFormat.LineSpacingRule)                                 | 返回或设置指定段落格式的行距。可读写 WdLineSpacing 类型。                                                                                                                                       |
|  24  | [LineUnitAfter](#ParagraphFormat.LineUnitAfter)                                     | 返回或设置指定段落的段后间距（以网格线为单位）。可读写 Single 类型。                                                                                                                            |
|  25  | [LineUnitBefore](#ParagraphFormat.LineUnitBefore)                                   | 返回或设置指定段落的段前间距（以网格线为单位）。可读写 Single 类型。                                                                                                                            |
|  26  | [MirrorIndents](#ParagraphFormat.MirrorIndents)                                     | 返回或设置一个 Long 类型的值，该值代表左缩进和右缩进的宽度是否相同。该属性值可以是 True 、 False 或 wdUndefined 。可读写。                                                                      |
|  27  | [NoLineNumber](#ParagraphFormat.NoLineNumber)                                       | 如果取消指定段的行号，则该属性值为 True 。该属性值可以是 True 、 False 或 wdUndefined 。可读写 Long 类型。                                                                                      |
|  28  | [OutlineLevel](#ParagraphFormat.OutlineLevel)                                       | 返回或设置指定段落的大纲级别。可读写 WdOutlineLevel 类型。                                                                                                                                      |
|  29  | [PageBreakBefore](#ParagraphFormat.PageBreakBefore)                                 | 如果在指定段落前插入了分页符，则该属性值为 True 。该属性值可以是 True 、 False 或 wdUndefined 。可读写 Long 类型。                                                                              |
|  30  | [Parent](#ParagraphFormat.Parent)                                                   | 返回一个 Object 类型值，该值代表指定 ParagraphFormat 对象的父对象。                                                                                                                             |
|  31  | [ReadingOrder](#ParagraphFormat.ReadingOrder)                                       | 返回或设置指定段落的读取次序而不改变其对齐方式。可读写 WdReadingOrder 类型。                                                                                                                    |
|  32  | [RightIndent](#ParagraphFormat.RightIndent)                                         | 返回或设置指定段落的右缩进量（以磅为单位）。可读写 Single 类型。                                                                                                                                |
|  33  | [Shading](#ParagraphFormat.Shading)                                                 | 返回一个 Shading 对象，该对象代表指定对象的底纹格式。                                                                                                                                           |
|  34  | [SpaceAfter](#ParagraphFormat.SpaceAfter)                                           | 返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 Single 类型。                                                                                                                       |
|  35  | [SpaceAfterAuto](#ParagraphFormat.SpaceAfterAuto)                                   | 如果 WPS 自动设置指定段落的段后间距，则该属性为 True 。可读/写 Long 类型。                                                                                                                      |
|  36  | [SpaceBefore](#ParagraphFormat.SpaceBefore)                                         | 返回或设置指定段落的段前间距（以磅为单位）。可读/写 Single 类型。                                                                                                                               |
|  37  | [SpaceBeforeAuto](#ParagraphFormat.SpaceBeforeAuto)                                 | 如果 WPS 自动设置指定段落的段前间距，则该属性为 True 。可读/写 Long 类型。                                                                                                                      |
|  38  | [Style](#ParagraphFormat.Style)                                                     | 返回或设置指定对象的样式。可读写 Variant 类型。                                                                                                                                                 |
|  39  | [TabStops](#ParagraphFormat.TabStops)                                               | 返回或设置一个 TabStops 集合，该集合代表指定段落中的所有自定义制表位。可读写。                                                                                                                  |
|  40  | [TextboxTightWrap](#ParagraphFormat.TextboxTightWrap)                               | 返回或设置一个 WdTextboxTightWrap 常量，该常量代表文本环绕形状或文本框的紧密程度。可读写。                                                                                                      |
|  41  | [WidowControl](#ParagraphFormat.WidowControl)                                       | 在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 True 。该属性值可以是 True 、 False 或 wdUndefined 。可读写 Long 类型。                                     |
|  42  | [WordWrap](#ParagraphFormat.WordWrap)                                               | 如果 WPS 在指定段落或文本框架的西文单词中间断字换行，则该属性值为 True 。可读写 Long 类型。                                                                                                     |

## 成员方法

### ParagraphFormat.CloseUp

删除指定段落格式中段落前的任何间距。

**语法**

**express.CloseUp ()**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*以下两个语句是等价的：*/
function test()
{
let pgfm = ActiveDocument.Paragraphs.Item(1).Format
pgfm.CloseUp()
pgfm.SpaceBefore = 0
}
```

### ParagraphFormat.IndentCharWidth

将一个或多个段落缩进指定的字符数。

**语法**

**express.IndentCharWidth ( *Count* )**

\*express是一个代表 ParagraphFormat 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                     |
|---------|-----------|----------|--------------------------|
| *Count* | 必选      | Integer  | 指定段落要缩进的字符数。 |

**说明**

此方法等效于单击 **“格式”** 工具栏上的 **“增加缩进量”** 按钮。

**示例**

``` JavaScript
/*以下示例将活动文档的第一段缩进 10 个字符。*/
ActiveDocument.Paragraphs.Item(1).Format.IndentCharWidth (10)
```

### ParagraphFormat.IndentFirstLineCharWidth

将一个或多个段落的首行缩进指定的字符数。

**语法**

**express.IndentFirstLineCharWidth ( *Count* )**

\*express是一个代表 ParagraphFormat 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                               |
|---------|-----------|----------|------------------------------------|
| *Count* | 必选      | Integer  | 每个指定段落的首行要缩进的字符数。 |

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的首行缩进 10 个字符。*/
Application.ActiveDocument.Paragraphs.Item(1).Format.IndentFirstLineCharWidth (10)
```

### ParagraphFormat.OpenOrCloseUp

切换指定段落的段前间距。

**语法**

**express.OpenOrCloseUp ()**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

如果指定段落的段前间距等于 0（零），则此方法将其设置为 12 磅；如果段落的段前间距大于 0（零），则此方法将其设置为 0（零）。

**示例**

``` JavaScript
/*以下示例切换活动文档中第一段的格式，或增加 12 磅的段前间距，或不设置段前间距。*/
ActiveDocument.Paragraphs.Item(1).Format.OpenOrCloseUp()
```

### ParagraphFormat.OpenUp

为指定段落设置 12 磅的段前间距。

**语法**

**express.OpenUp ()**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

您还可以使用 **SpaceBefore** 属性设置段落间距。下列两个语句是等效的：

``` JavaScript
function test()
{
Selection.ParagraphFormat.OpenUp()
Selection.ParagraphFormat.SpaceBefore = 12
}
```

**示例**

``` JavaScript
/*以下示例更改活动文档的第二段的格式，为该段落设置 12 磅的段前间距。*/
ActiveDocument.Paragraphs.Item(2).Format.OpenUp()
```

### ParagraphFormat.Reset

删除手动段落格式（不使用样式应用的格式）。

**语法**

**express.Reset ()**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

如果手动将段落设置为右对齐，而基本样式为其他对齐方式，则用 **Reset** 方法将手动对齐方式改为与基本样式相同。

**示例**

``` JavaScript
/*本示例删除活动文档第二段的手动设定的段落格式。*/
Application.ActiveDocument.Paragraphs.Item(2).Reset()
```

### ParagraphFormat.Space1

为指定段落设置单倍行距。

**语法**

**express.Space1 ()**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

精确间距取决于各段内最大字符的字号。

您还可以使用 **LineSpacingRule** 属性设置段落间距。下列两个语句是等效的：

``` JavaScript
function test()
{
Selection.ParagraphFormat.Space1()
Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
}
```

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的行距更改为单倍行距。*/
ActiveDocument.Paragraphs.Item(1).Format.Space1()
```

### ParagraphFormat.Space15

为指定段落设置 1.5 倍行距。

**语法**

**express.Space15 ()**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

精确间距为各段内最大字符的字号加上 6 磅。

您还可以使用 **LineSpacingRule** 属性设置段落间距。下列两个语句是等效的：

``` JavaScript
function test()
{
Selection.ParagraphFormat.Space15()
Selection.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
}
```

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的行距更改为 1.5 倍行距。*/
ActiveDocument.Paragraphs.Item(1).Format.Space15()
```

### ParagraphFormat.Space2

为指定段落设置 2 倍行距。

**语法**

**express.Space2 ()**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

精确间距为各段内最大字符的字号加上 12 磅。

您还可以使用 **LineSpacingRule** 属性设置段落间距。例如，下列两个语句是等效的：

``` JavaScript
function test()
{
Selection.ParagraphFormat.Space2()
Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceDouble
 


}
```

**示例**

``` JavaScript
/*以下示例将所选内容中第一段的行距更改为 2 倍行距。*/
ActiveDocument.Paragraphs.Item(1).Format.Space2()
```

### ParagraphFormat.TabHangingIndent

将悬挂缩进量设置为指定的制表位数。

**语法**

**express.TabHangingIndent ( *Count* )**

\*express是一个代表 ParagraphFormat 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                     |
|---------|-----------|----------|--------------------------------------------------------------------------|
| *Count* | 必选      | Integer  | 如果为正，则表示要缩进的制表位数；如果为负，则表示要删除的缩进制表位数。 |

**说明**

如果 *Count* 的值为负数，您还可以使用此方法删除悬挂缩进的制表位。

**示例**

``` JavaScript
/*以下示例将所选段落的悬挂缩进设置为第二个制表位。*/
Selection.ParagraphFormat.TabHangingIndent(2)
```

``` JavaScript
/*以下示例将所选段落的悬挂缩进量减少一个制表位。*/
Selection.ParagraphFormat.TabHangingIndent(-1)
```

### ParagraphFormat.TabIndent

将指定段落的左缩进量设置为指定的制表位数。

**语法**

**express.TabIndent ( *Count* )**

\*express是一个代表 ParagraphFormat 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                     |
|---------|-----------|----------|--------------------------------------------------------------------------|
| *Count* | 必选      | Integer  | 如果为正，则表示要缩进的制表位数；如果为负，则表示要删除的缩进制表位数。 |

**说明**

如果 *Count* 的值为负数，您还可以使用此方法删除缩进。

**示例**

``` JavaScript
/*以下示例将所选段落缩进到第二个制表位。*/
Selection.ParagraphFormat.TabIndent(2)
```

``` JavaScript
/*以下示例将所选段落的缩进量减少一个制表位。*/
Selection.ParagraphFormat.TabIndent(-1)
```

## 成员属性

### ParagraphFormat.AddSpaceBetweenFarEastAndAlpha

如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 **True** 。如果仅对于某些指定段落将该属性设置为 **True** ，则该属性会返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndAlpha**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置 WPS 在活动文档第一段的日文文字和拉丁文字之间自动添加空格。*/
ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndAlpha = true
```

### ParagraphFormat.AddSpaceBetweenFarEastAndDigit

如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 **True** 。如果仅对于某些指定段落将该属性设置为 **True** ，则该属性会返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndDigit**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例可实现的功能是：设定 WPS 自动在活动文档第一段的中文文字和数字之间添加空格。*/
ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndDigit = true
```

### ParagraphFormat.Alignment

返回或设置一个 **WdParagraphAlignment** 常量，该常量代表指定段落的对齐方式，可读写。

**语法**

**express.Alignment**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

由于您选择或安装的语言支持不同（例如美国英语），上述部分常量可能无法使用。

### ParagraphFormat.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### ParagraphFormat.AutoAdjustRightIndent

如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 **True** 。如果只将某些指定段落的 **AutoAdjustRightIndent** 属性设置为 **True** ，则该属性会返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.AutoAdjustRightIndent**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设定 WPS 根据指定的每行字符数，自动调整所选段落的右缩进。*/
Selection.ParagraphFormat.AutoAdjustRightIndent = true
```

### ParagraphFormat.BaseLineAlignment

返回或设置一个 **WdBaselineAlignment** 常量，该常量代表行中字体的垂直位置，可读写。

**语法**

**express.BaseLineAlignment**

\*express是一个代表 ParagraphFormat 对象的变量。

### ParagraphFormat.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### ParagraphFormat.CharacterUnitFirstLineIndent

返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。 **Single** 类型，可读写。

**语法**

**express.CharacterUnitFirstLineIndent**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中第一段的首行缩进设为一个字符。*/
ActiveDocument.Paragraphs.Item(1).CharacterUnitFirstLineIndent = 1
```

``` JavaScript
/*本示例将活动文档中第二段的悬挂缩进设为 1.5 个字符。*/
ActiveDocument.Paragraphs.Item(2).CharacterUnitFirstLineIndent = -1.5
```

### ParagraphFormat.CharacterUnitLeftIndent

该属性返回或设置指定段落的左缩进量（以字符为单位）。 **Single** 类

**语法**

**express.CharacterUnitLeftIndent**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中第一段的左缩进设为从左边距缩进一个字符。*/
ActiveDocument.Paragraphs.Item(1).CharacterUnitLeftIndent = 1
```

### ParagraphFormat.CharacterUnitRightIndent

该属性返回或设置指定段落的右缩进量（以字符为单位）。 **Single** 类型，可读写。

**语法**

**express.CharacterUnitRightIndent**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中的所有段落的右缩进设为从右边距缩进一个字符。*/
ActiveDocument.Paragraphs.CharacterUnitRightIndent = 1
```

### ParagraphFormat.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### ParagraphFormat.DisableLineHeightGrid

如果该属性的值为 **True** ，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 **DisableLineHeightGrid** 属性设置为 **True** ，则返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.DisableLineHeightGrid**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设定 WPS 在已指定每页行数的情况下，将选定段落中的字符与行网格对齐。*/
Selection.ParagraphFormat.DisableLineHeightGrid = true
```

### ParagraphFormat.Duplicate

返回一个只读 **ParagraphFormat** 对象，该对象代表指定段落的段落格式。

**语法**

**express.Duplicate**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

可以使用 **Duplicate** 属性来获取重复的 **ParagraphFormat** 对象所有属性的设置。可将 **Duplicate** 属性返回的对象赋给另一个相同类型的对象，从而一次应用该对象的所有设置。在将重复对象赋给另一个对象之前，可更改重复对象的任何属性，而不会影响源对象。

**示例**

``` JavaScript
/*本示例复制活动文档第一段的段落格式，并将该格式保存在 myDup 变量中，然后将 myDup 的左缩进量改为 1 英寸。再创建一篇新文档，插入文本，将 myDup 中保存的段落格式应用于该文本。*/
function test()
{
ActiveDocument.Range(0, 0).InsertAfter ("Paragraph Number 1")
let myDup = ActiveDocument.Paragraphs.Item(1).Format.Duplicate
myDup.LeftIndent = InchesToPoints(1)
Documents.Add()
Selection.InsertAfter ("This is a new paragraph.")
Selection.Paragraphs.Format = myDup
}
```

### ParagraphFormat.FarEastLineBreakControl

如果为 **True** ，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 **FarEastLineBreakControl** 属性设定为 **True** ，则返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.FarEastLineBreakControl**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 将东亚语言文字的换行规则应用于活动文档的第一段。*/
function test()
{
ActiveDocument.Paragraphs.Item(1).FarEastLineBreakControl = true
}
```

### ParagraphFormat.FirstLineIndent

返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。 **Single** 类型，可读写。

**语法**

**express.FirstLineIndent**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档的首段设置 1 英寸的首行缩进。*/
Application.ActiveDocument.Paragraphs.Item(1).FirstLineIndent = InchesToPoints(1)

/*本示例为活动文档的第二段设置 0.5 英寸的悬挂缩进。InchesToPoints 方法用来将英寸转化为磅值。*/
Application.ActiveDocument.Paragraphs.Item(2).FirstLineIndent = InchesToPoints(-0.5)
```

### ParagraphFormat.HalfWidthPunctuationOnTopOfLine

如果为 **True** ，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 **True** ，则此属性将返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.HalfWidthPunctuationOnTopOfLine**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 将活动文档第一段的行首标点符号改为半角字符。*/
ActiveDocument.Paragraphs.Item(1).HalfWidthPunctuationOnTopOfLine = true
```

### ParagraphFormat.HangingPunctuation

如果为 **True** ，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 **True** ，则返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.HangingPunctuation**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：允许活动文档第一段的标点可以溢出边界。*/
ActiveDocument.Paragraphs.Item(1).HangingPunctuation = true
```

### ParagraphFormat.Hyphenation

如果指定的段落进行自动断字，则该属性值为 **True** 。如果指定的段落不进行自动断字，则该属性值为 **False** 。可读写 **Long** 类型。

**语法**

**express.Hyphenation**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例关闭活动文档中具有“正文”样式的所有段落的自动断字功能。*/
ActiveDocument.Styles.Item("正文").ParagraphFormat.Hyphenation = false
```

### ParagraphFormat.KeepTogether

在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.KeepTogether**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

### ParagraphFormat.KeepWithNext

在 WPS 对文档重新分页时，如果指定段落与它的下一段位于同一页上，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.KeepWithNext**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

### ParagraphFormat.LeftIndent

返回或设置一个 **Single** 类型的值，该值代表指定段落格式的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

\*express是一个代表 ParagraphFormat 对象的变量。

### ParagraphFormat.LineSpacing

返回或设置指定段落的行距（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.LineSpacing**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

使用 **LinesToPoints** 方法可将行数转换为相应的值（以磅为单位）。例如， ` LinesToPoints(2) ` 返回值 24。

可以在 **LineSpacingRule** 属性已设置为下列值之后对 **LineSpacing** 属性进行设置：

- **wdLineSpaceAtLeast** 行距可以大于或等于指定的 **LineSpacing** 值，但绝不能小于该值。
- **wdLineSpaceExactly** 指定的 **LineSpacing** 行距值绝不能更改，即使段落中使用了较大的字体。
- **wdLineSpaceMultiple** 必须指定 **LineSpacing** 属性值，以磅为单位。

**示例**

``` JavaScript
/*以下示例将选定段落的行距设置为至少 24 磅。*/
function test() {
  let pgfm = Application.Selection.ParagraphFormat
  pgfm.LineSpacingRule = wdLineSpaceAtLeast
  pgfm.LineSpacing = 24
}
```

### ParagraphFormat.LineSpacingRule

返回或设置指定段落格式的行距。可读写 **WdLineSpacing** 类型。

**语法**

**express.LineSpacingRule**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

使用 **wdLineSpaceSingle** 、 **wdLineSpace1pt5** 或 **wdLineSpaceDouble** 可将行距设置为这些值之一。要将行距设置为固定磅值或多倍行距，还必须设置 **LineSpacing** 属性。

**示例**

``` JavaScript
/*以下示例为活动文档的第一段设置 2 倍行距。*/
ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceDouble
```

``` JavaScript
/*以下示例返回所选内容的第一段所用的行距标准。*/
let lrule = Selection.Paragraphs.Item(1).LineSpacingRule
```

### ParagraphFormat.LineUnitAfter

返回或设置指定段落的段后间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitAfter**

\*express是一个代表 ParagraphFormat 对象的变量。

### ParagraphFormat.LineUnitBefore

返回或设置指定段落的段前间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitBefore**

\*express是一个代表 ParagraphFormat 对象的变量。

### ParagraphFormat.MirrorIndents

返回或设置一个 **Long** 类型的值，该值代表左缩进和右缩进的宽度是否相同。该属性值可以是 **True** 、 **False** 或 **wdUndefined** 。可读写。

**语法**

**express.MirrorIndents**

\*express是一个代表 ParagraphFormat 对象的变量。

### ParagraphFormat.NoLineNumber

如果取消指定段的行号，则该属性值为 **True** 。该属性值可以是 **True** 、 **False** 或 **wdUndefined** 。可读写 **Long** 类型。

**语法**

**express.NoLineNumber**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

使用 **PageSetup** 对象的 **LineNumbering** 属性可打开行号。

### ParagraphFormat.OutlineLevel

返回或设置指定段落的大纲级别。可读写 **WdOutlineLevel** 类型。

**语法**

**express.OutlineLevel**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

如果段落应用了一种标题样式（从“标题 1”到“标题 9”），则大纲级别与该标题样式相同，并且不能更改。大纲级别仅在大纲视图或文档结构图窗格中可见。

### ParagraphFormat.PageBreakBefore

如果在指定段落前插入了分页符，则该属性值为 **True** 。该属性值可以是 **True** 、 **False** 或 **wdUndefined** 。可读写 **Long** 类型。

**语法**

**express.PageBreakBefore**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例在所选内容的第一段前插入分页符。*/
Selection.Paragraphs.Item(1).PageBreakBefore = true
```

### ParagraphFormat.Parent

返回一个 **Object** 类型值，该值代表指定 **ParagraphFormat** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ParagraphFormat 对象的变量。

### ParagraphFormat.ReadingOrder

返回或设置指定段落的读取次序而不改变其对齐方式。可读写 **WdReadingOrder** 类型。

**语法**

**express.ReadingOrder**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

使用 **Selection** 对象的 **LtrPara** 、 **LtrRun** 、 **RtlPara** 和 **RtlRun** 方法可更改段落的对齐方式和读取次序。

**示例**

``` JavaScript
/*以下示例将第一段的读取次序设置为从右向左。*/
ActiveDocument.Paragraphs.Item(1).ReadingOrder = wdReadingOrderRtl
```

### ParagraphFormat.RightIndent

返回或设置指定段落的右缩进量（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.RightIndent**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中所有段落的右缩进都设置为距右边距 1 英寸。InchesToPoints 方法用来将英寸转换为磅。*/
ActiveDocument.Paragraphs.RightIndent = InchesToPoints(1)
```

### ParagraphFormat.Shading

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 ParagraphFormat 对象的变量。

### ParagraphFormat.SpaceAfter

返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceAfter**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中所有段落的段后间距设置为 12 磅。*/
ActiveDocument.Range().ParagraphFormat.SpaceAfter = 12
```

### ParagraphFormat.SpaceAfterAuto

如果 WPS 自动设置指定段落的段后间距，则该属性为 **True** 。可读/写 **Long** 类型。

**语法**

**express.SpaceAfterAuto**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceAfterAuto** 属性设置为 **True** ，则返回 **wdUndefined** 。该属性可被设置为 **True** 或 **False** 。

如果将 **SpaceAfterAuto** 设置为 **True** ，则 **SpaceAfter** 属性将被忽略。

**示例**

``` JavaScript
/*本示例显示关于活动文档的 SpaceAfterAuto 属性设置的报告。*/
function test()
{
let saa = ActiveDocument.Range().ParagraphFormat.SpaceAfterAuto
switch(saa) {
    case 1:
    MsgBox("Spacing after paragraphs is handled automatically for all paragraphs.")
    case 0:
    MsgBox("Spacing after paragraphs is handled manually for all paragraphs.")
    case null:
    MsgBox("Spacing after paragraphs is handled automatically for some paragraphs,manually for some paragraphs.")
}
}
```

### ParagraphFormat.SpaceBefore

返回或设置指定段落的段前间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceBefore**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中所有段落的段前间距设置为 12 磅。*/
ActiveDocument.Range().ParagraphFormat.SpaceBefore = 12
```

### ParagraphFormat.SpaceBeforeAuto

如果 WPS 自动设置指定段落的段前间距，则该属性为 **True** 。可读/写 **Long** 类型。

**语法**

**express.SpaceBeforeAuto**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceBeforeAuto** 属性设置为 **True** ，则返回 **wdUndefined** 。该属性可被设置为 **True** 或 **False** 。

如果将 **SpaceBeforeAuto** 设置为 **True** ，则 **SpaceBefore** 属性将被忽略。

**示例**

``` JavaScript
/*本示例显示关于活动文档的 SpaceBeforeAuto 属性设置的报告。*/
function test()
{
let saa = ActiveDocument.Range().ParagraphFormat.SpaceBeforeAuto
switch(saa) {
    case 1:
    MsgBox( "Spacing after paragraphs is handled automatically for all paragraphs.")
    case 0:
    MsgBox( "Spacing after paragraphs is handled manually for all paragraphs.")
    case null:
    MsgBox("Spacing after paragraphs is handled automatically for some paragraphs,manually for some paragraphs.")
}
}
```

### ParagraphFormat.Style

返回或设置指定对象的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

要设置该属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。

### ParagraphFormat.TabStops

返回或设置一个 **TabStops** 集合，该集合代表指定段落中的所有自定义制表位。可读写。

**语法**

**express.TabStops**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例在活动文档各段的 2 英寸处添加居中的制表位。InchesToPoints 方法用来将英寸转换为磅。*/
ActiveDocument.Paragraphs.TabStops.Add(InchesToPoints(2), wdAlignTabCenter)
```

``` JavaScript
/*以下示例将文档中各段的制表位设置为与第一段的制表位一致。*/
function test()
{
let para1Tabs = ActiveDocument.Paragraphs.Item(1).TabStops
ActiveDocument.Paragraphs.TabStops = para1Tabs
}
```

### ParagraphFormat.TextboxTightWrap

返回或设置一个 **WdTextboxTightWrap** 常量，该常量代表文本环绕形状或文本框的紧密程度。可读写。

**语法**

**express.TextboxTightWrap**

\*express是一个代表 ParagraphFormat 对象的变量。

### ParagraphFormat.WidowControl

在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 **True** 。该属性值可以是 **True** 、 **False** 或 **wdUndefined** 。可读写 **Long** 类型。

**语法**

**express.WidowControl**

\*express是一个代表 ParagraphFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例给活动文档中的各段落设置格式，从而使段中的首行或末行有可能单独位于上页的页尾或下页的页首。*/
ActiveDocument.Paragraphs.WidowControl = false
```

### ParagraphFormat.WordWrap

如果 WPS 在指定段落或文本框架的西文单词中间断字换行，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.WordWrap**

\*express是一个代表 ParagraphFormat 对象的变量。

**说明**

如果只有一些指定段落或文本框架的此属性设置为 **True** ，则此属性返回 **wdUndefined** 。根据您所选择或安装的语言支持（例如，美国英语），这种用法可能不可用。

**示例**

``` JavaScript
/*以下示例设置 WPS 在活动文档的第一段中的西文单词中间断字换行。*/
ActiveDocument.Paragraphs.Item(1).WordWrap = true
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ParagraphFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Reviewer 对象

## 简介

代表已修订过的文档的单个审阅者。 Reviewer 对象是 Reviewers 集合的一个成员。

## 说明

使用 Reviewers ( *索引* ) （其中 *index* 是审阅者的名称或编号）可返回 Reviewer 对象。 使用 Visible 属性可以显示或隐藏文档中的单个审阅者。

## 属性

| 序号 | 名称                         | 说明               |
|:----:|:-----------------------------|:-------------------|
|  1   | [Visible](#Reviewer.Visible) | 指定的对象是否可见 |

## 成员属性

### Reviewer.Visible

指定的对象是否可见

**语法**

**express.Visible**

\*express是一个代表 Reviewer 对象的变量。

**示例**

``` JavaScript
/*
下面的代码示例隐藏名为"zhangsan"的审阅者，并显示名为"lisi"的审阅者。 这假定"zhangsan"和"lisi"是 Reviewers 集合 的成员。 如果没有这两个审阅者，您将收到错误。
*/
function test()
{
    let vw = Application.ActiveWindow.View 
    if (vw.Reviewers.Count < 2)
        return 
    vw.Reviewers.Item("chongge").Visible = false 
    vw.Reviewers("lisi").Visible = true
 }
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Reviewer](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Style 对象

## 简介

代表单个内置样式或用户定义的样式。 Style 对象将样式属性（如字体、字形、段落间距）添加为 Style 对象的属性。 Style 对象是 Styles 集合的成员。 Styles 集合包含指定文档中的所有样式。

## 说明

使用 Styles ( *Index* ) 可返回一个 Style 对象，其中 *Index* 为样式名、 WdBuiltinStyle 常量或索引号。样式名的拼写和间距必须完全匹配，但是其大小写不必完全匹配。以下示例更改活动文档中名为“Color”的用户定义的样式的字体名称。

``` JavaScript
Application.ActiveDocument.Styles.Item("Color").Font.Name = "Arial"
```

以下示例将内置“标题 1”样式设置为非加粗。

``` JavaScript
Application.ActiveDocument.Styles.Item(wdStyleHeading1).Font.Bold = false
```

样式索引编号代表样式在按字母顺序排列的样式名列表中的位置。请注意， Styles(1) 是按字母顺序排列的列表中的第一个样式。以下示例显示 Styles 集合中第一个样式的基准样式和样式名。

``` JavaScript
alert("Base style= " + Application.ActiveDocument.Styles.Item(1).BaseStyle + "\r" + "Style name= " + ActiveDocument.Styles.Item(1).NameLocal)
```

要将样式应用于范围、段落或多个段落，请将 Style 属性设为用户定义的样式名或内置样式名。以下示例对活动文档的前四段应用“正文”样式。

``` JavaScript
let myRange = Application.ActiveDocument.Range(ActiveDocument.Paragraphs.Item(1).Range.Start,ActiveDocument.Paragraphs.Item(4).Range.End)
myRange.Style = wdStyleNormal
```

以下示例对所选内容的第一段应用“标题 1”样式。

``` JavaScript
Application.Selection.Paragraphs.Item(1).Style = wdStyleHeading1
```

以下示例创建一个名为“Bolded”的字符样式，并将其应用于所选内容。

``` JavaScript
let myStyle = Application.ActiveDocument.Styles.Add("Bolded",wdStyleTypeCharacter)
myStyle.Font.Bold = true
Application.Selection.Range.Style = "Bolded"
```

使用 OrganizerCopy 方法可在文档和模板之间复制样式。使用 UpdateStyles 方法可更新活动文档中的样式，使其与附加模板中的样式定义相匹配。使用 OpenAsDocument 方法可将模板作为文档打开，以便可以修改该模板的样式。

## 方法

| 序号 | 名称                                            | 说明                                                                 |
|:----:|:------------------------------------------------|:---------------------------------------------------------------------|
|  1   | [Delete](#Style.Delete)                         | 删除指定的样式。                                                     |
|  2   | [LinkToListTemplate](#Style.LinkToListTemplate) | 将指定的样式与一个列表模板链接，如此该样式中的格式就可以应用于列表。 |

## 属性

| 序号 | 名称                                                                              | 说明                                                                                                              |
|:----:|:----------------------------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Style.Application)                                                 | 返回一个代表 WPS 应用程序的 Application 对象。                                                                    |
|  2   | [AutomaticallyUpdate](#Style.AutomaticallyUpdate)                                 | 如果样式根据选定内容重新定义，则该属性值为 True 。 Boolean 类型，可读写。                                         |
|  3   | [BaseStyle](#Style.BaseStyle)                                                     | 返回或设置一种现有样式，根据此样式可以设置其他样式。 Variant 类型，可读写。                                       |
|  4   | [Borders](#Style.Borders)                                                         | 返回一个 Borders 集合，该集合代表指定样式的所有边框。                                                             |
|  5   | [BuiltIn](#Style.BuiltIn)                                                         | 如果指定对象是 WPS 的内置样式或题注标签之一，则该属性值为 True 。 Boolean 类型，只读。                            |
|  6   | [Creator](#Style.Creator)                                                         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                      |
|  7   | [Description](#Style.Description)                                                 | 返回指定样式的说明。只读 String 类型。                                                                            |
|  8   | [Font](#Style.Font)                                                               | 返回或设置 Font 对象，该对象代表指定样式的字符格式。 Font 类型，可读写。                                          |
|  9   | [Frame](#Style.Frame)                                                             | 返回一个 Frame 对象，该对象代表指定样式的图文框格式。只读。                                                       |
|  10  | [InUse](#Style.InUse)                                                             | 如果为 True ，则指定的样式是一个已在文档中修改或应用的内置样式，或者是在文档中新建的样式。 Boolean 类型，只读。   |
|  11  | [LanguageID](#Style.LanguageID)                                                   | 返回或设置一个 WdLanguageID 常量，该常量代表指定范围的语言。可读写。                                              |
|  12  | [LanguageIDFarEast](#Style.LanguageIDFarEast)                                     | 返回或设置指定对象的东亚语言。可读写 WdLanguageID 类型。                                                          |
|  13  | [Linked](#Style.Linked)                                                           | 返回或设置 Boolean 值，该值表示某样式是否是段落和字符格式都可以使用的链接样式。只读。                             |
|  14  | [LinkStyle](#Style.LinkStyle)                                                     | 设置或返回一个 Variant 类型的值，该值代表段落和字符样式之间的链接。可读写。                                       |
|  15  | [ListLevelNumber](#Style.ListLevelNumber)                                         | 返回指定样式的列表级别。只读 Long 类型。                                                                          |
|  16  | [ListTemplate](#Style.ListTemplate)                                               | 返回一个 ListTemplate 对象，该对象代表指定 Style 对象的列表格式。                                                 |
|  17  | [Locked](#Style.Locked)                                                           | 如果不能更改或编辑样式，则该属性值为 True 。可读写 Boolean 类型。                                                 |
|  18  | [NameLocal](#Style.NameLocal)                                                     | 返回由用户语言表示的内置样式的名称。可读写 String 类型。                                                          |
|  19  | [NextParagraphStyle](#Style.NextParagraphStyle)                                   | 返回或设置自动应用于在格式设置为指定样式的段落之后插入的新段落的样式。 Variant 类型，可读写。                     |
|  20  | [NoProofing](#Style.NoProofing)                                                   | 如果拼写和语法检查程序忽略设置为此样式格式的文本，则该属性值为 True 。可读写 Long 类型。                          |
|  21  | [NoSpaceBetweenParagraphsOfSameStyle](#Style.NoSpaceBetweenParagraphsOfSameStyle) | 如果为 True ，则 WPS 会删除设置为使用相同样式的段落的间距。 Boolean 类型，可读写。                                |
|  22  | [ParagraphFormat](#Style.ParagraphFormat)                                         | 返回或设置一个 ParagraphFormat 对象，该对象代表指定样式的段落设置。可读/写。                                      |
|  23  | [Parent](#Style.Parent)                                                           | 返回一个 Object 类型值，该值代表指定 Style 对象的父对象。                                                         |
|  24  | [Priority](#Style.Priority)                                                       | 返回或设置一个 Long 类型的值，该值代表 “样式” 任务窗格中排序样式的优先级。可读写。                                |
|  25  | [QuickStyle](#Style.QuickStyle)                                                   | 返回或设置一个 Boolean 类型的值，该值代表样式是否对应于可用的快速样式。可读写。                                   |
|  26  | [Shading](#Style.Shading)                                                         | 返回一个 Shading 对象，该对象代表指定对象的底纹格式。                                                             |
|  27  | [Table](#Style.Table)                                                             | 返回一个 TableStyle 对象，该对象代表可通过使用表格样式应用于表格的属性。                                          |
|  28  | [Type](#Style.Type)                                                               | 返回样式的类型。只读 WdStyleType 类型。                                                                           |
|  29  | [UnhideWhenUsed](#Style.UnhideWhenUsed)                                           | 如果指定样式在用于文档中后显示为 WPS 2015 中 “样式” 和 “样式” 任务窗格中的推荐样式，则该属性值为 True 。可读/写。 |
|  30  | [Visibility](#Style.Visibility)                                                   | 如果指定样式显示为 “样式” 库和 “样式” 任务窗格中的推荐样式，则该属性值为 True 。可读/写。                         |

## 成员方法

### Style.Delete

删除指定的样式。

**语法**

**express.Delete ()**

\*express是一个代表 Style 对象的变量。

### Style.LinkToListTemplate

将指定的样式与一个列表模板链接，如此该样式中的格式就可以应用于列表。

**语法**

**express.LinkToListTemplate ( *ListTemplate* , *ListLevelNumber* )**

\*express是一个代表 Style 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型            | 说明                                                               |
|-------------------|-----------|---------------------|--------------------------------------------------------------------|
| *ListTemplate*    | 必选      | ListTemplate object | 与样式相链接的列表模板。                                           |
| *ListLevelNumber* | 可选      | Variant             | 对应样式所链接到的列表级的整数。如果省略该参数，则使用样式的级别。 |

**示例**

``` JavaScript
/*本示例创建一个新的列表模板，接着将 1 到 9 的标题样式与 1 到 9 的列表级别相链接。然后将新列表模板应用于文档。任何应用标题样式的段落都根据列表模板进行编号。*/
function test()
{
let Lis
let Sty
let ltTemp = ActiveDocument.ListTemplates.Add(true)
for(let i=1;i <= 9;i++) {
    Lis = ltTemp.ListLevels.Item(i)
    Lis.NumberStyle = wdListNumberStyleArabic
    Lis.NumberPosition = InchesToPoints(0.25 * (i - 1))
    Lis.TextPosition = InchesToPoints(0.25 * i)
    Lis.NumberFormat = "%" + i + "."
    Sty = ActiveDocument.Styles.Item("标题 " + i)
    Sty.LinkToListTemplate(ltTemp)
}
ActiveDocument.Content.ListFormat.ApplyListTemplate(ltTemp)
}
```

## 成员属性

### Style.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Style 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Style.AutomaticallyUpdate

如果样式根据选定内容重新定义，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutomaticallyUpdate**

\*express是一个代表 Style 对象的变量。

**说明**

如果将 **AutomaticallyUpdate** 属性设置为 **False** ，则 WPS 会在根据选定内容重新定义样式之前，提示确认。在将样式应用于具有相同样式、但手动格式设置不同的选定内容时，可以重新定义样式。

**示例**

``` JavaScript
/*本示例创建名为“Style1”的样式，该样式不需进行确认即可重新定义。*/
function test()
{
let docNew = Documents.Add()
let styleNew = docNew.Styles.Add("Style1")
styleNew.BaseStyle = docNew.Styles.Item(wdStyleNormal)
styleNew.ParagraphFormat.LineSpacingRule = wdLineSpaceDouble
styleNew.AutomaticallyUpdate = true
}
```

### Style.BaseStyle

返回或设置一种现有样式，根据此样式可以设置其他样式。 **Variant** 类型，可读写。

**语法**

**express.BaseStyle**

\*express是一个代表 Style 对象的变量。

**说明**

要设置 **BaseStyle** 属性，请指定基本样式的本地名称（一个整数或 **wdBuiltinStyle** 常量），或代表基本样式的对象。有关 **wdBuiltinStyle** 常量的列表，请参阅要设置的对象的 **Style** 属性。

**示例**

``` JavaScript
/*本示例新建一篇文档，然后为其新添一个名为“myHeading”的段落样式，并以样式“标题 1”为新样式的基本样式。然后将新样式的左缩进量设置为 1 英寸（72 磅）。*/
function test()
{
let docNew = Documents.Add()
let styleNew = docNew.Styles.Add("NewHeading1")
styleNew.BaseStyle = docNew.Styles.Item(wdStyleHeading1)
styleNew.ParagraphFormat.LeftIndent = 72
}
```

``` JavaScript
/*本示例返回文档正文段落样式所用的基本样式。*/
function test()
{
let styleBase = ActiveDocument.Styles.Item(wdStyleBodyText).BaseStyle
alert(styleBase)
}
```

### Style.Borders

返回一个 **Borders** 集合，该集合代表指定样式的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Style 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Style.BuiltIn

如果指定对象是 WPS 的内置样式或题注标签之一，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.BuiltIn**

\*express是一个代表 Style 对象的变量。

**说明**

可用 **WdBuiltinStyle** 常量指定所有语言通用的内置样式，也可用 WPS 的某个语言版本的样式名来指定某种语言使用的内置样式。例如，如果在 WPS Office的语言设置中指定美式英语，则下列语句是等效的：

``` JavaScript
function test()
{
ActiveDocument.Styles.Item(wdStyleHeading1)
ActiveDocument.Styles.Item("Heading 1")
}
```

**示例**

``` JavaScript
/*本示例实现的功能是：检查活动文档中的所有样式，如果检查到一个非内置样式，则显示该样式的名称。*/
function test()
{
for(let i=1;i <= ActiveDocument.Styles.Count;i++) {
    if(ActiveDocument.Styles.Item(i).BuiltIn == false) {
        alert(ActiveDocument.Styles.Item(i).NameLocal)
    }
}
}
```

### Style.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Style 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Style.Description

返回指定样式的说明。只读 **String** 类型。

**语法**

**express.Description**

\*express是一个代表 Style 对象的变量。

**说明**

样式说明的典型示例可以是：“正文+字体：Arial、12 磅、加粗、倾斜、段前 12 磅、段后 3 磅、KeepWithNext、2 级”。

**示例**

``` JavaScript
/*以下示例新建一个文档，并插入一个制表符分隔的活动文档样式及其说明的列表*/
function test() {
    let docActive = Application.ActiveDocument
    let docNew = Application.Documents.Add()
    for(let i=1;i <= docActive.Styles.Count;i++) {
        let Rng = docNew.Range()
        Rng.InsertAfter(ActiveDocument.Styles.Item(i).NameLocal + "\t" + ActiveDocument.Styles.Item(i).Description)
        Rng.InsertParagraphAfter()
    }
}
```

### Style.Font

返回或设置 **Font** 对象，该对象代表指定样式的字符格式。 **Font** 类型，可读写。

**语法**

**express.Font**

\*express是一个代表 Style 对象的变量。

**说明**

要设置该属性，需指定一个返回 **Font** 对象的表达式。

**示例**

``` JavaScript
/*本示例将取消活动文档的“标题 1”样式中的加粗格式*/
Application.ActiveDocument.Styles.Item(wdStyleHeading1).Font.Bold = false
```

### Style.Frame

返回一个 **Frame** 对象，该对象代表指定样式的图文框格式。只读。

**语法**

**express.Frame**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例创建一个带图文框格式的样式，然后将该样式用于选定内容的第一段。*/
function test()
{
let styleNew = ActiveDocument.Styles.Add("frame",wdStyleTypeParagraph)
let Fra = styleNew.Frame
Fra.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
Fra.HeightRule = wdFrameAuto
Fra.WidthRule = wdFrameAuto
Fra.TextWrap = true
Selection.Paragraphs.Item(1).Range.Style = "frame"
}
```

### Style.InUse

如果为 **True** ，则指定的样式是一个已在文档中修改或应用的内置样式，或者是在文档中新建的样式。 **Boolean** 类型，只读。

**语法**

**express.InUse**

\*express是一个代表 Style 对象的变量。

**说明**

**InUse** 属性并不一定表明该样式当前是否已应用于文档中的任何文本。例如，删除一个应用了该样式的文本后，该样式的 **InUse** 属性仍然为 **True** 。对于尚未在文档中使用过的内置样式，该属性返回 **False** 。

**示例**

``` JavaScript
/*本示例显示一个消息框，列出活动文档中当前正在使用的所有样式的名称。*/
function test()
{
let docActive = ActiveDocument
let strMessage = "Styles in use:" + "\r"
for(let i = 1;i < = docActive.Styles.Count; i++)  { 
    if(docActive.Styles.Item(i).InUse == true) {
        docActive.Content.Find
        docActive.Content.Find.ClearFormatting()
        docActive.Content.Find.Text = ""
        docActive.Content.Find.Style = docActive.Styles.Item(i)
        docActive.Content.Find.Execute(null,null,null,null,null,null,null,null,true)
        if(docActive.Content.Find.Found == true) {
            strMessage = strMessage + docActive.Styles.Item(i).Name + "\r"
        }
    }
}
alert(strMessage)
}
```

### Style.LanguageID

返回或设置一个 **WdLanguageID** 常量，该常量代表指定范围的语言。可读写。

**语法**

**express.LanguageID**

\*express是一个代表 Style 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdLanguageID** 常量可能不可用。

**示例**

``` JavaScript
/*以下示例将 Title 样式重定义为使用西班牙语校对工具，然后在消息框中显示新样式的说明。*/
function test()
{
ActiveDocument.Styles.Item("Title").LanguageID = wdSpanish
alert(ActiveDocument.Styles.Item("Title").Description)
}
```

### Style.LanguageIDFarEast

返回或设置指定对象的东亚语言。可读写 **WdLanguageID** 类型。

**语法**

**express.LanguageIDFarEast**

\*express是一个代表 Style 对象的变量。

**说明**

推荐使用本方法，以返回或设置在 WPS 东亚版中创建的文档中的东亚文字的语言。

**示例**

``` JavaScript
/*以下示例将所选内容的语言设置为朝鲜语。*/
Selection.LanguageIDFarEast = wdKorean
```

### Style.Linked

返回或设置 **Boolean** 值，该值表示某样式是否是段落和字符格式都可以使用的链接样式。只读。

**语法**

**express.Linked**

\*express是一个代表 Style 对象的变量。

**说明**

如果为 **True** ，表示 **“修改样式”** 对话框中的 **“样式类型”** 选项设置为 **“链接(段落和字符)”** 。

### Style.LinkStyle

设置或返回一个 **Variant** 类型的值，该值代表段落和字符样式之间的链接。可读写。

**语法**

**express.LinkStyle**

\*express是一个代表 Style 对象的变量。

**说明**

一旦将字符样式与段落样式链接起来，这两种样式就采取相同的字符格式。

**示例**

``` JavaScript
/*本示例创建并设置新字符样式的格式，并将字符样式与内置标题样式“标题 1”链接起来，这样“标题 1”样式就可以采用新创建样式的字符格式。*/
function LinkHeadStyle() {
    let styChar1 = ActiveDocument.Styles.Add("标题 1 Characters",wdStyleTypeCharacter)
    styChar1.Font.Name = "Verdana"
    styChar1.Font.Bold = true
    styChar1.Font.Shadow = true
    let Bor = styChar1.Font.Borders.Item(1)
    Bor.LineStyle = wdLineStyleDot
    Bor.LineWidth = wdLineWidth300pt
    Bor.Color = wdColorDarkRed
    ActiveDocument.Styles.Item("标题 1").LinkStyle = ActiveDocument.Styles.Item("标题 1 Characters")
    let Con = ActiveDocument.Content
    Con.InsertParagraphAfter()
    Con.InsertAfter("New Linked Style")
    Con.Select()
    Selection.Collapse(wdCollapseEnd)
    Selection.Style = ActiveDocument.Styles.Item("标题 1")
}
```

### Style.ListLevelNumber

返回指定样式的列表级别。只读 **Long** 类型。

**语法**

**express.ListLevelNumber**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*以下示例显示“标题 3”样式的列表级别。*/
alert(ActiveDocument.Styles.Item(wdStyleHeading3).ListLevelNumber)
```

### Style.ListTemplate

返回一个 **ListTemplate** 对象，该对象代表指定 **Style** 对象的列表格式。

**语法**

**express.ListTemplate**

\*express是一个代表 Style 对象的变量。

**说明**

一个列表模板包含定义特定列表的所有格式。 **“项目符号和编号”** 对话框（ **“格式”** 菜单）中的每个选项卡上有七种格式（不包括 **“无”** ），每一种格式对应于一个列表模板。文档和模板也可以包含列表模板的集合。

### Style.Locked

如果不能更改或编辑样式，则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.Locked**

\*express是一个代表 Style 对象的变量。

### Style.NameLocal

返回由用户语言表示的内置样式的名称。可读写 **String** 类型。

**语法**

**express.NameLocal**

\*express是一个代表 Style 对象的变量。

**说明**

通过设置该属性可为用户定义的样式重新命名，或者为内置样式添加别名。

**示例**

``` JavaScript
/*以下示例显示应用于选定段落的样式名（由用户语言表示）。如果应用于选定段落的样式不止一个，则显示第一个样式名。*/
alert(Selection.Paragraphs.Style.NameLocal)
```

``` JavaScript
/*以下示例在活动文档中为“标题 1”样式添加别名“MyH1”。*/
ActiveDocument.Styles.Item("标题 1").NameLocal = "MyH1"
```

``` JavaScript
/*以下示例将名为“Test”的样式重新命名为“Intro”。*/
ActiveDocument.Styles.Item("Test").NameLocal = "Intro"
```

### Style.NextParagraphStyle

返回或设置自动应用于在格式设置为指定样式的段落之后插入的新段落的样式。 **Variant** 类型，可读写。

**语法**

**express.NextParagraphStyle**

\*express是一个代表 Style 对象的变量。

**说明**

您可以使用样式的本地名称（一个整数或 **WdBuiltinStyle** 常量）或代表新样式的对象，来设置 **NextParagraphStyle** 属性。有关 **WdBuiltinStyle** 常量的列表，请参阅要设置的对象的 **Style** 属性。

**示例**

``` JavaScript
/*本示例在活动文档中设置“标题 1”样式后接“标题 2”样式。*/
ActiveDocument.Styles.Item(wdStyleHeading1).NextParagraphStyle = ActiveDocument.Styles.Item(wdStyleHeading2)
```

``` JavaScript
/*本示例新建一篇文档并添加一个名为“MyStyle”的段落样式。该新样式根据“正文”样式创建，后接“标题 3”样式，左缩进量为 1 英寸（72 磅），设为加粗格式。*/
function test()
{
let myDoc = Documents.Add()
let myStyle = myDoc.Styles.Add("MyStyle")
myStyle.BaseStyle = wdStyleNormal
myStyle.NextParagraphStyle = wdStyleHeading3
myStyle.ParagraphFormat.LeftIndent = 72
myStyle.Font.Bold = true
}
```

### Style.NoProofing

如果拼写和语法检查程序忽略设置为此样式格式的文本，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.NoProofing**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*以下示例设置拼写和语法检查程序，使其忽略活动文档中设置为“Test”样式格式的任何文字。*/
ActiveDocument.Styles.Item("Test").NoProofing = true
```

### Style.NoSpaceBetweenParagraphsOfSameStyle

如果为 **True** ，则 WPS 会删除设置为使用相同样式的段落的间距。 **Boolean** 类型，可读写。

**语法**

**express.NoSpaceBetweenParagraphsOfSameStyle**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
function NoSpace() {
    ActiveDocument.Styles.Item("List 1").NoSpaceBetweenParagraphsOfSameStyle = true
}
```

### Style.ParagraphFormat

返回或设置一个 **ParagraphFormat** 对象，该对象代表指定样式的段落设置。可读/写。

**语法**

**express.ParagraphFormat**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例修改活动文档的标题 2 样式。应用该样式的段落将根据第一个制表位缩进，并且采用 2 倍行距。*/
function test()
{
let Par = ActiveDocument.Styles.Item(wdStyleHeading2).ParagraphFormat
Par.TabIndent(1)
Par.Space2()
}
```

### Style.Parent

返回一个 **Object** 类型值，该值代表指定 **Style** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Style 对象的变量。

### Style.Priority

返回或设置一个 **Long** 类型的值，该值代表 **“样式”** 任务窗格中排序样式的优先级。可读写。

**语法**

**express.Priority**

\*express是一个代表 Style 对象的变量。

### Style.QuickStyle

返回或设置一个 **Boolean** 类型的值，该值代表样式是否对应于可用的快速样式。可读写。

**语法**

**express.QuickStyle**

\*express是一个代表 Style 对象的变量。

### Style.Shading

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Style 对象的变量。

### Style.Table

返回一个 **TableStyle** 对象，该对象代表可通过使用表格样式应用于表格的属性。

**语法**

**express.Table**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例新建一个表格样式，该样式仅为第一行、最后一行和最后一列指定四周边框以及特殊的边框和底纹。*/
function NewTableStyle() {
    let styTable = ActiveDocument.Styles.Add("TableStyle 1", wdStyleTypeTable)
    let Tab = styTable.Table
    //Apply borders around table, a double border to the heading row,
    //a double border to the last column, and shading to last row
    Tab.Borders.Item(wdBorderTop).LineStyle = wdLineStyleSingle
    Tab.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleSingle
    Tab.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleSingle
    Tab.Borders.Item(wdBorderRight).LineStyle = wdLineStyleSingle
    Tab.Condition(wdFirstRow).Borders.Item(wdBorderBottom).LineStyle = wdLineStyleDouble
    Tab.Condition(wdLastColumn).Borders.Item(wdBorderLeft).LineStyle = wdLineStyleDouble
    Tab.Condition(wdLastRow).Shading.BackgroundPatternColor = wdColorGray125
}
```

### Style.Type

返回样式的类型。只读 **WdStyleType** 类型。

**语法**

**express.Type**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条消息，显示活动文档中名为“SubTitle”的样式的样式类型。*/
function test()
{
if(ActiveDocument.Styles.Item("SubTitle").Type == wdStyleTypeParagraph) {
    alert("Paragraph style")
}
else if(ActiveDocument.Styles.Item("SubTitle").Type == wdStyleTypeCharacter) {
    alert("Character style")
}
}
```

### Style.UnhideWhenUsed

如果指定样式在用于文档中后显示为 WPS 2015 中 **“样式”** 和 **“样式”** 任务窗格中的推荐样式，则该属性值为 **True** 。可读/写。

**语法**

**express.UnhideWhenUsed**

\*express是一个代表 Style 对象的变量。

### Style.Visibility

如果指定样式显示为 **“样式”** 库和 **“样式”** 任务窗格中的推荐样式，则该属性值为 **True** 。可读/写。

**语法**

**express.Visibility**

\*express是一个代表 Style 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Style](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Styles 对象

## 简介

Style 对象的集合，这些对象代表文档中的内置样式和用户定义的样式。

## 方法

| 序号 | 名称                 | 说明                                                             |
|:----:|:---------------------|:-----------------------------------------------------------------|
|  1   | [Add](#Styles.Add)   | 返回一个 HeadingStyle 对象，该对象代表添加到文档中的新标题样式。 |
|  2   | [Item](#Styles.Item) | 返回集合中的单个 Style 对象。                                    |

## 属性

| 序号 | 名称                               | 说明                                                                         |
|:----:|:-----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Styles.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Styles.Count)             | 返回一个 Long 类型的值，该值代表集合中样式的数量。只读。                     |
|  3   | [Creator](#Styles.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#Styles.Parent)           | 返回一个 Object 类型值，该值代表指定 Styles 对象的父对象。                   |

## 成员方法

### Styles.Add

返回一个 **HeadingStyle** 对象，该对象代表添加到文档中的新标题样式。

**语法**

**express.Add ( *Style* , *Level* )**

\*express是一个代表 Styles 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                  |
|---------|-----------|----------|-----------------------------------------------------------------------|
| *Style* | 必选      | Variant  | 要添加的样式。要指定该参数，可以使用该样式的字符串名称或 Style 对象。 |
| *Level* | 必选      | Integer  | 代表标题级别的数值。                                                  |

**说明**

在编译目录或图表目录时，都会包含该新标题样式。

**示例**

``` JavaScript
/*本示例在活动文档的开始处添加一个目录，然后向用于创建目录的样式列表添加 Title 样式*/
let myToc = Application.ActiveDocument.TablesOfContents.Add(ActiveDocument.Range(0, 0), true, 1, 3)
myToc.HeadingStyles.Add("title", 2)
```

### Styles.Item

返回集合中的单个 **Style** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Styles 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

## 成员属性

### Styles.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Styles 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Styles.Count

返回一个 **Long** 类型的值，该值代表集合中样式的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Styles 对象的变量。

**说明**

### Styles.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Styles 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Styles.Parent

返回一个 **Object** 类型值，该值代表指定 **Styles** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Styles 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Styles](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Template 对象

## 简介

代表文档模板。 Template 对象是 Templates 集合的一个成员。 Templates 集合包含所有可用的 Template 对象。

## 说明

可以使用 Templates ( *Index* ) 返回单个 Template 对象，其中 *Index* 是模板名称或索引编号。如果 Memo2.dot 模板在 Templates 集合中，以下示例将保存该模板。

``` JavaScript
for(let i = 1; i <= Application.Templates.Count; i++){
    if(Application.Templates.Item(i).Name.toLowerCase() == "memo2.dot"){
        Application.Templates.Item(i).Save()
    }
}
```

索引号代表模板在 Templates 集合中的位置。以下示例打开 Templates 集合中的第一个模板。

``` JavaScript
Application.Templates.Item(1).OpenAsDocument()
```

Add 方法不可用于 Templates 集合。可以改用下列方式之一向 Templates 集合添加模板：

- 使用 Open 方法和 Documents 集合打开基于模板的文档或模板
- 使用 Add 方法和 Documents 集合打开基于模板的新文档
- 使用 Add 方法和 Addins 集合加载共用模板
- 使用 AttachedTemplate 属性和 Document 对象将模板附加到文档

使用 NormalTemplate 属性可返回代表 Normal 模板的模板对象。使用 AttachedTemplate 属性可返回附加到指定文档的模板。

可以使用 DefaultFilePath 属性返回或设置用户或工作组模板的位置（即，用于存储这些模板的文件夹）。以下示例在 “选项” 对话框（通过 “工具” 菜单打开）中的 “文件位置” 选项卡上显示用户模板文件夹。

``` JavaScript
alert(Application.Options.DefaultFilePath(wdUserTemplatesPath))
```

## 方法

| 序号 | 名称                                          | 说明                                             |
|:----:|:----------------------------------------------|:-------------------------------------------------|
|  1   | [OpenAsDocument](#Template.OpenAsDocument%20) | 将指定模板作为文档打开并返回一个 Document 对象。 |
|  2   | [Save](#Template.Save)                        | 保存指定的模板。                                 |

## 属性

| 序号 | 名称                                                             | 说明                                                                                                                                              |
|:----:|:-----------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Template.Application)                             | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                    |
|  2   | [BuildingBlockEntries](#Template.BuildingBlockEntries)           | 返回一个 BuildingBlockEntries 集合，该集合代表模板中构建基块条目的集合。只读。                                                                    |
|  3   | [BuildingBlockTypes](#Template.BuildingBlockTypes)               | 返回一个 BuildingBlockTypes 集合，该集合代表模板中包含的构建基块类型的集合。只读。                                                                |
|  4   | [BuiltInDocumentProperties](#Template.BuiltInDocumentProperties) | 返回一个 DocumentProperties 集合，该集合代表指定文档的所有内置的文档属性。                                                                        |
|  5   | [Creator](#Template.Creator)                                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                      |
|  6   | [CustomDocumentProperties](#Template.CustomDocumentProperties)   | 返回一个 DocumentProperties 集合，该集合代表指定文档的所有自定义文档属性。                                                                        |
|  7   | [FarEastLineBreakLanguage](#Template.FarEastLineBreakLanguage)   | 返回或设置在指定的文档或模板中对文本进行换行时使用的东亚语言。 WdFarEastLineBreakLanguageID 类型，可读写。                                        |
|  8   | [FarEastLineBreakLevel](#Template.FarEastLineBreakLevel)         | 返回或设置指定文档的换行控制级别。 WdFarEastLineBreakLevel 类型，可读写。                                                                         |
|  9   | [FullName](#Template.FullName)                                   | 指定模板的名称，其中包含驱动器或 Web 路径。 String 类型，只读。                                                                                   |
|  10  | [JustificationMode](#Template.JustificationMode)                 | 返回或设置指定模板的字符间距调整。可读写 WdJustificationMode 类型。                                                                               |
|  11  | [KerningByAlgorithm](#Template.KerningByAlgorithm)               | 如果 WPS 调整指定文档中的半角西文字符与标点符号的字距，则该属性值为 True 。可读写 Boolean 类型。                                                  |
|  12  | [LanguageID](#Template.LanguageID)                               | 返回或设置一个 WdLanguageID 常量，该常量代表指定范围的语言。可读写。                                                                              |
|  13  | [LanguageIDFarEast](#Template.LanguageIDFarEast)                 | 返回或设置指定对象的东亚语言。可读写 WdLanguageID 类型。                                                                                          |
|  14  | [ListTemplates](#Template.ListTemplates)                         | 返回一个 ListTemplates 集合，该集合代表指定文档、模板或列表库的所有列表格式。只读。                                                               |
|  15  | [Name](#Template.Name)                                           | 返回指定对象的名称。 String 类型，只读。                                                                                                          |
|  16  | [NoLineBreakAfter](#Template.NoLineBreakAfter)                   | 返回或设置 WPS 不在其后面进行换行的首尾字符。可读写 String 类型。                                                                                 |
|  17  | [NoLineBreakBefore](#Template.NoLineBreakBefore)                 | 返回或设置 WPS 不在其前面进行换行的首尾字符。可读写 String 类型。                                                                                 |
|  18  | [NoProofing](#Template.NoProofing)                               | 如果拼写和语法检查程序忽略基于此模板的文档，则该属性值为 True 。可读写 Long 类型。                                                                |
|  19  | [Parent](#Template.Parent)                                       | 返回一个 Object 类型值，该值代表指定 Template 对象的父对象。                                                                                      |
|  20  | [Path](#Template.Path)                                           | 返回指定文档模板的路径。只读 String 类型。                                                                                                        |
|  21  | [Saved](#Template.Saved)                                         | 如果自上次保存之后指定的模板未更改，则该属性值为 True 。如果关闭文档时 WPS 提示保存对文档所做的更改，则该属性值为 False 。 Boolean 类型，可读写。 |
|  22  | [Type](#Template.Type)                                           | 返回模板的类型。只读 WdTemplateType 类型。                                                                                                        |
|  23  | [VBProject](#Template.VBProject)                                 | 返回指定模板的 VBProject 对象。                                                                                                                   |

## 成员方法

### Template.OpenAsDocument

将指定模板作为文档打开并返回一个 **Document** 对象。

**语法**

**express.OpenAsDocument ()**

\*express是一个代表 Template 对象的变量。

**返回值**

Document

**说明**

将模板作为文档打开则允许编辑该模板的内容。当 **Template** 对象中的某属性或方法（例如 **Styles** 属性）无效时，就可能需要使用这种方法。

**示例**

``` JavaScript
/*本示例打开活动文档所选用的模板，如果该模板包含了不止一个段落标记的内容，则显示一个消息框，然后关闭该模板。*/
function test()
{
let docNew = ActiveDocument.AttachedTemplate.OpenAsDocument()

if(docNew.Content.Text != "\r"){
    MsgBox("Template is not empty")
}
else{
    MsgBox("Template is empty")
}

docNew.Close(wdDoNotSaveChanges)
}
```

``` JavaScript
/*本示例为 Normal 模板保存一个副本“Backup.dot”。*/
function test()
{
let docNew = NormalTemplate.OpenAsDocument()
docNew.SaveAs("Backup.dot")
docNew.Close(wdDoNotSaveChanges)
}
```

``` JavaScript
/*本示例更改活动文档所选用的模板中“标题 1”的样式设置。UpdateStyles 方法用来更新活动文档中的样式。*/
function test()
{
let docNew = ActiveDocument.AttachedTemplate.OpenAsDocument()
let fontStyle = docNew.Styles.Item(wdStyleHeading1).Font

fontStyle.Name = "Arial"
fontStyle.Size = 16
fontStyle.Bold = false

docNew.Close(wdSaveChanges)
ActiveDocument.UpdateStyles()
}
```

### Template.Save

保存指定的模板。

**语法**

**express.Save ()**

\*express是一个代表 Template 对象的变量。

**说明**

如果之前未保存过此模板，则 **“另存为”** 对话框将提示用户输入文件名。

## 成员属性

### Template.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Template 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Template.BuildingBlockEntries

返回一个 **BuildingBlockEntries** 集合，该集合代表模板中构建基块条目的集合。只读。

**语法**

**express.BuildingBlockEntries**

\*express是一个代表 Template 对象的变量。

### Template.BuildingBlockTypes

返回一个 **BuildingBlockTypes** 集合，该集合代表模板中包含的构建基块类型的集合。只读。

**语法**

**express.BuildingBlockTypes**

\*express是一个代表 Template 对象的变量。

### Template.BuiltInDocumentProperties

返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有内置的文档属性。

**语法**

**express.BuiltInDocumentProperties**

\*express是一个代表 Template 对象的变量。

**说明**

要返回单个代表特定内置文档属性的 **DocumentProperty** 对象，可使用 **BuiltinDocumentProperties** 属性。如果 WPS 没有为一个内置的文档属性定义一个值，则读取这个文档属性的 **Value** 属性时会产生一个错误。

用 **CustomDocumentProperties** 属性返回自定义文档属性的集合。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Template.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Template 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Template.CustomDocumentProperties

返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有自定义文档属性。

**语法**

**express.CustomDocumentProperties**

\*express是一个代表 Template 对象的变量。

**说明**

使用 **BuiltInDocumentProperties** 属性可返回内置文档属性的集合。 **msoPropertyTypeString** 类型的属性的长度不能超过 255 个字符。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Template.FarEastLineBreakLanguage

返回或设置在指定的文档或模板中对文本进行换行时使用的东亚语言。 **WdFarEastLineBreakLanguageID** 类型，可读写。

**语法**

**express.FarEastLineBreakLanguage**

\*express是一个代表 Template 对象的变量。

### Template.FarEastLineBreakLevel

返回或设置指定文档的换行控制级别。 **WdFarEastLineBreakLevel** 类型，可读写。

**语法**

**express.FarEastLineBreakLevel**

\*express是一个代表 Template 对象的变量。

**说明**

如果 **FarEastLineBreakControl** 属性设为 **False** ，则忽略该属性。

**示例**

``` JavaScript
/*本示例设置 WPS 在活动文档中第一级首尾字符处换行。*/
ActiveDocument.FarEastLineBreakLevel = wdJustificationModeCompressKana
```

### Template.FullName

指定模板的名称，其中包含驱动器或 Web 路径。 **String** 类型，只读。

**语法**

**express.FullName**

\*express是一个代表 Template 对象的变量。

**说明**

使用该属性与按顺序使用 **Path、PathSeparator** 和 **Name** 属性的作用相同。

**示例**

``` JavaScript
/*本示例显示活动文档附加模板的路径和文件名*/
function TemplateName(){
    alert(Application.ActiveDocument.AttachedTemplate.FullName)
}
```

### Template.JustificationMode

返回或设置指定模板的字符间距调整。可读写 **WdJustificationMode** 类型。

**语法**

**express.JustificationMode**

\*express是一个代表 Template 对象的变量。

**示例**

``` JavaScript
/*以下示例设置 WPS 在调整字符间距时只压缩标点符号。*/
NormalTemplate.JustificationMode = wdJustificationModeCompressKana
```

### Template.KerningByAlgorithm

如果 WPS 调整指定文档中的半角西文字符与标点符号的字距，则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.KerningByAlgorithm**

\*express是一个代表 Template 对象的变量。

**示例**

``` JavaScript
/*以下示例设置 WPS 调整 Normal 模板中的半角西文字符与标点符号的字距。*/
NormalTemplate.KerningByAlgorithm = true
```

### Template.LanguageID

返回或设置一个 **WdLanguageID** 常量，该常量代表指定范围的语言。可读写。

**语法**

**express.LanguageID**

\*express是一个代表 Template 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdLanguageID** 常量可能不可用。

### Template.LanguageIDFarEast

返回或设置指定对象的东亚语言。可读写 **WdLanguageID** 类型。

**语法**

**express.LanguageIDFarEast**

\*express是一个代表 Template 对象的变量。

**说明**

推荐使用本方法，以返回或设置在 WPS 东亚版中创建的文档中的东亚文字的语言。

**示例**

``` JavaScript
/*以下示例将所选内容的语言设置为朝鲜语。*/
NormalTemplate.LanguageIDFarEast = wdKorean
```

### Template.ListTemplates

返回一个 **ListTemplates** 集合，该集合代表指定文档、模板或列表库的所有列表格式。只读。

**语法**

**express.ListTemplates**

\*express是一个代表 Template 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Template.Name

返回指定对象的名称。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 Template 对象的变量。

### Template.NoLineBreakAfter

返回或设置 WPS 不在其后面进行换行的首尾字符。可读写 **String** 类型。

**语法**

**express.NoLineBreakAfter**

\*express是一个代表 Template 对象的变量。

**示例**

``` JavaScript
/*以下示例实现的功能是：在活动文档中，设置“$”、“(”、“[”、“\”和“{”作为首尾字符时，WPS 不在这些首尾字符的后面进行换行。*/
ActiveDocument.NoLineBreakAfter = "$([\{"
```

### Template.NoLineBreakBefore

返回或设置 WPS 不在其前面进行换行的首尾字符。可读写 **String** 类型。

**语法**

**express.NoLineBreakBefore**

\*express是一个代表 Template 对象的变量。

**示例**

``` JavaScript
/*以下示例实现的功能是：在活动文档中，设置“!”、“)”和“]”作为首尾字符时， WPS 不在这些首尾字符的前面进行换行。*/
NormalTemplate.NoLineBreakBefore = "!)]"
```

### Template.NoProofing

如果拼写和语法检查程序忽略基于此模板的文档，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.NoProofing**

\*express是一个代表 Template 对象的变量。

### Template.Parent

返回一个 **Object** 类型值，该值代表指定 **Template** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Template 对象的变量。

### Template.Path

返回指定文档模板的路径。只读 **String** 类型。

**语法**

**express.Path**

\*express是一个代表 Template 对象的变量。

**说明**

该路径不包括尾部字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符。使用 **Name** 属性可返回不带路径的文件名，而使用 **FullName** 属性可同时返回文件名和路径。

### Template.Saved

如果自上次保存之后指定的模板未更改，则该属性值为 **True** 。如果关闭文档时 WPS 提示保存对文档所做的更改，则该属性值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.Saved**

\*express是一个代表 Template 对象的变量。

**示例**

``` JavaScript
/*以下示例将 Normal 模板的状态设置为未更改。如果更改了 Normal 模板，则退出 WPS 时不保存所做的更改。*/
function test()
{
NormalTemplate.Saved = true
Application.Quit()
}
```

### Template.Type

返回模板的类型。只读 WdTemplateType 类型。

**语法**

**express.Type**

\*express是一个代表 Template 对象的变量。

### Template.VBProject

返回指定模板的 **VBProject** 对象。

**语法**

**express.VBProject**

\*express是一个代表 Template 对象的变量。

**说明**

使用该属性可访问代码模块和用户窗体。

要在“对象浏览器”中查看 **VBProject** 对象，必须选中“Visual Basic 编辑器”的 **“引用”** 对话框（ **“工具”** 菜单）中的 **“示例代码 Extensibility”** 复选框。

**示例**

``` JavaScript
/*以下示例显示 Normal 模板的 Visual Basic 项目的名称。*/
function test()
{
Set normProj = NormalTemplate.VBProject
MsgBox normProj.Name
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Template](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AddIn 对象

## 简介

代表已加载或未加载的单个加载项。

## 说明

AddIn 对象是 AddIns 集合的成员。 AddIns 集合包含与 WPP 相关的所有可用加载项，无论这些加载项是否加载。该集合不包含 组件对象模型 (COM) 加载项 （COM 加载项：通过添加自定义命令和指定的功能来扩展 WPS Office程序的功能的补充程序。COM 加载项可在一个或多个 Office 程序中运行。COM 加载项使用文件扩展名 .dll 或 .exe。） 。

## 属性

| 序号 | 名称                              | 说明                                                                                                                  |
|:----:|:----------------------------------|:----------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#AddIn.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                               |
|  2   | [AutoLoad](#AddIn.AutoLoad)       | 决定指定的加载宏是否在每次启动WPP 时自动加载。MsoTriState 类型，可读/写。                                             |
|  3   | [FullName](#AddIn.FullName)       | 返回指定的加载宏或保存的演示文稿的名称，包括路径、当前文件系统分隔符和文件扩展名。只读。String 类型。                 |
|  4   | [Loaded](#AddIn.Loaded)           | 确定是否已加载指定加载项。MsoTriState 类型，可读/写。                                                                 |
|  5   | [Name](#AddIn.Name)               | 已注册的文件类型的加载项名称（标题）。String 类型，只读。                                                             |
|  6   | [Parent](#AddIn.Parent)           | 返回指定对象的父对象。                                                                                                |
|  7   | [Path](#AddIn.Path)               | 返回 String，该值代表到指定的 AddIn、Application 或 Presentation 对象的路径，或者后跟 MotionEffect 对象的路径。只读。 |
|  8   | [Registered](#AddIn.Registered)   | 返回指定的加载宏是否注册到 Windows 注册表中。MsoTriState 类型，可读/写。                                              |

## 成员属性

### AddIn.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 AddIn 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path & "\\Added Slide")
}
```

``` JavaScript
/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test(){
　　　　for(let i=1;i <= Application.ActivePresentation.Slides.Item(1).Shapes.Count;i++) {
   　　　　 if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {
    　　　　    alert(ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)
  　　　　  }
　　　　}
}
```

### AddIn.AutoLoad

决定指定的加载宏是否在每次启动WPP 时自动加载。MsoTriState 类型，可读/写。

**语法**

**express.AutoLoad**

\*express是一个代表 AddIn 对象的变量。

**说明**

如果将该属性设置为 **msoTrue** ，则自动将 **Registered** 属性设置为 **msoTrue** 。

|                                                |
|------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。  |
| msoCTrue                                       |
| msoFalse                                       |
| msoTriStateMixed                               |
| msoTriStateToggle                              |
| msoTrue 每次启动WPP 时，自动加载指定的加载宏。 |

**示例**

``` JavaScript
/*本示例显示每次启动WPP 时自动加载的每个加载宏的名称。*/
function test(){
    let afound
    for(let i=1;i <= Application.AddIns.Count;i++) {
 　　　　   if(Application.AddIns.Item(i).AutoLoad) {
  　　　　      alert(Application.AddIns.Item(i).Name)
   　　　　     afound = true
   　　　　 }
　　　　}
　　　　if(afound != true) {
　　　　    alert("No add-ins were loaded automatically.")
　　　　}
}
```

本示例指定名为“MyTools”的加载宏在每次启动WPP 时自动加载。

``` JavaScript
Application.AddIns.Item("mytools").AutoLoad = msoTrue
```

### AddIn.FullName

返回指定的加载宏或保存的演示文稿的名称，包括路径、当前文件系统分隔符和文件扩展名。只读。String 类型。

**语法**

**express.FullName**

\*express是一个代表 AddIn 对象的变量。

**说明**

此属性与 Path 属性等效，附加上当前文件系统分隔符和 Name 属性。

**示例**

``` JavaScript
/*以下示例显示每一个可用加载宏的路径及文件名。*/
function test(){
　　　　for(let i=1;i <= Application.AddIns.Count;i++) {
   　　　　 alert(Application.AddIns.Item(i).FullName)
　　　　}
}
```

以下示例显示当前演示文稿的路径和文件名（假设演示文稿已保存）。

``` JavaScript
alert(Application.ActivePresentation.FullName)
```

### AddIn.Loaded

确定是否已加载指定加载项。MsoTriState 类型，可读/写。

**语法**

**express.Loaded**

\*express是一个代表 AddIn 对象的变量。

**说明**

|                                                |
|------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。  |
| msoCTrue                                       |
| msoFalse                                       |
| msoTriStateMixed                               |
| msoTriStateToggle                              |
| msoTrue 指定加载项已加载。在“加载项”选项卡中。 |

**示例**

以下示例将 MyTools.ppa 添加到 **“加载项”** 选项卡中的列表，然后加载它。

``` JavaScript
Application.AddIns.Add("c:\\my documents\\mytools.ppa").Loaded = msoTrue
```

以下示例卸载名为“MyTools”的加载项。

``` JavaScript
Application.AddIns.Item("mytools").Loaded = msoFalse
```

### AddIn.Name

已注册的文件类型的加载项名称（标题）。String 类型，只读。

**语法**

**express.Name**

\*express是一个代表 AddIn 对象的变量。

**说明**

如果包含某对象的集合的 Item 方法采用 Variant 类型的参数，则可以将该对象的名称与 Item 方法联合使用来返回对该对象的引用。例如，如果某个形状的 Name 属性值为“Rectangle 2”，则 .Shapes.Item("Rectangle 2") 将返回对该形状的引用。

### AddIn.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 AddIn 对象的变量。

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

### AddIn.Path

返回 String，该值代表到指定的 AddIn、Application 或 Presentation 对象的路径，或者后跟 MotionEffect 对象的路径。只读。

**语法**

**express.Path**

\*express是一个代表 AddIn 对象的变量。

**说明**

如果使用此属性返回一个尚未保存的演示文稿的路径，则会得到一个空字符串。如果使用此属性返回未加载的加载宏的路径，则会导致错误。 路径中不包括最后的反斜线 (\\ 或指定对象的名称。使用 Presentation 对象的 Name 属性可返回不带路径的文件名，使用 FullName 属性则可返回带路径的文件名。 MotionEffect 对象返回的 String 类型值是一个特殊路径，它使用与 VML 路径说明相同的语法描述移动效果在 From 和 To 之间沿行的路径。

**示例**

``` JavaScript
/*以下示例将当前演示文稿保存到与WPP 相同的文件夹中。 */
function test(){
　　　　let fName = Application.Path + "\\test presentation"
　　　　Application.ActivePresentation.SaveAs(fName)
}
```

### AddIn.Registered

返回指定的加载宏是否注册到 Windows 注册表中。MsoTriState 类型，可读/写。

**语法**

**express.Registered**

\*express是一个代表 AddIn 对象的变量。

**说明**

|                                                 |
|-------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。   |
| msoCTrue                                        |
| msoFalse                                        |
| msoTriStateMixed                                |
| msoTriStateToggle                               |
| msoTrue 将指定的加载宏注册到 Windows 注册表中。 |

**示例**

本示例将名为“MyTools”的加载宏注册到 Windows 注册表中。

``` JavaScript
Application.AddIns.Item("MyTools").Registered = msoTrue
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/AddIn](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AnimationSettings 对象

## 简介

代表幻灯片放映时应用于指定形状的动画的特殊效果。

## 属性

| 序号 | 名称                                                            | 说明                                                                                                                                                                                                                                                                            |
|:----:|:----------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AdvanceMode](#AnimationSettings.AdvanceMode)                   | 返回或设置一个值，该值指示指定形状的动画是仅在被单击时换页还是在经过指定时间后自动换页。可读/写 PpAdvanceMode 类型。如果该形状不会有动画效果，请确保将 TextLevelEffect 属性设置为除 ppAnimateLevelNone 之外的值，并且将 Animate 属性设置为 True 。                              |
|  2   | [AdvanceTime](#AnimationSettings.AdvanceTime)                   | 返回或设置以秒为单位的时间，在该时间过后指定的形状将开始播放动画。可读/写                                                                                                                                                                                                       |
|  3   | [AfterEffect](#AnimationSettings.AfterEffect)                   | 返回或设置 PpAfterEffect 常量，该常量指示指定的形状在生成之后是变暗、隐藏还是不更改。可读/写。                                                                                                                                                                                  |
|  4   | [Animate](#AnimationSettings.Animate)                           | 决定在幻灯片放映过程中指定的形状是否被赋予动画效果。可读/写 MsoTriState 类型。                                                                                                                                                                                                  |
|  5   | [AnimateBackground](#AnimationSettings.AnimateBackground)       | 如果指定的对象是自选图形，若该形状独立于其所包含的文本单独显示动画，则值为 msoTrue 。如果指定的形状是图形对象，若指定的图形对象的背景（坐标轴和网格线）被赋予动画效果，则值为 msoTrue 。仅适用于图形对象或带有经多个步骤才能显示的文本的自选图形。 MsoTriState 类型，可读/写 。 |
|  6   | [AnimateTextInReverse](#AnimationSettings.AnimateTextInReverse) | 决定指定形状是否按相反顺序构造。仅应用于具有多个创建步骤的形状（例如，包含列表的形状）。可读/写。 MsoTriState 类型。                                                                                                                                                            |
|  7   | [AnimationOrder](#AnimationSettings.AnimationOrder)             | 返回或设置一个整数，该整数代表指定形状在动画形状集合中的位置。可读/写 Long 类型。                                                                                                                                                                                               |
|  8   | [Application](#AnimationSettings.Application)                   | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                         |
|  9   | [ChartUnitEffect](#AnimationSettings.ChartUnitEffect)           | 返回或设置一个值，该值指示图形区域是按序列、类别还是按元素播放动画。可读/写 PpChartUnitEffect 类型。                                                                                                                                                                            |
|  10  | [DimColor](#AnimationSettings.DimColor)                         | 返回或设置一个 ColorFormat 对象，该对象代表指定形状生成后的颜色。只读。                                                                                                                                                                                                         |
|  11  | [EntryEffect](#AnimationSettings.EntryEffect)                   | 对于 AnimationSettings 对象，此属性返回或设置应用于指定形状中动画的特殊效果。对于 SlideShowTransition 对象，此属性返回或设置应用于指定幻灯片切换的特殊效果。 PpEntryEffect 类型，可读写。                                                                                       |
|  12  | [Parent](#AnimationSettings.Parent)                             | 返回指定对象的父对象。                                                                                                                                                                                                                                                          |
|  13  | [PlaySettings](#AnimationSettings.PlaySettings)                 | 返回一个 PlaySettings 对象，该对象包含指定的媒体剪辑在幻灯片放映中的播放方式的信息。                                                                                                                                                                                            |
|  14  | [SoundEffect](#AnimationSettings.SoundEffect)                   | 返回一个 SoundEffect 对象，该对象代表切换到指定幻灯片时播放的声音。                                                                                                                                                                                                             |
|  15  | [TextLevelEffect](#AnimationSettings.TextLevelEffect)           | 设置或返回一个 PpTextLevelEffect 常量，该常量表示指定形状中的文本是否由第一级段落、第二级段落或其他级段落（最高为第五级段落）激活动画显示。可读/写。                                                                                                                            |
|  16  | [TextUnitEffect](#AnimationSettings.TextUnitEffect)             | 返回或设置一个 PpTextUnitEffect 常量，该常量指示指定形状中的文本是按段落、按字或是按字母被赋予动画效果。可读/写。                                                                                                                                                               |

## 成员属性

### AnimationSettings.AdvanceMode

返回或设置一个值，该值指示指定形状的动画是仅在被单击时换页还是在经过指定时间后自动换页。可读/写 **PpAdvanceMode** 类型。如果该形状不会有动画效果，请确保将 **TextLevelEffect** 属性设置为除 **ppAnimateLevelNone** 之外的值，并且将 **Animate** 属性设置为 **True** 。

**语法**

**express.AdvanceMode**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

| PpAdvanceMode 可以是下列 PpAdvanceMode 常量之一。 |
|---------------------------------------------------|
| ppAdvanceModeMixed                                |
| ppAdvanceOnClick                                  |
| ppAdvanceOnTime                                   |

**示例**

``` JavaScript
/*本示例将当前演示文稿第一张幻灯片的第二个形状设为五秒后自动播放动画。*/
function test(){
　　　　let Ani = Appliaction.ActivePresentation.Slides.Item(1).Shapes.Item(2).AnimationSettings
　　　　Ani.AdvanceMode = ppAdvanceOnTime
　　　　Ani.AdvanceTime = 5
　　　　Ani.TextLevelEffect = ppAnimateByAllLevels
　　　　Ani.Animate = true
}
```

### AnimationSettings.AdvanceTime

返回或设置以秒为单位的时间，在该时间过后指定的形状将开始播放动画。可读/写

**语法**

**express.AdvanceTime**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

在指定的时间过后，指定的幻灯片动画不会自动开始，除非将动画的 AdvanceMode 属性设置为 ppAdvanceOnTime。

**示例**

``` JavaScript
/*本示例将当前演示文稿第一张幻灯片上的第二个形状设为五秒后自动播放动画。 */
function test(){
　　　　let Ani = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).AnimationSettings
　　　　Ani.AdvanceMode = ppAdvanceOnTime
　　　　Ani.AdvanceTime = 5
　　　　Ani.TextLevelEffect = ppAnimateByAllLevels
　　　　Ani.Animate = true
}
```

### AnimationSettings.AfterEffect

返回或设置 **PpAfterEffect** 常量，该常量指示指定的形状在生成之后是变暗、隐藏还是不更改。可读/写。

**语法**

**express.AfterEffect**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

除非带有动画播放后效果的形状播放动画，且幻灯片上至少有另一个形状在其播放后播放动画；否则无法看到为该形状设置的动画播放后效果。要给一个形状赋予动画效果，必须将 **AnimationSettings** 对象的 **TextLevelEffect** 属性设置为 **ppAnimateLevelNone** 以外的值，并且 **EntryEffect** 属性设置为除 **ppEffectNone** 以外的常量。另外，还必须将 **Animate** 属性必须设置为 **True** 。若要更改幻灯片上形状的创建顺序，请使用 **AnimationOrder** 属性。

| PpAfterEffect 可以是下列 PpAfterEffect 常量之一。 |
|---------------------------------------------------|
| ppAfterEffectDim                                  |
| ppAfterEffectHide                                 |
| ppAfterEffectHideOnClick                          |
| ppAfterEffectMixed                                |
| ppAfterEffectNothing                              |

**示例**

``` JavaScript
/*本示例指定当前演示文稿中第一张幻灯片的标题在创建后变暗。但是，如果它是第一张幻灯片中最后的或唯一的形状，文本将不会变暗。*/
function test(){
　　　　let Ani = Application.ActivePresentation.Slides.Item(1).Shapes.Title.AnimationSettings
　　　　Ani.Animate = true
　　　　Ani.TextLevelEffect = ppAnimateByAllLevels
　　　　Ani.AfterEffect = ppAfterEffectDim
}
```

### AnimationSettings.Animate

决定在幻灯片放映过程中指定的形状是否被赋予动画效果。可读/写 **MsoTriState** 类型。

**语法**

**express.Animate**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

要为一个形状赋予动画效果，则该形状的 **AnimationSettings** 对象的 **TextLevelEffect** 属性必须设置为除 **ppAnimateLevelNone** 之外的值，并且必须将 **Animate** 属性设置为 **True** ，或将 **EntryEffect** 属性设置为除 **ppEffectNone** 之外的常量。

| MsoTriState 可以是下列 MsoTriState 常量之一。        |
|------------------------------------------------------|
| msoCTrue                                             |
| msoFalse                                             |
| msoTriStateMixed                                     |
| msoTriStateToggle                                    |
| msoTrue 在幻灯片放映过程中指定的形状被赋予动画效果。 |

**示例**

``` JavaScript
/*本示例指定活动演示文稿第二张幻灯片的标题在创建后变暗。如果该标题是第二张幻灯片中的最后一个或唯一创建的形状，则该文本不变暗。*/
function test(){
　　　　let Ani = Application.ActivePresentation.Slides.Item(2).Shapes.Title.AnimationSettings
　　　　Ani.TextLevelEffect = ppAnimateByAllLevels
　　　　Ani.AfterEffect = ppAfterEffectDim
　　　　Ani.Animate = msoTrue
}
```

### AnimationSettings.AnimateBackground

如果指定的对象是自选图形，若该形状独立于其所包含的文本单独显示动画，则值为 **msoTrue** 。如果指定的形状是图形对象，若指定的图形对象的背景（坐标轴和网格线）被赋予动画效果，则值为 **msoTrue** 。仅适用于图形对象或带有经多个步骤才能显示的文本的自选图形。 **MsoTriState** 类型，可读/写 。

**语法**

**express.AnimateBackground**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

使用 **TextLevelEffect** 和 **TextUnitEffect** 属性可以控制附加到指定形状的文本的动画效果。

如果将该属性设置为 **MsoTrue** 并将 **TextLevelEffect** 属性设置为 **ppAnimateByAllLevels** ，则形状及其文本将被同时赋予动画效果。如果将该属性设置为 **MsoTrue** 并将 **TextLevelEffect** 属性设置为除 **ppAnimateByAllLevels** 之外的值，则形状将在文本被赋予动画效果之前直接被赋予动画效果。

除非为指定形状赋予动画效果，否则看不到设置该属性的效果。要给一个形状赋予动画效果，必须将 **TextLevelEffect** 属性设置为 **ppAnimateLevelNone** 以外的值，并且将 **Animate** 属性设置为 **MsoTrue** ，或者将 **EntryEffect** 属性设置为除 **ppEffectNone** 之外的常量。

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue                                       |

**示例**

``` JavaScript
/*本示例将创建一个包含文本的矩形。本示例还指定该形状应从右下角飞入，应从第一级段落显示文本，并且该形状应独立于其所包含的文本单独显示动画。在本示例中，EntryEffect 属性用于打开动画效果。*/
function AnimateTextBox() {
    let addShp = Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeRectangle, 50, 200, 200, 200)
        addShp.TextFrame.TextRange.Text = "Reason 1" + "\r" + "Reason 2" + "\r" + "Reason 3"
    let Ani = addShp.AnimationSettings
        Ani.EntryEffect = ppEffectFlyFromBottomRight
        Ani.TextLevelEffect = ppAnimateByFirstLevel
        Ani.TextUnitEffect = ppAnimateByParagraph
        Ani.AnimateBackground = msoTrue
}
```

### AnimationSettings.AnimateTextInReverse

决定指定形状是否按相反顺序构造。仅应用于具有多个创建步骤的形状（例如，包含列表的形状）。可读/写。 **MsoTriState** 类型。

**语法**

**express.AnimateTextInReverse**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue：指定形状按相反顺序构造。             |

指定形状必须被赋予动画效果，您才能看到设置该属性的效果。要给一个形状赋予动画效果，必须将 **AnimationSettings** 对象的 **TextLevelEffect** 属性设置为 **ppAnimateLevelNone** 以外的值，并且 **Animate** 属性必须设置为 **True** 。

**示例**

``` JavaScript
/*以下示例在当前演示文稿的第一张幻灯片后添加一张幻灯片，然后设置标题文本，并为文本占位符添加三个项的列表，并设置该列表按相反顺序构造。*/
function test(){
　　　　let myShape = Application.ActivePresentation.Slides.Add(2, ppLayoutText).Shapes
   　　　　 myShape.Item(1).TextFrame.TextRange.Text = "Top Three Reasons"
　　　　let shape2 = myShape.Item(2)
   　　　　 shape2.TextFrame.TextRange.Text = "Reason 1" + "\r" + "Reason 2" + "\r" + "Reason 3"
　　　　let Ani = shape2.AnimationSettings
  　　　　  Ani.Animate = msoTrue
  　　　　  Ani.TextLevelEffect = ppAnimateByFirstLevel
  　　　　  Ani.AnimateTextInReverse = msoTrue
}
```

### AnimationSettings.AnimationOrder

返回或设置一个整数，该整数代表指定形状在动画形状集合中的位置。可读/写 **Long** 类型。

**语法**

**express.AnimationOrder**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

指定形状必须被赋予动画效果，您才能看到设置该属性的效果。要给一个形状赋予动画效果，必须将 **AnimationSettings** 对象的 **TextLevelEffect** 属性设置为 **ppAnimateLevelNone** 以外的值，并且 **Animate** 属性必须设置为 **True** 。

**示例**

本示例指定活动演示文稿中第二张幻灯片的第二个形状在所有动画中第二个播放。

``` JavaScript
Application.ActivePresentation.Slides.Item(2).Shapes.Item(2).AnimationSettings.AnimationOrder = 2
```

### AnimationSettings.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 AnimationSettings 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres)
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test(){
　　　　for(let i=1;i <= ActivePresentation.Slides.Item(1).Shapes.Count;i++) {
   　　　　 if(ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {
      　　　　  MsgBox(ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)
　　　　    }
　　　　}
}
```

### AnimationSettings.ChartUnitEffect

返回或设置一个值，该值指示图形区域是按序列、类别还是按元素播放动画。可读/写 **PpChartUnitEffect** 类型。

**语法**

**express.ChartUnitEffect**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

如果图形未显示动画效果，请确保 **Animate** 属性已设置为 **True** 。

| PpChartUnitEffect 可以是下列 PpChartUnitEffect 常量之一。 |
|-----------------------------------------------------------|
| ppAnimateByCategory                                       |
| ppAnimateByCategoryElements                               |
| ppAnimateBySeries                                         |
| ppAnimateBySeriesElements                                 |
| ppAnimateChartAllAtOnce                                   |
| ppAnimateChartMixed                                       |

**示例**

``` JavaScript
/*本示例设置活动演示文稿第三张幻灯片的第二个形状按序列动画显示。为使本示例运行，第二个形状必须为一个图表。*/
function test(){
　　　　let myShape = Application.ActivePresentation.Slides.Item(3).Shapes.Item(2)
　　　　let Ani = myShape.AnimationSettings
 　　　　   Ani.ChartUnitEffect = ppAnimateBySeries
　　　　    Ani.EntryEffect = ppEffectFlyFromLeft
 　　　　   Ani.Animate = true
}
```

### AnimationSettings.DimColor

返回或设置一个 **ColorFormat** 对象，该对象代表指定形状生成后的颜色。只读。

**语法**

**express.DimColor**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

如果没有达到预期的效果，请检查其他的创建设置。必须将 **AnimationSettings** 对象的 **TextLevelEffect** 属性设为非 **ppAnimateLevelNone** 的值， **AfterEffect** 属性设为 **ppAfterEffectDim** ，且 **Animate** 属性设为 **True** 才能看到 **DimColor** 属性的效果。另外，如果指定形状是要在幻灯片上创建的唯一项目或最后一个项目，则该形状无法变暗。若要更改幻灯片上形状的创建顺序，请使用 **AnimationOrder** 属性。

**示例**

``` JavaScript
/*本示例在当前演示文稿中添加一张幻灯片。该幻灯片包含一个标题和一个三项目的列表。本示例设置标题和列表在创建后变暗，并设置它们变暗的颜色。*/
function test(){
　　　　let myShape = Application.ActivePresentation.Slides.Add(2, ppLayoutText).Shapes
　　　　let Shape1 = myShape.Item(1)
 　　　　   Shape1.TextFrame.TextRange.Text = "Sample title"
　　　　let Ani1 = Shape1.AnimationSettings
 　　　　   Ani1.TextLevelEffect = ppAnimateByAllLevels
  　　　　  Ani1.AfterEffect = ppAfterEffectDim
  　　　　  Ani1.DimColor.SchemeColor = ppShadow
 　　　　   Ani1.Animate = true
　　　　let Shape2 = myShape.Item(2)
   　　　　 Shape2.TextFrame.TextRange.Text = "Item one" + "\r" + "Item two"
　　　　let Ani2 = Shape2.AnimationSettings
　　　　    Ani2.TextLevelEffect = ppAnimateByFirstLevel
  　　　　  Ani2.AfterEffect = ppAfterEffectDim
　　　　    Ani2.DimColor.RGB = 100, 150, 130
  　　　　  Ani2.Animate = true
}
```

### AnimationSettings.EntryEffect

对于 **AnimationSettings** 对象，此属性返回或设置应用于指定形状中动画的特殊效果。对于 **SlideShowTransition** 对象，此属性返回或设置应用于指定幻灯片切换的特殊效果。 **PpEntryEffect** 类型，可读写。

**语法**

**express.EntryEffect**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

如果将指定形状的 **TextLevelEffect** 属性设为 **ppAnimateLevelNone** （默认值）或将 **Animate** 属性设为 **False** ，您将看不到通过 **EntryEffect** 属性应用的特殊效果。

**示例**

``` JavaScript
/*本示例在当前演示文稿中添加一张标题幻灯片，并设置该标题在幻灯片放映中动画时从右侧飞入。*/
function test(){
　　　　let myShape = Application.ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes.Item(1)
   　　　　 myShape.TextFrame.TextRange.Text = "Sample title"
　　　　let Ani = myShape.AnimationSettings
　　　　    Ani.TextLevelEffect = ppAnimateByAllLevels
 　　　　   Ani.EntryEffect = ppEffectFlyFromRight
    　　　　Ani.Animate = true
}
```

### AnimationSettings.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 AnimationSettings 对象的变量。

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

### AnimationSettings.PlaySettings

返回一个 **PlaySettings** 对象，该对象包含指定的媒体剪辑在幻灯片放映中的播放方式的信息。

**语法**

**express.PlaySettings**

\*express是一个代表 AnimationSettings 对象的变量。

**示例**

``` JavaScript
/*以下示例在当前演示文稿的第一张幻灯片中插入名为 Clock.avi 的影片，设置它在幻灯片切换后自动播放，并指定该影片对象在不播放时隐藏。*/
function test(){
　　　　let ole = Application.ActivePresentation.Slides.Item(1).Shapes.AddOLEObject(10, 10, 250, 250, undefined, "c:\\winnt\\Clock.avi")
　　　　let ps = ole.AnimationSettings.PlaySettings
　　　　ps.PlayOnEntry = true
　　　　ps.HideWhileNotPlaying = true
}
```

### AnimationSettings.SoundEffect

返回一个 **SoundEffect** 对象，该对象代表切换到指定幻灯片时播放的声音。

**语法**

**express.SoundEffect**

\*express是一个代表 AnimationSettings 对象的变量。

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

### AnimationSettings.TextLevelEffect

设置或返回一个 **PpTextLevelEffect** 常量，该常量表示指定形状中的文本是否由第一级段落、第二级段落或其他级段落（最高为第五级段落）激活动画显示。可读/写。

**语法**

**express.TextLevelEffect**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

要使 **TextLevelEffect** 属性生效，必须将 **Animate** 属性设置为 **True** 。

| PpTextLevelEffect 可以是下列 PpTextLevelEffect 常量之一。 |
|-----------------------------------------------------------|
| ppAnimateByAllLevels                                      |
| ppAnimateByFifthLevel                                     |
| ppAnimateByFirstLevel                                     |
| ppAnimateByFourthLevel                                    |
| ppAnimateBySecondLevel                                    |
| ppAnimateByThirdLevel                                     |
| ppAnimateLevelMixed                                       |
| ppAnimateLevelNone                                        |

**示例**

``` JavaScript
/*本示例在当前演示文稿中添加标题幻灯片及标题文本，并将该标题的字母设置为逐个显示。*/
function test(){
　　　　let myShape = Application.ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes.Item(1)
   　　　　 myShape.TextFrame.TextRange.Text = "Sample title"
　　　　let Ani = myShape.AnimationSettings
  　　　　  Ani.Animate = true
   　　　　 Ani.TextLevelEffect = ppAnimateByFirstLevel
  　　　　  Ani.TextUnitEffect = ppAnimateByCharacter
   　　　　 Ani.EntryEffect = ppEffectFlyFromLeft
}
```

### AnimationSettings.TextUnitEffect

返回或设置一个 **PpTextUnitEffect** 常量，该常量指示指定形状中的文本是按段落、按字或是按字母被赋予动画效果。可读/写。

**语法**

**express.TextUnitEffect**

\*express是一个代表 AnimationSettings 对象的变量。

**说明**

| PpTextUnitEffect 可以是下列 PpTextUnitEffect 常量之一。 |
|---------------------------------------------------------|
| ppAnimateByCharacter                                    |
| ppAnimateByParagraph                                    |
| ppAnimateByWord                                         |
| ppAnimateUnitMixed                                      |

为使 **TextUnitEffect** 属性设置生效，必须将指定形状的 **TextLevelEffect** 属性设置为除 **ppAnimateLevelNone** 或 **ppAnimateByAllLevels** 之外的值，并且必须将 **Animate** 属性设置为 **True** 。

**示例**

``` JavaScript
/*本示例在当前演示文稿中添加标题幻灯片及标题文本，并将该标题的字母设置为逐个显示。*/
function test(){
　　　　let myShape = Application.ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes.Item(1)
  　　　　  myShape.TextFrame.TextRange.Text = "Sample title"
　　　　let Ani = myShape.AnimationSettings
   　　　　 Ani.Animate = true
  　　　　  Ani.TextLevelEffect = ppAnimateByFirstLevel
   　　　　 Ani.TextUnitEffect = ppAnimateByCharacter
   　　　　 Ani.EntryEffect = ppEffectFlyFromLeft
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/AnimationSettings](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DocumentWindows 对象

## 简介

WPP中当前所有打开的 DocumentWindow 对象的集合。该集合不包含打开的幻灯片放映窗口，这些窗口包含在 SlideShowWindows 集合中。

## 方法

| 序号 | 名称                                | 说明                     |
|:----:|:------------------------------------|:-------------------------|
|  1   | [Arrange](#DocumentWindows.Arrange) | 指定是层叠还是平铺窗口   |
|  2   | [Item](#DocumentWindows.Item)       | 从指定集合中返回单个对象 |

## 属性

| 序号 | 名称                                        | 说明                                                |
|:----:|:--------------------------------------------|:----------------------------------------------------|
|  1   | [Application](#DocumentWindows.Application) | 返回一个Application对象，该对象表示指定对象的创建者 |
|  2   | [Count](#DocumentWindows.Count)             | 返回指定集合中的对象数目。                          |
|  3   | [Parent](#DocumentWindows.Parent)           | 返回指定对象的父对象                                |

## 成员方法

### DocumentWindows.Arrange

指定是层叠还是平铺窗口

**语法**

**express.Arrange ( *arrangeStyle* )**

\*express是一个代表 DocumentWindows 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型       | 说明                   |
|----------------|-----------|----------------|------------------------|
| *arrangeStyle* | 可选      | PpArrangeStyle | 指定是层叠还是平铺窗口 |

### DocumentWindows.Item

从指定集合中返回单个对象

**语法**

**express.Item ( *Index* )**

\*express是一个代表 DocumentWindows 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                           |
|---------|-----------|----------|--------------------------------|
| *Index* | 必选      | Long     | 集合中要返回的单个对象的索引号 |

**示例**

``` JavaScript
/*以下示例关闭第二个文档窗口*/
Application.Windows.Item(2).Close()
```

## 成员属性

### DocumentWindows.Application

返回一个Application对象，该对象表示指定对象的创建者

**语法**

**express.Application**

\*express是一个代表 DocumentWindows 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个Presentation对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行的WPP的文件夹中。*/
function test(){
     let pptPres = Application.ActivePresentation
     pptPres.Slides.Add(1,1)
     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

### DocumentWindows.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 DocumentWindows 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口*/
function test(){
    let dWindows = Application.Windows

    for（let i = 2; i<= dWindows.Count; i++){

         dWindows.Item(i).Close()

  }
}
```

### DocumentWindows.Parent

返回指定对象的父对象

**语法**

**express.Parent**

\*express是一个代表 DocumentWindows 对象的变量。

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

> WPS 开发文档： [WPS 基础接口/演示 API 参考/DocumentWindows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LinkFormat 对象

## 简介

包含应用于链接 OLE 对象的属性和方法。

## 说明

OLEFormat 对象包含应用于链接和非链接 OLE 对象的属性和方法。 PictureFormat 对象包含应用于图片和 OLE 对象的属性和方法。

## 方法

| 序号 | 名称                         | 说明                                                                               |
|:----:|:-----------------------------|:-----------------------------------------------------------------------------------|
|  1   | [Update](#LinkFormat.Update) | 更新指定的链接 OLE 对象。要一次更新演示文稿中的所有链接，可使用 UpdateLinks 方法。 |

## 属性

| 序号 | 名称                                         | 说明                                                               |
|:----:|:---------------------------------------------|:-------------------------------------------------------------------|
|  1   | [Application](#LinkFormat.Application)       | 返回一个 Application 对象，该对象表示指定对象的创建者。            |
|  2   | [AutoUpdate](#LinkFormat.AutoUpdate)         | 返回或设置更新链接的方法。可读/写 PpUpdateOption 类型。            |
|  3   | [Parent](#LinkFormat.Parent)                 | 返回指定对象的父对象。                                             |
|  4   | [SourceFullName](#LinkFormat.SourceFullName) | 返回或设置链接 OLE 对象的源文件的名称和路径。可读/写 String 类型。 |

## 成员方法

### LinkFormat.Update

更新指定的链接 OLE 对象。要一次更新演示文稿中的所有链接，可使用 **UpdateLinks** 方法。

**语法**

**express.Update ()**

\*express是一个代表 LinkFormat 对象的变量。

**示例**

``` JavaScript
/*本示例更新当前演示文稿中所有链接的 OLE 对象。*/
function test(){
　　　　for(let i = 1; i <= Application.ActivePresentation.Slides.Count; i++){
    　　　　for(let ii = 1; ii <= Application.ActivePresentation.Slides.Item(i).Shapes.Count; ii++){
       　　　　 if(Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).Type == msoLinkedOLEObject){
　　　　            Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).Type.LinkFormat.Update()
　　　　        }
　　　　    }
　　　　}}
```

## 成员属性

### LinkFormat.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 LinkFormat 对象的变量。

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

### LinkFormat.AutoUpdate

返回或设置更新链接的方法。可读/写 **PpUpdateOption** 类型。

**语法**

**express.AutoUpdate**

\*express是一个代表 LinkFormat 对象的变量。

**说明**

| PpUpdateOption 可以是下列 PpUpdateOption 常量之一。              |
|------------------------------------------------------------------|
| ppUpdateOptionAutomatic 每次打开演示文稿或更改源文件时更新链接。 |
| ppUpdateOptionManual 仅当用户特别要求更新演示文稿时更新链接。    |
| ppUpdateOptionMixed                                              |

**示例**

``` JavaScript
/*本示例浏览活动演示文稿中所有幻灯片的所有形状，并将所有链接的 ET 工作表设为手动更新。*/
function test(){
　　　　for(let i = 1; i<= Application.ActivePresentation.Slides.Count; i++){
  　　　　  for(let ii = 1; ii <= Application.ActivePresentation.Slides.Item(i).Shapes.Count; ii++){
　　　　        if(Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).Type == msoLinkedOLEObject){
　　　　            if(Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).OLEFormat.ProgID == "Excel.Sheet"){
 　　　　               Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).LinkFormat.AutoUpdate = ppUpdateOptionManual
　　　　            }
 　　　　       }
 　　　　   }
　　　　}
}
```

### LinkFormat.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 LinkFormat 对象的变量。

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

### LinkFormat.SourceFullName

返回或设置链接 OLE 对象的源文件的名称和路径。可读/写 **String** 类型。

**语法**

**express.SourceFullName**

\*express是一个代表 LinkFormat 对象的变量。

**示例**

``` JavaScript
/*本示例将当前演示文稿第一张幻灯片第一个形状的源文件设为 Wordtest.doc，并指定该对象的图像自动更新。*/
function test(){
　　　　let sha = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
   　　　　 if(sha.Type == msoLinkedOLEObject){
    　　　　    let lin = sha.LinkFormat
    　　　　    lin.SourceFullName = "C:\\my documents\\wordtest.doc"
     　　　　   lin.AutoUpdate = ppUpdateOptionAutomatic
  　　　　 }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/LinkFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Pane 对象

## 简介

一个对象，它代表普通视图中的三个窗格之一，或文档窗口中任意其他视图的某个窗格。

## 说明

使用 Panes ( *index* ) 返回单个 Pane 对象，其中 *index* 是窗格的索引号。下表列出了普通视图中窗格的名称及其相应的索引号。

| 窗格   | 索引号 |
|--------|--------|
| 大纲   | 1      |
| 幻灯片 | 2      |
| 备注   | 3      |

当使用文档窗口视图而非普通视图时，使用 Panes (1) 可引用单个 Pane 对象。

使用 Activate 方法可激活指定窗格。

使用 ViewType 属性可决定活动窗格。

普通视图是唯一具有多窗格的视图。所有其他文档窗口视图都只有一个窗格，即文档窗口。

``` JavaScript
/*以下示例使用 ViewType 属性判断幻灯片窗格是否为活动窗格。如果幻灯片窗格是活动窗格，则使用 Activate 方法激活备注窗格。*/
function test() {
    let window = Application.ActiveWindow
    if(window.ActivePane.ViewType == ppViewSlide){
        window.Panes.Item(3).Activate()
    }
}
```

## 方法

| 序号 | 名称                       | 说明             |
|:----:|:---------------------------|:-----------------|
|  1   | [Activate](#Pane.Activate) | 激活指定的对象。 |

## 属性

| 序号 | 名称                             | 说明                                                                                                                                 |
|:----:|:---------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Active](#Pane.Active)           | 返回指定的窗格或窗口是否处于激活的状态。只读。 MsoTriState 类型。                                                                    |
|  2   | [Application](#Pane.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                              |
|  3   | [Parent](#Pane.Parent)           | 返回指定对象的父对象。                                                                                                               |
|  4   | [ViewType](#Pane.ViewType)       | 在带有窗格的任一视图中，返回指定窗格的视图类型。用于不带窗格的视图时，则返回父 DocumentWindow 对象的视图类型。只读 PpViewType 类型。 |

## 成员方法

### Pane.Activate

激活指定的对象。

**语法**

**express.Activate ()**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
/*本示例按顺序激活紧挨在活动窗口之后的文档窗口。*/
Application.Windows.Item(2).Activate()
```

## 成员属性

### Pane.Active

返回指定的窗格或窗口是否处于激活的状态。只读。 **MsoTriState** 类型。

**语法**

**express.Active**

\*express是一个代表 Pane 对象的变量。

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
/*以下示例检查演示文稿文件 "test.ppt" 是否位于当前窗口。如果不是，则将当前处于激活状态的演示文稿的名称保存在变量 oldWin 中，并激活 "test.ppt" 演示文稿。*/
function test() {
  let mypane = Application.Presentations.Item("C:\\Users\\Administrator\\Desktop\\text.ppt").Windows.Item(1)
  if(mypane.Active == false){
      let oldWin = Application.ActiveWindow
      mypane.Activate()
  }
}
```

### Pane.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Pane 对象的变量。

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

### Pane.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Pane 对象的变量。

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

### Pane.ViewType

在带有窗格的任一视图中，返回指定窗格的视图类型。用于不带窗格的视图时，则返回父 **DocumentWindow** 对象的视图类型。只读 **PpViewType** 类型。

**语法**

**express.ViewType**

\*express是一个代表 Pane 对象的变量。

**说明**

|                                             |
|---------------------------------------------|
| PpViewType 可以是下列 PpViewType 常量之一。 |
| **ppViewHandoutMaster**                     |
| **ppViewMasterThumbnails**                  |
| **ppViewNormal**                            |
| **ppViewNotesMaster**                       |
| **ppViewNotesPage**                         |
| **ppViewOutline**                           |
| **ppViewPrintPreview**                      |
| **ppViewSlide**                             |
| **ppViewSlideMaster**                       |
| **ppViewSlideSorter**                       |
| **ppViewThumbnails**                        |
| **ppViewTitleMaster**                       |

**示例**

``` JavaScript
/*如果当前窗格中的视图是幻灯片视图，则本示例使备注窗格成为当前窗格。备注窗格是 Panes 集合中的第三个成员。*/
function test() {
  let mypane = Application.ActiveWindow
  if(mypane.ActivePane.ViewType == ppViewSlide){
     mypane.Panes.Item(3).Activate()
  }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Pane](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Placeholders 对象

## 简介

代表指定幻灯片上占位符的所有 Shape 对象的集合。

## 说明

Placeholders 集合中的每个 Shape 对象代表一个占位符，占位符可以是文本、图表、表格、组织结构图或其他类型的对象。如果幻灯片有标题，则标题是集合中的第一个占位符。

使用 Delete 方法可以删除各个占位符，而且可以使用 AddPlaceholder 方法恢复删除的占位符，但添加占位符时不能使占位符数多于幻灯片创建时所具有的数目。要更改给定幻灯片上的占位符数目，请设置 Layout 属性。

## 方法

| 序号 | 名称                                   | 说明                                                                        |
|:----:|:---------------------------------------|:----------------------------------------------------------------------------|
|  1   | [FindByName](#Placeholders.FindByName) | 可在 Placeholders 集合中指定的索引位置或以指定的名称查找占位符。。 可读写。 |
|  2   | [Item](#Placeholders.Item)             | 从指定集合中返回单个对象。                                                  |

## 属性

| 序号 | 名称                                     | 说明                                                    |
|:----:|:-----------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#Placeholders.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#Placeholders.Count)             | 返回指定集合中的对象数目。                              |
|  3   | [Parent](#Placeholders.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### Placeholders.FindByName

可在 Placeholders 集合中指定的索引位置或以指定的名称查找占位符。。 可读写。

**语法**

**express.FindByName ( *Index* )**

\*express是一个代表 Placeholders 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                   |
|---------|-----------|----------|------------------------|
| *Index* | 必选      | Variant  | 要查找的占位符的索引。 |

**返回值**

Shape

### Placeholders.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Placeholders 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                             |
|---------|-----------|----------|----------------------------------|
| *Index* | 必选      | Long     | 集合中要返回的单个对象的索引号。 |

**返回值**

Shape

## 成员属性

### Placeholders.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Placeholders 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres) {
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

### Placeholders.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 Placeholders 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口。*/
function test(){
　　　　let mydocument = Application.Windows  
       for(let i = 2;i <= mydocument.Count;i++){  
   　　　　 mudocument.Item(i).Close()
　　　　}
}
```

### Placeholders.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Placeholders 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let mydocument = myShapes.AddShape(msoShapeOval,50,50,300,150).TextFrame
  　　　　 mydocument.TextRange.Text = "Test text"
  　　　　 mydocument.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Placeholders](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ShapeNode 对象

## 简介

代表用户定义的任意多边形中节点的几何图形以及几何图形编辑属性。

## 说明

节点包括任意多边形的线段之间的顶点以及曲线段的控制点。 ShapeNode 对象是 ShapeNodes 集合的成员。 ShapeNodes 集合包含任意多边形中的所有节点。

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                               |
|:----:|:--------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ShapeNode.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                            |
|  2   | [Creator](#ShapeNode.Creator)         | 返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。                      |
|  3   | [EditingType](#ShapeNode.EditingType) | 如果指定的节点是顶点，则此属性的返回值表示对节点所做的更改将如何影响与该节点相连的两条线段。如果该节点是曲线段的控制点，则此属性返回相邻顶点的编辑类型。只读 MsoEditingType 类型。 |
|  4   | [Parent](#ShapeNode.Parent)           | 返回指定对象的父对象。                                                                                                                                                             |
|  5   | [Points](#ShapeNode.Points)           | 返回一个 Variant 类型的值，该值以坐标对形式表示指定节点的位置。每个坐标以磅为单位表示。可以使用 SetPosition 方法设置该属性的值。只读。                                             |
|  6   | [SegmentType](#ShapeNode.SegmentType) | 返回一个值，该值代表与指定节点相关联的线段是直线段还是曲线段。只读 MsoSegmentType 类型。                                                                                           |

## 成员属性

### ShapeNode.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 ShapeNode 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres){
    pptPres.Slides.Add (1, 1)
    pptPres.SaveAs (pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test(){
　　　　let shpOle = ActivePresentation.Slides.Item(1).Shapes
       for(let i = 1; i <= shpOle.Count; i++){
    　　　　if(shpOle.Item(i).Type == msoLinkedOLEObject){
       　　　　 MsgBox (shpOle.Item(i).OLEFormat.Application.Name)
　　　　    }
　　　　}
}
```

### ShapeNode.Creator

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 ShapeNode 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

``` JavaScript
/*以下示例显示关于 myObject 创建者的消息。*/
function test(){
　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
　　　　if(myObject.Creator == "&h50575054"){
　　　　    MsgBox ("This is aWPP object")
　　　　}
　　　　else{
　　　　    MsgBox ("This is not aWPP object")
　　　　}
}
```

### ShapeNode.EditingType

如果指定的节点是顶点，则此属性的返回值表示对节点所做的更改将如何影响与该节点相连的两条线段。如果该节点是曲线段的控制点，则此属性返回相邻顶点的编辑类型。只读 **MsoEditingType** 类型。

**语法**

**express.EditingType**

\*express是一个代表 ShapeNode 对象的变量。

**说明**

此属性为只读。可用 **SetEditingType** 方法设置此属性的值。

| MsoEditingType 可以是下列 MsoEditingType 常量之一。 |
|-----------------------------------------------------|
| msoEditingAuto                                      |
| msoEditingCorner                                    |
| msoEditingSmooth                                    |
| msoEditingSymmetric                                 |

**示例**

``` JavaScript
/*本示例将 myDocument 中第三个形状中的所有角结点更改为光滑结点。第三个形状必须是任意多边形。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　let mynodes = myDocument.Shapes.Item(3).Nodes
    for(let n = 1; n <= mynodes.Count; n++){
   　　　　 if(mynodes.Item(n).EditingType == msoEditingCorner){
  　　　　      mynodes.Item(n).SetEditingType (n, msoEditingSmooth)
　　　　    }
　　　　}
}
```

### ShapeNode.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ShapeNode 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes
　　　　let myshape = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
　　　　myshape.TextRange.Text = "Test text"
　　　　myshape.Parent.Rotation = 45
}
```

### ShapeNode.Points

返回一个 **Variant** 类型的值，该值以坐标对形式表示指定节点的位置。每个坐标以磅为单位表示。可以使用 **SetPosition** 方法设置该属性的值。只读。

**语法**

**express.Points**

\*express是一个代表 ShapeNode 对象的变量。

**示例**

``` JavaScript
/*本示例将当前演示文稿中第三个形状的第二个结点向右移 200 磅，并向下移 300 磅。第三个形状必须是任意多边形。*/
function test(){
　　　　let mynodes = ActivePresentation.Slides.Item(1).Shapes.Item(3).Nodes
　　　　let pointsArray = mynodes.Item(2).Points
　　　　let currXvalue = pointsArray(1, 1)
　　　　let currYvalue = pointsArray(1, 2)
　　　　mynodes.SetPosition (2, currXvalue + 200, currYvalue + 300)
}
```

### ShapeNode.SegmentType

返回一个值，该值代表与指定节点相关联的线段是直线段还是曲线段。只读 **MsoSegmentType** 类型。

**语法**

**express.SegmentType**

\*express是一个代表 ShapeNode 对象的变量。

**说明**

此属性为只读。可使用 **SetSegmentType** 方法设置此属性的值。

| MsoSegmentType 可以是下列 MsoSegmentType 常量之一。                         |
|-----------------------------------------------------------------------------|
| msoSegmentCurve 如果指定节点是曲线段的控制点，则 SegmentType 属性返回该值。 |
| msoSegmentLine                                                              |

**示例**

``` JavaScript
/*本示例将 myDocument 中第三个形状的所有直线段更改为曲线段。第三个形状必须是任意多边形。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　let mynodes = myDocument.Shapes.Item(3).Nodes
　　　　let n = 1
　　　　while (n <= mynodes.Count){
   　　　　 if(mynodes.Item(n).SegmentType == msoSegmentLine){
　　　　        mynodes.SetSegmentType (n, msoSegmentCurve)
　　　　    }
　　　　    n = n + 1
　　　　}
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/ShapeNode](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Shapes 对象

## 简介

指定的幻灯片上所有 Shape 对象的集合。

## 说明

每个 Shape 对象代表绘图层中的一个对象，例如自选图形、任意多边形、OLE 对象或图片。

| 注释                                                                                                                                                                                                        |
|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 若要使用文档中的形状的子集（例如，只对文档中的自选图形或选定的形状进行操作），则必须构造一个包含要使用的形状的 ShapeRange 集合。有关如何使用一个形状或同时使用多个形状的概述，请参阅 使用形状（图形对象）。 |

## 示例

使用 Shapes 属性返回 Shapes 集合。以下示例选择当前演示文稿中的所有形状。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Shapes.SelectAll()
```

| 注释                                                                                                                                                                                                                   |
|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 若要同时对文档中的所有形状进行某种操作（例如删除或设置一个属性），可不带参数使用 [Range](#Shapes.Range) 方法创建一个 ShapeRange 对象，该对象包含 Shapes 集合中的所有形状，然后对 ShapeRange 对象应用相应的属性或方法。 |

使用 [AddCallout](#Shapes.AddCallout) 、 [AddComment](#Shapes.AddComment) 、 [AddConnector](#Shapes.AddConnector) 、 [AddCurve](#Shapes.AddCurve) 、 [AddLabel](#Shapes.AddLabel) 、 [AddLine](#Shapes.AddLine) 、 [AddMediaObject](#Shapes.AddMediaObject) 、 [AddOLEObject](#Shapes.AddOLEObject) 、 [AddPicture](#Shapes.AddPicture) 、 [AddPlaceholder](#Shapes.AddPlaceholder) 、 [AddPolyline](#Shapes.AddPolyline) 、 [AddShape](#Shapes.AddShape) 、 [AddTable](#Shapes.AddTable) 、 [AddTextbox](#Shapes.AddTextbox) 、 [AddTextEffect](#Shapes.AddTextEffect) 或 [AddTitle](#Shapes.AddTitle) 方法创建一个新形状并将其添加到 Shapes 集合中。联合使用 BuildFreeform 方法和 ConvertToShape 方法创建一个新任意多边形并将其添加到集合中。以下示例在活动演示文稿中添加一个矩形。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)
```

使用 Shapes ( *index* ) 返回单个 Shape 对象，其中 *index* 是形状的名称或索引号。

使用 Shapes.Range ( *index* ) 返回一个 ShapeRange 集合，该集合代表 Shapes 集合的一个子集，其中 *index* 是形状的名称或索引号，或由形状的名称或索引号组成的数组。以下示例设置活动演示文稿中第一个和第三个形状的填充图案。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Shapes.Range([1, 3]).Fill .Patterned(msoPatternHorizontalBrick)
```

使用 Shapes.Placeholders ( *index* ) 返回一个代表占位符的 Shape 对象，其中 *index* 是占位符编号。如果指定的幻灯片有标题，则使用 Shapes.Placeholders(1) 或 Shapes.Title 返回标题占位符。以下示例在活动演示文稿中添加一张幻灯片并为标题和副标题添加文本（副标题是此版式的幻灯片中的第二个占位符）。

``` JavaScript
function test()
{
    let myShape = Application.ActivePresentation.Slides.Add(1, ppLayoutTitle).Shapes
    myShape.Title.TextFrame.TextRange.Text = "This is the title text"
    myShape.Placeholders.Item(2).TextFrame.TextRange.Text = "This is subtitle text" 
} 
```

## 方法

| 序号 | 名称                                     | 说明                                                                                                                                                                                                                                                                                                               |
|:----:|:-----------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddChart](#Shapes.AddChart)             | 向演示文稿中添加图表。                                                                                                                                                                                                                                                                                             |
|  2   | [AddConnector](#Shapes.AddConnector)     | 创建一个连接符。返回一个代表新连接符的 Shape 对象。添加连接符时，它没有连接任何对象。使用 BeginConnect 和 EndConnect 方法可将连接符的头和尾连接到文档中的其他形状上。                                                                                                                                              |
|  3   | [AddLine](#Shapes.AddLine)               | 创建线条。返回一个代表新线条的 Shape 对象。                                                                                                                                                                                                                                                                        |
|  4   | [AddMediaObject](#Shapes.AddMediaObject) | 创建多媒体对象。返回一个代表新多媒体对象的 Shape 对象。                                                                                                                                                                                                                                                            |
|  5   | [AddOLEObject](#Shapes.AddOLEObject)     | 创建 OLE 对象。返回一个代表新 OLE 对象的 Shape 对象。                                                                                                                                                                                                                                                              |
|  6   | [AddPicture](#Shapes.AddPicture)         | 从现有文件创建图片。返回一个代表新图片的 Shape 对象。                                                                                                                                                                                                                                                              |
|  7   | [AddPlaceholder](#Shapes.AddPlaceholder) | 恢复幻灯片上以前已删除的占位符。返回一个 Shape 代表已恢复占位符的对象。                                                                                                                                                                                                                                            |
|  8   | [AddShape](#Shapes.AddShape)             | 创建一个自选形状。返回一个代表新自选形状的 Shape 对象。                                                                                                                                                                                                                                                            |
|  9   | [AddTable](#Shapes.AddTable)             | 将表格形状添加到幻灯片中。                                                                                                                                                                                                                                                                                         |
|  10  | [AddTextEffect](#Shapes.AddTextEffect)   | 创建艺术字对象。返回一个代表新艺术字对象的 Shape 对象。                                                                                                                                                                                                                                                            |
|  11  | [AddTextbox](#Shapes.AddTextbox)         | 创建一个文本框。返回一个代表新文本框的 Shape 对象。                                                                                                                                                                                                                                                                |
|  12  | [Item](#Shapes.Item)                     | 从指定集合中返回单个对象。                                                                                                                                                                                                                                                                                         |
|  13  | [Paste](#Shapes.Paste)                   | 在 z 顺序的最上端将剪贴板上的形状、幻灯片或文本粘贴到指定的 Shapes 集合中。粘贴的每个对象都会成为指定的 Shapes 集合的成员。如果剪贴板包含全部幻灯片，则这些幻灯片将作为包含幻灯片图像的形状粘贴。如果剪贴板包含文本范围，则该文本将被粘贴到一个新创建的 TextFrame 形状中。返回一个代表粘贴对象的 ShapeRange 对象。 |
|  14  | [Range](#Shapes.Range)                   | 返回一个代表 Shapes 集合中的形状子集的 ShapeRange 对象。                                                                                                                                                                                                                                                           |
|  15  | [SelectAll](#Shapes.SelectAll)           | 选定所有形状（ Shapes 集合中）或所有图表节点（ DiagramNodes 或 DiagramNodeChildren 集合中）。                                                                                                                                                                                                                      |

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                   |
|:----:|:-------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Shapes.Application)   | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                |
|  2   | [Count](#Shapes.Count)               | 返回指定集合中的对象数目。                                                                                                             |
|  3   | [HasTitle](#Shapes.HasTitle)         | 返回指定幻灯片上的对象集合是否包含标题占位符。只读 MsoTriState 类型。                                                                  |
|  4   | [Placeholders](#Shapes.Placeholders) | 返回一个 Placeholders 集合，该集合代表幻灯片上所有占位符的集合。集合中的每个占位符可包含文本、图表、表格、组织机构图或其他对象。只读。 |
|  5   | [Title](#Shapes.Title)               | 返回一个代表幻灯片标题的 Shape 对象。只读。                                                                                            |

## 成员方法

### Shapes.AddChart

向演示文稿中添加图表。

**语法**

**express.AddChart ( *Type* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明         |
|----------|-----------|----------|--------------|
| *Type*   | 可选      | Variant  | 图表的类型。 |
| *Left*   | 可选      | Number   | 左侧尺寸。   |
| *Top*    | 可选      | Number   | 顶端尺寸。   |
| *Width*  | 可选      | Number   | 图表的宽度。 |
| *Height* | 可选      | Number   | 图表的高度。 |

### Shapes.AddConnector

创建一个连接符。返回一个代表新连接符的 **Shape** 对象。添加连接符时，它没有连接任何对象。使用 **BeginConnect** 和 **EndConnect** 方法可将连接符的头和尾连接到文档中的其他形状上。

**语法**

**express.AddConnector ( *Type* , *BeginX* , *BeginY* , *EndX* , *EndY* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                     |
|----------|-----------|----------|----------------------------------------------------------|
| *Type*   | 必选      | Number   | 连接符的类型。                                           |
| *BeginX* | 必选      | Number   | 连接符的起点相对于幻灯片左边缘的水平位置（以磅为单位）。 |
| *BeginY* | 必选      | Number   | 连接符的起点相对于幻灯片上边缘的垂直位置（以磅为单位）。 |
| *EndX*   | 必选      | Number   | 连接符的终点相对于幻灯片左边缘的水平位置（以磅为单位）。 |
| *EndY*   | 必选      | Number   | 连接符的终点相对于幻灯片上边缘的垂直位置（以磅为单位）。 |

**说明**

将一个连接符连接到某个形状时，如果必要，该连接符的长度和位置会自动调整。因此，如果要将一个连接符连接到其他形状，则与添加该连接符时指定的位置和长度无关。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加两个矩形，然后用曲线连接符将它们连接起来。*/
/*请注意，将连接符连接到矩形上时，连接符的长度和位置会自动调整；因此，它与添加标注时指定的位置和长度是无关的（长度不能为零）。*/
function test() 
{
    let shpShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let shpFirst = shpShapes.AddShape(msoShapeRectangle, 100, 50, 200, 100)
    let shpSecond = shpShapes.AddShape(msoShapeRectangle, 300, 300, 200, 100)
    let myFormat = shpShapes.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat
    myFormat.BeginConnect(shpFirst, 1)
    myFormat.EndConnect(shpSecond, 1)
    myFormat.Parent.RerouteConnections()
}
```

### Shapes.AddLine

创建线条。返回一个代表新线条的 **Shape** 对象。

**语法**

**express.AddLine ( *BeginX* , *BeginY* , *EndX* , *EndY* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                 |
|----------|-----------|----------|------------------------------------------------------|
| *BeginX* | 必选      | Number   | 线条起点相对于幻灯片左边缘的水平位置（以磅为单位）。 |
| *BeginY* | 必选      | Number   | 线条起点相对于幻灯片上边缘的垂直位置（以磅为单位）。 |
| *EndX*   | 必选      | Number   | 线条终点相对于幻灯片左边缘的水平位置（以磅为单位）。 |
| *EndY*   | 必选      | Number   | 线条终点相对于幻灯片上边缘的垂直位置（以磅为单位）。 |

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加一条蓝色虚线。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myLine = myDocument.Shapes.AddLine(10, 10, 250, 250).Line
    myLine.DashStyle = msoLineDashDotDot
    myLine.ForeColor.RGB = (50, 0, 128)
}
```

### Shapes.AddMediaObject

创建多媒体对象。返回一个代表新多媒体对象的 Shape 对象。

**语法**

**express.AddMediaObject ( *FileName* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                              |
|------------|-----------|----------|-----------------------------------------------------------------------------------|
| *FileName* | 必选      | String   | 必选。String 类型。从中创建媒体对象的文件。如果未指定路径，则使用当前工作文件夹。 |
| *Left*     | 可选      | Number   | 媒体对象边界框左上角相对于文档左上角的位置（以磅为单位）。                        |
| *Top*      | 可选      | Number   | 媒体对象边界框左上角相对于文档左上角的位置（以磅为单位）。                        |
| *Width*    | 可选      | Number   | 媒体对象边界框的宽度（以磅为单位）。                                              |
| *Height*   | 可选      | Number   | 媒体对象边界框的高度（以磅为单位）。                                              |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加名为“Clock.avi”的影片。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.AddMediaObject("C:\\WINNT\\clock.avi", 5, 5, 100, 100)
}
```

### Shapes.AddOLEObject

创建 OLE 对象。返回一个代表新 OLE 对象的 **Shape** 对象。

**语法**

**express.AddOLEObject ( *Left* , *Top* , *Width* , *Height* , *ClassName* , *FileName* , *DisplayAsIcon* , *IconFileName* , *IconIndex* , *IconLabel* , *Link* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型    | 说明                                                                                                                                                                         |
|-----------------|-----------|-------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Left*          | 可选      | Number      | 新对象左上角相对于幻灯片左上角的位置（以磅为单位）。默认值为 0（零）。                                                                                                       |
| *Top*           | 可选      | Number      | 新对象左上角相对于幻灯片左上角的位置（以磅为单位）。默认值为 0（零）。                                                                                                       |
| *Width*         | 可选      | Number      | OLE 对象的初始宽度（以磅为单位）。                                                                                                                                           |
| *Height*        | 可选      | Number      | OLE 对象的初始高度（以磅为单位）。                                                                                                                                           |
| *ClassName*     | 可选      | String      | OLE 长类名或待创建对象的 ProgID。必须为该对象指定 ClassName 或 FileName 参数，但不能同时指定两者。                                                                           |
| *FileName*      | 可选      | String      | 要创建的对象的源文件。如果未指定路径，则使用当前工作文件夹。必须为该对象指定 ClassName 或 FileName 参数，但不能同时指定两者。                                                |
| *DisplayAsIcon* | 可选      | MsoTriState | 决定是否将 OLE 对象显示为图标。                                                                                                                                              |
| *IconFileName*  | 可选      | String      | 包含将要显示的图标的文件。                                                                                                                                                   |
| *IconIndex*     | 可选      | Number      | IconFileName 中的图标索引。文件中第一个图标的索引号是 0（零）。如果 IconFileName 中不存在已知索引号的图标，则使用索引号为 1 的图标（文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选      | String      | 显示在图标下面的标签（标题）。                                                                                                                                               |
| *Link*          | 可选      | MsoTriState | 决定是否将 OLE 对象链接到从中创建该对象的文件。如果已指定 ClassName 的值，则此参数必须是 msoFalse。                                                                          |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加命令按钮。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.AddOLEObject(100, 100, 150, 50, "Forms.CommandButton.1")
}
```

### Shapes.AddPicture

从现有文件创建图片。返回一个代表新图片的 **Shape** 对象。

**语法**

**express.AddPicture ( *FileName* , *LinkToFile* , *SaveWithDocument* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型    | 说明                                                                                                  |
|--------------------|-----------|-------------|-------------------------------------------------------------------------------------------------------|
| *FileName*         | 必选      | String      | 要在其中创建 OLE 对象的文件                                                                           |
| *LinkToFile*       | 必选      | MsoTriState | 确定是否将图片链接到从中创建该图片的文件。                                                            |
| *SaveWithDocument* | 必选      | MsoTriState | 确定是否将已链接的图片与其插入到的文档一起保存。如果 LinkToFile 为 msoFalse，则此参数必须为 msoTrue。 |
| *Left*             | 必选      | Number      | 图片左边缘相对于幻灯片左边缘的位置（以磅为单位）。                                                    |
| *Top*              | 必选      | Number      | 图片上边缘相对于幻灯片上边缘的位置（以磅为单位）。                                                    |
| *Width*            | 可选      | Number      | 图片的宽度（以磅为单位）。                                                                            |
| *Height*           | 可选      | Number      | 图片的高度（以磅为单位）。                                                                            |

### Shapes.AddPlaceholder

恢复幻灯片上以前已删除的占位符。返回一个 **Shape** 代表已恢复占位符的对象。

**语法**

**express.AddPlaceholder ( *Type* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型          | 说明                                                                                                                                                                                                                                                                                                                                                                            |
|----------|-----------|-------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Type*   | 必选      | PpPlaceholderType | 占位符的类型。只有在版式类型为 ppLayoutVerticalText、ppLayoutClipArtAndVerticalText、ppLayoutVerticalTitleAndText 或 ppLayoutVerticalTitleAndTextOverChart 的幻灯片中才能找到类型为 ppPlaceholderVerticalBody 或 ppPlaceholderVerticalTitle 的占位符。不能在用户界面中使用上述任何版式创建幻灯片；必须通过使用 Add 方法或通过设置现有幻灯片的 Layout 属性，以编程方式创建它们。 |
| *Left*   | 可选      | Number            | 占位符左上角相对于文档左上角的位置（以磅为单位）。                                                                                                                                                                                                                                                                                                                              |
| *Top*    | 可选      | Number            | 占位符左上角相对于文档左上角的位置（以磅为单位）。                                                                                                                                                                                                                                                                                                                              |
| *Width*  | 可选      | Number            | 占位符的宽度（以磅为单位）。                                                                                                                                                                                                                                                                                                                                                    |
| *Height* | 可选      | Number            | 占位符的高度（以磅为单位）。                                                                                                                                                                                                                                                                                                                                                    |

**说明**

如果从幻灯片中删除了多个指定类型的占位符， **AddPlaceholder** 方法将它们依次放回到幻灯片中，从原来索引号最小的占位符开始。

### Shapes.AddShape

创建一个自选形状。返回一个代表新自选形状的 **Shape** 对象。

**语法**

**express.AddShape ( *Type* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型         | 说明                                                   |
|----------|-----------|------------------|--------------------------------------------------------|
| *Type*   | 必选      | MsoAutoShapeType | 指定要创建的自选形状的类型。                           |
| *Left*   | 必选      | Number           | 自选形状左边缘相对于幻灯片左边缘的位置（以磅为单位）。 |
| *Top*    | 必选      | Number           | 自选形状上边缘相对于幻灯片上边缘的位置（以磅为单位）。 |
| *Width*  | 必选      | Number           | 自选形状的宽度（以磅为单位）。                         |
| *Height* | 必选      | Number           | 自选形状的高度（以磅为单位）。                         |

**说明**

若要更改所添加的自选形状的类型，可设置其 **AutoShapeType** 属性。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加矩形。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)
}
```

### Shapes.AddTable

将表格形状添加到幻灯片中。

**语法**

**express.AddTable ( *NumRows* , *NumColumns* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                 |
|--------------|-----------|----------|------------------------------------------------------|
| *NumRows*    | 必选      | Number   | 表格的行数。                                         |
| *NumColumns* | 必选      | Number   | 表格的列数。                                         |
| *Left*       | 可选      | Number   | 从幻灯片左边缘到表格左边缘之间的距离（以磅为单位）。 |
| *Top*        | 可选      | Number   | 从幻灯片上边缘到表格上边缘之间的距离（以磅为单位）。 |
| *Width*      | 可选      | Number   | 新表格的宽度（以磅为单位）。                         |
| *Height*     | 可选      | Number   | 新表格的高度（以磅为单位）。                         |

**示例**

``` JavaScript
/*本示例在当前演示文稿的第二张幻灯片上创建一张新表格。*/
/*该表格具有三个行，四个列。它到幻灯片左、上边缘的距离均为 10 磅。*/
/*该新表格的宽度为 288 磅，这样，表格中四个列的宽度都是一英寸（一英寸等于 72 磅）。表格的高度设为 216 磅，这样，表格中三个行的高度都是一英寸。*/

Application.ActivePresentation.Slides.Item(2).Shapes .AddTable(3, 4, 10, 10, 288, 216)
```

### Shapes.AddTextEffect

创建艺术字对象。返回一个代表新艺术字对象的 **Shape** 对象。

**语法**

**express.AddTextEffect ( *PresetTextEffect* , *Text* , *FontName* , *FontSize* , *FontBold* , *FontItalic* , *Left* , *Top* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型            | 说明                                                       |
|--------------------|-----------|---------------------|------------------------------------------------------------|
| *PresetTextEffect* | 必选      | MsoPresetTextEffect | 预置文字效果。                                             |
| *Text*             | 必选      | String              | 艺术字中的文字。                                           |
| *FontName*         | 必选      | String              | 艺术字中所用字体的名称。                                   |
| *FontSize*         | 必选      | Number              | 艺术字中所用字体的大小（以磅为单位）。                     |
| *FontBold*         | 必选      | MsoTriState         | 确定艺术字中使用的字体是否设为粗体。                       |
| *FontItalic*       | 必选      | MsoTriState         | 确定艺术字中使用的字体是否设为斜体。                       |
| *Left*             | 必选      | Number              | 艺术字边界框左边缘相对于幻灯片左边缘的位置（以磅为单位）。 |
| *Top*              | 必选      | Number              | 艺术字边界框上边缘相对于幻灯片上边缘的位置（以磅为单位）。 |

**说明**

向文档添加艺术字对象时，该对象的高度和宽度将自动根据所指定的文字的大小和数量来设置。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加包含文本“Test”的艺术字。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let newWordArt = myDocument.Shapes.AddTextEffect(msoTextEffect1, "Test", "Arial Black", 36, msoFalse, msoFalse, 10, 10)
}
```

### Shapes.AddTextbox

创建一个文本框。返回一个代表新文本框的 **Shape** 对象。

**语法**

**express.AddTextbox ( *Orientation* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型           | 说明                                                                                   |
|---------------|-----------|--------------------|----------------------------------------------------------------------------------------|
| *Orientation* | 必选      | MsoTextOrientation | 文本方向。这些常量中的某些可能不可用，这取决于选择或安装的语言支持（例如，中国中文）。 |
| *Left*        | 必选      | Nubmer             | 文本框左边缘相对于幻灯片左边缘的位置（以磅为单位）。                                   |
| *Top*         | 必选      | Nubmer             | 文本框上边缘相对于幻灯片上边缘的位置（以磅为单位）。                                   |
| *Width*       | 必选      | Nubmer             | 文本框的宽度（以磅为单位）。                                                           |
| *Height*      | 必选      | Nubmer             | 文本框的高度（以磅为单位）。                                                           |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加包含文本“Test Box”的文本框。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 50).TextFrame.TextRange.Text = "Test Box"
}
```

### Shapes.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

**示例**

``` JavaScript
/*本示例将活动演示文稿第一张幻灯片中名为“Rectangle 1”的形状的前景色设置为红色。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item("rectangle 1").Fill.ForeColor.RGB = (128, 0, 0)
```

### Shapes.Paste

在 z 顺序的最上端将剪贴板上的形状、幻灯片或文本粘贴到指定的 **Shapes** 集合中。粘贴的每个对象都会成为指定的 **Shapes** 集合的成员。如果剪贴板包含全部幻灯片，则这些幻灯片将作为包含幻灯片图像的形状粘贴。如果剪贴板包含文本范围，则该文本将被粘贴到一个新创建的 **TextFrame** 形状中。返回一个代表粘贴对象的 **ShapeRange** 对象。

**语法**

**express.Paste ()**

\*express是一个代表 Shapes 对象的变量。

**说明**

将剪贴板内容粘贴到窗口中之前，请使用 **ViewType** 属性设置窗口视图。下表显示了可以粘贴到每个视图中的内容。

| 视图                   | 可以从剪贴板中粘贴以下内容                                                                                                                                                                                                                                                                                                |
|------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 幻灯片视图或备注页视图 | 形状、文本或整张幻灯片。如果从剪贴板粘贴一个幻灯片，该幻灯片的图像将作为嵌入对象被插入到幻灯片、母版或备注页中；如果选中形状，粘贴的文本将附加到形状文本之后；如果选中文本，粘贴的文本将替换选中的文本；如果选中任何其他对象，粘贴的文本将被放到它自己的文本框中。粘贴的形状将被放到 z 顺序的最上端且不会替换选中的形状。 |
| 大纲视图               | 文本或整张幻灯片。不能向大纲视图粘贴形状。粘贴的幻灯片将被插到光标所在的幻灯片之前。                                                                                                                                                                                                                                      |
| 幻灯片浏览视图         | 整张幻灯片。不能向幻灯片浏览视图粘贴形状或文本。粘贴的幻灯片将被插到光标处或演示文稿中最后选中的一张幻灯片之后。                                                                                                                                                                                                          |

**示例**

``` JavaScript
/*本示例将活动演示文稿第一张幻灯片的第一个形状复制到剪贴板中，然后将其粘贴到第二张幻灯片中。*/
function test() 
{
    let myPre = Application.ActivePresentation
    myPre.Slides.Item(1).Shapes.Item(1).Copy()
    myPre.Slides.Item(2).Shapes.Paste()
}
```

### Shapes.Range

返回一个代表 **Shapes** 集合中的形状子集的 **ShapeRange** 对象。

**语法**

**express.Range ( *Indexx* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                                                                       |
|----------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Indexx* | 可选      | Variant  | 要包含在范围中的各个形状。可以是指定形状的索引号的 Integer、指定形状名称的 String，或者是包含整数或字符串的数组。如果省略此参数，则 Range 方法将返回指定集合中的所有对象。 |

**说明**

虽然可以使用 **Range** 方法返回任意数目的形状或幻灯片，但如果只需要返回该集合的一个成员，则使用 **Item** 方法会更为简单。例如，Shapes(1) 比 Shapes.Range(1) 简单，Slides(2) 比 Slides.Range(2) 简单。

若要为 **Index** 指定一个整数或字符串数组，可以使用 **Array** 函数。

**示例**

``` JavaScript
/*本示例设置 myDocument 中第一个形状和第三个形状的填充图案。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick)
}

/*本示例设置第一张幻灯片中名为 Oval 4 和 Rectangle 5 的形状的填充图案。*/
function test() 
{
    let myArray = ["Oval 4", "Rectangle 5"]
    let myRange = Application.ActivePresentation.Slides.Item(1).Shapes.Range(myArray)
    myRange.Fill.Patterned(msoPatternHorizontalBrick)
}

/*本示例设置第一张幻灯片中所有形状的填充图案。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Range().Fill.Patterned(msoPatternHorizontalBrick)

/*本示例设置第一张幻灯片中第一个形状的填充图案。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myRange = myDocument.Shapes.Range(1)
    myRange.Fill.Patterned(msoPatternHorizontalBrick)
}
```

### Shapes.SelectAll

选定所有形状（ **Shapes** 集合中）或所有图表节点（ **DiagramNodes** 或 **DiagramNodeChildren** 集合中）。

**语法**

**express.SelectAll ()**

\*express是一个代表 Shapes 对象的变量。

**示例**

``` JavaScript
/*本示例选择 myDocument 中的所有形状。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.SelectAll()
}
```

## 成员属性

### Shapes.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Shapes 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres) 
{
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test()
{
    let myShape = Application.ActivePresentation.Slides.Item(1).Shapes
    for(let shpOle = 1; shpOle <= myShape.Count; shpOle++) 
    {
        if(myShape.Item(shpOle).Type == msoLinkedOLEObject)
        {
            alert(myShape.Item(shpOle).OLEFormat.Application.Name)
        }
    }
}
```

### Shapes.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 Shapes 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口。*/
function test()
{
    let myWindow = Application.Windows
    for(let i = 2; i <= myWindow.Count; i++) 
    {
        myWindow.Item(2).Close()
    }
}
```

### Shapes.HasTitle

返回指定幻灯片上的对象集合是否包含标题占位符。只读 **MsoTriState** 类型。

**语法**

**express.HasTitle**

\*express是一个代表 Shapes 对象的变量。

**说明**

为ture，表示 定的幻灯片上的对象集合包含标题占位符。

### Shapes.Placeholders

返回一个 **Placeholders** 集合，该集合代表幻灯片上所有占位符的集合。集合中的每个占位符可包含文本、图表、表格、组织机构图或其他对象。只读。

**语法**

**express.Placeholders**

\*express是一个代表 Shapes 对象的变量。

**示例**

``` JavaScript
/*本示例在当前演示文稿中添加一张幻灯片，然后为该幻灯片标题（幻灯片的第一个占位符）和副标题添加文本。*/ 
function test()
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let mySlide =  Application.ActivePresentation.Slides .Add(1, ppLayoutTitle).Shapes.Placeholders
    mySlide.Item(1).TextFrame.TextRange.Text = "This is the title text"
    mySlide.Item(2).TextFrame.TextRange.Text = "This is subtitle text"
}
```

### Shapes.Title

返回一个代表幻灯片标题的 **Shape** 对象。只读。

**语法**

**express.Title**

\*express是一个代表 Shapes 对象的变量。

**说明**

也可以使用 Shapes 或 **Placeholders** 集合的 **[Item](#Shapes.Item)** 方法返回幻灯片标题。

**示例**

``` JavaScript
/*本示例设置 myDocument 上的标题文本。*/
function test()
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.Title.TextFrame.TextRange.Text = "Welcome!"
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Shapes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Slide 对象

## 简介

代表一个幻灯片。 Slides 集合包含演示文稿中的所有 Slide 对象。

## 说明

| 注释                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 如果试图返回对单张幻灯片的引用却得到了一个 SlideRange 对象，请不要奇怪。单张幻灯片既可以由 Slide 对象表示也可以由只包含一张幻灯片的 SlideRange 集合表示，这取决于返回该幻灯片引用的方式。例如，如果使用 A dd 方法创建并返回对幻灯片的引用，则该幻灯片由 Slide 对象表示。然而，如果使用 Duplicate 方法创建并返回对幻灯片的引用，则该幻灯片由包含单张幻灯片的 SlideRange 集合表示。因为应用于 Slide 对象的所有属性和方法也可应用于包含单张幻灯片的 SlideRange 集合，所以可对返回的幻灯片进行相同的操作，而不管它是由 Slide 对象还是 SlideRange 集合表示。 |

以下示例说明如何执行下列操作：

- 返回一个通过名称、索引号或幻灯片 ID 号指定的幻灯片
- 返回选定范围中的幻灯片
- 返回指定的任意文档窗口或幻灯片放映窗口中当前显示的幻灯片
- 新建幻灯片

## 示例

使用 Slides ( *index* )（其中 *index* 是幻灯片名称或索引号）或者使用 Slides.FindBySlideID ( *index* )（其中 *index* 是幻灯片 ID 号）返回单个 Slide 对象。以下示例设置活动演示文稿中第一张幻灯片的版式。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Layout = ppLayoutTitle
```

使用 Slides ( *index* )（其中 *index* 是幻灯片名称或索引号）或使用 Slides.FindBySlideID ( *index* )（其中 *index* 是幻灯片 ID 号）返回单个 Slide 对象。以下示例设置活动演示文稿中第一张幻灯片的版式。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Layout = ppLayoutTitle
```

以下示例设置 ID 号为 265 的幻灯片的版式。

``` JavaScript
Application.ActivePresentation.Slides.FindBySlideID(265).Layout = ppLayoutTitle
```

使用 Selection.SlideRange ( *index* ) 返回单个 Slide 对象，其中 *index* 是选定范围中的幻灯片的名称或索引号。以下示例设置活动窗口选定范围中的第一张幻灯片的版式（假设至少选定一张幻灯片）。

``` JavaScript
Application.ActiveWindow.Selection.SlideRange.Item(1).Layout = ppLayoutTitle 
```

如果只选定了一张幻灯片，可以使用 Selection.SlideRange 返回包含选定幻灯片的 SlideRange 集合。以下示例设置当前窗口当前所选对象中第一张幻灯片的版式（假设正好只选定一张幻灯片）

``` JavaScript
Application.ActiveWindow.Selection.SlideRange.Layout = ppLayoutTitle 
```

使用 Slide 属性返回指定文档窗口或幻灯片放映窗口视图中当前显示的幻灯片。以下示例将第二个文档窗口中当前显示的幻灯片复制到剪贴板。

``` JavaScript
Application.Windows.Item(2).View.Slide.Copy()
```

使用 Add 方法新建幻灯片并添加到演示文稿中。以下示例在当前演示文稿的开头添加一个标题幻灯片。

``` JavaScript
Application.ActivePresentation.Slides.Add(1, ppLayoutTitleOnly)
```

## 方法

| 序号 | 名称                                                  | 说明                                                                             |
|:----:|:------------------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [ApplyTemplate](#Slide.ApplyTemplate)                 | 将设计模板应用于指定的演示文稿。                                                 |
|  2   | [ApplyTheme](#Slide.ApplyTheme)                       | 对指定幻灯片应用主题或设计模板。                                                 |
|  3   | [ApplyThemeColorScheme](#Slide.ApplyThemeColorScheme) | 对指定幻灯片应用配色方案。                                                       |
|  4   | [Copy](#Slide.Copy)                                   | 将指定对象复制到剪贴板。                                                         |
|  5   | [Cut](#Slide.Cut)                                     | 删除指定对象并将其放到剪贴板中。                                                 |
|  6   | [Delete](#Slide.Delete)                               | 删除指定的对象。                                                                 |
|  7   | [Export](#Slide.Export)                               | 使用指定的图形筛选器导出幻灯片或幻灯片范围，并将导出的文件以指定的文件名保存。   |
|  8   | [MoveTo](#Slide.MoveTo)                               | 将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。 |
|  9   | [Select](#Slide.Select)                               | 选择指定的对象。                                                                 |

## 属性

| 序号 | 名称                                                | 说明                                                                                                                                                                                                                                                                                                                                                                             |
|:----:|:----------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Slide.Application)                   | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                                                                                                                          |
|  2   | [Background](#Slide.Background)                     | 返回代表幻灯片背景的 ShapeRange 对象。                                                                                                                                                                                                                                                                                                                                           |
|  3   | [BackgroundStyle](#Slide.BackgroundStyle)           | 设置或返回指定对象的背景样式。可读/写。                                                                                                                                                                                                                                                                                                                                          |
|  4   | [ColorScheme](#Slide.ColorScheme)                   | 返回或设置 ColorScheme 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。                                                                                                                                                                                                                                                                                   |
|  5   | [Comments](#Slide.Comments)                         | 返回一个 Comments 对象，该对象代表批注的集合。                                                                                                                                                                                                                                                                                                                                   |
|  6   | [Design](#Slide.Design)                             | 返回代表设计的 Design 对象。                                                                                                                                                                                                                                                                                                                                                     |
|  7   | [HeadersFooters](#Slide.HeadersFooters)             | 返回一个 HeadersFooters 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。                                                                                                                                                                                                                                                     |
|  8   | [HyperLinks](#Slide.HyperLinks)                     | 返回一个代表指定幻灯片中的所有超链接的 HyperLinks 集合。只读。                                                                                                                                                                                                                                                                                                                   |
|  9   | [Layout](#Slide.Layout)                             | 返回或设置一个 PpSlideLayout 常量，该常量代表幻灯片版式。可读/写。                                                                                                                                                                                                                                                                                                               |
|  10  | [Master](#Slide.Master)                             | 返回一个 Master 对象，该对象代表幻灯片母版。只读。                                                                                                                                                                                                                                                                                                                               |
|  11  | [Name](#Slide.Name)                                 | 向一个演示文稿插入幻灯片时，Powerpoint 自动为幻灯片分配 Sliden 形式的名称，其中 n 是一个整数，代表幻灯片在演示文稿中创建的顺序。例如，在某个演示文稿中插入的第一张幻灯片被自动命名为 Slide1。如果将幻灯片从一个演示文稿复制到另一个演示文稿，则该幻灯片会失去它在第一个演示文稿中的名称并在第二个演示文稿中被自动分配新名称。幻灯片范围必须至少包含一张幻灯片。 String 可读/写。 |
|  12  | [NotesPage](#Slide.NotesPage)                       | 返回一个 SlideRange 对象，该对象代表指定幻灯片或幻灯片区域的备注页。只读。                                                                                                                                                                                                                                                                                                       |
|  13  | [Shapes](#Slide.Shapes)                             | 返回一个 Shapes 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。                                                                                                                               |
|  14  | [SlideID](#Slide.SlideID)                           | 返回指定幻灯片的唯一 ID 号。只读。 Number 类型。                                                                                                                                                                                                                                                                                                                                 |
|  15  | [SlideIndex](#Slide.SlideIndex)                     | 返回 Slides 集合内指定幻灯片的索引号。只读。 Number 类型。                                                                                                                                                                                                                                                                                                                       |
|  16  | [SlideNumber](#Slide.SlideNumber)                   | 返回幻灯片编号。只读 Number 类型。                                                                                                                                                                                                                                                                                                                                               |
|  17  | [SlidesShowTransition](#Slide.SlidesShowTransition) | 返回一个 SlidesShowTransition 对象，该对象代表指定幻灯片切换的特殊效果。只读。                                                                                                                                                                                                                                                                                                   |
|  18  | [TimeLin](#Slide.TimeLin)                           | 返回 TimeLine 对象，该对象代表幻灯片的动画时间线。                                                                                                                                                                                                                                                                                                                               |

## 成员方法

### Slide.ApplyTemplate

将设计模板应用于指定的演示文稿。

**语法**

**express.ApplyTemplate ( *FileName* )**

\*express是一个代表 Slide 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                 |
|------------|-----------|----------|----------------------|
| *FileName* | 必选      | String   | 指定设计模板的名称。 |

**说明**

如果在字符串中引用未安装的演示文稿设计模板，则将产生运行时错误。无论 **Application.FeatureInstall** 属性设置是什么，该模板都不自动安装。要对当前未安装的模板使用 **ApplyTemplate** 方法，必须先安装附加设计模板。为此，请通过运行 WPS Office安装程序（单击 Windows 控制面板中的 **“添加/删除程序”** 图标）安装WPP 的附加设计模板。

**示例**

``` JavaScript
/*本示例将“Professional”设计模板应用于当前演示文稿。*/
Application.ActivePresentation.ApplyTemplate("c:\\program files\\office\\templates" + "\\presentationdesigns\\professional.pot")
```

### Slide.ApplyTheme

对指定幻灯片应用主题或设计模板。

**语法**

**express.ApplyTheme ( *themeName* )**

\*express是一个代表 Slide 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                     |
|-------------|-----------|----------|--------------------------------------------------------------------------|
| *themeName* | 必选      | String   | 应用于 Slide 对象的主题文件 (.thmx) 或设计模板文件 (.pot) 的路径和名称。 |

**示例**

``` JavaScript
/*以下示例对指定幻灯片应用保存的主题：*/
Application.ActivePresentation.Slides.Item(1).ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.thmx")

/*以下示例对指定幻灯片应用保存的设计模板：*/
Application.ActivePresentation.Slides.Item(1).ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.pot")
```

### Slide.ApplyThemeColorScheme

对指定幻灯片应用配色方案。

**语法**

**express.ApplyThemeColorScheme ( *themeColorSchemeName* )**

\*express是一个代表 Slide 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型 | 说明                                     |
|------------------------|-----------|----------|------------------------------------------|
| *themeColorSchemeName* | 必选      | String   | 应用于幻灯片的配色方案文件的路径和名称。 |

### Slide.Copy

将指定对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 Slide 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

``` JavaScript
/*本示例将当前演示文稿的第一张幻灯片复制到剪贴板。*/
Application.ActivePresentation.Slides.Item(1).Copy()

/*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()
```

### Slide.Cut

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut ()**

\*express是一个代表 Slide 对象的变量。

**示例**

``` JavaScript
/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中。*/
Application.Windows.Item(1).Selection.Cut()

/*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中。*/
Application.ActivePresentation.Slides.Item(1).Cut()

/*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut()
```

### Slide.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 Slide 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

### Slide.Export

使用指定的图形筛选器导出幻灯片或幻灯片范围，并将导出的文件以指定的文件名保存。

**语法**

**express.Export ( *FileName* , *FilterName* , *ScaleWidth* , *ScaleHeight* )**

\*express是一个代表 Slide 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                        |
|---------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName*    | 必选      | String   | 将导出并保存到磁盘的文件的名称。可以包括完整路径；如果不包括完整路径，WPP 就会在当前文件夹中创建一个文件。                                                                                                                  |
| *FilterName*  | 必选      | Single   | 要导出幻灯片的图形格式。指定的图形格式必须在 Windows 注册表中注册导出筛选器。可以指定注册的扩展名或注册的筛选器名称。WPP将首先在注册表中搜索匹配的扩展名。如果没有找到与指定字符串匹配的扩展名，WPP将查找匹配的筛选器名称。 |
| *ScaleWidth*  | 可选      | Number   | 导出幻灯片的宽度（以像素为单位）。                                                                                                                                                                                          |
| *ScaleHeight* | 可选      | Number   | 导出幻灯片的高度（以像素为单位）。                                                                                                                                                                                          |

**说明**

导出一个演示文稿不会将演示文稿的 **Saved** 属性设置为 **True** 。

WPP使用指定的图形筛选器保存每张单独的幻灯片。导出并保存到磁盘的幻灯片的名称由WPP 决定。通常情况下，这些文件保存为 Slide1.wmf、Slide2.wmf 等。保存文件的路径在 **FileName** 参数中指定。

**示例**

``` JavaScript
/*本示例将活动演示文稿中第三张幻灯片以 JPEG 图形格式导出到磁盘上。该幻灯片将保存为 Slide 3 of Annual Sales.jpg。*/
function test()
{
    let mySlide = Application.ActivePresentation.Slides.Item(3)
    mySlide.Export("c:\\my documents\\Graphic Format\\" + "Slide 3 of Annual Sales", "JPG")
}
```

### Slide.MoveTo

将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。

**语法**

**express.MoveTo ( *toPos* )**

\*express是一个代表 Slide 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                     |
|---------|-----------|----------|--------------------------|
| *toPos* | 必选      | Number   | 动画效果要移动到的索引。 |

**示例**

``` JavaScript
/*本示例将当前演示文稿中的第二张幻灯片移动到第一张幻灯片的位置上。*/
function test() 
{
    Application.ActivePresentation.Slides.Item(2).MoveTo(1)
}
```

### Slide.Select

选择指定的对象。

**语法**

**express.Select ()**

\*express是一个代表 Slide 对象的变量。

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片

**示例**

``` JavaScript
/*本示例选择活动演示文稿中第一张幻灯片标题的前五个字符。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.Characters(1, 5).Select()

/*本示例选择活动演示文稿中的第一张幻灯片。*/
Application.ActivePresentation.Slides.Item(1).Select()

/*本示例选取一个已添加到演示文稿的新幻灯片中的表格。该表格具有三行和三列。*/
function test()
{
    let mySlide = Application.Presentations.Add().Slides
    mySlide.Add(1, ppLayoutBlank).Shapes.AddTable(3, 3).Select()
}
```

## 成员属性

### Slide.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Slide 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(){
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(s.Application.Path + "\\Added Slide")
}
/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test()
{
    for(let shpOle = 1; shpOle <= Application.ActivePresentation.Slides.Item(1).Shapes.Count; shpOle++) 
    {
        if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(shpOle).Type == msoLinkedOLEObject) 
        {
            alert(Application.ActivePresentation.Slides.Item(1).Shapes.Item(shpOle).OLEFormat.Application.Name)
        }
    }
}
```

### Slide.Background

返回代表幻灯片背景的 **ShapeRange** 对象。

**语法**

**express.Background**

\*express是一个代表 Slide 对象的变量。

**说明**

如果要使用 **Background** 属性设置单张幻灯片的背景而不更改幻灯片母版，则该幻灯片的 **FollowMasterBackground** 属性必须设置为 **False** 。

### Slide.BackgroundStyle

设置或返回指定对象的背景样式。可读/写。

**语法**

**express.BackgroundStyle**

\*express是一个代表 Slide 对象的变量。

### Slide.ColorScheme

返回或设置 **ColorScheme** 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。

**语法**

**express.ColorScheme**

\*express是一个代表 Slide 对象的变量。

### Slide.Comments

返回一个 **Comments** 对象，该对象代表批注的集合。

**语法**

**express.Comments**

\*express是一个代表 Slide 对象的变量。

### Slide.Design

返回代表设计的 **Design** 对象。

**语法**

**express.Design**

\*express是一个代表 Slide 对象的变量。

### Slide.HeadersFooters

返回一个 **HeadersFooters** 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。

**语法**

**express.HeadersFooters**

\*express是一个代表 Slide 对象的变量。

**示例**

``` JavaScript
/*以下示例在当前演示文稿的备注母版中设置页脚文本、日期和时间的格式，并将日期和时间设定为可自动修改。*/
function test()
{
    let notesmaster = Application.ActivePresentation.NotesMaster.HeadersFooters
    notesmaster.Footer.Text = "Regional Sales"
    let newnotesmaster = notesmaster2.DateAndTime
    newnotesmaster.UseFormat = true
    newnotesmaster.Format = ppDateTimeHmmss
}
```

### Slide.HyperLinks

返回一个代表指定幻灯片中的所有超链接的 **HyperLinks** 集合。只读。

**语法**

**express.HyperLinks**

\*express是一个代表 Slide 对象的变量。

**示例**

``` JavaScript
/*以下示例允许用户更新当前演示文稿中所有超链接中已过时的 Internet 地址。*/
function test()
{
    let oldAddr = "Old Internet address"
    let newAddr = "New Internet address"
    for(let s = 1; s <= Application.ActivePresentation.Slides.Count; s++) 
    {
        for(let h = 1; h <= Application.ActivePresentation.Slides.Item(s).Hyperlinks.Count; h++) 
        {
            if(Application.ActivePresentation.Slides.Item(s).Hyperlinks.Item(h).Address.toLowerCase() == oldAddr) 
            {
                Application.ActivePresentation.Slides.Item(s).Hyperlinks.Item(h).Address = newAddr
            }
        }
    }
}
```

### Slide.Layout

返回或设置一个 **PpSlideLayout** 常量，该常量代表幻灯片版式。可读/写。

**语法**

**express.Layout**

\*express是一个代表 Slide 对象的变量。

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

**示例**

``` JavaScript
/*本示例中，如果当前演示文稿的第一张幻灯片原来只有一个标题，则改变该幻灯片的版式以包含一个标题和一个副标题。*/
function test()
{
    let slides = Application.ActivePresentation.Slides.Item(1)
    if(slides.Layout == ppLayoutTitleOnly) 
    {
        slides.Layout = ppLayoutTitle
    }
}
```

### Slide.Master

返回一个 **Master** 对象，该对象代表幻灯片母版。只读。

**语法**

**express.Master**

\*express是一个代表 Slide 对象的变量。

### Slide.Name

向一个演示文稿插入幻灯片时，Powerpoint 自动为幻灯片分配 Sliden 形式的名称，其中 n 是一个整数，代表幻灯片在演示文稿中创建的顺序。例如，在某个演示文稿中插入的第一张幻灯片被自动命名为 Slide1。如果将幻灯片从一个演示文稿复制到另一个演示文稿，则该幻灯片会失去它在第一个演示文稿中的名称并在第二个演示文稿中被自动分配新名称。幻灯片范围必须至少包含一张幻灯片。 **String** 可读/写。

**语法**

**express.Name**

\*express是一个代表 Slide 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。

### Slide.NotesPage

返回一个 **SlideRange** 对象，该对象代表指定幻灯片或幻灯片区域的备注页。只读。

**语法**

**express.NotesPage**

\*express是一个代表 Slide 对象的变量。

**说明**

对 **SlideRange** 对象代表的备注页使用下列属性和方法无效： **Copy** 方法、 **Cut** 方法、 **Delete** 方法、 **Duplicate** 方法、 **HeadersFooters** 属性、 **Hyperlinks** 属性、 **Layout** 属性、 **PrintSteps** 属性、 **SlideShowTransition** 属性。 **NotesPage** 属性返回一张或一组幻灯片的备注页，且只允许对这些备注页进行修改。如果要进行会影响所有备注页的更改，请使用 **NotesMaster** 返回代表备注母版的 **Slide** 对象。

### Slide.Shapes

返回一个 Shapes 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。

**语法**

**express.Shapes**

\*express是一个代表 Slide 对象的变量。

**示例**

``` JavaScript
/*以下示例添加一个宽 100 磅、高 50 磅的矩形，它的左上角距当前演示文稿第一张幻灯片的左边 5 磅、上边 25 磅。*/
function test()
{
    let firstSlide = Application.ActivePresentation.Slides.Item(1)
    firstSlide.Shapes.AddShape(msoShapeRectangle, 5, 25, 100, 50)
}

/*以下示例设置当前演示文稿第一张幻灯片的第三个形状的填充纹理。*/
function test()
{
    let newRect = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
    newRect.Fill.PresetTextured(msoTextureOak)
}

/*假设当前演示文稿第一张幻灯片包含一个标题，以下示例的第二行和第三行设置该演示文稿第一张幻灯片的标题文本。*/
function test()
{
    let firstSl = Application.ActivePresentation.Slides.Item(1)
    firstSl.Shapes.Title.TextFrame.TextRange.Text = "Some title text"
    firstSl.Shapes.Item(1).TextFrame.TextRange.Text = "Other title text"
}

/*假设当前演示文稿第二张幻灯片中的第二个形状包含文本框架，以下示例向该幻灯片添加一系列的段落。请注意：Chr(13) 用于在该文本中插入段落标记。*
function test()
{
    let tShape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(2)
    tShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item"
}

/*对于大多数幻灯片版式，第一个形状为文本占位符。以下示例与上例完成相同功能。*/
function test()
{
    let testShape = Application.ActivePresentation.Slides.Item(2).Shapes.Placeholders.Item(2)
    testShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item"    
}
```

### Slide.SlideID

返回指定幻灯片的唯一 ID 号。只读。 **Number** 类型。

**语法**

**express.SlideID**

\*express是一个代表 Slide 对象的变量。

**说明**

与 SlideIndex 属性不同，在演示文稿中添加或重新排列幻灯片时， Slide 对象的 SlideID 属性不会改变。因此，与使用具有幻灯片索引号的 Item 方法相比，使用具有幻灯片 ID 号的 FindBySlideID 方法是从 Slides 集合返回特定 Slide 对象的更可靠方法。

**示例**

``` JavaScript
/*以下示例示范如何检索 Slide 对象的唯一 ID 号，并使用此编号从 Slide 集合返回该 Slides 对象。*/
function test()
{
    let gslides = Application.ActivePresentation.Slides

    //Get slide ID
    let graphSlideID = gslides.Add(2, ppLayoutChart).SlideID
    gslides.FindBySlideID(graphSlideID).SlideShowTransition.EntryEffect = ppEffectCoverLeft
    //Use ID to return specific slide
}
```

### Slide.SlideIndex

返回 **Slides** 集合内指定幻灯片的索引号。只读。 **Number** 类型。

**语法**

**express.SlideIndex**

\*express是一个代表 Slide 对象的变量。

**说明**

与 SlideID 属性不同，在演示文稿中添加或重新排列幻灯片时， Slide 对象的 SlideIndex 属性会改变。因此，与使用具有幻灯片索引号的 Item 方法相比，使用具有幻灯片 ID 号的 **FindBySlideID** 方法是从 Slides 集合返回特定 Slide 对象的更可靠方法。

**示例**

``` JavaScript
/*以下示例显示第一个幻灯片放映窗口中当前放映幻灯片的索引号。*/
alert(Application.SlideShowWindows.Item(1).View.Slide.SlideIndex)
```

### Slide.SlideNumber

返回幻灯片编号。只读 **Number** 类型。

**语法**

**express.SlideNumber**

\*express是一个代表 Slide 对象的变量。

**说明**

Slide 对象的 SlideNumber 属性是显示幻灯片编号时在幻灯片右下角出现的实际编号。此编号由幻灯片在演示文稿中的编号（ SlideIndex 属性值）以及演示文稿的起始幻灯片编号（ FirstSlideNumber 属性值）决定。幻灯片编号始终等于起始幻灯片编号 + 幻灯片索引号 - 1。

**示例**

``` JavaScript
/*本示例显示第一张幻灯片编号的变化对某特定幻灯片编号的影响。*/
function test()
{
    let slides2 = Application.ActivePresentation
    slides2.PageSetup.FirstSlideNumber = 1        //starts slide numbering at 1
    alert(slides2.Slides.Item(2).SlideNumber)    //returns 2

    slides2.PageSetup.FirstSlideNumber = 10       //starts slide numbering at 10
    alert(slides2.Slides.Item(2).SlideNumber)    //returns 11
}
```

### Slide.SlidesShowTransition

返回一个 **SlidesShowTransition** 对象，该对象代表指定幻灯片切换的特殊效果。只读。

**语法**

**express.SlidesShowTransition**

\*express是一个代表 Slide 对象的变量。

**示例**

``` JavaScript
/*以下示例设置在幻灯片放映过程中，当前演示文稿的第二张幻灯片在 5 秒后自动换片，并在幻灯片切换时播放犬吠声。*/
function test()
{
    let slides2 = Application.ActivePresentation.Slides.Item(2).SlideShowTransition
    slides2.AdvanceOnTime = true
    slides2.AdvanceTime = 5
    slides2.SoundEffect.ImportFromFile("c:\\windows\\media\\dogbark.wav")

    Application.ActivePresentation.SlideShowSettings.AdvanceMode = ppSlideShowUseSlideTimings
}
```

### Slide.TimeLin

返回 **TimeLine** 对象，该对象代表幻灯片的动画时间线。

**语法**

**express.TimeLin**

\*express是一个代表 Slide 对象的变量。

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Slide](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# COMAddIns 对象

## 简介

COMAddIn 对象的集合，这些对象提供有关在 Windows 注册表中注册的 COM 加载项的信息。

## 说明

使用 Application 对象的 COMAddIns 属性可返回 WPS Office宿主应用程序的 COMAddIns 集合。如下面的示例所示，此集合包含了某个给定的 Office 宿主应用程序中所有可用的 COM 加载项，而 COMAddins 集合的 Count 属性可返回可用的 COM 加载项的数目。

``` JavaScript
alert(Application.COMAddIns.Count)
```

如下面的示例所示，使用 COMAddins 集合的 Update 方法可刷新 Windows 注册表中 COM 加载项的列表。

``` JavaScript
Application.COMAddIns.Update()
```

可以使用 COMAddIns.Item(index) ，其中 *index* 既可以是一个能够返回 COMAddIns 集合中相应位置的加载项的序数值，也可以是代表指定 COM 加载项的 ProgID 的 String 值。下面的示例在消息框中显示一个 COM 加载项的说明文本和 ProgID (" msodraa9.ShapeSelect ")。

``` JavaScript
alert(Application.COMAddIns.Item("msodraa9.ShapeSelect").Description)
```

## 方法

| 序号 | 名称                        | 说明                                                               |
|:----:|:----------------------------|:-------------------------------------------------------------------|
|  1   | [Item](#COMAddIns.Item)     | 获取指定的 COMAddIns 集合的成员。                                  |
|  2   | [Update](#COMAddIns.Update) | 根据保存在 Windows 注册表中的加载项列表更新 COMAddIns 集合的内容。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                             |
|:----:|:--------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#COMAddIns.Application) | 获取一个 Application 对象，代表 COMAddIns 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Count](#COMAddIns.Count)             | 获取宿主应用程序中 COM 加载项的数量。只读。                                                                                      |
|  3   | [Creator](#COMAddIns.Creator)         | 获取一个 32 位整数，指示创建 COMAddIns 对象时所使用的应用程序。只读。                                                            |
|  4   | [Parent](#COMAddIns.Parent)           | 获取 COMAddIns 对象的 Parent 对象。只读。                                                                                        |

## 成员方法

### COMAddIns.Item

获取指定的 **COMAddIns** 集合的成员。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 COMAddIns 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                   |
|---------|-----------|----------|------------------------|
| *Index* | 必选      | Long     | 代表集合中成员的位置。 |

### COMAddIns.Update

根据保存在 Windows 注册表中的加载项列表更新 COMAddIns 集合的内容。

**语法**

**express.Update ()**

\*express是一个代表 COMAddIns 对象的变量。

**说明**

在 WPS Office应用程序中使用特定 COM 加载项之前，必须先用相应的组件类别 ID 在 Windows 注册表中将该加载项注册为 COM 组件。通常，COM 加载项的安装程序会向注册表中添加必要的项。

**示例**

``` JavaScript
//下面的示例将根据保存在 Windows 注册表中的加载项列表更新 COMAddIns 集合的内容。
Application.COMAddIns.Update()
```

## 成员属性

### COMAddIns.Application

获取一个 **Application** 对象，代表 **COMAddIns** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 COMAddIns 对象的变量。

### COMAddIns.Count

获取宿主应用程序中 COM 加载项的数量。只读。

**语法**

**express.Count**

\*express是一个代表 COMAddIns 对象的变量。

### COMAddIns.Creator

获取一个 32 位整数，指示创建 **COMAddIns** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 COMAddIns 对象的变量。

### COMAddIns.Parent

获取 **COMAddIns** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 COMAddIns 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/COMAddIns](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# EffectParameter 对象

## 简介

描述单个图片效果参数。

## 说明

图片效果被处理为由各个项构成的链，以便创建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。效果参数指定这些效果的属性。

示例 下面的代码将设置 WPP 幻灯片中某个形状的几个图片效果填充属性。

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

| 序号 | 名称                                        | 说明                                                                            |
|:----:|:--------------------------------------------|:--------------------------------------------------------------------------------|
|  1   | [Application](#EffectParameter.Application) | 获取一个 Application 对象，该对象代表 EffectParameter 对象的容器应用程序。只读  |
|  2   | [Creator](#EffectParameter.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 EffectParameter 对象的应用程序。只读 |
|  3   | [Name](#EffectParameter.Name)               | 检索 EffectParameter 参数的字符串名称。只读                                     |
|  4   | [Value](#EffectParameter.Value)             | 检索或设置 EffectParameter 对象的值。可读/写                                    |

## 成员属性

### EffectParameter.Application

获取一个 **Application** 对象，该对象代表 **EffectParameter** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 EffectParameter 对象的变量。

### EffectParameter.Creator

获取一个 32 位整数，该整数指示在其中创建了 **EffectParameter** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 EffectParameter 对象的变量。

### EffectParameter.Name

检索 **EffectParameter** 参数的字符串名称。只读

**语法**

**express.Name**

\*express是一个代表 EffectParameter 对象的变量。

### EffectParameter.Value

检索或设置 **EffectParameter** 对象的值。可读/写

**语法**

**express.Value**

\*express是一个代表 EffectParameter 对象的变量。

**示例**

``` JavaScript
//下面的代码将 PictureEffect 对象的第一个参数设置为色温。
picEffect.EffectParameters.Item(1).Value = Application.Enum.msoEffectColorTemperature
```

> WPS 开发文档： [WPS 基础接口/通用 API 参考/EffectParameter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SignatureSet 对象

## 简介

对应于附加到文档的数字签名的 Signature 对象的集合。

## 说明

使用 Document 对象的 Signatures 属性可返回 SignatureSet 集合，例如：

``` JavaScript
let sigs = Application.ActiveDocument.Signatures
```

可以使用 Add 方法向 SignatureSet 集合中添加一个 Signature 对象，并使用 Item 方法返回现有的成员。 AddSignatureLine 方法也可向该集合中添加一个 Signature 对象。另请参见 Subset 属性，它可充当一个筛选器，用于决定该集合中是否出现某些 Signature 对象。要从 SignatureSet 集合中删除一个 Signature 对象，请使用 Signature 对象的 Delete 方法。

``` JavaScript
/*下面的示例提示用户选择要用于签署 WPS 中活动文档的数字签名。要使用此示例，请在 WPS 中打开一个文档，并向此函数传递与“数字证书”对话框中数字证书的“颁发者”和“颁发给”字段相匹配的证书颁发者名称和证书签署者名称。此示例将进行测试，以确保在将新签名提交到磁盘之前用户选择的数字签名满足特定条件，例如没有过期。*/
function test(strIssuer, strSigner) {
    //Display the dialog box that lets the
    //user select a digital signature.
    //If the user selects a signature, then
    //it is added to the Signatures
    //collection. If the user doesn't, then
    //an error is returned.
    var sig = Application.ActiveWorkbook.Signatures.Add();

    //Test several properties before committing the Signature object to disk.
    if(sig.Issuer == strIssuer && sig.IsCertificateExpired == false && sig.IsCertificateRevoked == false && sig.IsValid == true) {
      alert("Signed");
      AddSignature = true;
    }
    //Otherwise, remove the Signature object from the SignatureSet collection.
    else {
        sig.Delete();
        alert("Not signed");
        AddSignature = false ;
    }
}
```

## 方法

| 序号 | 名称                                                           | 说明                                                                  |
|:----:|:---------------------------------------------------------------|:----------------------------------------------------------------------|
|  1   | [AddNonVisibleSignature](#SignatureSet.AddNonVisibleSignature) | 在以数字方式签署文档时创建一个签名数据包。                            |
|  2   | [AddSignatureLine](#SignatureSet.AddSignatureLine)             | 向在其中收集签名的文档中添加行。                                      |
|  3   | [Item](#SignatureSet.Item)                                     | 获取一个与文档当前所签署的数字签名之一相对应的 Signature 对象。只读。 |

## 属性

| 序号 | 名称                                                     | 说明                                                                                                                                |
|:----:|:---------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SignatureSet.Application)                 | 获取一个 Application 对象，代表 SignatureSet 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [CanAddSignatureLine](#SignatureSet.CanAddSignatureLine) | 获取一个 Boolean 类型的值，指示是否可向文档中添加签名行。只读。                                                                     |
|  3   | [Count](#SignatureSet.Count)                             | 获取一个 Long 类型的值，指示 SignatureSet 对象中的项数。只读。                                                                      |
|  4   | [Creator](#SignatureSet.Creator)                         | 获取一个 32 位整数，指示创建 SignatureSet 对象时所使用的应用程序。只读。                                                            |
|  5   | [Parent](#SignatureSet.Parent)                           | 获取 SignatureSet 对象的 Parent 对象。只读。                                                                                        |
|  6   | [ShowSignaturesPane](#SignatureSet.ShowSignaturesPane)   | 获取或设置一个 Boolean 类型的值，指示是否应显示 “签名” 任务窗格。可读/写。                                                          |
|  7   | [Subset](#SignatureSet.Subset)                           | 获取或设置一个可充当文档可用 Signature 对象上的筛选器的值。可读/写。                                                                |

## 成员方法

### SignatureSet.AddNonVisibleSignature

在以数字方式签署文档时创建一个签名数据包。

**语法**

**express.AddNonVisibleSignature ( *varSigProv* )**

\*express是一个代表 SignatureSet 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                    |
|--------------|-----------|----------|-------------------------|
| *varSigProv* | 可选      | Variant  | 代表签名提供程序的 ID。 |

**返回值**

Signature

**说明**

要提供用于触发此方法的入口点，您需要创建包含签名提供程序加载项的用户界面。此入口点通常以菜单项的形式提供给用户。

**示例**

``` JavaScript
/*在以数字方式签署文档时，以下函数使用签名提供程序 ID 参数来创建签名数据包。*/
function test(varSigProviderID) {
    var objSignatureSet = Application.ActiveWorkbook.Signatures;
    var objSignature = Application.ActiveWorkbook.Signatures.Item(1);
        objSignature = objSignatureSet.AddNonVisibleSignature(varSigProviderID);
    return objSignature;
}
```

### SignatureSet.AddSignatureLine

向在其中收集签名的文档中添加行。

**语法**

**express.AddSignatureLine ( *varSigProv* )**

\*express是一个代表 SignatureSet 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                    |
|--------------|-----------|----------|-------------------------|
| *varSigProv* | 可选      | Variant  | 代表签名提供程序的 ID。 |

**返回值**

Signature

**说明**

添加了行后，文档的作者可以添加必要的信息，以使每个签名行都显示应签名人员的姓名和（可选）职称。当用户打开文档时，WPS Office可识别出存在一个或多个签名行，但签名行为空白。随后，Office 将向用户发出警告，告知他们需要签署此文档，并帮助用户找到所请求的签名在文档中的位置。

**示例**

``` JavaScript
/*以下示例中的过程将接收签名提供程序的 ID，并添加签名行（只要文档不为只读）。*/
function test(SignProviderID) {
    if(Application.ActiveWorkbook.Signatures.CanAddSignatureLine) {
    var objSignature = Application.ActiveWorkbook.Signatures.AddSignatureLine(SignProviderID); 
    }
    return objSignature;
}
```

### SignatureSet.Item

获取一个与文档当前所签署的数字签名之一相对应的 **Signature** 对象。只读。

**语法**

**express.Item ( *iSig* )**

\*express是一个代表 SignatureSet 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                          |
|--------|-----------|----------|-------------------------------|
| *iSig* | 必选      | Long     | 确定要返回的 Signature 对象。 |

## 成员属性

### SignatureSet.Application

获取一个 **Application** 对象，代表 **SignatureSet** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 SignatureSet 对象的变量。

### SignatureSet.CanAddSignatureLine

获取一个 **Boolean** 类型的值，指示是否可向文档中添加签名行。只读。

**语法**

**express.CanAddSignatureLine**

\*express是一个代表 SignatureSet 对象的变量。

### SignatureSet.Count

获取一个 **Long** 类型的值，指示 **SignatureSet** 对象中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 SignatureSet 对象的变量。

### SignatureSet.Creator

获取一个 32 位整数，指示创建 **SignatureSet** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SignatureSet 对象的变量。

### SignatureSet.Parent

获取 **SignatureSet** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SignatureSet 对象的变量。

### SignatureSet.ShowSignaturesPane

获取或设置一个 **Boolean** 类型的值，指示是否应显示 **“签名”** 任务窗格。可读/写。

**语法**

**express.ShowSignaturesPane**

\*express是一个代表 SignatureSet 对象的变量。

### SignatureSet.Subset

获取或设置一个可充当文档可用 **Signature** 对象上的筛选器的值。可读/写。

**语法**

**express.Subset**

\*express是一个代表 SignatureSet 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SignatureSet](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
