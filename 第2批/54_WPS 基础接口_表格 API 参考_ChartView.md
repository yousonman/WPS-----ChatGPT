# ChartView 对象

## 简介

代表图表的视图。

## 说明

ChartView 对象是可由 SheetViews 集合（类似于 Sheets 集合）返回的对象之一。 ChartView 对象只适用于图表工作表。

``` JavaScript
/*以下示例返回 ChartView 对象。*/

ActiveWindow.SheetViews.Item(1)
```

``` JavaScript
/*以下示例返回 Chart 对象。*/
ActiveWindow.SheetViews.Item(1).Sheet
```

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                               |
|:----:|:--------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartView.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#ChartView.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Parent](#ChartView.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  4   | [Sheet](#ChartView.Sheet)             | 返回指定的 ChartView 对象的工作表名称。只读。                                                                                                                      |

## 成员属性

### ChartView.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ChartView 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### ChartView.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartView 对象的变量。

### ChartView.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartView 对象的变量。

### ChartView.Sheet

返回指定的 **ChartView** 对象的工作表名称。只读。

**语法**

**express.Sheet**

\*express是一个代表 ChartView 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
