# SheetViews 对象

## 简介

指定的或活动工作簿窗口中所有工作表视图的集合。

## 说明

以下示例将返回活动窗口的视图计数。

``` JavaScript
Application.ActiveWindow.SheetViews.Count
```

## 方法

| 序号 | 名称                     | 说明                                                    |
|:----:|:-------------------------|:--------------------------------------------------------|
|  1   | [Item](#SheetViews.Item) | 返回一个 SheetView 对象，该对象代表工作簿的视图。只读。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                               |
|:----:|:---------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SheetViews.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#SheetViews.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#SheetViews.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Parent](#SheetViews.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员方法

### SheetViews.Item

返回一个 **SheetView** 对象，该对象代表工作簿的视图。只读。

**语法**

**express.Item ( *index* )**

\*express是一个代表 SheetViews 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *index* | 必选      | Variant  | 视图的索引值。 |

**示例**

## 成员属性

### SheetViews.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SheetViews 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### SheetViews.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 SheetViews 对象的变量。

### SheetViews.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 SheetViews 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### SheetViews.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SheetViews 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/SheetViews](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
