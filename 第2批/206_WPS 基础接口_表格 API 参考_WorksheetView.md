# WorksheetView 对象

## 简介

指定的或活动工作簿中所有 WorksheetView 对象的集合。

## 说明

通过使用 DisplayFormulas、DisplayGridlines 和 DisplayHeadings 等属性控制应用程序或工作簿级视图的外观和风格。

## 属性

| 序号 | 名称                                                | 说明                                                                                                                                                               |
|:----:|:----------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#WorksheetView.Application)           | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [DisplayFormulas](#WorksheetView.DisplayFormulas)   | 返回或设置在当前工作表视图中是显示还是隐藏公式。可读/写 Boolean 类型。                                                                                             |
|  3   | [DisplayGridlines](#WorksheetView.DisplayGridlines) | 如果显示网格线，则为 True 。可读/写 Boolean 类型。                                                                                                                 |
|  4   | [DisplayHeadings](#WorksheetView.DisplayHeadings)   | 如果同时显示行标题和列标题，则为 True ；如果未显示标题，则为 False 。可读/写 Boolean 类型。                                                                        |
|  5   | [DisplayOutline](#WorksheetView.DisplayOutline)     | 如果显示分级显示符号，则为 True 。可读/写 Boolean 类型。                                                                                                           |
|  6   | [DisplayZeros](#WorksheetView.DisplayZeros)         | 如果显示零值，则为 True 。可读/写 Boolean 类型。                                                                                                                   |
|  7   | [Sheet](#WorksheetView.Sheet)                       | 返回指定 WorksheetView 对象的工作表名称。只读。                                                                                                                    |

## 成员属性

### WorksheetView.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### WorksheetView.DisplayFormulas

返回或设置在当前工作表视图中是显示还是隐藏公式。可读/写 **Boolean** 类型。

**语法**

**express.DisplayFormulas**

\*express是一个代表 WorksheetView 对象的变量。

### WorksheetView.DisplayGridlines

如果显示网格线，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.DisplayGridlines**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

此属性仅影响显示的网格线。使用 **PrintGridlines** 属性可以控制网格线的打印。

### WorksheetView.DisplayHeadings

如果同时显示行标题和列标题，则为 **True** ；如果未显示标题，则为 **False** 。可读/写 **Boolean** 类型。

**语法**

**express.DisplayHeadings**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

此属性仅影响显示的标题。使用 **PrintHeadings** 属性可以控制标题的打印。

### WorksheetView.DisplayOutline

如果显示分级显示符号，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.DisplayOutline**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

### WorksheetView.DisplayZeros

如果显示零值，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.DisplayZeros**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

### WorksheetView.Sheet

返回指定 **WorksheetView** 对象的工作表名称。只读。

**语法**

**express.Sheet**

\*express是一个代表 WorksheetView 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/WorksheetView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
