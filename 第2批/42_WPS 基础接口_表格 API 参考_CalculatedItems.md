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
