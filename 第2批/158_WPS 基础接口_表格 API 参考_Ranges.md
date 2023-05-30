# Ranges 对象

## 简介

由 Range 对象组成的集合。

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                               |
|:----:|:-----------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Ranges.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#Ranges.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#Ranges.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Item](#Ranges.Item)               | 返回一个 Range 对象，该对象代表工作簿中的项的区域。只读。                                                                                                          |
|  5   | [Parent](#Ranges.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员属性

### Ranges.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Ranges 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### Ranges.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 Ranges 对象的变量。

### Ranges.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Ranges 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Ranges.Item

返回一个 **Range** 对象，该对象代表工作簿中的项的区域。只读。

**语法**

**express.Item**

\*express是一个代表 Ranges 对象的变量。

### Ranges.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Ranges 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Ranges](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
