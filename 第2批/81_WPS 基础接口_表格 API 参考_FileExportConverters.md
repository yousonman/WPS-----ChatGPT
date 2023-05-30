# FileExportConverters 对象

## 简介

FileExportConverter 对象的集合，这些对象代表可供保存文件时使用的所有文件转换器。

## 属性

| 序号 | 名称                                             | 说明                                                                             |
|:----:|:-------------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [Application](#FileExportConverters.Application) | 返回代表 ET 应用程序的 Application 对象。只读。                                  |
|  2   | [Count](#FileExportConverters.Count)             | 返回一个 Long 类型的值，该值代表集合中文件转换器的数量。只读。                   |
|  3   | [Creator](#FileExportConverters.Creator)         | 返回表示在其中创建特定对象的应用程序的 32 位整数。 Long 类型，只读。             |
|  4   | [Item](#FileExportConverters.Item)               | 返回集合中的单个 FileExportConverter 对象。                                      |
|  5   | [Parent](#FileExportConverters.Parent)           | 返回一个 Object 类型的值，该值代表指定 FileExportConverters 对象的父对象。只读。 |

## 成员属性

### FileExportConverters.Application

返回代表 ET 应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 FileExportConverters 对象的变量。

### FileExportConverters.Count

返回一个 **Long** 类型的值，该值代表集合中文件转换器的数量。只读。

**语法**

**express.Count**

\*express是一个代表 FileExportConverters 对象的变量。

### FileExportConverters.Creator

返回表示在其中创建特定对象的应用程序的 32 位整数。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 FileExportConverters 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回十六进制值 5843454C，该值表示字符串“XCEL”。 **Creator** 属性设计为在 ET for the Macintosh 中使用，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### FileExportConverters.Item

返回集合中的单个 **FileExportConverter** 对象。

**语法**

**express.Item**

\*express是一个代表 FileExportConverters 对象的变量。

### FileExportConverters.Parent

返回一个 **Object** 类型的值，该值代表指定 **FileExportConverters** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FileExportConverters 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/FileExportConverters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
