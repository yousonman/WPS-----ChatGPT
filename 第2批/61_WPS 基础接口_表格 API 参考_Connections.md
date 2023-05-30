# Connections 对象

## 简介

指定工作簿的 Connection 对象的集合。

## 说明

``` JavaScript
/*以下示例演示如何从现有文件向工作簿添加连接。*/
Application.ActiveWorkbook.Connections.AddFromFile(
    "C:\\Documents and Settings\\myComputer\\My Documents\\My Data Sources\\Northwind 2007 Customers.odc")
```

## 方法

| 序号 | 名称                                    | 说明                   |
|:----:|:----------------------------------------|:-----------------------|
|  1   | [Add](#Connections.Add)                 | 添加到工作簿的新连接。 |
|  2   | [AddFromFile](#Connections.AddFromFile) | 添加从指定文件的连接。 |
|  3   | [Item](#Connections.Item)               | 此方法创建一个连接项。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                               |
|:----:|:----------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Connections.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#Connections.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#Connections.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Parent](#Connections.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员方法

### Connections.Add

添加到工作簿的新连接。

**语法**

**express.Add ( *Name* , *Description* , *ConnectionString* , *CommandText* , *lCmdtype* )**

\*express是一个代表 Connections 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                     |
|--------------------|-----------|----------|--------------------------|
| *Name*             | 必选      | String   | 连接的名称。             |
| *Description*      | 必选      | String   | 连接的简要说明。         |
| *ConnectionString* | 必选      | String   | 连接字符串。             |
| *CommandText*      | 必选      | String   | 用于创建连接的命令文本。 |
| *lCmdtype*         | 可选      | String   | 命令类型。               |

**返回值**

WorkbookConnection

### Connections.AddFromFile

添加从指定文件的连接。

**语法**

**express.AddFromFile ( *Filename* )**

\*express是一个代表 Connections 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明         |
|------------|-----------|----------|--------------|
| *Filename* | 必选      | String   | 文件的名称。 |

**返回值**

WorkbookConnection

### Connections.Item

此方法创建一个连接项。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Connections 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明         |
|---------|-----------|----------|--------------|
| *Index* | 必选      | Variant  | 项的索引值。 |

**返回值**

WorkbookConnection

## 成员属性

### Connections.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Connections 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### Connections.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 Connections 对象的变量。

### Connections.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Connections 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Connections.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Connections 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Connections](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
