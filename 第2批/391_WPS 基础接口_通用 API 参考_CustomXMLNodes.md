# CustomXMLNodes 对象

## 简介

包含 CustomXMLNodes 对象的集合，这些对象代表文档中的 XML 节点。

## 说明

Attributes 和 ChildNodes 属性返回此类型节点的集合。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

## 属性

| 序号 | 名称                                       | 说明                                                                       |
|:----:|:-------------------------------------------|:---------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLNodes.Application) | 获取一个代表 CustomXMLNodes 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [Count](#CustomXMLNodes.Count)             | 获取 CustomXMLNodes 集合中 CustomXMLNode 对象的数目。只读。                |
|  3   | [Creator](#CustomXMLNodes.Creator)         | 获取一个 32 位整数，指示创建 CustomXMLNodes 对象时所使用的应用程序。只读。 |
|  4   | [Item](#CustomXMLNodes.Item)               | 获取 CustomXMLNodes 集合中的一个 CustomXMLNode 对象。只读。                |
|  5   | [Parent](#CustomXMLNodes.Parent)           | 获取 CustomXMLNodes 对象的 Parent 对象。只读。                             |

## 成员属性

### CustomXMLNodes.Application

获取一个代表 **CustomXMLNodes** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLNodes 对象的变量。

### CustomXMLNodes.Count

获取 **CustomXMLNodes** 集合中 **CustomXMLNode** 对象的数目。只读。

**语法**

**express.Count**

\*express是一个代表 CustomXMLNodes 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNodes.Creator

获取一个 32 位整数，指示创建 **CustomXMLNodes** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLNodes 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNodes.Item

获取 **CustomXMLNodes** 集合中的一个 **CustomXMLNode** 对象。只读。

**语法**

**express.Item**

\*express是一个代表 CustomXMLNodes 对象的变量。

### CustomXMLNodes.Parent

获取 **CustomXMLNodes** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLNodes 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLNodes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
