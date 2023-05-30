# CustomXMLPrefixMapping 对象

## 简介

代表一个命名空间前缀。

## 说明

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

以下示例通过向 CustomXMLPrefixMapping 集合添加命名空间和前缀，创建一个 CustomXMLPrefixMapping 对象。

``` JavaScript
let objNamespace = Application.ActivePresentation.CustomXMLParts.Item(1).NamespaceManager.AddNamespace("xs", "urn:invoice:namespace")
```

## 属性

| 序号 | 名称                                                 | 说明                                                                               |
|:----:|:-----------------------------------------------------|:-----------------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLPrefixMapping.Application)   | 获取一个代表 CustomXMLPrefixMapping 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [Creator](#CustomXMLPrefixMapping.Creator)           | 获取一个 32 位整数，指示创建 CustomXMLPrefixMapping 对象时所使用的应用程序。只读。 |
|  3   | [NamespaceURI](#CustomXMLPrefixMapping.NamespaceURI) | 获取 CustomXMLPrefixMapping 对象的命名空间的唯一地址标识符。只读。                 |
|  4   | [Parent](#CustomXMLPrefixMapping.Parent)             | 获取 CustomXMLPrefixMapping 对象的 Parent 对象。只读。                             |
|  5   | [Prefix](#CustomXMLPrefixMapping.Prefix)             | 获取 CustomXMLPrefixMapping 对象的前缀。只读。                                     |

## 成员属性

### CustomXMLPrefixMapping.Application

获取一个代表 **CustomXMLPrefixMapping** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMapping.Creator

获取一个 32 位整数，指示创建 **CustomXMLPrefixMapping** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMapping.NamespaceURI

获取 **CustomXMLPrefixMapping** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMapping.Parent

获取 **CustomXMLPrefixMapping** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMapping.Prefix

获取 **CustomXMLPrefixMapping** 对象的前缀。只读。

**语法**

**express.Prefix**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLPrefixMapping](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
