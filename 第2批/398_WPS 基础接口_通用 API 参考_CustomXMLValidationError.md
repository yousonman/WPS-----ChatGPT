# CustomXMLValidationError 对象

## 简介

代表 CustomXMLValidationErrors 集合中的单个验证错误。

## 说明

当根据架构验证操作时（如添加节点时），验证错误可能被触发，或者在不满足某个条件时（例如，如果开始日期晚于结束日期），可能由用户触发。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

## 方法

| 序号 | 名称                                       | 说明                                                   |
|:----:|:-------------------------------------------|:-------------------------------------------------------|
|  1   | [Delete](#CustomXMLValidationError.Delete) | 删除代表数据验证错误的 CustomXMLValidationError 对象。 |

## 属性

| 序号 | 名称                                                 | 说明                                                                                                        |
|:----:|:-----------------------------------------------------|:------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLValidationError.Application) | 获取一个代表 CustomXMLValidationError 对象的容器应用程序的 Application 对象。只读。                         |
|  2   | [Creator](#CustomXMLValidationError.Creator)         | 获取一个 32 位整数，指示创建 CustomXMLValidationError 对象时所使用的应用程序。只读。                        |
|  3   | [ErrorCode](#CustomXMLValidationError.ErrorCode)     | 获取一个代表 CustomXMLValidationError 对象中的验证错误的数字。只读。                                        |
|  4   | [Name](#CustomXMLValidationError.Name)               | 获取 CustomXMLValidationError 对象中的错误名称。如果不存在错误，则该属性将返回 Nothing 。只读。             |
|  5   | [Node](#CustomXMLValidationError.Node)               | 获取 CustomXMLValidationError 对象中的一个节点（如果存在）。如果不存在节点，则该属性将返回 Nothing 。只读。 |
|  6   | [Parent](#CustomXMLValidationError.Parent)           | 获取 CustomXMLValidationError 对象的 Parent 对象。只读。                                                    |
|  7   | [Text](#CustomXMLValidationError.Text)               | 获取 CustomXMLValidationError 对象中的文本。只读。                                                          |
|  8   | [Type](#CustomXMLValidationError.Type)               | 获取 CustomXMLValidationError 对象中所生成的错误的类型。只读。                                              |

## 成员方法

### CustomXMLValidationError.Delete

删除代表数据验证错误的 **CustomXMLValidationError** 对象。

**语法**

**express.Delete ()**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

|                                                                                                                                                                      |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例将删除验证错误。
Application.ActivePresentation.CustomXMLParts.Item(1).Errors.Item(1).Delete()
```

## 成员属性

### CustomXMLValidationError.Application

获取一个代表 **CustomXMLValidationError** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Creator

获取一个 32 位整数，指示创建 **CustomXMLValidationError** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.ErrorCode

获取一个代表 **CustomXMLValidationError** 对象中的验证错误的数字。只读。

**语法**

**express.ErrorCode**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Name

获取 **CustomXMLValidationError** 对象中的错误名称。如果不存在错误，则该属性将返回 **Nothing** 。只读。

**语法**

**express.Name**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Node

获取 **CustomXMLValidationError** 对象中的一个节点（如果存在）。如果不存在节点，则该属性将返回 **Nothing** 。只读。

**语法**

**express.Node**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Parent

获取 **CustomXMLValidationError** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Text

获取 **CustomXMLValidationError** 对象中的文本。只读。

**语法**

**express.Text**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Type

获取 **CustomXMLValidationError** 对象中所生成的错误的类型。只读。

**语法**

**express.Type**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLValidationError](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
