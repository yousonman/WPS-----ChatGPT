# CustomXMLSchema 对象

## 简介

代表 CustomXMLSchemaCollection 集合中的一个架构。

## 说明

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

## 方法

| 序号 | 名称                              | 说明                                    |
|:----:|:----------------------------------|:----------------------------------------|
|  1   | [Delete](#CustomXMLSchema.Delete) | 从 CustomXMLSchema 集合中删除指定架构。 |
|  2   | [Reload](#CustomXMLSchema.Reload) | 从文件重新加载架构。                    |

## 属性

| 序号 | 名称                                          | 说明                                                                        |
|:----:|:----------------------------------------------|:----------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLSchema.Application)   | 获取一个代表 CustomXMLSchema 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [Creator](#CustomXMLSchema.Creator)           | 获取一个 32 位整数，指示创建 CustomXMLSchema 对象时所使用的应用程序。只读。 |
|  3   | [Location](#CustomXMLSchema.Location)         | 获取一个 String 类型的值，代表计算机上架构的位置。只读。                    |
|  4   | [NamespaceURI](#CustomXMLSchema.NamespaceURI) | 获取 CustomXMLSchema 对象的命名空间的唯一地址标识符。只读。                 |
|  5   | [Parent](#CustomXMLSchema.Parent)             | 获取 CustomXMLSchema 对象的 Parent 对象。只读。                             |

## 成员方法

### CustomXMLSchema.Delete

从 **CustomXMLSchema** 集合中删除指定架构。

**语法**

**express.Delete ()**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

如果尝试对已经过验证或附加到数据流的集合中的架构进行此操作，则不会执行操作并会显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例将架构添加到集合，然后删除该架构。
function test()
{
     try 
    {
        // Adds a schema to the collection.
    　　Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Add("urn:invoice:namespace") 
        // Deletes the schema.
        Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Item(1).Delete()
    }            
    // Exception handling. Show the message and resume.
    catch(exception) 
    {
        alert(exception.Description)
    }
}
```

### CustomXMLSchema.Reload

从文件重新加载架构。

**语法**

**express.Reload ()**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

通常，此方法用于更新架构的位置或确定架构是否仍然有效。它还用于重新加载频繁更改的架构。如果尝试对集合中的已经过验证或绑定到数据流的架构执行此操作，则不执行操作，并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例指定了架构的位置，然后重新加载它。
function test()
{
    try{
        let objCustomXMLSchema = Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Item(1)
        // Reload the schema.
        bjCustomXMLSchema.Reload()
    }catch(exception){
    alert(exception.Description)
    }
}
```

## 成员属性

### CustomXMLSchema.Application

获取一个代表 **CustomXMLSchema** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLSchema.Creator

获取一个 32 位整数，指示创建 **CustomXMLSchema** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLSchema.Location

获取一个 **String** 类型的值，代表计算机上架构的位置。只读。

**语法**

**express.Location**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLSchema.NamespaceURI

获取 **CustomXMLSchema** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLSchema.Parent

获取 **CustomXMLSchema** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLSchema](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
