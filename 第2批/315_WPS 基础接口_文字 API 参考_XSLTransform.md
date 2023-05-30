# XSLTransform 对象

## 简介

代表一个已注册的可扩展样式表语言转换 (XSLT)。

## 方法

| 序号 | 名称                           | 说明                                                      |
|:----:|:-------------------------------|:----------------------------------------------------------|
|  1   | [Delete](#XSLTransform.Delete) | 删除可用的可扩展样式表语言转换 (XSLT) 列表中指定的 XSLT。 |

## 属性

| 序号 | 名称                                     | 说明                                                                                      |
|:----:|:-----------------------------------------|:------------------------------------------------------------------------------------------|
|  1   | [Alias](#XSLTransform.Alias)             | 返回一个 String 类型的值，该值代表指定对象的显示名称。                                    |
|  2   | [Application](#XSLTransform.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                            |
|  3   | [Creator](#XSLTransform.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。              |
|  4   | [ID](#XSLTransform.ID)                   | 返回 String 类型的值，该值包含分配给当前 XSLTransform 对象的 GUID。只读。                 |
|  5   | [Location](#XSLTransform.Location)       | 返回或设置 String 值，该值表示 XML 架构命名空间的 XSL 转换在架构库中的物理位置。可读/写。 |
|  6   | [Parent](#XSLTransform.Parent)           | 返回一个 Object 类型值，该值代表指定 XSLTransform 对象的父对象。                          |

## 成员方法

### XSLTransform.Delete

删除可用的可扩展样式表语言转换 (XSLT) 列表中指定的 XSLT。

**语法**

**express.Delete ()**

\*express是一个代表 XSLTransform 对象的变量。

## 成员属性

### XSLTransform.Alias

返回一个 **String** 类型的值，该值代表指定对象的显示名称。

**语法**

**express.Alias**

\*express是一个代表 XSLTransform 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动文档所附加的第一个架构的显示名称。*/
MsgBox(Application.XMLNamespaces.Item(1).Alias)
```

### XSLTransform.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 XSLTransform 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XSLTransform.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XSLTransform 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XSLTransform.ID

返回 **String** 类型的值，该值包含分配给当前 XSLTransform 对象的 GUID。只读。

**语法**

**express.ID**

\*express是一个代表 XSLTransform 对象的变量。

### XSLTransform.Location

返回或设置 **String** 值，该值表示 XML 架构命名空间的 XSL 转换在架构库中的物理位置。可读/写。

**语法**

**express.Location**

\*express是一个代表 XSLTransform 对象的变量。

### XSLTransform.Parent

返回一个 **Object** 类型值，该值代表指定 **XSLTransform** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XSLTransform 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XSLTransform](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
