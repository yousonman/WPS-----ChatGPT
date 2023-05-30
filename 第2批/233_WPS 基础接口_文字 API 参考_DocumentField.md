# DocumentField 对象

## 简介

代表文档中的一个公文域。

## 方法

| 序号 | 名称                            | 说明           |
|:----:|:--------------------------------|:---------------|
|  1   | [Delete](#DocumentField.Delete) | 删除指定的域。 |
|  2   | [Select](#DocumentField.Select) | 选择指定的域。 |

## 属性

| 序号 | 名称                                      | 说明                                                                          |
|:----:|:------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#DocumentField.Application) | 返回一个代表 WPS 应用程序的 [Application](Application) [](Application) 对象。 |
|  2   | [Creator](#DocumentField.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。  |
|  3   | [Hidden](#DocumentField.Hidden)           | 如果为 True ，则隐藏文档中的公文域。 Boolean 类型，可读写。                   |
|  4   | [Name](#DocumentField.Name)               | 返回或设置指定对象的名称。可读/写 String 类型。                               |
|  5   | [Parent](#DocumentField.Parent)           | 返回一个 Object 类型值，该值代表指定 DocumentField 对象的父对象。             |
|  6   | [PrintOut](#DocumentField.PrintOut)       | 该公文域对象是否可以打印 。                                                   |
|  7   | [Range](#DocumentField.Range)             | 返回指定对象的区域 Range。                                                    |
|  8   | [ReadOnly](#DocumentField.ReadOnly)       | 如果为True，代表该公文域值只读。 Boolean 类型， 可读/写                       |
|  9   | [StoryType](#DocumentField.StoryType)     | 返回一个Long类型值，该值代表此公文域的类型。只读。                            |
|  10  | [Value](#DocumentField.Value)             | 返回一个字符串，该字符串为此公文域中的内容。 String 类型，可读/写。           |

## 成员方法

### DocumentField.Delete

删除指定的域。

**语法**

**express.Delete ( *DeleteWithContent* )**

\*express是一个代表 DocumentField 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                              |
|---------------------|-----------|----------|-----------------------------------|
| *DeleteWithContent* | 可选      | Boolean  | 是否删除该域中的内容，默认为false |

### DocumentField.Select

选择指定的域。

**语法**

**express.Select ()**

\*express是一个代表 DocumentField 对象的变量。

## 成员属性

### DocumentField.Application

返回一个代表 WPS 应用程序的 **[Application](Application)** [](Application) 对象。

**语法**

**express.Application**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DocumentField 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### DocumentField.Hidden

如果为 **True** ，则隐藏文档中的公文域。 **Boolean** 类型，可读写。

**语法**

**express.Hidden**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Name

返回或设置指定对象的名称。可读/写 **String** 类型。

**语法**

**express.Name**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Parent

返回一个 **Object** 类型值，该值代表指定 **DocumentField** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.PrintOut

该公文域对象是否可以打印 。

**语法**

**express.PrintOut**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Range

返回指定对象的区域 Range。

**语法**

**express.Range**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.ReadOnly

如果为True，代表该公文域值只读。 **Boolean** 类型， 可读/写

**语法**

**express.ReadOnly**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.StoryType

返回一个Long类型值，该值代表此公文域的类型。只读。

**语法**

**express.StoryType**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Value

返回一个字符串，该字符串为此公文域中的内容。 **String** 类型，可读/写。

**语法**

**express.Value**

\*express是一个代表 DocumentField 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/DocumentField](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
