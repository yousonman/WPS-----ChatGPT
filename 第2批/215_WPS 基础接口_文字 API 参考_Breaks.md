# Breaks 对象

## 简介

页面中分页符、分栏符和分节符的集合。使用 Breaks 集合及相关对象和属性可通过编程方式定义文档的页面版式。

## 方法

| 序号 | 名称                 | 说明                          |
|:----:|:---------------------|:------------------------------|
|  1   | [Item](#Breaks.Item) | 返回集合中的单个 Break 对象。 |

## 属性

| 序号 | 名称                               | 说明                                                                         |
|:----:|:-----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Breaks.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                         |
|  2   | [Count](#Breaks.Count)             | 返回 Breaks 集合中的项目数。 Long 类型，只读。                               |
|  3   | [Creator](#Breaks.Creator)         | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。 |
|  4   | [Parent](#Breaks.Parent)           | 返回一个 Object 类型的值，该值代表指定 Breaks 集合中的父对象。               |

## 成员方法

### Breaks.Item

返回集合中的单个 **Break** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Breaks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

## 成员属性

### Breaks.Application

返回一个 Application 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Breaks 对象的变量。

### Breaks.Count

返回 **Breaks** 集合中的项目数。 **Long** 类型，只读。

**语法**

**express.Count**

\*express是一个代表 Breaks 对象的变量。

### Breaks.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Breaks 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

### Breaks.Parent

返回一个 **Object** 类型的值，该值代表指定 **Breaks** 集合中的父对象。

**语法**

**express.Parent**

\*express是一个代表 Breaks 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Breaks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
