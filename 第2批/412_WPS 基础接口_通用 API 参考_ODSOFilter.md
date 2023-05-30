# ODSOFilter 对象

## 简介

表示应用于附加的邮件合并数据源的筛选。 ODSOFilter 对象为 ODSOFilters 对象的成员。

## 说明

每个筛选器均为查询字符串中的一行。使用 Column 、 Comparison 、 CompareTo 和 Conjunction 属性可返回或设置数据源查询条件。

## 属性

| 序号 | 名称                                   | 说明                                                                                                                              |
|:----:|:---------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ODSOFilter.Application) | 获取一个 Application 对象，代表 ODSOFilter 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Column](#ODSOFilter.Column)           | 获取或设置一个 String 类型的值，该值代表邮件合并数据源中要用于筛选的字段名。可读/写。                                             |
|  3   | [CompareTo](#ODSOFilter.CompareTo)     | 获取或设置一个 String 类型的值，该值代表查询筛选条件中要比较的文本。可读/写。                                                     |
|  4   | [Comparison](#ODSOFilter.Comparison)   | 获取或设置一个代表 Column 和 CompareTo 的比较方式的 MsoFilterComparison 常量。可读/写。                                           |
|  5   | [Conjunction](#ODSOFilter.Conjunction) | 获取或设置一个 MsoFilterConjunction 常量，该常量代表 ODSOFilters 对象中的某个筛选条件与其中的其他筛选条件的关系。可读/写。        |
|  6   | [Creator](#ODSOFilter.Creator)         | 获取一个 32 位整数，指示创建 ODSOFilter 对象时所使用的应用程序。只读。                                                            |
|  7   | [Index](#ODSOFilter.Index)             | 获取一个 Long 类型的值，代表集合中 ODSOFilter 对象的索引号。只读。                                                                |
|  8   | [Parent](#ODSOFilter.Parent)           | 获取 ODSOFilter 对象的 Parent 对象。只读。                                                                                        |

## 成员属性

### ODSOFilter.Application

获取一个 **Application** 对象，代表 **ODSOFilter** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 ODSOFilter 对象的变量。

### ODSOFilter.Column

获取或设置一个 **String** 类型的值，该值代表邮件合并数据源中要用于筛选的字段名。可读/写。

**语法**

**express.Column**

\*express是一个代表 ODSOFilter 对象的变量。

**示例**

### ODSOFilter.CompareTo

获取或设置一个 **String** 类型的值，该值代表查询筛选条件中要比较的文本。可读/写。

**语法**

**express.CompareTo**

\*express是一个代表 ODSOFilter 对象的变量。

**示例**

### ODSOFilter.Comparison

获取或设置一个代表 **Column** 和 **CompareTo** 的比较方式的 **MsoFilterComparison** 常量。可读/写。

**语法**

**express.Comparison**

\*express是一个代表 ODSOFilter 对象的变量。

**示例**

### ODSOFilter.Conjunction

获取或设置一个 **MsoFilterConjunction** 常量，该常量代表 **ODSOFilters** 对象中的某个筛选条件与其中的其他筛选条件的关系。可读/写。

**语法**

**express.Conjunction**

\*express是一个代表 ODSOFilter 对象的变量。

**示例**

### ODSOFilter.Creator

获取一个 32 位整数，指示创建 **ODSOFilter** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 ODSOFilter 对象的变量。

### ODSOFilter.Index

获取一个 **Long** 类型的值，代表集合中 **ODSOFilter** 对象的索引号。只读。

**语法**

**express.Index**

\*express是一个代表 ODSOFilter 对象的变量。

### ODSOFilter.Parent

获取 **ODSOFilter** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ODSOFilter 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/ODSOFilter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
