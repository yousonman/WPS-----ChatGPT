# ODSOColumns 对象

## 简介

代表邮件合并数据源中的数据字段的 ODSOColumn 对象的集合。

## 说明

## 方法

| 序号 | 名称                      | 说明                                        |
|:----:|:--------------------------|:--------------------------------------------|
|  1   | [Item](#ODSOColumns.Item) | 指定 ODSOColumns 集合中的 ODSOColumn 对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                               |
|:----:|:----------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ODSOColumns.Application) | 获取一个 Application 对象，代表 ODSOColumns 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Count](#ODSOColumns.Count)             | 获取一个 Long 类型的值，指示 ODSOColumns 集合中的项数。只读。                                                                      |
|  3   | [Creator](#ODSOColumns.Creator)         | 获取一个 32 位整数，指示创建 ODSOColumns 对象时所使用的应用程序。只读。                                                            |
|  4   | [Parent](#ODSOColumns.Parent)           | 获取 ODSOColumns 对象的 Parent 对象。只读。                                                                                        |

## 成员方法

### ODSOColumns.Item

指定 **ODSOColumns** 集合中的 **ODSOColumn** 对象。

**语法**

**express.Item ( *varIndex* )**

\*express是一个代表 ODSOColumns 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明           |
|------------|-----------|----------|----------------|
| *varIndex* | 必选      | Variant  | 项目的索引号。 |

**示例**

## 成员属性

### ODSOColumns.Application

获取一个 **Application** 对象，代表 **ODSOColumns** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 ODSOColumns 对象的变量。

### ODSOColumns.Count

获取一个 **Long** 类型的值，指示 **ODSOColumns** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 ODSOColumns 对象的变量。

### ODSOColumns.Creator

获取一个 32 位整数，指示创建 **ODSOColumns** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 ODSOColumns 对象的变量。

### ODSOColumns.Parent

获取 **ODSOColumns** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ODSOColumns 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/ODSOColumns](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
