# Lists 对象

## 简介

List 对象的集合，这些对象代表指定文档中的所有列表。

## 方法

| 序号 | 名称                | 说明 |
|:----:|:--------------------|:-----|
|  1   | [Item](#Lists.Item) |      |

## 属性

| 序号 | 名称                              | 说明                                                       |
|:----:|:----------------------------------|:-----------------------------------------------------------|
|  1   | [Application](#Lists.Application) | 返回一个代表 WPS 应用程序的 Application 对象               |
|  2   | [Count](#Lists.Count)             | 返回一个 Number 类型的值，该值代表集合中列表的数目。只读。 |

## 成员方法

### Lists.Item

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Lists 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                           |
|---------|-----------|----------|----------------------------------------------------------------|
| *Index* | 必选      | Number   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

## 成员属性

### Lists.Application

返回一个代表 WPS 应用程序的 **Application** 对象

**语法**

**express.Application**

\*express是一个代表 Lists 对象的变量。

### Lists.Count

返回一个 **Number** 类型的值，该值代表集合中列表的数目。只读。

**语法**

**express.Count**

\*express是一个代表 Lists 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Lists](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
