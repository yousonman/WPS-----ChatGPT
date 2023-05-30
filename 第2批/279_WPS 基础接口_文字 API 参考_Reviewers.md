# Reviewers 对象

## 简介

Reviewer 对象 [的集合](https://docs.microsoft.com/zh-cn/office/vba/api/word.reviewer) ，这些对象代表一个或多个文档的审阅者。 审阅者 集合中包含的所有审阅者审阅或编辑在一台计算机上打开的文档的名称。

## 说明

使用 Reviewers (Index) （其中 Index 是审阅者的名称或索引号）可返回 Reviewers 集合中的单个审阅者。

## 方法

| 序号 | 名称                    | 说明                           |
|:----:|:------------------------|:-------------------------------|
|  1   | [Item](#Reviewers.Item) | 返回集合中的单个 Reviewer 对象 |

## 属性

| 序号 | 名称                      | 说明                                                     |
|:----:|:--------------------------|:---------------------------------------------------------|
|  1   | [Count](#Reviewers.Count) | 返回一 个 Long ，代表集合中审阅者的数量。 此为只读属性。 |

## 成员方法

### Reviewers.Item

返回集合中的单个 **Reviewer** 对象

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Reviewers 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                              |
|---------|-----------|----------|-----------------------------------------------------------------------------------|
| *Index* | 可选      | Variant  | 要返回的单个对象。 可以是 Long 类型，表示序号位置或 字符串 ，表示单个对象的名称。 |

**示例**

``` JavaScript
//获取当前激活窗口第一个审阅者
Application.ActiveWindow.View.Reviewers.Item(1)
```

## 成员属性

### Reviewers.Count

返回一 个 Long ，代表集合中审阅者的数量。 此为只读属性。

**语法**

**express.Count**

\*express是一个代表 Reviewers 对象的变量。

**示例**

``` JavaScript
//获取审阅者的数量。 
Application.ActiveWindow.View.Reviewers.Count
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Reviewers](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
