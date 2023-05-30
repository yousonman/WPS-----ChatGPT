# Pages 对象

## 简介

文档页面的集合。使用 Pages 集合及相关对象和属性可通过编程方式定义工作簿的页面布局。

## 说明

使用 Pages 属性可返回 Pages 集合。下面的示例访问活动工作表的所有页面。

``` JavaScript
let objPage = Application.ActiveWindow.Panes.Item(1).Pages
```

使用 Item 方法可访问代表工作表中单个页面的单个 Page 对象。下面的示例访问活动工作表的第一页。

``` JavaScript
let objPage = Application.ActiveWindow.Panes.Item(1).Pages.Item(1)
```

## 方法

| 序号 | 名称                | 说明                                                     |
|:----:|:--------------------|:---------------------------------------------------------|
|  1   | [Item](#Pages.Item) | 返回一个 Page 对象，该对象代表工作簿中的页的集合。只读。 |

## 属性

| 序号 | 名称                  | 说明                                   |
|:----:|:----------------------|:---------------------------------------|
|  1   | [Count](#Pages.Count) | 返回集合中对象的数目。只读 Long 类型。 |

## 成员方法

### Pages.Item

返回一个 **Page** 对象，该对象代表工作簿中的页的集合。只读。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Pages 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明         |
|---------|-----------|----------|--------------|
| *Index* | 必选      | Variant  | 页的索引值。 |

## 成员属性

### Pages.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 Pages 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Pages](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
