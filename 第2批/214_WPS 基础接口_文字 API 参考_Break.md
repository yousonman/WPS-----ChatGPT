# Break 对象

## 简介

代表页面中的单个分页符、分栏符和分节符。使用 Break 对象及相关方法和属性 可通过编程方式定义文档的页面版式。

## 说明

使用 Breaks 集合的 Item 方法返回特定 Break 对象。以下示例返回活动文档第一页中的第一个分隔符。

``` JavaScript
Application.ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1).Breaks.Item(1)
```

## 属性

| 序号 | 名称                              | 说明                                                                         |
|:----:|:----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Break.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                         |
|  2   | [Creator](#Break.Creator)         | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。 |
|  3   | [PageIndex](#Break.PageIndex)     | 返回一个 Long 类型的值，该值代表指定分隔符出现的页码。                       |
|  4   | [Parent](#Break.Parent)           | 返回一个 Object 类型的值，该值代表指定 Break 对象的父对象。                  |
|  5   | [Range](#Break.Range)             | 返回一个 Range 对象，该对象代表指定对象所含的部分文档。                      |

## 成员属性

### Break.Application

返回一个 Application 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Break 对象的变量。

### Break.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Break 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

### Break.PageIndex

返回一个 **Long** 类型的值，该值代表指定分隔符出现的页码。

**语法**

**express.PageIndex**

\*express是一个代表 Break 对象的变量。

**示例**

``` JavaScript
/*下列代码返回指定分隔符出现的页码。*/
Application.ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1).Breaks.Item(1).PageIndex
```

### Break.Parent

返回一个 **Object** 类型的值，该值代表指定 **Break** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Break 对象的变量。

### Break.Range

返回一个 Range 对象，该对象代表指定对象所含的部分文档。

**语法**

**express.Range**

\*express是一个代表 Break 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Break](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
