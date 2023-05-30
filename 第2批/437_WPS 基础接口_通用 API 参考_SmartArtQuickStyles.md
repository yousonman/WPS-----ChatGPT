# SmartArtQuickStyles 对象

## 简介

代表 Smart Art 快速样式的集合。

## 说明

``` JavaScript
/*下面的代码更改 WPP 中的 Smart Art 图表的快速样式。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(1)
```

## 方法

| 序号 | 名称                              | 说明                                                               |
|:----:|:----------------------------------|:-------------------------------------------------------------------|
|  1   | [Item](#SmartArtQuickStyles.Item) | 在指定的索引处或通过指定的唯一 ID 来检索 SmartArtQuickStyle 对象。 |

## 属性

| 序号 | 名称                                            | 说明                                                                                |
|:----:|:------------------------------------------------|:------------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtQuickStyles.Application) | 获取一个 Application 对象，该对象代表 SmartArtQuickStyles 对象的容器应用程序。只读  |
|  2   | [Count](#SmartArtQuickStyles.Count)             | 检索包含在 SmartArtQuickStyles 集合中的 SmartArtQuickStyle 对象数的计数。只读       |
|  3   | [Creator](#SmartArtQuickStyles.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtQuickStyles 对象的应用程序。只读 |
|  4   | [Parent](#SmartArtQuickStyles.Parent)           | 返回调用对象。只读                                                                  |

## 成员方法

### SmartArtQuickStyles.Item

在指定的索引处或通过指定的唯一 ID 来检索 **SmartArtQuickStyle** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                   |
|---------|-----------|----------|------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 指定一个代表索引的整数或一个代表 SmartArtQuickStyle 对象位置的字符串。 |

**返回值**

SmartArtQuickStyle

## 成员属性

### SmartArtQuickStyles.Application

获取一个 **Application** 对象，该对象代表 **SmartArtQuickStyles** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

### SmartArtQuickStyles.Count

检索包含在 SmartArtQuickStyles 集合中的 SmartArtQuickStyle 对象数的计数。只读

**语法**

**express.Count**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

### SmartArtQuickStyles.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtQuickStyles** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

### SmartArtQuickStyles.Parent

返回调用对象。只读

**语法**

**express.Parent**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtQuickStyles](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
