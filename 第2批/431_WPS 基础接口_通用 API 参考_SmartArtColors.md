# SmartArtColors 对象

## 简介

SmartArt 颜色样式的集合。

## 说明

模拟 WPS Office Fluent 功能区用户界面中“SmartArt 工具”上“设计”组内的“更改颜色”命令上的命令。

``` JavaScript
/*下面的代码设置 Smart Art 图表的配色方案。*/Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Color = Application.SmartArtColors.Item(1)
```

## 方法

| 序号 | 名称                         | 说明                                                          |
|:----:|:-----------------------------|:--------------------------------------------------------------|
|  1   | [Item](#SmartArtColors.Item) | 检索位于指定索引处或具有指定的唯一 Id 的 SmartArtColor 对象。 |

## 属性

| 序号 | 名称                                       | 说明                                                                             |
|:----:|:-------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtColors.Application) | 获取一个 Application 对象，该对象代表 SmartArtColors 对象的容器应用程序。只读。  |
|  2   | [Count](#SmartArtColors.Count)             | 检索包含在 SmartArtColors 集合中的 SmartArtColor 对象数的计数。只读。            |
|  3   | [Creator](#SmartArtColors.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtColors 对象的应用程序。只读。 |
|  4   | [Parent](#SmartArtColors.Parent)           | 返回调用对象。只读。                                                             |

## 成员方法

### SmartArtColors.Item

检索位于指定索引处或具有指定的唯一 Id 的 **SmartArtColor** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SmartArtColors 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                            |
|---------|-----------|----------|-----------------------------------------------------------------|
| *Index* | 必选      | Variant  | 指定一个代表索引的整数或一个代表 Smart Art 颜色的 Id 的字符串。 |

**返回值**

SmartArtColor

## 成员属性

### SmartArtColors.Application

获取一个 **Application** 对象，该对象代表 **SmartArtColors** 对象的容器应用程序。只读。

**语法**

**express.Application**

\*express是一个代表 SmartArtColors 对象的变量。

### SmartArtColors.Count

检索包含在 **SmartArtColors** 集合中的 **SmartArtColor** 对象数的计数。只读。

**语法**

**express.Count**

\*express是一个代表 SmartArtColors 对象的变量。

### SmartArtColors.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtColors** 对象的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SmartArtColors 对象的变量。

### SmartArtColors.Parent

返回调用对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SmartArtColors 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtColors](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
