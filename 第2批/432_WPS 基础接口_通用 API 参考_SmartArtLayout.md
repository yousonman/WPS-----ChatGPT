# SmartArtLayout 对象

## 简介

代表 Smart Art 图表。

## 说明

选项包括基本列表、图片题注列表、垂直项目符号列表等。

``` JavaScript
/*下面的代码更改 WPP 中的 Smart Art 图表的图表样式。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Layout = Application.SmartArtLayouts.Item(1)
```

## 属性

| 序号 | 名称                                       | 说明                                                                             |
|:----:|:-------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtLayout.Application) | 获取一个 Application 对象，该对象代表 SmartArtLayout 对象的容器应用程序。只读。  |
|  2   | [Category](#SmartArtLayout.Category)       | 检索与 SmartArt 布局关联的主类别名称。只读。                                     |
|  3   | [Creator](#SmartArtLayout.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtLayout 对象的应用程序。只读。 |
|  4   | [Description](#SmartArtLayout.Description) | 检索 SmartArt 布局的说明。只读。                                                 |
|  5   | [Id](#SmartArtLayout.Id)                   | 检索关联的 SmartArt 布局的唯一 Id。只读。                                        |
|  6   | [Name](#SmartArtLayout.Name)               | 检索 SmartArt 布局的字符串名称。只读。                                           |
|  7   | [Parent](#SmartArtLayout.Parent)           | 返回调用对象。只读。                                                             |

## 成员属性

### SmartArtLayout.Application

获取一个 **Application** 对象，该对象代表 **SmartArtLayout** 对象的容器应用程序。只读。

**语法**

**express.Application**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Category

检索与 SmartArt 布局关联的主类别名称。只读。

**语法**

**express.Category**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtLayout** 对象的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Description

检索 SmartArt 布局的说明。只读。

**语法**

**express.Description**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Id

检索关联的 SmartArt 布局的唯一 Id。只读。

**语法**

**express.Id**

\*express是一个代表 SmartArtLayout 对象的变量。

**说明**

与此属性关联的 ID 区分大小写。

### SmartArtLayout.Name

检索 SmartArt 布局的字符串名称。只读。

**语法**

**express.Name**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Parent

返回调用对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SmartArtLayout 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtLayout](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
