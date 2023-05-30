# SmartArtQuickStyle 对象

## 简介

代表 Smart Art 快速样式

## 说明

``` JavaScript
/*下面的代码更改 WPP 中的 Smart Art 图表的快速样式。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(1)
```

## 方法

| 序号 | 名称                                   | 说明 |
|:----:|:---------------------------------------|:-----|
|  1   | [Reserve](#SmartArtQuickStyle.Reserve) |      |

## 属性

| 序号 | 名称                                           | 说明                                                                               |
|:----:|:-----------------------------------------------|:-----------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtQuickStyle.Application) | 获取一个 Application 对象，该对象代表 SmartArtQuickStyle 对象的容器应用程序。只读  |
|  2   | [Category](#SmartArtQuickStyle.Category)       | 检索与 SmartArt 快速样式关联的主类别名称。只读                                     |
|  3   | [Creator](#SmartArtQuickStyle.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtQuickStyle 对象的应用程序。只读 |
|  4   | [Description](#SmartArtQuickStyle.Description) | 检索 SmartArt 快速样式的说明。只读。                                               |
|  5   | [Id](#SmartArtQuickStyle.Id)                   | 检索关联的 SmartArt 快速样式的唯一 Id。只读。                                      |
|  6   | [Name](#SmartArtQuickStyle.Name)               | 检索 SmartArt 快速样式的字符串名称。只读。                                         |
|  7   | [Parent](#SmartArtQuickStyle.Parent)           | 返回调用对象。只读                                                                 |

## 成员方法

### SmartArtQuickStyle.Reserve

**语法**

**express.Reserve ()**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

## 成员属性

### SmartArtQuickStyle.Application

获取一个 **Application** 对象，该对象代表 **SmartArtQuickStyle** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Category

检索与 SmartArt 快速样式关联的主类别名称。只读

**语法**

**express.Category**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtQuickStyle** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Description

检索 SmartArt 快速样式的说明。只读。

**语法**

**express.Description**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Id

检索关联的 SmartArt 快速样式的唯一 Id。只读。

**语法**

**express.Id**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

**说明**

与此属性关联的 ID 区分大小写。

### SmartArtQuickStyle.Name

检索 SmartArt 快速样式的字符串名称。只读。

**语法**

**express.Name**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Parent

返回调用对象。只读

**语法**

**express.Parent**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtQuickStyle](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
