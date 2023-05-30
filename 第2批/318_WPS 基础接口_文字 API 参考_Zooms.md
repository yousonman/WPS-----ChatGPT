# Zooms 对象

## 简介

Zoom 对象的集合，该集合代表每个视图（如大纲视图、普通视图或页面视图）的缩放选项。

## 方法

| 序号 | 名称                | 说明                              |
|:----:|:--------------------|:----------------------------------|
|  1   | [Item](#Zooms.Item) | 返回指定的 Zoom 对象。返回 Zoom值 |

## 属性

| 序号 | 名称                              | 说明                                                                         |
|:----:|:----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Zooms.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Creator](#Zooms.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  3   | [Parent](#Zooms.Parent)           | 返回一个 Object 类型值，该值代表指定 Zooms 对象的父对象。                    |

## 成员方法

### Zooms.Item

返回指定的 **Zoom** 对象。返回 Zoom值

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Zooms 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型   | 说明                 |
|---------|-----------|------------|----------------------|
| *Index* | 可选      | WdViewType | 指定的显示比例类型。 |

## 成员属性

### Zooms.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 Zooms 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Zooms.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Zooms 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 Creator 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Zooms.Parent

返回一个 **Object** 类型值，该值代表指定 **Zooms** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Zooms 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Zooms](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
