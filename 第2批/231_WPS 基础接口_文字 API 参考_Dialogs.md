# Dialogs 对象

## 简介

WPS 中的 Dialog 对象的集合。每个 Dialog 对象代表一个内置的 WPS 对话框。

## 方法

| 序号 | 名称                  | 说明                      |
|:----:|:----------------------|:--------------------------|
|  1   | [Item](#Dialogs.Item) | 返回 WPS 中的一个对话框。 |

## 属性

| 序号 | 名称                                | 说明                                                                         |
|:----:|:------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Dialogs.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Dialogs.Count)             | 返回一个 Number 类型的值，该值代表集合中的对话框数。只读。                   |
|  3   | [Creator](#Dialogs.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#Dialogs.Parent)           | 返回一个 Object 类型值，该值代表指定 Dialogs 对象的父对象。                  |

## 成员方法

### Dialogs.Item

返回 WPS 中的一个对话框。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Dialogs 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明               |
|---------|-----------|--------------|--------------------|
| *Index* | 必选      | WdWordDialog | 指定对话框的常量。 |

**示例**

``` JavaScript
/* 本示例显示“页面设置”对话框 */
function test(){
  Application.Dialogs.Item(wdDialogFileDocumentLayout).Display()
}
```

## 成员属性

### Dialogs.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Dialogs 对象的变量。

### Dialogs.Count

返回一个 **Number** 类型的值，该值代表集合中的对话框数。只读。

**语法**

**express.Count**

\*express是一个代表 Dialogs 对象的变量。

### Dialogs.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Dialogs 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Dialogs.Parent

返回一个 **Object** 类型值，该值代表指定 **Dialogs** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Dialogs 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Dialogs](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
