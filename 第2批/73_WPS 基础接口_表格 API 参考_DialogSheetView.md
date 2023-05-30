# DialogSheetView 对象

## 简介

代表工作簿中的当前 对话框 工作表视图。

## 说明

若要访问此对象，活动工作簿中必须有已开发的对话框工作表。如果没有对话框工作表，该对象的视图属性将返回空字符串值。

``` JavaScript
/*以下示例将为活动工作簿启用对话框工作表视图。*/
Application.Worksheets.Item("Sheet1").DialogSheetView.Visible = true
```

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                               |
|:----:|:--------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#DialogSheetView.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#DialogSheetView.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Parent](#DialogSheetView.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  4   | [Sheet](#DialogSheetView.Sheet)             | 返回指定的 DialogSheetView 对象的工作表名称。只读。                                                                                                                |

## 成员属性

### DialogSheetView.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 DialogSheetView 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### DialogSheetView.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DialogSheetView 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### DialogSheetView.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DialogSheetView 对象的变量。

### DialogSheetView.Sheet

返回指定的 **DialogSheetView** 对象的工作表名称。只读。

**语法**

**express.Sheet**

\*express是一个代表 DialogSheetView 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/DialogSheetView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
