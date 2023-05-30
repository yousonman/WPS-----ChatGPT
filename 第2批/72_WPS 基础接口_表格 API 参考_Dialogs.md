# Dialogs 对象

## 简介

ET 中所有 Dialog 对象的集合。

## 说明

每个 Dialog 对象代表一个内置对话框。不能在集合中新建或添加内置对话框。 Dialog 对象只能在 Show 方法中用来显示相应的对话框。

ET Visual Basic 对象库包含许多内置对话框的内置常量。每个常量的前缀均为“xlDialog”，随后即为对话框的名称。例如， “应用名称” 对话框常量为 xlDialogApplyNames ， “查找文件” 对话框常量为 xlDialogFindFile 。这些常量是 XlBuiltinD ialog 枚举类型的成员。

使用 Dialogs 属性可返回 Dialogs 集合。下例显示可用的 ET 内置对话框的个数。

``` JavaScript
alert(Application.Dialogs.Count)
```

使用 Dialogs ( *index* )（其中 *index* 是用于标识对话框的内置常量）可返回单个 Dialog 对象。下例运行内置的 “打开文件” 对话框。

``` JavaScript
let dlgAnswer = Application.Dialogs.Item(xlDialogOpen).Show()
```

## 方法

| 序号 | 名称                  | 说明                   |
|:----:|:----------------------|:-----------------------|
|  1   | [Item](#Dialogs.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Dialogs.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Dialogs.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Dialogs.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Dialogs.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Dialogs.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Dialogs 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型        | 说明                          |
|---------|-----------|-----------------|-------------------------------|
| *Index* | 必选      | XlBuiltInDialog | Variant。对象的名称或索引号。 |

**示例**

``` JavaScript
/* 此示例显示“打开”对话框，并选定“只读”选项*/
Application.Dialogs.Item(xlDialogOpen).Show(null, null, true)
```

## 成员属性

### Dialogs.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Dialogs 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
function test()
{
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### Dialogs.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Dialogs 对象的变量。

### Dialogs.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Dialogs 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Dialogs.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Dialogs 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Dialogs](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
