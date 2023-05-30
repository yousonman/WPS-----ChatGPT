# CustomProperty 对象

## 简介

代表标识符信息。标识符信息可用于 XML 的元数据。

## 说明

使用 Add 方法或 CustomProper ties 集合的 It em 属性可返回 CustomProperty 对象。

返回 CustomProperty 对象后，可在 Add 方法中使用 CustomPr operties 属性向工作表中添加元数据。

在本示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值。

``` JavaScript
function test() {
    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
    alert(cusProperties.Name + "\t" + cusProperties.Value)
}
```

## 方法

| 序号 | 名称                             | 说明       |
|:----:|:---------------------------------|:-----------|
|  1   | [Delete](#CustomProperty.Delete) | 删除对象。 |

## 属性

| 序号 | 名称                                       | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CustomProperty.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#CustomProperty.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Name](#CustomProperty.Name)               | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                    |
|  4   | [Parent](#CustomProperty.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [Value](#CustomProperty.Value)             | Borders.LineStyle 的同义词。                                                                                                                                                                                                    |

## 成员方法

### CustomProperty.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 CustomProperty 对象的变量。

**说明**

可删除自定义文档属性，但是无法删除内置文档属性。

## 成员属性

### CustomProperty.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomProperty 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
function test(){
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### CustomProperty.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CustomProperty 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### CustomProperty.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 CustomProperty 对象的变量。

### CustomProperty.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomProperty 对象的变量。

### CustomProperty.Value

**Borders.LineStyle** 的同义词。

**语法**

**express.Value**

\*express是一个代表 CustomProperty 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/CustomProperty](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
