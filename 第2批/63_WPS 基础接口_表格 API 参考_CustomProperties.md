# CustomProperties 对象

## 简介

由代表附加信息的 CustomProperty 对象组成的集合，这些信息可用作 XML 的元数据。

## 说明

使用 Worksheet 对象的 CustomProperties 属性返回 CustomProperties 集合。

返回 CustomProperties 集合后，可根据选择向工作表和智能标记中添加元数据。

若要向工作表添加元数据，请在 Add 方法中使用 CustomProperties 属性。

下例演示了该功能。在此示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值。

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

| 序号 | 名称                           | 说明                                                                    |
|:----:|:-------------------------------|:------------------------------------------------------------------------|
|  1   | [Add](#CustomProperties.Add)   | 添加自定义属性信息，返回 一个代表自定义属性信息的 CustomProperty 对象。 |
|  2   | [Item](#CustomProperties.Item) | 从集合中返回一个对象。                                                  |

## 属性

| 序号 | 名称                                         | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CustomProperties.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#CustomProperties.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#CustomProperties.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#CustomProperties.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### CustomProperties.Add

添加自定义属性信息，返回 一个代表自定义属性信息的 **CustomProperty** 对象。

**语法**

**express.Add ( *Name* , *Value* )**

\*express是一个代表 CustomProperties 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明               |
|---------|-----------|----------|--------------------|
| *Name*  | 必选      | String   | 自定义属性的名称。 |
| *Value* | 必选      | Variant  | 自定义属性的值。   |

**示例**

``` JavaScript
/* 此示例向活动工作表添加标识符信息，并将其名称和值返回给用户*/
function test() {

    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
    alert(cusProperties.Name + "\t" + cusProperties.Value)

}
```

### CustomProperties.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 CustomProperties 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
/* 下例演示了该功能。在此示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值*/
function test() {

    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
        alert(cusProperties.Name + "\t" + cusProperties.Value)

}
```

## 成员属性

### CustomProperties.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomProperties 对象的变量。

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

### CustomProperties.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 CustomProperties 对象的变量。

### CustomProperties.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CustomProperties 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### CustomProperties.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomProperties 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/CustomProperties](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
