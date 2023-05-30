# CustomView 对象

## 简介

代表自定义工作簿视图。

## 说明

CustomView 对象是 CustomViews 集合的成员。

使用 CustomViews ( *index* )（其中 *index* 为自定义视图的名称或索引号）可返回 CustomView 对象。

``` JavaScript
/*下例显示名为“Current Inventory”自定义视图。*/
Application.ActiveWorkbook.CustomViews.Item("Current Inventory").Show()
```

## 方法

| 序号 | 名称                         | 说明             |
|:----:|:-----------------------------|:-----------------|
|  1   | [Delete](#CustomView.Delete) | 删除对象。       |
|  2   | [Show](#CustomView.Show)     | 显示自定义视图。 |

## 属性

| 序号 | 名称                                         | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CustomView.Application)       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#CustomView.Creator)               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Name](#CustomView.Name)                     | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  4   | [Parent](#CustomView.Parent)                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [PrintSettings](#CustomView.PrintSettings)   | 如果在自定义视图中包含打印设置，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                       |
|  6   | [RowColSettings](#CustomView.RowColSettings) | 如果自定义视图包括对隐藏行和隐藏列（包括筛选信息）的设置，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                             |

## 成员方法

### CustomView.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 CustomView 对象的变量。

### CustomView.Show

显示自定义视图。

**语法**

**express.Show ()**

\*express是一个代表 CustomView 对象的变量。

## 成员属性

### CustomView.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomView 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
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

### CustomView.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CustomView 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### CustomView.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 CustomView 对象的变量。

### CustomView.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomView 对象的变量。

### CustomView.PrintSettings

如果在自定义视图中包含打印设置，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.PrintSettings**

\*express是一个代表 CustomView 对象的变量。

**示例**

``` JavaScript
/*本示例创建活动工作簿中自定义视图及其打印设置和行列设置列表*/
function test(){
Application.Worksheets.Item(1).Cells.Item(1,1).Value2 = "Name"
Application.Worksheets.Item(1).Cells.Item(1,2).Value2 = "Print Settings"
Application.Worksheets.Item(1).Cells.Item(1,3).Value2 = "RowColSettings"
let rw = 0
let myView = Application.ActiveWorkbook.CustomViews
for(let v = 1; v <= myView.Count;v++) {
    rw = rw + 1
    Application.Worksheets.Item(1).Cells.Item(rw, 1).Value2 = myView.Item(v).Name
    Application.Worksheets.Item(1).Cells.Item(rw, 2).Value2 = myView.Item(v).PrintSettings
    Application.Worksheets.Item(1).Cells.Item(rw, 3).Value2 = myView.Item(v).RowColSettings
}
}
```

### CustomView.RowColSettings

如果自定义视图包括对隐藏行和隐藏列（包括筛选信息）的设置，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.RowColSettings**

\*express是一个代表 CustomView 对象的变量。

**示例**

``` JavaScript
/*本示例创建活动工作簿中自定义视图及其打印设置和行列设置列表*/
function test(){
Application.Worksheets.Item(1).Cells.Item(1,1).Value2 = "Name"
Application.Worksheets.Item(1).Cells.Item(1,2).Value2 = "Print Settings"
Application.Worksheets.Item(1).Cells.Item(1,3).Value2 = "RowColSettings"
let rw = 0
let myView = Application.ActiveWorkbook.CustomViews
for(let v = 1; v <= myView.Count;v++) {
    rw = rw + 1
    Application.Worksheets.Item(1).Cells.Item(rw, 1).Value2 = myView.Item(v).Name
    Application.Worksheets.Item(1).Cells.Item(rw, 2).Value2 = myView.Item(v).PrintSettings
    Application.Worksheets.Item(1).Cells.Item(rw, 3).Value2 = myView.Item(v).RowColSettings
}
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/CustomView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
