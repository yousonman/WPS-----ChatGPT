# Areas 对象

## 简介

由选定区域内的多个子区域或连续单元格块组成的集合。

## 方法

| 序号 | 名称                | 说明                   |
|:----:|:--------------------|:-----------------------|
|  1   | [Item](#Areas.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                              | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Areas.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Areas.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Areas.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Areas.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Areas.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Areas 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 必选      | Long     | 对象的索引号。 |

**示例**

``` JavaScript
/* 此示例检查当前选定区域是否为多重选定区域，如果是，则清除其中的第一个子区域的内容*/
function test() {
  if(Application.Selection.Areas.Count != 1){
      Application.Selection.Areas.Item(1).Clear()
  }
}
```

## 成员属性

### Areas.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Areas 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
function test() {
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### Areas.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Areas 对象的变量。

**示例**

``` JavaScript
/* 此示例显示 Sheet1 上选定区域中的列数。此示例还将检测选定区域中是否包含多重选定区域，如果包含，则对多重选定区域中每一子区域进行循环*/
function test(){
    Application.Worksheets.Item("Sheet1").Activate()
    let iAreaCount = Application.Selection.Areas.Count

    if(iAreaCount <= 1){
        alert("The selection contains " + Application.Selection.Columns.Count + " columns.")
    }
    else{
        for(let i = 1; i <= iAreaCount; i++){
            alert("Area " + i + " of the selection contains " + Application.Selection.Areas.Item(i).Columns.Count + " columns.")
        }
    }
}
```

### Areas.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Areas 对象的变量。

### Areas.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Areas 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Areas](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
