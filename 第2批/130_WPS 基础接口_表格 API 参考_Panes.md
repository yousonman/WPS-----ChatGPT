# Panes 对象

## 简介

指定的窗口中所有 Pane 对象的集合。

## 说明

Pane 对象只对于工作表和 ET 4.0 宏表存在。

使用 Panes 属性返回 Panes 集合。下例中，如果当前活动窗口包含若干窗格，则冻结这些窗格。

``` JavaScript
if(Application.ActiveWindow.Panes.Count > 1) { 
    Application.ActiveWindow.FreezePanes = true
}
```

使用 Panes ( *index* )（其中 *index* 是窗格索引号）可返回一个 Pane 对象。下例滚动 Sheet1 所在窗口中左上角的窗格。

``` JavaScript
function test(){
    Application.Worksheets.Item("sheet1").Activate()
    Application.Windows.Item(1).Panes.Item(1).LargeScroll(1)
}
```

## 方法

| 序号 | 名称                | 说明                   |
|:----:|:--------------------|:-----------------------|
|  1   | [Item](#Panes.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                              | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Panes.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Panes.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Panes.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Panes.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Panes.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Panes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明         |
|---------|-----------|----------|--------------|
| *Index* | 必选      | Long     | 对象的索引号 |

**示例**

``` JavaScript
function test(){  
  /*此示例拆分第一张工作表所在的窗口，然后滚动窗口左下角的窗格，直至第五行到达此窗格的顶部。*/
  Application.Worksheets.Item(1).Activate()
  Application.ActiveWindow.Split = true
  Application.ActiveWindow.Panes.Item(3).ScrollRow = 5
}
```

## 成员属性

### Panes.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Panes 对象的变量。

**示例**

``` JavaScript
function test(){  
  /*本示例显示一条有关创建 myObject 的应用程序的消息。*/
  if(Application.ActiveWorkbook.Application.Value == "ET") {
      alert("This is an ET Application object.")
  }                             
  else {
      alert("This is not an ET Application object.")
  }
}
```

### Panes.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Panes 对象的变量。

### Panes.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Panes 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Panes.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Panes 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Panes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
