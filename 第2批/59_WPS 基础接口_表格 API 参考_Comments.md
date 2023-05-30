# Comments 对象

## 简介

由单元格批注组成的集合。

## 说明

每个批注都由一个 Comment 对象代表。

使用 Comments 属性可返回 Comments 集合。下例隐藏第一张工作表中的所有批注。

``` JavaScript
function test() {
    let cmt = Application.Worksheets.Item(1).Comments
    for(let c = 1;c <= cmt.Count;c++) {
        cmt.Item(c).Visible = false
    }
} 
```

使用 AddCom ment 方法可在区域内添加批注。下例在第一张工作表的单元格 E5 中添加批注。

``` JavaScript
function test() {
    let myComment = Application.Worksheets.Item(1).Range("E5").AddComment()
    myComment.Visible = false
    myComment.Text("reviewed on " + Date())
} 
```

使用 Comments ( *index* )（其中 *index* 为批注号）可返回 Comments 集合中的单条批注。下例隐藏第一张工作表中的第二条批注。

``` JavaScript
Application.Worksheets.Item(1).Comments.Item(2).Visible = false
```

## 方法

| 序号 | 名称                   | 说明                                                     |
|:----:|:-----------------------|:---------------------------------------------------------|
|  1   | [Item](#Comments.Item) | 从集合中返回一个对象， 包含在集合中的一个 Comment 对象。 |

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Comments.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Comments.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Comments.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Comments.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Comments.Item

从集合中返回一个对象， 包含在集合中的一个 **Comment** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Comments 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 必选      | Long     | 对象的索引号。 |

**示例**

``` JavaScript
/* 本示例隐藏第二条批注*/
Application.Worksheets.Item(1).Comments.Item(2).Visible = false
```

## 成员属性

### Comments.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Comments 对象的变量。

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

### Comments.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Comments 对象的变量。

### Comments.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Comments 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Comments.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Comments 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Comments](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
