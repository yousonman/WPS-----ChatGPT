# Sort 对象

## 简介

代表数据区域的排序方式。

## 说明

以下示例在活动工作表中生成一个数据区域并对这些数据排序。

``` JavaScript
function test(){
    //Building data to sort on the active sheet.
    Range("A1").Value2 = "Name"
    Range("A2").Value2 = "Bill"
    Range("A3").Value2 = "Rod"
    Range("A4").Value2 = "John"
    Range("A5").Value2 = "Paddy"
    Range("A6").Value2 = "Kelly"
    Range("A7").Value2 = "William"
    Range("A8").Value2 = "Janet"
    Range("A9").Value2 = "Florence"
    Range("A10").Value2 = "Albert"
    Range("A11").Value2 = "Mary"
    alert("The list is out of order.  Hit Ok to continue..." + jsInformation)
                                    
    //Selecting a cell within the range.
    Range("A2").Select()
                                    
    //Applying sort.
    let sor = Application.ActiveWorkbook.Worksheets.Item(ActiveSheet.Name).Sort
    sor.SortFields.Clear()
    sor.SortFields.Add(Range("A2:A11"), xlSortOnValues, xlAscending, xlSortNormal)
    sor.SetRange(Range("A1:A11"))
    sor.Header = xlYes
    sor.MatchCase = false
    sor.Orientation = xlTopToBottom
    sor.SortMethod = xlPinYin
    sor.Apply()
    alert("Sort complete." + jsInformation)                              
}
```

## 方法

| 序号 | 名称                       | 说明                                       |
|:----:|:---------------------------|:-------------------------------------------|
|  1   | [Apply](#Sort.Apply)       | 根据当前应用的排序状态对区域进行排序。     |
|  2   | [SetRange](#Sort.SetRange) | 设置 Sort 对象的起始字符和结束字符的位置。 |

## 属性

| 序号 | 名称                             | 说明                                                                                                                                                               |
|:----:|:---------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Sort.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#Sort.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Header](#Sort.Header)           | 指定第一行是否包含标题信息。可读/写 XlYesNoGuess 类型。                                                                                                            |
|  4   | [MatchCase](#Sort.MatchCase)     | 如果设置为 True ，则执行区分大小写的排序，如果设置为 False ，则执行不区分大小写的排序。可读/写。                                                                   |
|  5   | [Orientation](#Sort.Orientation) | 指定排序方向。可读/写 XlSortOrientation 类型。                                                                                                                     |
|  6   | [Parent](#Sort.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  7   | [Rng](#Sort.Rng)                 | 返回要执行排序的值的区域。只读。                                                                                                                                   |
|  8   | [SortFields](#Sort.SortFields)   | 可用来在工作簿、列表和自动筛选上存储排序状态。只读 SortFields 类型。                                                                                               |
|  9   | [SortMethod](#Sort.SortMethod)   | 指定中文排序方法。可读/写 XlSortMethod 类型。                                                                                                                      |

## 成员方法

### Sort.Apply

根据当前应用的排序状态对区域进行排序。

**语法**

**express.Apply ()**

\*express是一个代表 Sort 对象的变量。

### Sort.SetRange

设置 **Sort** 对象的起始字符和结束字符的位置。

**语法**

**express.SetRange ( *Rng* )**

\*express是一个代表 Sort 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明                   |
|-------|-----------|----------|------------------------|
| *Rng* | 必选      | Range    | 指定 Sort 集合的区域。 |

**说明**

仅当对工作表区域应用排序时才能使用 **SetRange** ，当区域在表内时不能使用它。

## 成员属性

### Sort.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Sort 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### Sort.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Sort 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Sort.Header

指定第一行是否包含标题信息。可读/写 **XlYesNoGuess** 类型。

**语法**

**express.Header**

\*express是一个代表 Sort 对象的变量。

**说明**

默认值为 **xlNo** 。如果希望 ET 确定标题，可以指定 **xlGuess** 。

### Sort.MatchCase

如果设置为 **True** ，则执行区分大小写的排序，如果设置为 **False** ，则执行不区分大小写的排序。可读/写。

**语法**

**express.MatchCase**

\*express是一个代表 Sort 对象的变量。

### Sort.Orientation

指定排序方向。可读/写 **XlSortOrientation** 类型。

**语法**

**express.Orientation**

\*express是一个代表 Sort 对象的变量。

**说明**

|                                                              |
|--------------------------------------------------------------|
| XlSortOrientation 可以是下列 **XlSortOrientation** 常量之一: |
| **xlSortColumns**                                            |
| **xlSortRows**                                               |

### Sort.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Sort 对象的变量。

### Sort.Rng

返回要执行排序的值的区域。只读。

**语法**

**express.Rng**

\*express是一个代表 Sort 对象的变量。

### Sort.SortFields

可用来在工作簿、列表和自动筛选上存储排序状态。只读 **SortFields** 类型。

**语法**

**express.SortFields**

\*express是一个代表 Sort 对象的变量。

### Sort.SortMethod

指定中文排序方法。可读/写 **XlSortMethod** 类型。

**语法**

**express.SortMethod**

\*express是一个代表 Sort 对象的变量。

**说明**

|                                                  |
|--------------------------------------------------|
| XlSortMethod 可以是以下 **SortMethod** 常量之一: |
| **xlStroke**                                     |
| **xlPinYin**                                     |

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Sort](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
