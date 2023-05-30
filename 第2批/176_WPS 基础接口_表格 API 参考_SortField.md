# SortField 对象

## 简介

SortField 对象包含 Worksheet 、 Lists 和 AutoFilter 对象的所有排序信息。

## 说明

开发人员可以使用 BeforeSort 事件重写 ET 的默认行为，将自己的排序算法写入应用程序中。

## 方法

| 序号 | 名称                              | 说明                                            |
|:----:|:----------------------------------|:------------------------------------------------|
|  1   | [Delete](#SortField.Delete)       | 从 SortFields 集合中删除指定的 SortField 对象。 |
|  2   | [ModifyKey](#SortField.ModifyKey) | 修改字段中按其排序的键值。                      |
|  3   | [SetIcon](#SortField.SetIcon)     | 为 SortField 对象设置图标。                     |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                               |
|:----:|:--------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SortField.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#SortField.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [CustomOrder](#SortField.CustomOrder) | 指定对字段进行排序的自定义次序。可读/写 Variant 类型。                                                                                                             |
|  4   | [DataOption](#SortField.DataOption)   | 指定 SortField 对象所指定的区域中的文本排序方式。可读/写 XlSortDataOption 类型。                                                                                   |
|  5   | [Key](#SortField.Key)                 | 指定排序字段，该字段确定要排序的值。只读。                                                                                                                         |
|  6   | [Order](#SortField.Order)             | 确定关键字所指定的值的排序次序。可读/写。                                                                                                                          |
|  7   | [Parent](#SortField.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  8   | [Priority](#SortField.Priority)       | 指定排序字段的优先级。可读/写。                                                                                                                                    |
|  9   | [SortOn](#SortField.SortOn)           | 返回或设置要排序的单元格属性。可读/写 XlSortOn 类型。                                                                                                              |
|  10  | [SortOnValue](#SortField.SortOnValue) | 返回指定 SortField 对象执行排序的值。只读。                                                                                                                        |

## 成员方法

### SortField.Delete

从 **SortFields** 集合中删除指定的 **SortField** 对象。

**语法**

**express.Delete ()**

\*express是一个代表 SortField 对象的变量。

### SortField.ModifyKey

修改字段中按其排序的键值。

**语法**

**express.ModifyKey ( *Key* )**

\*express是一个代表 SortField 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明             |
|-------|-----------|----------|------------------|
| *Key* | 必选      | Range    | 指定要修改的键。 |

### SortField.SetIcon

为 **SortField** 对象设置图标。

**语法**

**express.SetIcon ( *Icon* )**

\*express是一个代表 SortField 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Icon* | 必选      | Icon     | 要设置的图标。 |

## 成员属性

### SortField.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SortField 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### SortField.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 SortField 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### SortField.CustomOrder

指定对字段进行排序的自定义次序。可读/写 **Variant** 类型。

**语法**

**express.CustomOrder**

\*express是一个代表 SortField 对象的变量。

### SortField.DataOption

指定 **SortField** 对象所指定的区域中的文本排序方式。可读/写 XlSortDataOption 类型。

**语法**

**express.DataOption**

\*express是一个代表 SortField 对象的变量。

**说明**

**xlSortDataOption** 可以是 **xlSortNormal** 或 **xlSortTextAsNumbers** 。

### SortField.Key

指定排序字段，该字段确定要排序的值。只读。

**语法**

**express.Key**

\*express是一个代表 SortField 对象的变量。

**说明**

关键字可以是区域名称（字符串）或 **Range** 对象。

### SortField.Order

确定关键字所指定的值的排序次序。可读/写。

**语法**

**express.Order**

\*express是一个代表 SortField 对象的变量。

### SortField.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SortField 对象的变量。

### SortField.Priority

指定排序字段的优先级。可读/写。

**语法**

**express.Priority**

\*express是一个代表 SortField 对象的变量。

### SortField.SortOn

返回或设置要排序的单元格属性。可读/写 **XlSortOn** 类型。

**语法**

**express.SortOn**

\*express是一个代表 SortField 对象的变量。

**说明**

|                                             |
|---------------------------------------------|
| XlSortOn 可以是下列 **XlSortOn** 常量之一： |
| **SortOnCellColor**                         |
| **SortOnFontColor**                         |
| **SortOnIcon**                              |
| **SortOnValues**                            |

### SortField.SortOnValue

返回指定 **SortField** 对象执行排序的值。只读。

**语法**

**express.SortOnValue**

\*express是一个代表 SortField 对象的变量。

**说明**

**SortOnValue** 可与值、单元格颜色、字体颜色和单元格图标一起使用。

**示例**

``` JavaScript
function test(){
  /*本示例对 B 列和 Sheet1 中的数据按字体颜色以升序排序*/
  Application.ActiveWorkbook.Worksheets.Item("Sheet1").Sort.SortFields.Clear()
  Application.ActiveWorkbook.Worksheets.Item("Sheet1").Sort.SortFields.Add(Range("B1:B25"), xlSortOnFontColor, xlAscending, " ", xlSortNormal).SortOnValue.Color.RGB = (0, 0, 0)
                                    
  let sor = Application.ActiveWorkbook.Worksheets.Item("Sheet1").Sort
  sor.SetRange(Range("A1:B25"))
  sor.Header = xlGuess
  sor.MatchCase = false
  sor.Orientation = xlTopToBottom
  sor.SortMethod = xlPinYin
  sor.Apply()
  
  /*单元格颜色*/
  let SortOn = xlSortOnCellColor
  SortOnValue.Color.RGB = (255, 255, 0)
  
  /*字体颜色*/
  SortOn = xlSortOnFontColor
  SortOnValue.Color.RGB = (255, 255, 0)
  
  /*图标*/
  let SortOn = xlSortOnIcon
  SortOnValue.Color.RGB = (255, 255, 0)
  SortField.SetIcon(ActiveWorkbook.IconSets(1).Item(3))
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/SortField](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
