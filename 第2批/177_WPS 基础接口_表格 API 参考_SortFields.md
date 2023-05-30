# SortFields 对象

## 简介

SortFields 集合是 SortField 对象的集合。开发人员可以使用该集合存储工作簿、列表和自动筛选的排序状态。

## 说明

该对象具有添加、计数、排序和删除 SortField 对象的属性。

``` JavaScript
function test(){
    Application.ActiveSheet.Sort.SortFields.Add(Range("A1"),null,xlDescending) 
    Application.ActiveSheet.Sort.SortFields.Add(Range("B1"),null,xlDescending)  
}
```

## 方法

| 序号 | 名称                       | 说明                                                                    |
|:----:|:---------------------------|:------------------------------------------------------------------------|
|  1   | [Add](#SortFields.Add)     | 创建新的排序字段，并返回一个 SortFields 对象。返回 SortField值          |
|  2   | [Clear](#SortFields.Clear) | 清除所有 Sort Fields 对象。                                             |
|  3   | [Item](#SortFields.Item)   | 返回一个 SortField 对象，该对象代表可以存储在工作簿中的项的集合。只读。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                               |
|:----:|:---------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SortFields.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#SortFields.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#SortFields.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Parent](#SortFields.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员方法

### SortFields.Add

创建新的排序字段，并返回一个 **SortFields** 对象。返回 SortField值

**语法**

**express.Add ( *Key* , *SortOn* , *Order* , *CustomOrder* , *DataOption* )**

\*express是一个代表 SortFields 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                           |
|---------------|-----------|----------|--------------------------------|
| *Key*         | 必选      | Range    | 指定用于排序的键值。           |
| *SortOn*      | 可选      | Variant  | 要进行排序的字段。             |
| *Order*       | 可选      | Variant  | 指定排序次序。                 |
| *CustomOrder* | 可选      | Variant  | 指定是否应使用自定义排序次序。 |
| *DataOption*  | 可选      | Variant  | 指定数据选项。                 |

### SortFields.Clear

清除所有 **Sort Fields** 对象。

**语法**

**express.Clear ()**

\*express是一个代表 SortFields 对象的变量。

### SortFields.Item

返回一个 **SortField** 对象，该对象代表可以存储在工作簿中的项的集合。只读。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SortFields 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明               |
|---------|-----------|----------|--------------------|
| *Index* | 必选      | Variant  | 排序字段的索引值。 |

## 成员属性

### SortFields.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SortFields 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### SortFields.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 SortFields 对象的变量。

### SortFields.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 SortFields 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### SortFields.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SortFields 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/SortFields](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
