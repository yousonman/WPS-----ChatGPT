# PluginStorage 对象

## 简介

通过 PluginStorage 对象可以实现一个 WPS 加载项中的多个网页之间的数据共享。 注意通过 PluginStorage 对象只能传递简单的数据类型，JSON 数据建议转成字符串后再进行共享。

## 说明

PluginStorage 对象是 Application 对象的子对象，每一个 WPS 加载项都只有一个 PluginStorage 实例。 注意 PluginStorage 对象不能持久化，所有数据只在关闭 WPS 加载项之前有效。 数据持久化可以将数据写入本地文件中，参考 FileSystem 对象 。

``` JavaScript
let ps = Application.PluginStorage
```

下列代码示例往 PluginStorage 中存储一个key为"count", 值为5的代码，如果 PluginStorage 中key为count的代码已存在，则会用新值进行覆盖。

``` JavaScript
Application.PluginStorage.setItem("count", 5)
```

以下示例枚举出 PluginStorage 中所有的key。

``` JavaScript
let ps = Application.PluginStorage
let itemCounts = ps.length; for(let i = 0;  i < itemCounts; ++i){     let itemKey = ps.Key(i) //如果要取到对应的value，使用ps.GetItem(itemKey)     console.log(itemKey) }
```

有关 PluginStorage 的详细方法和属性，可以参考 PluginStorage 方法和属性相关介绍

## 方法

| 序号 | 名称                                    | 说明                        |
|:----:|:----------------------------------------|:----------------------------|
|  1   | [clear](#PluginStorage.clear)           | 清空容器中的所有键值对。    |
|  2   | [getItem](#PluginStorage.getItem)       | 返回key对应的value。        |
|  3   | [key](#PluginStorage.key)               | 返回index对应的key。        |
|  4   | [removeItem](#PluginStorage.removeItem) | 删除容器中key对应的键值对。 |
|  5   | [setItem](#PluginStorage.setItem)       | 向容器中添加键值对。        |

## 属性

| 序号 | 名称                            | 说明                                |
|:----:|:--------------------------------|:------------------------------------|
|  1   | [length](#PluginStorage.length) | 返回PluginStorage容器中键值对的个数 |

## 成员方法

### PluginStorage.clear

清空容器中的所有键值对。

**语法**

**express.clear ()**

\*express是一个代表 PluginStorage 对象的变量。

**示例**

``` JavaScript
下列代码示例清空PluginStorage中的元素。
Application.PluginStorage.clear()
```

### PluginStorage.getItem

返回key对应的value。

**语法**

**express.getItem ( *key* )**

\*express是一个代表 PluginStorage 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型     | 说明                                                                                          |
|-------|-----------|--------------|-----------------------------------------------------------------------------------------------|
| *key* | 必选      | 整数、字符串 | key的数据类型只能是简单类型，具体来说可以是整数或字符串，推荐使用具有命名意义的字符串作为key. |

**示例**

``` JavaScript
下列代码获取key为count的value并在控制台输出。
let value = Application.PluginStorage.getItem("count")
if (value){
    console.log(value)
}
```

### PluginStorage.key

返回index对应的key。

**语法**

**express.key ( *index* )**

\*express是一个代表 PluginStorage 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                      |
|---------|-----------|----------|-------------------------------------------------------------------------------------------|
| *index* | 必选      | 整数     | index的起始值从0开始，常见的用法是先用length属性取到元素个数，然后用此方法来得到对应的key |

**示例**

``` JavaScript
下列代码获取容器中第1个元素的key并在控制台输出。
let count = Application.PluginStorage.length
if (count > 0){
    let firstkey = Application.PluginStorage.key(0)
    console.log(firstkey)                           
}
```

### PluginStorage.removeItem

删除容器中key对应的键值对。

**语法**

**express.removeItem ( *key* )**

\*express是一个代表 PluginStorage 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型     | 说明                                          |
|-------|-----------|--------------|-----------------------------------------------|
| *key* | 必选      | 整数、字符串 | 如果容器中找不到key对应的键值对，则什么都不做 |

**示例**

``` JavaScript
下列代码用于删除key为"count"的键值对。
Application.PluginStorage.removeItem("count")
```

### PluginStorage.setItem

向容器中添加键值对。

**语法**

**express.setItem ( *key* , *value* )**

\*express是一个代表 PluginStorage 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明                                                                                          |
|---------|-----------|--------------|-----------------------------------------------------------------------------------------------|
| *key*   | 必选      | 整数、字符串 | key的数据类型只能是简单类型，具体来说可以是整数或字符串，推荐使用具有命名意义的字符串作为key. |
| *value* | 必选      | Variant      | {description}                                                                                 |

**示例**

``` JavaScript
下列代码用于将key为"count"，值为5的键值对存入容器。
Application.PluginStorage.setItem("count", 5)
```

## 成员属性

### PluginStorage.length

返回PluginStorage容器中键值对的个数

**语法**

**express.length**

\*express是一个代表 PluginStorage 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/加载项数据/PluginStorage](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# VLayout 对象

## 简介

将内部的控件按照垂直方向排布，一行一个。

## 说明

将内部的控件按照垂直方向排布，一行一个。

## 方法

| 序号 | 名称                  | 说明                       |
|:----:|:----------------------|:---------------------------|
|  1   | [Move](#VLayout.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                              | 说明                                                 |
|:----:|:--------------------------------------------------|:-----------------------------------------------------|
|  1   | [Enabled](#VLayout.Enabled)                       | 您可以通过该属性设置或返回是否启用该控件             |
|  2   | [Height](#VLayout.Height)                         | 获取或设置指定对象的高度                             |
|  3   | [LayoutBottomMargin](#VLayout.LayoutBottomMargin) | 获取或设置容器内部的控件与容器底部的距离             |
|  4   | [LayoutLeftMargin](#VLayout.LayoutLeftMargin)     | 获取或设置容器内部的控件与容器左边缘的距离           |
|  5   | [LayoutRightMargin](#VLayout.LayoutRightMargin)   | 获取或设置容器内部的控件与容器右边缘的距离           |
|  6   | [LayoutSpacing](#VLayout.LayoutSpacing)           | 获取或设置容器内部多个控件之间的水平间距             |
|  7   | [LayoutTopMargin](#VLayout.LayoutTopMargin)       | 获取或设置容器内部的控件与容器顶部的距离             |
|  8   | [Left](#VLayout.Left)                             | 您可以通过该属性指定控件在对象或报表中的位置         |
|  9   | [Name](#VLayout.Name)                             | 您可以通过该属性指定或确定表示对象名称的字符串，只读 |
|  10  | [Top](#VLayout.Top)                               | 您可以通过该属性指定控件在对象或报表中的位置         |
|  11  | [Visible](#VLayout.Visible)                       | 返回或设置该对象是否可见                             |
|  12  | [Width](#VLayout.Width)                           | 获取或设置指定对象的宽度                             |

## 成员方法

### VLayout.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 VLayout 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                               |
|----------|-----------|----------|------------------------------------|
| *Left*   | 必选      | Number   | 对象移动后相对于窗口的左边缘的距离 |
| *Top*    | 可选      | Number   | 对象移动后相对于窗口的上边缘的距离 |
| *Width*  | 可选      | Number   | 对象移动后的预期宽度               |
| *Height* | 可选      | Number   | 对象移动后的预期长度               |

**说明**

将对象移动到参数指定的位置

**示例**

``` JavaScript
/*将VLayout1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.VLayout1.Move(100,100,100,100)
}
```

## 成员属性

### VLayout.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 VLayout 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.VLayout1.Enabled = true
}
```

### VLayout.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将VLayout1控件的高度设置为100*/
function func1()
{
    UserForm1.VLayout1.Height = 100
}
```

### VLayout.LayoutBottomMargin

获取或设置容器内部的控件与容器底部的距离

**语法**

**express.LayoutBottomMargin**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器底部的距离

**示例**

``` JavaScript
/*将VLayout1中控件到容器底部的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutBottomMargin = 10
}
```

### VLayout.LayoutLeftMargin

获取或设置容器内部的控件与容器左边缘的距离

**语法**

**express.LayoutLeftMargin**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器左边缘的距离

**示例**

``` JavaScript
/*将VLayout1中控件到容器左边缘的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutLeftMargin = 10
}
```

### VLayout.LayoutRightMargin

获取或设置容器内部的控件与容器右边缘的距离

**语法**

**express.LayoutRightMargin**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器右边缘的距离

**示例**

``` JavaScript
/*将VLayout1中控件到容器右边缘的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutRightMargin = 10
}
```

### VLayout.LayoutSpacing

获取或设置容器内部多个控件之间的水平间距

**语法**

**express.LayoutSpacing**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部多个控件之间的水平间距

**示例**

``` JavaScript
/*将VLayout1中多个控件之间的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutSpacing = 10
}
```

### VLayout.LayoutTopMargin

获取或设置容器内部的控件与容器顶部的距离

**语法**

**express.LayoutTopMargin**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器顶部的距离

**示例**

``` JavaScript
/*将VLayout1中控件到容器顶部的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutTopMargin = 10
}
```

### VLayout.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 VLayout 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将VLayout1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.VLayout1.Left = 100
}
```

### VLayout.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 VLayout 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### VLayout.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 VLayout 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将VLayout1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.VLayout1.Top = 100
}
```

### VLayout.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 VLayout 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将VLayout1设置为可见*/
function func1()
{
    UserForm1.VLayout1.Visible = true
}
```

### VLayout.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将VLayout1的宽度设置为100*/
function func1()
{
    UserForm1.VLayout1.Width = 100
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/VLayout](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# AutoRecover 对象

## 简介

代表工作簿的自动恢复功能。

## 说明

AutoRecover 对象的属性确定备份所有文件的路径和时间间隔。

使用 Application 对象的 AutoRecover 属性可返回 AutoRecover 对象。

使用 AutoRecover 对象的 Path 属性可设置“自动恢复”文件的保存路径。

示例

``` JavaScript
/*以下示例将“自动恢复”文件的路径设置为驱动器 C。*/
function test(){
    Application.AutoRecover.Path = "C:\\"
}
```

使用 AutoRecover 对象的 Time 属性可设置备份所有文件的时间间隔。

| 注释                    |
|-------------------------|
| Time 属性的单位是分钟。 |

``` JavaScript
function SetTime(){
    Application.AutoRecover.Time = 5
}
```

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#AutoRecover.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#AutoRecover.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Enabled](#AutoRecover.Enabled)         | 如果启用对象，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                               |
|  4   | [Parent](#AutoRecover.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [Path](#AutoRecover.Path)               | 返回或设置一个 String 值，它代表 ET 将要存储“自动恢复”临时文件的位置的完整路径。                                                                                                                                                |
|  6   | [Time](#AutoRecover.Time)               | 设置或返回 AutoRecover 对象的时间间隔。允许的值为从 1 到 120 分钟的整数。默认值为 10 分钟。 Long 类型，可读写。                                                                                                                 |

## 成员属性

### AutoRecover.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 AutoRecover 对象的变量。

### AutoRecover.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 AutoRecover 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### AutoRecover.Enabled

如果启用对象，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Enabled**

\*express是一个代表 AutoRecover 对象的变量。

### AutoRecover.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 AutoRecover 对象的变量。

### AutoRecover.Path

返回或设置一个 **String** 值，它代表 ET 将要存储“自动恢复”临时文件的位置的完整路径。

**语法**

**express.Path**

\*express是一个代表 AutoRecover 对象的变量。

**示例**

``` JavaScript
/*此示例将“自动恢复”文件的路径设置为驱动器 C。*/
function test(){
    Application.AutoRecover.Path = "C:\\"
}
```

### AutoRecover.Time

设置或返回 **AutoRecover** 对象的时间间隔。允许的值为从 1 到 120 分钟的整数。默认值为 10 分钟。 **Long** 类型，可读写。

**语法**

**express.Time**

\*express是一个代表 AutoRecover 对象的变量。

**说明**

输入的小数值将舍入为最接近的整数。例如，输入的 5.5 与 6 等价。

如果输入有效区域以外的时间值，则 ET 将使用上一次使用的时间值。

**示例**

``` JavaScript
/*下例将“自动恢复”的时间间隔设置为 5 分钟，并通知用户。*/
function test(){
    Application.AutoRecover.Time = 5
    alert("The AutoRecover time interval is set at " + Application.AutoRecover.Time + " minutes.")
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AutoRecover](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Chart 对象

## 简介

代表工作簿中的图表。

## 说明

此图表既可以是嵌入的图表（包含在 ChartObject 对象中），也可以是单独的图表工作表。

示例部分中描述了以下用于返回 Chart 对象的属性和方法：

- Charts 方法
- ActiveChart 属性
- ActiveSheet 属性

## 示例

Charts 集合包含工作簿中每个图表工作表的 Chart 对象。使用 Charts ( *index* ) 可以返回单个 Chart 对象，其中 index 为图表工作表的索引号或名称。图表索引号表示图表工作表在工作簿标签栏上的位置。 *Charts(1)* 是工作簿中第一个（最左边的）图表； *Charts(Charts.Count)* 是最后一个（最右边的）图表。所有图表工作表均包括在索引计数中，即便是隐藏工作表也是如此。图表工作表名称显示在图表工作簿标签上。您可以使用 Name 属性设置或返回图表名称。本示例将更改第一张图表工作表中第一个系列的颜色。

``` JavaScript
Charts.Item(1).SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed
```

本示例将名为“Sales”的图表移至活动工作簿的尾部。

``` JavaScript
Charts.Item("Sales").Move(null,Sheets.Item(Sheets.Count))
```

Chart 对象也是 Sheets 集合的成员，此集合包含工作簿中的所有工作表（图表工作表和工作表）。使用 Sheets ( *index* ) 可以返回单个工作表，其中 *index* 是工作表索引号或名称。

当图表是活动对象时，您可以使用 ActiveChart 属性引用它。如果用户选择了图表工作表，或者用 Chart 对象的 Activate 方法或 ChartObject 对象的 Activate 方法激活了它，则该图表工作表处于活动状态。本示例将激活图表工作表 1，然后设置图表类型和名称。

``` JavaScript
Charts.Item(1).Activate()
let chart = ActiveChart
    chart.Type = xlLine
    chart.HasTitle = true
    chart.ChartTitle.Text = "January Sales"
```

如果用户选择了嵌入图表，或者用 Activate 方法激活了包含该嵌入图表的 ChartObject 对象，则该嵌入图表处于活动状态。本示例将激活工作表 1 上的嵌入图表 1，然后设置图表类型和名称。请注意，在激活了嵌入图表之后，本示例中的代码与前一个示例中的代码相同。通过使用 ActiveChart 属性，您可以编写能够引用嵌入图表或图表工作表（视哪一个处于活动状态而定）的 Visual Basic 代码。

``` JavaScript
Worksheets.Item(1).ChartObjects(1).Activate()
ActiveChart.ChartType = xlLine
ActiveChart.HasTitle = true
ActiveChart.ChartTitle.Text = "January Sales"
```

当图表工作表为活动工作表时，可以使用 ActiveSheet 属性来引用它。本示例使用 Activate 方法激活名为 Chart1 的图表工作表，然后将图表中系列 1 的内部颜色设置为蓝色。

``` JavaScript
Charts.Item("chart1").Activate()
ActiveSheet.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbBlue
```

## 方法

| 序号 | 名称                                                | 说明                                                                                                                                                                |
|:----:|:----------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Activate](#Chart.Activate)                         | 使当前图表成为活动图表。                                                                                                                                            |
|  2   | [ApplyChartTemplate](#Chart.ApplyChartTemplate)     | 将标准图表类型或自定义图表类型应用于图表。                                                                                                                          |
|  3   | [ApplyDataLabels](#Chart.ApplyDataLabels)           | 将数据标签应用到图表中的所有系列。                                                                                                                                  |
|  4   | [ApplyLayout](#Chart.ApplyLayout)                   | 应用功能区中显示的版式。                                                                                                                                            |
|  5   | [Axes](#Chart.Axes)                                 | 返回一个代表图表上单个坐标轴或坐标轴集合的某个对象。                                                                                                                |
|  6   | [ChartGroups](#Chart.ChartGroups)                   | 返回一个对象，该对象表示图表中单个图表组（ ChartGroup 对象）或所有图表组的集合（ ChartGroups 对象）。返回的集合中包括每种类型的图表组。                             |
|  7   | [ChartObjects](#Chart.ChartObjects)                 | 返回一个对象，它代表工作表上的一个嵌入式图表（ ChartObject 对象）或所有嵌入式图表的集合（ ChartObjects 对象）。                                                     |
|  8   | [ChartWizard](#Chart.ChartWizard)                   | 修改给定图表的属性。可使用本方法快速设置图表的格式，而不必逐个设置所有属性。本方法是非交互式的，并且仅更改指定的属性。                                              |
|  9   | [CheckSpelling](#Chart.CheckSpelling)               | 检查对象的拼写。                                                                                                                                                    |
|  10  | [ClearToMatchStyle](#Chart.ClearToMatchStyle)       | 清除图表元素格式以改为自动格式。                                                                                                                                    |
|  11  | [Copy](#Chart.Copy)                                 | 将工作表复制到工作簿的另一位置。                                                                                                                                    |
|  12  | [CopyPicture](#Chart.CopyPicture)                   | 将所选对象作为图片复制到剪贴板。                                                                                                                                    |
|  13  | [Delete](#Chart.Delete)                             | 删除对象。                                                                                                                                                          |
|  14  | [Evaluate](#Chart.Evaluate)                         | 将一个 ET 名称转换为一个对象或者一个值。                                                                                                                            |
|  15  | [Export](#Chart.Export)                             | 以图形格式导出图表。                                                                                                                                                |
|  16  | [ExportAsFixedFormat](#Chart.ExportAsFixedFormat)   | 导出为指定格式的文件。                                                                                                                                              |
|  17  | [GetChartElement](#Chart.GetChartElement)           | 返回指定的 X 坐标和 Y 坐标上图表元素的信息。本方法稍有与众不同之处：调用时只须指定前两个参数，在本方法执行期间，ET 为其余参数赋值，本方法返回后应检验这些参数的值。 |
|  18  | [Location](#Chart.Location)                         | 将图表移动到新位置。                                                                                                                                                |
|  19  | [Move](#Chart.Move)                                 | 将图表移到工作簿的另一位置。                                                                                                                                        |
|  20  | [OLEObjects](#Chart.OLEObjects)                     | 返回一个对象，它代表图表或工作表上的单个 OLE 对象 ( OLEObject ) 或所有 OLE 对象的集合（ OLEObjects 集合）。只读。                                                   |
|  21  | [Paste](#Chart.Paste)                               | 将剪贴板中的图表数据粘贴到指定的图表中。                                                                                                                            |
|  22  | [PrintOut](#Chart.PrintOut)                         | 打印对象。                                                                                                                                                          |
|  23  | [PrintPreview](#Chart.PrintPreview)                 | 按对象打印后的外观效果显示对象的预览。                                                                                                                              |
|  24  | [Protect](#Chart.Protect)                           | 保护图表使其不被修改。                                                                                                                                              |
|  25  | [Refresh](#Chart.Refresh)                           | 立即重新绘制指定的图表。                                                                                                                                            |
|  26  | [SaveAs](#Chart.SaveAs)                             | 将对图表或工作表的更改保存到另一不同文件中。                                                                                                                        |
|  27  | [SaveChartTemplate](#Chart.SaveChartTemplate)       | 向可用图表模板的列表添加自定义图表模板。                                                                                                                            |
|  28  | [Select](#Chart.Select)                             | 选择对象。                                                                                                                                                          |
|  29  | [SeriesCollection](#Chart.SeriesCollection)         | 返回一个对象，它代表图表或图表组中的一个系列（ Series 对象）或所有系列的集合（ SeriesCollection 集合）。                                                            |
|  30  | [SetBackgroundPicture](#Chart.SetBackgroundPicture) | 为图表设置背景图形。                                                                                                                                                |
|  31  | [SetDefaultChart](#Chart.SetDefaultChart)           | 指定 ET 新建图表时使用的图表模板的名称。                                                                                                                            |
|  32  | [SetElement](#Chart.SetElement)                     | 设置图表上的图表元素。可读/写 MsoChartElementType 类型。                                                                                                            |
|  33  | [SetSourceData](#Chart.SetSourceData)               | 为指定图表设置源数据区域。                                                                                                                                          |
|  34  | [Unprotect](#Chart.Unprotect)                       | 取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。                                                                                        |

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.AutoScaling">AutoScaling</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 ET 对三维图表进行缩放，使之与等效的二维图表的大小相近，则该属性值为 True 。 RightAngleAxes 属性必须为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.BackWall">BackWall</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Walls 对象，该对象允许用户单独对三维图表的背景墙进行格式设置。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.BarShape">BarShape</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置用于三维条形图或柱形图的形状。 XlBarShape 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ChartArea">ChartArea</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 ChartArea 对象，该对象表示图表的整个图表区。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ChartStyle">ChartStyle</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置图表的图表样式。可读/写 Variant 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ChartTitle">ChartTitle</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 ChartTitle 对象，该对象表示指定图表的标题。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ChartType">ChartType</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置图表类型。 XlChartType 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.CodeName">CodeName</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回对象的代码名。 String 型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.DataTable">DataTable</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 DataTable 对象，该对象表示图表的模拟运算表。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.DepthPercent">DepthPercent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置三维图表的深度，以图表宽度的百分比表示（有效范围从 20% 到 2000%）。 Long 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.DisplayBlanksAs">DisplayBlanksAs</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置在图表中绘制空白单元格的方式。可以是 XlDisplayBlanksAs 常量之一。 Long 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Elevation">Elevation</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置三维图表视图的仰角（以角度为单位）。 Long 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Floor">Floor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Floor 对象，该对象表示三维图表的基底。只读。</p>
<p>返回值类型为 Floor。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.GapDepth">GapDepth</a></td>
<td style="text-align: left;" data-valian="middle"><p>以数据标志宽度的形式返回或设置三维图表中数据系列之间的距离，本属性的值必须在 0 和 500 之间。 Long 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.HasAxis">HasAxis</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置图表上显示的坐标轴。 Variant 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.HasDataTable">HasDataTable</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果图表有模拟运算表，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.HasLegend">HasLegend</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果图表有图例，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.HasTitle">HasTitle</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果坐标轴或图表有可见标题，则为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">21</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.HeightPercent">HeightPercent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置三维图表的高度，以图表宽度的百分比表示（有效范围从 5% 到 500%）。 Long 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">22</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Hyperlinks">Hyperlinks</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Hyperlinks 集合，它代表图表的超链接。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">23</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Index">Index</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">24</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Legend">Legend</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Legend 对象，该对象表示指定图表的图例。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">25</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.MailEnvelope">MailEnvelope</a></td>
<td style="text-align: left;" data-valian="middle"><p>代表文档的电子邮件标题。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">26</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Name">Name</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 值，它代表对象的名称。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">27</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Next">Next</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回代表下一个工作表的 Worksheet 对象。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">28</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.PageSetup">PageSetup</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 PageSetup 对象，它包含用于指定对象的所有页面设置。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">29</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">30</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Perspective">Perspective</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Long 值，它代表三维图表视图的透视系数。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">31</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.PivotLayout">PivotLayout</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 PivotLayout 对象，该对象表示数据透视表中字段的位置以及数据透视图中坐标轴的位置。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">32</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.PlotArea">PlotArea</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 PlotArea 对象，该对象表示图表的绘图区。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">33</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.PlotBy">PlotBy</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置行或列在图表中作为数据系列使用的方式。可为以下 XlRowCol 常量之一： xlColumns 或 xlRows 。 Long 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">34</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.PlotVisibleOnly">PlotVisibleOnly</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果仅绘制可见单元格，则该值为 True 。如果可见单元格和隐藏单元格都绘制，则该值为 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">35</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Previous">Previous</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回代表下一个工作表的 Worksheet 对象。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">36</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.PrintedCommentPages">PrintedCommentPages</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回将为当前图表打印的批注页的数量。只读。</p>
<p>返回类型类型为 Long。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">37</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ProtectContents">ProtectContents</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果工作表内容是受保护的，则为 True 。对于图表，这样会保护整个图表。要打开内容保护，请使用 Protect 方法，并将 <em>Contents</em> 参数设置为 True 。 Boolean 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">38</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ProtectData">ProtectData</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果用户不能更改系列公式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">39</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ProtectDrawingObjects">ProtectDrawingObjects</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果形状是受保护的，则为 True 。要打开形状保护，请使用 Protect 方法，并将 <em>DrawingObjects</em> 参数设置为 True 。 Boolean 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">40</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ProtectFormatting">ProtectFormatting</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果用户不能更改格式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">41</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ProtectSelection">ProtectSelection</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不能选定图表元素，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">42</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ProtectionMode">ProtectionMode</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果启用了用户界面专用保护，则为 True 。要打开用户界面保护，请使用 Protect 方法，并将 <em>UserInterfaceOnly</em> 参数设置为 True 。 Boolean 类型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">43</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.RightAngleAxes">RightAngleAxes</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果图表的坐标轴为直角，并与图表的转角或仰角无关，则该值为 True 。仅应用于三维折线图、柱形图和条形图。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">44</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Rotation">Rotation</a></td>
<td style="text-align: left;" data-valian="middle"><p>以度为单位返回或设置三维图表视图的转角（绘图区绕 Z 轴的转角）。此属性的取值必须介于 0 到 360 之间，三维条形图除外（从 0 到 44 之间）。默认值是 20。仅适用于三维图表。 Variant 型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">45</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Shapes">Shapes</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Shapes 集合，它代表图表工作表上所有的形状。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">46</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ShowAllFieldButtons">ShowAllFieldButtons</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置是否在数据透视图上显示所有字段按钮。可读/写。</p>
<p>返回值类型为 Boolean。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">47</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ShowAxisFieldButtons">ShowAxisFieldButtons</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置是否要在数据透视图上显示坐标轴字段按钮。可读/写。</p>
<p>返回值类型为 Boolean。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">48</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ShowDataLabelsOverMaximum">ShowDataLabelsOverMaximum</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个布尔值，该值表示在数值大于数值轴上的最大值时是否显示数据标签。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">49</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ShowLegendFieldButtons">ShowLegendFieldButtons</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置是否要在数据透视图上显示图例字段按钮。可读/写。</p>
<p>返回值类型为 Boolean。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">50</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ShowReportFilterFieldButtons">ShowReportFilterFieldButtons</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置是否要在数据透视图上显示报表筛选字段按钮。可读/写。</p>
<p>返回值类型为 Boolean。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">51</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.ShowValueFieldButtons">ShowValueFieldButtons</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置是否要在数据透视图上显示值字段按钮。可读/写。</p>
<p>返回值类型为 Boolean。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">52</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.SideWall">SideWall</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Walls 对象，该对象允许用户单独对三维图表的侧面墙进行格式设置。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">53</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Tab">Tab</a></td>
<td style="text-align: left;" data-valian="middle"><p>为图表返回一个 Tab 对象。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">54</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Visible">Visible</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 XlSheetVisibility 值，它确定对象是否可见。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">55</td>
<td style="text-align: left;" data-valian="middle"><a href="#Chart.Walls">Walls</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Walls 对象，该对象表示三维图表的背景墙。只读。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### Chart.Activate

使当前图表成为活动图表。

**语法**

**express.Activate ()**

\*express是一个代表 Chart 对象的变量。

### Chart.ApplyChartTemplate

将标准图表类型或自定义图表类型应用于图表。

**语法**

**express.ApplyChartTemplate ()**

\*express是一个代表 Chart 对象的变量。

**说明**

此方法不支持采用原图表或组合图的常量。.

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**           |
|------------|---------------|--------------|--------------------|
| *Filename* | 必选          | **String**   | 图表模板的文件名。 |

### Chart.ApplyDataLabels

将数据标签应用到图表中的所有系列。

**语法**

**express.ApplyDataLabels ( *Type* , *LegendKey* , *AutoText* , *HasLeaderLines* , *ShowSeriesName* , *ShowCategoryName* , *ShowValue* , *ShowPercentage* , *ShowBubbleSize* , *Separator* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型         | 说明                                                         |
|--------------------|-----------|------------------|--------------------------------------------------------------|
| *Type*             | 可选      | XlDataLabelsType | 要应用的数据标签的类型。                                     |
| *LegendKey*        | 可选      | Variant          | 如果为 True，则在数据点旁边显示图例项标示。默认值为 False。  |
| *AutoText*         | 可选      | Variant          | 如果对象根据内容自动生成相应的文字，则为 True。              |
| *HasLeaderLines*   | 可选      | Variant          | 对于 Chart 和 Series 对象，如果数据系列有引导线，则为 True。 |
| *ShowSeriesName*   | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的系列名称。                   |
| *ShowCategoryName* | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的分类名称。                   |
| *ShowValue*        | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的值。                         |
| *ShowPercentage*   | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的百分比。                     |
| *ShowBubbleSize*   | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的气泡大小。                   |
| *Separator*        | 可选      | Variant          | 数据标签的分隔符。                                           |

**说明**

**注释：** When using the **Separator** parameter. If you use a string, you will get a string as the separator. If you use **xlDataLabelSeparatorDefault** (= 1), you will get the default data label separator, which is either a comma or a newline, depending on the data label. When a value of "1" is returned, it indicates that the user has not changed the default separator which is a comma ",". You can also pass a value of "1" to change the separator back to the default separator. The chart must first be active before you can access the data labels programmatically; otherwise a run-time error occurs.

|                                                                                              |
|----------------------------------------------------------------------------------------------|
| **XlDataLabelsType** 可为以下 **XlDataLabelsType** 常量之一。                                |
| **xlDataLabelsShowBubbleSizes** 。数据标签的气泡尺寸。                                       |
| **xlDataLabelsShowLabelAndPercent** 。占总数的百分比及数据点所属的分类。仅用于饼图和圆环图。 |
| **xlDataLabelsShowPercent** 。占总数的百分比。仅用于饼图和圆环图。                           |
| **xlDataLabelsShowLabel** 。数据点所属的分类。                                               |
| **xlDataLabelsShowNone** 。无数据标签。                                                      |
| **xlDataLabelsShowValue** 。默认。数据点的值（若未指定此参数，则假定为此值）。               |

**示例**

``` JavaScript
/*本示例对 Chart1 上的第一个数据系列应用分类标签。*/
Application.Charts.Item("Chart1").SeriesCollection(1).ApplyDataLabels(xlDataLabelsShowLabel)
```

### Chart.ApplyLayout

应用功能区中显示的版式。

**语法**

**express.ApplyLayout ( *Layout* , *ChartType* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型    | 说明                                          |
|-------------|-----------|-------------|-----------------------------------------------|
| *Layout*    | 必选      | Long        | 指定版式类型。版式类型由 1 到 10 的数字表示。 |
| *ChartType* | 可选      | XlChartType | 图表的类型。                                  |

**说明**

在当前图表类型上使用版式时，会将介于 1 到 10 之间的数字应用于该图表类型。还可以将某一图表类型的版式应用于另一个图表类型。例如，可以将折线图中的可用版式应用于柱形图。版式只会添加可用于该特定图表类型的图表元素。

### Chart.Axes

返回一个代表图表上单个坐标轴或坐标轴集合的某个对象。

**语法**

**express.Axes ( *Type* , *AxisGroup* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型    | 说明                                                                                                                     |
|-------------|-----------|-------------|--------------------------------------------------------------------------------------------------------------------------|
| *Type*      | 可选      | Variant     | 指定要返回的坐标轴。可为以下 XlAxisType 常量之一：xlValue、xlCategory 或 xlSeriesAxis（xlSeriesAxis 仅对三维图表有效）。 |
| *AxisGroup* | 可选      | XlAxisGroup | 指定坐标轴组。如果省略该参数，则使用主坐标轴组。三维图表仅有一个坐标轴组。                                               |

**示例**

``` JavaScript
/*本示例为 Chart1 的分类轴添加轴标签。*/
function test()
{
    let axes = Application.Charts.Item("Chart1").Axes(xlCategory)
    axes.HasTitle = true
    axes.AxisTitle.Text = "July Sales"
}

/*本示例将关闭 Chart1 中分类轴的主要网格线。*/
Application.Charts.Item("Chart1").Axes(xlCategory).HasMajorGridlines = false

/*本示例将关闭 Chart1 中所有坐标轴的网格线。*/
function test()
{
    for(let i = 1; i <= Application.Charts.Item("Chart1").Axes().Count; i++ ){
        Application.Charts.Item("Chart1").Axes().Item(i).HasMajorGridlines = false
        Application.Charts.Item("Chart1").Axes().Item(i).HasMinorGridlines = false
    }
}
```

### Chart.ChartGroups

返回一个对象，该对象表示图表中单个图表组（ **ChartGroup** 对象）或所有图表组的集合（ **ChartGroups** 对象）。返回的集合中包括每种类型的图表组。

**语法**

**express.ChartGroups ( *Index* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明         |
|---------|-----------|----------|--------------|
| *Index* | 可选      | Variant  | 图表组编号。 |

**示例**

``` JavaScript
/*本示例显示 Chart1 中第一个图表组的涨跌柱线，并对其颜色进行设置。本示例应在包含两个有一个或多个交叉数据点的系列的二维折线图上运行。*/
function test()
{
    let chart = Application.Charts.Item("Chart1").ChartGroups(1)
    chart.HasUpDownBars = true
    chart.DownBars.Interior.ColorIndex = 3
    chart.UpBars.Interior.ColorIndex = 5
}
```

### Chart.ChartObjects

返回一个对象，它代表工作表上的一个嵌入式图表（ **ChartObject** 对象）或所有嵌入式图表的集合（ **ChartObjects** 对象）。

**语法**

**express.ChartObjects ( *Index* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明          |
|---------|-----------|----------|---------------|
| *Index* | 可选      | Variant  | {description} |

**说明**

该方法不等同于 **Charts** 属性，它返回嵌入式图表，而 **Charts** 属性返回图表工作表。使用 **Chart** 属性可返回嵌入式图表的 **Chart** 对象。

**示例**

``` JavaScript
/*本示例向工作表 Sheet1 上的第一个嵌入式图表中添加标题。*/
function test()
{
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.HasTitle = true
    chart.ChartTitle.Text = "1995 Rainfall Totals by Month"
}

/*本示例在 Sheet1 上的第一个嵌入式图表中新建一个系列，新系列的数据源为 Sheet1 的 B1:B10 区域。*/
function test()
{
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
    Application.ActiveChart.SeriesCollection.Add(Worksheets.Item("Sheet1").Range("B1:B10"))
}

/*本示例清除工作表 Sheet1 上第一个嵌入式图表的格式设置。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats()
```

### Chart.ChartWizard

修改给定图表的属性。可使用本方法快速设置图表的格式，而不必逐个设置所有属性。本方法是非交互式的，并且仅更改指定的属性。

**语法**

**express.ChartWizard ( *Source* , *Gallery* , *Format* , *PlotBy* , *CategoryLabels* , *SeriesLabels* , *HasLegend* , *Title* , *CategoryTitle* , *ValueTitle* , *ExtraTitle* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                           |
|------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------|
| *Source*         | 可选      | Variant  | 包含新图表源数据的区域。如果省略本参数，ET 将编辑活动图表工作表或活动工作表上处于选定状态的图表。                              |
| *Gallery*        | 可选      | Variant  | 用于指定图表类型的 XlChartType 的常量之一。                                                                                    |
| *Format*         | 可选      | Variant  | 内置自动套用格式的选项编号。可为从 1 到 10 的数字，其取值取决于库的类型。如果省略此参数，ET 将根据库的类型和数据源选择默认值。 |
| *PlotBy*         | 可选      | Variant  | 指定每个系列的数据是来自行还是来自列。可以是以下 XlRowCol 常量之一：xlRows 或 xlColumns。                                      |
| *CategoryLabels* | 可选      | Variant  | 指定包含分类标签的源范围内的行数或列数的整数。合法值为从 0（零）至小于相应分类或系列的最大个数间的某一数字。                   |
| *SeriesLabels*   | 可选      | Variant  | 指定包含系列标志的源范围内的行数或列数的整数。合法值为从 0（零）至小于相应分类或系列的最大个数间的某一数字。                   |
| *HasLegend*      | 可选      | Variant  | 若要包括图例，则为 True。                                                                                                      |
| *Title*          | 可选      | Variant  | 图表标题文字。                                                                                                                 |
| *CategoryTitle*  | 可选      | Variant  | 分类轴标题文字。                                                                                                               |
| *ValueTitle*     | 可选      | Variant  | 数值轴标题文字。                                                                                                               |
| *ExtraTitle*     | 可选      | Variant  | 三维图表的系列轴标题，或二维图表的次数值轴标题。                                                                               |

**说明**

如果省略 *Source* ，并且选定内容不是活动工作表中的嵌入图表或者活动工作表不是当前存在的图表，则该方法失效并产生错误。

**示例**

``` JavaScript
/*本示例重新设置 Chart1 的格式，将其改为折线图，添加图例，并添加分类轴标题和数值轴标题。*/
Application.Charts.Item("Chart1").ChartWizard(undefined, xlLine, undefined, undefined, undefined, undefined, true, undefined, "Year", "Sales")
```

### Chart.CheckSpelling

检查对象的拼写。

**语法**

**express.CheckSpelling ( *CustomDictionary* , *IgnoreUppercase* , *AlwaysSuggest* , *SpellLang* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                                                                   |
|--------------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------|
| *CustomDictionary* | 可选      | Variant  | 一个字符串，它表示自定义词典的文件名，如果在主词典中找不到单词，则会到此词典中查找。如果省略此参数，则使用当前指定的词典。             |
| *IgnoreUppercase*  | 可选      | Variant  | 如果为 True，则 ET 忽略所有字母都是大写的单词。如果为 False，则 ET 检查所有字母都是大写的单词。如果省略此参数，将使用当前的设置。      |
| *AlwaysSuggest*    | 可选      | Variant  | 如果为 True，则 ET 在找到不正确拼写时显示建议的替换拼写列表。如果为 False，ET 将等待输入正确的拼写。如果省略此参数，将使用当前的设置。 |
| *SpellLang*        | 可选      | Variant  | 当前所用词典的语言。它可以是由 LanguageID 属性使用的 MsoLanguageID 值之一。                                                            |

**说明**

此方法没有返回值；ET 显示 **“拼写检查”** 对话框。

要检查工作表中的页眉、页脚和对象，可对 **Worksheet** 对象使用此方法。

### Chart.ClearToMatchStyle

清除图表元素格式以改为自动格式。

**语法**

**express.ClearToMatchStyle ()**

\*express是一个代表 Chart 对象的变量。

**说明**

使用此方法将图表元素格式重置为自动格式。如果在图表上使用此方法，则将重置所有格式（包括替代项）。

### Chart.Copy

将工作表复制到工作簿的另一位置。

**语法**

**express.Copy ( *Before* , *After* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                        |
|----------|-----------|----------|-----------------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 将要在其之前放置所复制工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选      | Variant  | 将要在其之后放置所复制工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After* ，则 ET 将新建一个工作簿，其中包含复制的工作表。

**Copy** 方法不支持图表对象。

**示例**

``` JavaScript
/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之后。*/
Application.Worksheets.Item("Sheet1").Copy(null,Worksheets.Item("Sheet3"))
```

### Chart.CopyPicture

将所选对象作为图片复制到剪贴板。

**语法**

**express.CopyPicture ( *Appearance* , *Format* , *Size* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型            | 说明                                                                                                     |
|--------------|-----------|---------------------|----------------------------------------------------------------------------------------------------------|
| *Appearance* | 可选      | XlPictureAppearance | 指定图片的复制方式。默认值为 xlScreen。                                                                  |
| *Format*     | 可选      | XlCopyPictureFormat | 图片的格式。默认值为 xlPicture。                                                                         |
| *Size*       | 可选      | XlPictureAppearance | 当对象是图表工作表中的图表（不是工作表中的嵌入式图表）时，此参数代表复制图片的大小。默认值为 xlPrinter。 |

### Chart.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Chart 对象的变量。

### Chart.Evaluate

将一个 ET 名称转换为一个对象或者一个值。

**语法**

**express.Evaluate ( *Name* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                         |
|--------|-----------|----------|------------------------------|
| *Name* | 必选      | Variant  | 使用 ET 命名约定的对象名称。 |

**说明**

该方法可使用下列 ET 名称类型：

- A1 格式引用。可以通过 A1 格式表示法引用单个单元格。所有引用均视为绝对引用。
- 区域。在引用中可以使用区域、交集和联合运算符（分别为冒号、空格和逗号）。
- 定义的名称。可用宏语言指定任何名称。
- 外部引用。可以使用 ! 运算符引用另一工作簿中的单元格或已定义的名称，例如， ` Evaluate("[BOOK1.XLS]Sheet1!A1") ` 。
- 图表对象。可以指定任何图表对象名称（如“Legend”、“Plot Area”或“Series 1”），以访问该对象的属性和方法。例如， ` Charts("Chart1").Evaluate("Legend").Font.Name ` 返回图例中所用字体的名称。

**注释** ： 使用方括号（例如，“\[A1:C5\]”，JsApi不支持方括号的用法）、或直接调用对应对象与用字符串参数调用 **Evaluate** 方法是等效的。例如，下列表达式对是等效的。

``` JavaScript
Application.Rnage("a1").Value2 = 25
Application.Evaluate("A1").Value2 = 25
        
trigVariable = SIN(45)
trigVariable = Application.Evaluate("SIN(45)")
        
Set firstCellInSheet = Application.Workbooks.Item("BOOK1.XLS").Sheets.Item(4).Range("A1")
Set firstCellInSheet = Application.Workbooks.Item("BOOK1.XLS").Sheets.Item(4).Evaluate("A1")
```

使用方括号的优点在于代码较短。使用 **Evaluate** 的优点在于参数是字符串，这样您既可以在代码中构造该字符串，也可以使用 Visual Basic 变量。

**示例**

``` JavaScript
/*此示例将工作表 Sheet1 上 A1 单元格的字体设置为加粗。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    boldCell = "A1"
    Application.Evaluate(boldCell).Font.Bold = true
}
```

### Chart.Export

以图形格式导出图表。

**语法**

**express.Export ( *Filename* , *FilterName* , *Interactive* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                    |
|---------------|-----------|----------|---------------------------------------------------------------------------------------------------------|
| *Filename*    | 必选      | String   | 被导出的文件的名称。                                                                                    |
| *FilterName*  | 可选      | Variant  | 图形过滤器出现在注册表中时独立于语言的名称。                                                            |
| *Interactive* | 可选      | Variant  | 如果为 True，则显示包含筛选器特定选项的对话框。如果为 False，则 ET 使用筛选器的默认值。默认值是 False。 |

**示例**

``` JavaScript
/*此示例以 GIF 文件格式导出第一个图表。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Export("current_sales.gif", "GIF")
```

### Chart.ExportAsFixedFormat

导出为指定格式的文件。

**语法**

**express.ExportAsFixedFormat ( *Type* , *Filename* , *Quality* , *IncludeDocProperties* , *IgnorePrintAreas* , *From* , *To* , *OpenAfterPublish* , *FixedFormatExtClassPtr* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型          | 说明                                                                         |
|--------------------------|-----------|-------------------|------------------------------------------------------------------------------|
| *Type*                   | 必选      | XlFixedFormatType | 要导出为的文件格式类型。                                                     |
| *Filename*               | 可选      | Variant           | 要保存的文件的文件名。可以包括完整路径，否则 ET 会将文件保存在当前文件夹中。 |
| *Quality*                | 可选      | Variant           | 可选 XlFixedFormatQuality。指定已发布文件的质量。                            |
| *IncludeDocProperties*   | 可选      | Variant           | 若要包括文档属性，则为 True；否则为 False。                                  |
| *IgnorePrintAreas*       | 可选      | Variant           | 若要忽略发布时设置的任何打印区域，则为 True；否则为 False。                  |
| *From*                   | 可选      | Variant           | 发布的起始页码。如果省略此参数，则从起始位置开始发布。                       |
| *To*                     | 可选      | Variant           | 发布的终止页码。如果省略此参数，则发布至最后一页                             |
| *OpenAfterPublish*       | 可选      | Variant           | 若要在发布文件后在查看器中显示文件，则为 True；否则为 False。                |
| *FixedFormatExtClassPtr* | 可选      | Variant           | 指向 FixedFormatExt 类的指针。                                               |

**说明**

此方法还支持初始化加载项，以将文件导出为固定格式的文件。例如，如果存在转换器，则 ET 将执行文件格式转换。转换通常由用户执行。

### Chart.GetChartElement

返回指定的 X 坐标和 Y 坐标上图表元素的信息。本方法稍有与众不同之处：调用时只须指定前两个参数，在本方法执行期间，ET 为其余参数赋值，本方法返回后应检验这些参数的值。

**语法**

**express.GetChartElement ( *x* , *y* , *ElementID* , *Arg1* , *Arg2* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                          |
|-------------|-----------|----------|-----------------------------------------------------------------------------------------------|
| *x*         | 必选      | Long     | 图表元素的 X 坐标。                                                                           |
| *y*         | 必选      | Long     | 图表元素的 Y 坐标。                                                                           |
| *ElementID* | 必选      | Long     | 该方法返回时，此参数包含指定坐标的图表元素的 XLChartItem 值。有关详细信息，请参阅“备注”部分。 |
| *Arg1*      | 必选      | Long     | 该方法返回时，该参数包含与图表元素相关的信息。有关详细信息，请参阅“备注”部分。                |
| *Arg2*      | 必选      | Long     | 该方法返回时，该参数包含与图表元素相关的信息。有关详细信息，请参阅“备注”部分。                |

**说明**

本方法返回的 *ElementID* 值决定了 *Arg1* 和 *Arg2* 所包含的信息，如下表所示。

| ElementID 常量              | 常量值 | Arg1         | Arg2            |
|-----------------------------|--------|--------------|-----------------|
| **xlAxis**                  | 21     | AxisIndex    | AxisType        |
| **xlAxisTitle**             | 17     | AxisIndex    | AxisType        |
| **xlDisplayUnitLabel**      | 30     | AxisIndex    | AxisType        |
| **xlMajorGridlines**        | 15     | AxisIndex    | AxisType        |
| **xlMinorGridlines**        | 16     | AxisIndex    | AxisType        |
| **xlPivotChartDropZone**    | 32     | DropZoneType | 无              |
| **xlPivotChartFieldButton** | 31     | DropZoneType | PivotFieldIndex |
| **xlDownBars**              | 20     | GroupIndex   | 无              |
| **xlDropLines**             | 26     | GroupIndex   | 无              |
| **xlHiLoLines**             | 25     | GroupIndex   | 无              |
| **xlRadarAxisLabels**       | 27     | GroupIndex   | 无              |
| **xlSeriesLines**           | 22     | GroupIndex   | 无              |
| **xlUpBars**                | 18     | GroupIndex   | 无              |
| **xlChartArea**             | 2      | 无           | 无              |
| **xlChartTitle**            | 4      | 无           | 无              |
| **xlCorners**               | 6      | 无           | 无              |
| **xlDataTable**             | 7      | 无           | 无              |
| **xlFloor**                 | 23     | 无           | 无              |
| **xlLeaderLines**           | 29     | 无           | 无              |
| **xlLegend**                | 24     | 无           | 无              |
| **xlNothing**               | 28     | 无           | 无              |
| **xlPlotArea**              | 19     | 无           | 无              |
| **xlWalls**                 | 5      | 无           | 无              |
| **xlDataLabel**             | 7      | SeriesIndex  | PointIndex      |
| **xlErrorBars**             | 9      | SeriesIndex  | 无              |
| **xlLegendEntry**           | 12     | SeriesIndex  | 无              |
| **xlLegendKey**             | 13     | SeriesIndex  | 无              |
| **xlSeries**                | 3      | SeriesIndex  | PointIndex      |
| **xlShape**                 | 14     | ShapeIndex   | 无              |
| **xlTrendline**             | 8      | SeriesIndex  | TrendLineIndex  |
| **xlXErrorBars**            | 10     | SeriesIndex  | 无              |
| **xlYErrorBars**            | 11     | SeriesIndex  | 无              |

下表说明了本方法返回后，参数 *Arg1* 和 *Arg2* 的含义。

| 参数            | 说明                                                                                                                                                                                                         |
|-----------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| AxisIndex       | 指定坐标轴是主坐标轴还是次坐标轴。可为以下 **XlAxisGroup** 常量之一： **xlPrimary** 或 **xlSecondary** 。                                                                                                    |
| AxisType        | 指定坐标轴类型。可为以下 **XlAxisType** 常量之一： **xlCategory** 、 **xlSeriesAxis** 或 **xlValue** 。                                                                                                      |
| DropZoneType    | 指定拖放区的类型：列、数据、页或行字段。可为以下 **XlPivotFieldOrientation** 常量之一： **xlColumnField** 、 **xlDataField** 、 **xlPageField** 或 **xlRowField** 。列和行字段常量分别指定了系列和分类字段。 |
| GroupIndex      | 指定特定图表组在 **ChartGroups** 集合内的偏移量。                                                                                                                                                            |
| PivotFieldIndex | 指定特定列（数据系列）、数据、页或行（类别）字段在 **PivotFields** 集合中的偏移量。如果拖放区域类型为 **xlDataField** ，则返回 -1。                                                                          |
| PointIndex      | 指定系列中的特定点在 **Points** 集合中的偏移量。-1 表示选定了所有数据点。                                                                                                                                    |
| SeriesIndex     | 指定特定系列在 **Series** 集合中的偏移量。                                                                                                                                                                   |
| ShapeIndex      | 指定特定形状在 **Shapes** 集合中的偏移量。                                                                                                                                                                   |
| TrendlineIndex  | 指定某个系列中特定趋势线在 **Trendlines** 集合中的偏移量。                                                                                                                                                   |

**示例**

``` JavaScript
/*当鼠标移动到图表图例上时，本示例发出警告。*/
function test()
{
    let IDNum 
    let a
    let b
    Application.ActiveChart.GetChartElement(X, Y, IDNum, a, b)
    if(IDNum == xlLegendEntry) {
        alert("WARNING: Move away from the legend")
    }
}
```

### Chart.Location

将图表移动到新位置。

**语法**

**express.Location ( *Where* , *Name* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型        | 说明                                                                                                                                                                                       |
|---------|-----------|-----------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Where* | 必选      | XlChartLocation | 图表移动的目标位置。                                                                                                                                                                       |
| *Name*  | 可选      | Variant         | 如果 Where 为 xlLocationAsObject，则该参数为必选参数。如果 Where 为 xlLocationAsObject，则该参数为嵌入该图表的工作表的名称。如果 Where 为 xlLocationAsNewSheet，则该参数为新工作表的名称。 |

**示例**

``` JavaScript
/*本示例将嵌入图表移至新图表工作表“Monthly Sales”。*/
Application.Worksheets.Item(1).ChartObjects(1).Chart.Location(xlLocationAsNewSheet, "Monthly Sales")
```

### Chart.Move

将图表移到工作簿的另一位置。

**语法**

**express.Move ( *Before* , *After* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                      |
|----------|-----------|----------|---------------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 将要在其之前放置所移动图表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选      | Variant  | 将要在其之后放置所移动图表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After* ，则 ET 将新建一个工作簿并将要移动的图表移到新工作簿中。

### Chart.OLEObjects

返回一个对象，它代表图表或工作表上的单个 OLE 对象 ( **OLEObject** ) 或所有 OLE 对象的集合（ **OLEObjects** 集合）。只读。

**语法**

**express.OLEObjects ( *Index* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                   |
|---------|-----------|----------|------------------------|
| *Index* | 可选      | Variant  | OLE 对象的名称或编号。 |

**示例**

``` JavaScript
/*此示例创建工作表 Sheet1 上 OLE 对象的链接类型列表。该列表将出现在此示例新建的工作表中。*/
function test()
{
    let newSheet = Application.Worksheets.Add()
    let j = 2
    newSheet.Range("A1").Value2 = "Name"
    newSheet.Range("B1").Value2 = "Link Type"
    for(let i = 1; i <= Application.Worksheets.Item("Sheet1").OLEObjects().Count; i++){
        newSheet.Cells.Item(j, 1).Value2 = Application.Worksheets.Item("Sheet1").OLEObjects().Item(i).Name
        if(Application.Worksheets.Item("Sheet1").OLEObjects().Item(i).OLEType == xlOLELink){
            newSheet.Cells.Item(j, 2) = "Linked"
        }
        else{
            newSheet.Cells.Item(j, 2) = "Embedded"
        }
        j++
    }
}
```

### Chart.Paste

将剪贴板中的图表数据粘贴到指定的图表中。

**语法**

**express.Paste ( *Type* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                    |
|--------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Type* | 可选      | Variant  | 指定要粘贴的图表信息（如果剪贴板中有一个图表）。可以为以下 XlPasteType 常量之一：xlPasteFormats、xlPasteFormulas 或 xlPasteAll。默认值为 xlPasteAll。如果剪贴板上的数据不是图表数据，则不能使用该参数。 |

**示例**

``` JavaScript
/*本示例将工作表 Sheet1 上单元格区域 B1:B5 中的数据粘贴到 Chart1 中。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("B1:B5").Copy()
    Application.Charts.Item("Chart1").Paste()
}
```

### Chart.PrintOut

打印对象。

**语法**

**express.PrintOut ( *From* , *To* , *Copies* , *Preview* , *ActivePrinter* , *PrintToFile* , *Collate* , *PrToFileName* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                              |
|-----------------|-----------|----------|---------------------------------------------------------------------------------------------------|
| *From*          | 可选      | Variant  | 打印的开始页号。如果省略此参数，则从起始位置开始打印。                                            |
| *To*            | 可选      | Variant  | 打印的终止页号。如果省略此参数，则打印至最后一页。                                                |
| *Copies*        | 可选      | Variant  | 打印份数。如果省略此参数，则只打印一份。                                                          |
| *Preview*       | 可选      | Variant  | 如果为 True，ET 将在打印对象之前调用打印预览。如果为 False（或省略该参数），则立即打印对象。      |
| *ActivePrinter* | 可选      | Variant  | 设置活动打印机的名称。                                                                            |
| *PrintToFile*   | 可选      | Variant  | 如果为 True，则打印到文件。如果没有指定 PrToFileName，ET 将提示用户输入要使用的输出文件的文件名。 |
| *Collate*       | 可选      | Variant  | 如果为 True，则逐份打印多个副本。                                                                 |
| *PrToFileName*  | 可选      | Variant  | 如果 PrintToFile 设为 True，则该参数指定要打印到的文件名。                                        |

**说明**

*From* 和 *To* 所描述的“页”指的是要打印的页，并非指定工作表或工作簿中的全部页。

**示例**

``` JavaScript
/*此示例打印当前活动工作表。*/
Application.ActiveSheet.PrintOut()
```

### Chart.PrintPreview

按对象打印后的外观效果显示对象的预览。

**语法**

**express.PrintPreview ( *EnableChanges* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                          |
|-----------------|-----------|----------|-------------------------------------------------------------------------------|
| *EnableChanges* | 可选      | Variant  | 传递 Boolean 值，以指定用户是否可更改边距和打印预览中可用的其他页面设置选项。 |

**示例**

``` JavaScript
/*此示例在打印预览中显示工作表 Sheet1。*/
Application.Worksheets.Item("Sheet1").PrintPreview()
```

### Chart.Protect

保护图表使其不被修改。

**语法**

**express.Protect ( *Password* , *DrawingObjects* , *Contents* , *Scenarios* , *UserInterfaceOnly* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                     |
|---------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Password*          | 可选      | Variant  | 一个字符串，该字符串为工作表或工作簿指定区分大小写的密码。如果省略此参数，不用密码就可以取消对工作表或工作簿的保护。否则，必须指定密码才能取消对工作表或工作簿的保护。如果忘记了密码，就无法取消对工作表或工作簿的保护。 |
| *DrawingObjects*    | 可选      | Variant  | 如果为 True，则保护形状。默认值是 True。                                                                                                                                                                                 |
| *Contents*          | 可选      | Variant  | 如果为 True，则保护内容。对于图表，这样会保护整个图表。对于工作表，这样会保护锁定的单元格。默认值是 True。                                                                                                               |
| *Scenarios*         | 可选      | Variant  | 如果为 True，则保护方案。此参数仅对工作表有效。默认值是 True。                                                                                                                                                           |
| *UserInterfaceOnly* | 可选      | Variant  | 如果为 True，则保护用户界面，但不保护宏。如果省略此参数，则既保护宏也保护用户界面。                                                                                                                                      |

### Chart.Refresh

立即重新绘制指定的图表。

**语法**

**express.Refresh ()**

\*express是一个代表 Chart 对象的变量。

### Chart.SaveAs

将对图表或工作表的更改保存到另一不同文件中。

**语法**

**express.SaveAs ( *Filename* , *FileFormat* , *Password* , *WriteResPassword* , *ReadOnlyRecommended* , *CreateBackup* , *AddToMru* , *TextCodepage* , *TextVisualLayout* , *Local* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称                  | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                  |
|-----------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Filename*            | 必选      | String   | Variant。一个字符串，表示要保存的文件名。可包含完整路径。如果不指定路径，ET 将文件保存到当前文件夹中。                                                                                                                                |
| *FileFormat*          | 可选      | Variant  | 保存文件时使用的文件格式。要查看有效的选项列表，请参阅 FileFormat 属性。对于现有文件，默认采用上一次指定的文件格式；对于新文件，默认采用当前所用 ET 版本的格式。                                                                      |
| *Password*            | 可选      | Variant  | 它是一个区分大小写的字符串（最长不超过 15 个字符），用于指定文件的保护密码。                                                                                                                                                          |
| *WriteResPassword*    | 可选      | Variant  | 一个表示文件写保护密码的字符串。如果文件保存时带有密码，但打开文件时不输入密码，则该文件以只读方式打开。                                                                                                                              |
| *ReadOnlyRecommended* | 可选      | Variant  | 如果为 True，则在打开文件时显示一条消息，提示该文件以只读方式打开。                                                                                                                                                                   |
| *CreateBackup*        | 可选      | Variant  | 如果为 True，则创建备份文件。                                                                                                                                                                                                         |
| *AddToMru*            | 可选      | Variant  | 如果为 True，则将该工作簿添加到最近使用的文件列表中。默认值是 False。                                                                                                                                                                 |
| *TextCodepage*        | 可选      | Variant  | 不在美国英语版的 ET 中使用。                                                                                                                                                                                                          |
| *TextVisualLayout*    | 可选      | Variant  | 不在美国英语版的 ET 中使用。                                                                                                                                                                                                          |
| *Local*               | 可选      | Variant  | 如果为 True，则以 ET（包括控制面板设置）的语言保存文件。如果为 False（默认值），则以 示例代码 (VBA) 的语言保存文件，其中 示例代码 (VBA) 通常为美国英语版本，除非从中运行 Workbooks.Open 的 VBA 项目是旧的已国际化的 XL5/95 VBA 项目。 |

**说明**

请使用同时包含大小写字母、数字和符号的强密码。弱密码不混合使用这些元素。强密码：Y6dh!et5。弱密码：House27。请使用您可以记住的强密码，这样就不必将它写下来。

### Chart.SaveChartTemplate

向可用图表模板的列表添加自定义图表模板。

**语法**

**express.SaveChartTemplate ( *Filename* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明             |
|------------|-----------|----------|------------------|
| *Filename* | 必选      | String   | 图表模板的名称。 |

**说明**

默认情况下，此方法将活动图表保存到用户的图表模板库中。如果指定了 UNC 或 URL，则图表将保存到指定位置。

**示例**

``` JavaScript
/*此示例基于活动图表添加一个新图表模板。*/
Application.ActiveChart.SaveChartTemplate("Presentation Chart")
```

### Chart.Select

选择对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                              |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

**说明**

要选择单元格或单元格区域，请使用 **Select** 方法。要使单个单元格成为活动单元格，请使用 **Activate** 方法。

### Chart.SeriesCollection

返回一个对象，它代表图表或图表组中的一个系列（ **Series** 对象）或所有系列的集合（ **SeriesCollection** 集合）。

**语法**

**express.SeriesCollection ( *Index* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                   |
|---------|-----------|----------|------------------------|
| *Index* | 可选      | Variant  | 数据系列的名称或编号。 |

**示例**

``` JavaScript
/*此示例打开 Chart1 中第一个数据系列的数据标签。*/
Application.Charts.Item("Chart1").SeriesCollection(1).HasDataLabels = true
```

### Chart.SetBackgroundPicture

为图表设置背景图形。

**语法**

**express.SetBackgroundPicture ( *Filename* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明         |
|------------|-----------|----------|--------------|
| *Filename* | 必选      | String   | 图形文件名。 |

### Chart.SetDefaultChart

指定 ET 新建图表时使用的图表模板的名称。

**语法**

**express.SetDefaultChart ( *Name* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                                                       |
|--------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Name* | 必选      | Variant  | 指定新建图表时将使用的默认图表模板的名称。此名称可以是字符串，指定库中图表，进而指定用户定义的模板，也可以是一个特殊常量 xlBuiltIn，用于指定内置图表模板。 |

**示例**

``` JavaScript
/*此示例将默认图表模板设置为名为“Monthly Sales”的自定义图表。*/
Application.ActiveChart.SetDefaultChart("Monthly Sales")
```

### Chart.SetElement

设置图表上的图表元素。可读/写 **MsoChartElementType** 类型。

**语法**

**express.SetElement ( *Element* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型            | 说明               |
|-----------|-----------|---------------------|--------------------|
| *Element* | 必选      | MsoChartElementType | 指定图表元素类型。 |

**说明**

对于图表， **“布局”** 选项卡中的以下命令对应于 **SetElement** 方法：

- **“标签”** 组中的全部内容。
- **“坐标轴”** 组中的全部内容。
- **“分析”** 组中的全部内容。
- **“绘图区”** 、 **“图表背景墙”** 和 **“图表基底”** 按钮。

**MsoChartElementType** 是常量枚举，引用上述所有命令。

function test() { }

**示例**

``` JavaScript
/*此示例使用不同常量值向活动图表设置图表元素。*/
function test()
{
    Application.ActiveChart.Axes(xlValue).MajorGridlines.Select()
    Application.ActiveChart.SetElement(msoElementChartTitleCenteredOverlay)
    Application.ActiveChart.SetElement(msoElementPrimaryCategoryGridLinesMinor)
    Application.ActiveChart.Walls.Select()
    Application.Application.CommandBars.Item("Clip Art").Visible = false
    Application.ActiveChart.SetElement(msoElementChartFloorShow)
}
```

### Chart.SetSourceData

为指定图表设置源数据区域。

**语法**

**express.SetSourceData ( *Source* , *PlotBy* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                |
|----------|-----------|----------|---------------------------------------------------------------------|
| *Source* | 必选      | Range    | 包含源数据的区域。                                                  |
| *PlotBy* | 可选      | Variant  | 指定数据绘制方式。可为以下 XlRowCol 常量之一：xlColumns 或 xlRows。 |

**示例**

``` JavaScript
/*本示例为第一个图表设置源数据区域。*/
Application.Charts.Item(1).SetSourceData(Sheets.Item(1).Range("a1:a10"), xlColumns)
```

### Chart.Unprotect

取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。

**语法**

**express.Unprotect ( *Password* )**

\*express是一个代表 Chart 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                   |
|------------|-----------|----------|----------------------------------------------------------------------------------------|
| *Password* | 可选      | Variant  | 指定用于解除图表保护的密码，此密码是区分大小写的。如果图表不设密码保护，则忽略该参数。 |

**说明**

如果您忘记了密码，将不能取消工作表或工作簿的保护。建议将密码和对应文档名妥善保存。

## 成员属性

### Chart.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    if(Application.ActiveWorkbook.Application.Value == "ET"){
        alert("This is an ET Application object.")
    }
    else{
        alert("This is not an ET Application object.")
    }
}
```

### Chart.AutoScaling

如果 ET 对三维图表进行缩放，使之与等效的二维图表的大小相近，则该属性值为 **True** 。 **RightAngleAxes** 属性必须为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoScaling**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例自动对 Chart1 进行缩放。本示例应在三维图表上运行。*/
function test()
{
    let charts = Application.Charts.Item("Chart1")
    charts.RightAngleAxes = true
    charts.AutoScaling = true
}
```

### Chart.BackWall

返回 **Walls** 对象，该对象允许用户单独对三维图表的背景墙进行格式设置。只读。

**语法**

**express.BackWall**

\*express是一个代表 Chart 对象的变量。

### Chart.BarShape

返回或设置用于三维条形图或柱形图的形状。 **XlBarShape** 类型，可读写。

**语法**

**express.BarShape**

\*express是一个代表 Chart 对象的变量。

### Chart.ChartArea

返回一个 **ChartArea** 对象，该对象表示图表的整个图表区。只读。

**语法**

**express.ChartArea**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 的图表区域内部颜色设置为红色，并将其边框颜色设置为蓝色。*/
function test()
{
    let area = Application.Charts.Item("Chart1").ChartArea
    area.Interior.ColorIndex = 3
    area.Border.ColorIndex = 5
}
```

### Chart.ChartStyle

返回或设置图表的图表样式。可读/写 **Variant** 类型。

**语法**

**express.ChartStyle**

\*express是一个代表 Chart 对象的变量。

**说明**

您可以使用介于 1 到 48 之间的数字设置图表样式。

### Chart.ChartTitle

返回一个 **ChartTitle** 对象，该对象表示指定图表的标题。只读。

**语法**

**express.ChartTitle**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例设置 Chart1 的标题文本。*/
function test()
{
    let charts = Application.Charts.Item("Chart1")
    charts.HasTitle = true
    charts.ChartTitle.Text = "First Quarter Sales"
}
```

### Chart.ChartType

返回或设置图表类型。 **XlChartType** 类型，可读写。

**语法**

**express.ChartType**

\*express是一个代表 Chart 对象的变量。

**说明**

某些图表类型对数据透视图报表无效。

**示例**

``` JavaScript
/*当图表为二维气泡图时，本示例将第一个图表组中的气泡大小设置为默认大小的 200%。*/
function test()
{
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    if(chart.ChartType == xlBubble){
        chart.ChartGroups(1).BubbleScale = 200
    }
}
```

### Chart.CodeName

返回对象的代码名。 **String** 型，只读。

**语法**

**express.CodeName**

\*express是一个代表 Chart 对象的变量。

**说明**

在“属性” 窗口中“（名称）” 右边的单元格中显示的值是所选对象的代码名。可以在设计过程中通过更改该值来改变对象的代码名。不能在运行过程中更改该属性。

对于一个返回指定对象的表达式，该表达式可使用对象的代码名。例如，如果第一张工作表的代码名为 Sheet1，则下列表达式是等价的。

``` JavaScript
Application.Worksheets.Item(1).Range("a1")
Sheet1.Range("a1")
```

工作表的名称可以与其代码名不同。创建一张工作表时，其工作表名称和代码名是相同的，不过，更改工作表名称时并不影响其代码名，并且，更改工作表代码名（在 Visual Basic 编辑器中使用“属性” 窗口）也不影响其名称。

**示例**

``` JavaScript
/*此示例显示第一张工作表的代码名。*/
alert(Application.Worksheets.Item(1).CodeName)
```

### Chart.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Chart 对象的变量。

### Chart.DataTable

返回一个 **DataTable** 对象，该对象表示图表的模拟运算表。只读。

**语法**

**express.DataTable**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例向嵌入图表添加带有外框的模拟运算表。*/
function test()
{
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.HasDataTable = true
    chart.DataTable.HasBorderOutline = true
}
```

### Chart.DepthPercent

返回或设置三维图表的深度，以图表宽度的百分比表示（有效范围从 20% 到 2000%）。 **Long** 类型，可读写。

**语法**

**express.DepthPercent**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 的深度设置为其宽度的 50%。本示例应在三维图表上运行（DepthPercent 属性在二维图表中无效）。*/
Application.Charts.Item("Chart1").DepthPercent = 50
```

### Chart.DisplayBlanksAs

返回或设置在图表中绘制空白单元格的方式。可以是 **XlDisplayBlanksAs** 常量之一。 **Long** 类型，可读写。

**语法**

**express.DisplayBlanksAs**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例设置 ET 在 Chart1 中不绘制空白单元格。*/
Application.Charts.Item("Chart1").DisplayBlanksAs = xlNotPlotted
```

### Chart.Elevation

返回或设置三维图表视图的仰角（以角度为单位）。 **Long** 类型，可读写。

**语法**

**express.Elevation**

\*express是一个代表 Chart 对象的变量。

**说明**

图表的仰角以角度为单位表示观察图表的高度。对于绝大多数图表类型，其默认值为 15。本属性的值必须介于 -90 到 90 之间，但三维条形图除外，对于三维条形图，该值必须介于 0 到 44 之间。

**示例**

``` JavaScript
/*本示例将 Chart1 的仰角设为 34 度。本示例应在三维图表上运行（Elevation 属性在二维图表上无效）。*/
Application.Charts.Item("Chart1").Elevation = 34
```

### Chart.Floor

返回一个 **Floor** 对象，该对象表示三维图表的基底。只读。

返回值类型为 Floor。

**语法**

**express.Floor**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 的基底颜色设为蓝色。本示例应在三维图表上运行（Floor 属性在二维图表无效）。*/
Application.Charts.Item("Chart1").Floor.Interior.ColorIndex = 5
```

### Chart.GapDepth

以数据标志宽度的形式返回或设置三维图表中数据系列之间的距离，本属性的值必须在 0 和 500 之间。 **Long** 类型，可读写。

**语法**

**express.GapDepth**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 的数据系列之间的距离设为数据标志宽度的 200%。本示例应在三维图表上运行（GapDepth 属性在二维图表上无效）。*/
Application.Charts.Item("Chart1").GapDepth = 2000
```

### Chart.HasAxis

返回或设置图表上显示的坐标轴。 **Variant** 类型，可读写。

**语法**

**express.HasAxis**

\*express是一个代表 Chart 对象的变量。

**说明**

设置该属性时，必须至少输入一个参数的值。

如果更改了图表类型或 **Axis.AxisGroup** 、 **Chart.AxisGroup** 或 **Series.AxisGroup** 属性，ET 可能要创建或删除坐标轴。

参数

| **名称** | **必选/可选** | **数据类型** | **说明**                                                                 |
|----------|---------------|--------------|--------------------------------------------------------------------------|
| *Index1* | 必选          | **Variant**  | 坐标轴类型。系列坐标轴仅适用于三维图表。可以是 **XlAxisType** 常量之一。 |
| *Index2* | 可选          | **Variant**  | 坐标轴组。三维图表只有一组坐标轴。可以是 **XlAxisGroup** 常量之一。      |

**示例**

``` JavaScript
/*本示例显示 Chart1 中的主数值轴。*/
Application.Charts.Item("Chart1").HasAxis(xlValue, xlPrimary) = true
```

### Chart.HasDataTable

如果图表有模拟运算表，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasDataTable**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例使嵌入图表模拟运算表显示时带有外边框，但无单元格边框。*/
function test()
{
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.HasDataTable = true
    let newchart = chart.DataTable
        newchart.HasBorderHorizontal = false
        newchart.HasBorderVertical = false
        newchart.HasBorderOutline = true
}
```

### Chart.HasLegend

如果图表有图例，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasLegend**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例显示 Chart1 的图例，然后将图例的字体颜色设为蓝色。*/
function test()
{
    let charts = Application.Charts.Item("Chart1")
    charts.HasLegend = true
    charts.Legend.Font.ColorIndex = 5
}
```

### Chart.HasTitle

如果坐标轴或图表有可见标题，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasTitle**

\*express是一个代表 Chart 对象的变量。

**说明**

图表标题由 **ChartTitle** 对象代表。

### Chart.HeightPercent

返回或设置三维图表的高度，以图表宽度的百分比表示（有效范围从 5% 到 500%）。 **Long** 类型，可读写。

**语法**

**express.HeightPercent**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 的高度设为其宽度的 80%。本示例应在三维图表上运行。*/
Application.Charts.Item("Chart1").HeightPercent = 80
```

### Chart.Hyperlinks

返回一个 **Hyperlinks** 集合，它代表图表的超链接。

**语法**

**express.Hyperlinks**

\*express是一个代表 Chart 对象的变量。

### Chart.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 Chart 对象的变量。

### Chart.Legend

返回一个 **Legend** 对象，该对象表示指定图表的图例。只读。

**语法**

**express.Legend**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例显示 Chart1 的图例，然后将图例的字体颜色设为蓝色。*/
function test()
{
    Application.Charts.Item("Chart1").HasLegend = true
    Application.Charts.Item("Chart1").Legend.Font.ColorIndex = 5
}
```

### Chart.MailEnvelope

代表文档的电子邮件标题。

**语法**

**express.MailEnvelope**

\*express是一个代表 Chart 对象的变量。

### Chart.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Chart 对象的变量。

**说明**

此属性对图表对象（嵌入式图表）而言是只读的。

### Chart.Next

返回代表下一个工作表的 **Worksheet** 对象。

**语法**

**express.Next**

\*express是一个代表 Chart 对象的变量。

**说明**

如果该对象为区域，则此属性会模拟 Tab 键，但会返回下一个单元格，而不选中它。

在处于保护状态的工作表中，此属性返回下一个未锁定单元格。在未保护的工作表中，此属性总是返回紧靠指定单元格右边的单元格。

### Chart.PageSetup

返回一个 **PageSetup** 对象，它包含用于指定对象的所有页面设置。只读。

**语法**

**express.PageSetup**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*此示例设置 Chart1 的中央页眉的文字。*/
Application.Charts.Item("Chart1").PageSetup.CenterHeader = "December Sales"
```

### Chart.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Chart 对象的变量。

### Chart.Perspective

返回或设置一个 **Long** 值，它代表三维图表视图的透视系数。

**语法**

**express.Perspective**

\*express是一个代表 Chart 对象的变量。

**说明**

此属性的值必须在 0 到 100 之间。如果 **RightAngleAxes** 属性是 **True** ，则此属性会被忽略。

**示例**

``` JavaScript
/*此示例设置 Chart1 的透视系数为 70。此示例应在三维图表上运行。*/
function test()
{
    Application.Charts.Item("Chart1").RightAngleAxes = false
    Application.Charts.Item("Chart1").Perspective = 70
}
```

### Chart.PivotLayout

返回一个 **PivotLayout** 对象，该对象表示数据透视表中字段的位置以及数据透视图中坐标轴的位置。只读。

**语法**

**express.PivotLayout**

\*express是一个代表 Chart 对象的变量。

**说明**

如果指定图表不是一个数据透视图，则该属性值为 **Nothing** 。

**示例**

``` JavaScript
/*本示例创建第一个数据透视图中使用的所有数据透视表字段的名称列表。*/
function test()
{
    let objNewSheet = Application.Worksheets.Add()
    objNewSheet.Activate()
    let intRow = 1
    for(let i =1; i <= Application.Charts.Item("Chart1").PivotLayout.PivotFields().Count; i++){
        objNewSheet.Cells.Item(intRow, 1).Value2 = Application.Charts.Item("Chart1").PivotLayout.PivotFields(i).Caption
        intRow ++
    }
}
```

### Chart.PlotArea

返回一个 **PlotArea** 对象，该对象表示图表的绘图区。只读。

**语法**

**express.PlotArea**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 绘图区的内部颜色设置为蓝绿色。*/
Application.Charts.Item("Chart1").PlotArea.Interior.ColorIndex = 8
```

### Chart.PlotBy

返回或设置行或列在图表中作为数据系列使用的方式。可为以下 **XlRowCol** 常量之一： **xlColumns** 或 **xlRows** 。 **Long** 类型，可读写。

**语法**

**express.PlotBy**

\*express是一个代表 Chart 对象的变量。

**说明**

对于数据透视图，该属性是只读的，并且总返回 **xlColumns** 。

**示例**

``` JavaScript
/*本示例使嵌入图表按列绘制数据。*/
Application.Worksheets.Item(1).ChartObjects(1).Chart.PlotBy = xlColumns
```

### Chart.PlotVisibleOnly

如果仅绘制可见单元格，则该值为 **True** 。如果可见单元格和隐藏单元格都绘制，则该值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.PlotVisibleOnly**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例使 ET 在 Chart1 中仅绘制可见单元格。*/
Application.Charts.Item("Chart1").PlotVisibleOnly = true
```

### Chart.Previous

返回代表下一个工作表的 **Worksheet** 对象。

**语法**

**express.Previous**

\*express是一个代表 Chart 对象的变量。

**说明**

如果指定对象为区域，则此属性的作用是仿效 Shift+Tab；但此属性只是返回上一单元格，并不选定它。

在保护工作表上，该属性返回上一个未锁定的单元格。在未保护的工作表上，该属性通常返回指定单元格左侧相邻的单元格。

### Chart.PrintedCommentPages

返回将为当前图表打印的批注页的数量。只读。

返回类型类型为 **Long。**

**语法**

**express.PrintedCommentPages**

\*express是一个代表 Chart 对象的变量。

**说明**

由于图表和图表工作表不支持批注，因此 **Chart** 对象的 **PrintedCommentPages** 属性将总是返回零。

### Chart.ProtectContents

如果工作表内容是受保护的，则为 **True** 。对于图表，这样会保护整个图表。要打开内容保护，请使用 **Protect** 方法，并将 *Contents* 参数设置为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.ProtectContents**

\*express是一个代表 Chart 对象的变量。

### Chart.ProtectData

如果用户不能更改系列公式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ProtectData**

\*express是一个代表 Chart 对象的变量。

**说明**

该属性在文件被保存后将改变。如果将该属性设置为 **True** 然后重新打开文件，则它不再为 **True** 。

**示例**

``` JavaScript
/*本示例保护第一张工作表上嵌入的第一个图表中的数据。*/
Application.Worksheets.Item(1).ChartObjects(1).Chart.ProtectData = true
```

### Chart.ProtectDrawingObjects

如果形状是受保护的，则为 **True** 。要打开形状保护，请使用 **Protect** 方法，并将 *DrawingObjects* 参数设置为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.ProtectDrawingObjects**

\*express是一个代表 Chart 对象的变量。

### Chart.ProtectFormatting

如果用户不能更改格式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ProtectFormatting**

\*express是一个代表 Chart 对象的变量。

**说明**

该属性在文件被保存后将改变。如果将该属性设置为 **True** 然后重新打开文件，则它不再为 **True** 。

**示例**

``` JavaScript
/*本示例保护第一张工作表上嵌入的第一个图表的格式。*/
Application.Worksheets.Item(1).ChartObjects(1).Chart.ProtectFormatting = true
```

### Chart.ProtectSelection

如果不能选定图表元素，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ProtectSelection**

\*express是一个代表 Chart 对象的变量。

**说明**

该属性为 **True** 时，不能向图表中添加形状，而且图表元素的 **Click** 和 **DoubleClick** 事件不会发生。

该属性在文件被保存后将改变。如果将该属性设置为 **True** 然后重新打开文件，则它不再为 **True** 。

**示例**

``` JavaScript
/*本示例使第一张工作表上嵌入的第一个图表中的图表元素不能被选定。*/
Application.Worksheets.Item(1).ChartObjects(1).Chart.ProtectSelection = true
```

### Chart.ProtectionMode

如果启用了用户界面专用保护，则为 **True** 。要打开用户界面保护，请使用 **Protect** 方法，并将 *UserInterfaceOnly* 参数设置为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.ProtectionMode**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*此示例显示 ProtectionMode 属性的状态。*/
alert(Application.ActiveSheet.ProtectionMode)
```

### Chart.RightAngleAxes

如果图表的坐标轴为直角，并与图表的转角或仰角无关，则该值为 **True** 。仅应用于三维折线图、柱形图和条形图。 **Boolean** 类型，可读写。

**语法**

**express.RightAngleAxes**

\*express是一个代表 Chart 对象的变量。

**说明**

如果该属性值为 **True** ，则忽略 **Perspective** 属性。

**示例**

``` JavaScript
/*本示例设置 Chart1 中坐标轴夹角为直角。本示例应在三维图表上运行。*/
Application.Charts.Item("Chart1").RightAngleAxes = true
```

### Chart.Rotation

以度为单位返回或设置三维图表视图的转角（绘图区绕 Z 轴的转角）。此属性的取值必须介于 0 到 360 之间，三维条形图除外（从 0 到 44 之间）。默认值是 20。仅适用于三维图表。 **Variant** 型，可读写。

**语法**

**express.Rotation**

\*express是一个代表 Chart 对象的变量。

**说明**

旋转量始终四舍五入到最接近的整数。

**示例**

``` JavaScript
/*此示例将第一张图表的转角设置为 30 度。此示例应在三维图表上运行。*/
Application.Charts.Item("Chart1").Rotation = 30
```

### Chart.Shapes

返回一个 **Shapes** 集合，它代表图表工作表上所有的形状。只读。

**语法**

**express.Shapes**

\*express是一个代表 Chart 对象的变量。

### Chart.ShowAllFieldButtons

返回或设置是否在数据透视图上显示所有字段按钮。可读/写。

返回值类型为 **Boolean。**

**语法**

**express.ShowAllFieldButtons**

\*express是一个代表 Chart 对象的变量。

**说明**

将 **ShowAllFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示所有字段按钮。将该属性设置为 **False** 将隐藏所有字段按钮。

**ShowAllFieldButtons** 属性对应于 **“分析”** 选项卡（在选择数据透视图时可用）的 **“字段按钮”** 下拉列表中的 **“全部隐藏”** 命令。

**示例**

``` JavaScript
/*下面的代码示例将 Chart 1 设置为显示所有字段按钮。*/
function test()
{
    Application.ActiveSheet.ChartObjects("Chart 1").Activate()
    Application.ActiveChart.ShowAllFieldButtons = true
}
```

### Chart.ShowAxisFieldButtons

返回或设置是否要在数据透视图上显示坐标轴字段按钮。可读/写。

返回值类型为 **Boolean。**

**语法**

**express.ShowAxisFieldButtons**

\*express是一个代表 Chart 对象的变量。

**说明**

将 **ShowAxisFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示坐标轴字段按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowAxisFieldButtons** 属性对应于 **“分析”** 选项卡（在选择数据透视图时可用）的 **“字段按钮”** 下拉列表中的 **“显示坐标轴字段按钮”** 命令。

**示例**

``` JavaScript
/*下面的代码示例将 Chart 1 设置为隐藏坐标轴字段按钮。*/
function test()
{
    Application.ActiveSheet.ChartObjects("Chart 1").Activate()
    Application.ActiveChart.ShowAxisFieldButtons = false
}
```

### Chart.ShowDataLabelsOverMaximum

返回或设置一个布尔值，该值表示在数值大于数值轴上的最大值时是否显示数据标签。可读/写 **Boolean** 类型。

**语法**

**express.ShowDataLabelsOverMaximum**

\*express是一个代表 Chart 对象的变量。

**说明**

如果您更改数值轴，使其变得比数据点的大小还小，则可以使用此属性设置是否显示数据标签。此属性只适用于二维图表。

### Chart.ShowLegendFieldButtons

返回或设置是否要在数据透视图上显示图例字段按钮。可读/写。

返回值类型为 **Boolean。**

**语法**

**express.ShowLegendFieldButtons**

\*express是一个代表 Chart 对象的变量。

**说明**

将 **ShowLegendFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示图例字段按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowLegendFieldButtons** 属性对应于 **“分析”** 选项卡（在选择数据透视图时可用）的 **“字段按钮”** 下拉列表中的 **“显示图例字段按钮”** 命令。

**示例**

``` JavaScript
/*下面的代码示例将 Chart 1 设置为显示图例字段按钮。*/
function test()
{
    Application.ActiveSheet.ChartObjects("Chart 1").Activate()
    Application.ActiveChart.ShowLegendFieldButtons = true
}
```

### Chart.ShowReportFilterFieldButtons

返回或设置是否要在数据透视图上显示报表筛选字段按钮。可读/写。

返回值类型为 **Boolean。**

**语法**

**express.ShowReportFilterFieldButtons**

\*express是一个代表 Chart 对象的变量。

**说明**

将 **ShowReportFilterFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示 **“报表筛选字段”** 按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowReportFilterFieldButtons** 属性对应于 **“分析”** 选项卡（在选择数据透视图时可用）的 **“字段按钮”** 下拉列表中的 **“显示报表筛选字段按钮”** 命令。

**示例**

``` JavaScript
/*下面的代码示例将 Chart 1 设置为显示报表筛选字段按钮。*/
function test()
{
    Application.ActiveSheet.ChartObjects("Chart 1").Activate()
    Application.ActiveChart.ShowReportFilterFieldButtons = true
}
```

### Chart.ShowValueFieldButtons

返回或设置是否要在数据透视图上显示值字段按钮。可读/写。

返回值类型为 **Boolean。**

**语法**

**express.ShowValueFieldButtons**

\*express是一个代表 Chart 对象的变量。

**说明**

将 **ShowValueFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示 **“值字段”** 按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowValueFieldButtons** 属性对应于 **“分析”** 选项卡（在选择数据透视图时可用）的 **“字段按钮”** 下拉列表中的 **“显示值字段按钮”** 命令。

**示例**

``` JavaScript
/*下面的代码示例将 Chart 1 设置为显示值字段按钮。*/
function test()
{
    Application.ActiveSheet.ChartObjects("Chart 1").Activate()
    Application.ActiveChart.ShowValueFieldButtons = true
}
```

### Chart.SideWall

返回 **Walls** 对象，该对象允许用户单独对三维图表的侧面墙进行格式设置。只读。

**语法**

**express.SideWall**

\*express是一个代表 Chart 对象的变量。

### Chart.Tab

为图表返回一个 **Tab** 对象。

**语法**

**express.Tab**

\*express是一个代表 Chart 对象的变量。

### Chart.Visible

返回或设置一个 **XlSheetVisibility** 值，它确定对象是否可见。

**语法**

**express.Visible**

\*express是一个代表 Chart 对象的变量。

### Chart.Walls

返回一个 Walls 对象，该对象表示三维图表的背景墙。只读。

**语法**

**express.Walls**

\*express是一个代表 Chart 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 的背景墙边框颜色设置为红色。本示例应在三维图表上运行。*/
Application.Charts.Item("Chart1").Walls.Border.ColorIndex = 3
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Chart](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ChartGroups 对象

## 简介

代表图表中用同一格式绘制的一个或多个数据系列。

## 说明

ChartGroups 集合是指定图表中的所有 ChartGroup 对象的集合。每张图表包含一个或多个图表组，每个图表组包含一个或多个数据系列，每个数据系列包含一个或多个数据点。例如，单张图表可能既包含折线图图表组（其中包含所有用折线图格式绘制的数据系列），也包含条形图图表组（其中包含所有用条形图格式绘制的数据系列）。

使用 ChartGroups 方法可返回 ChartGroups 集合。以下示例显示工作表 1 上嵌入图表 1 中图表组的个数。

``` JavaScript
alert(Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups().Count)
```

使用 ChartGroups ( *index* )（其中 *index* 是图表组的索引号）可以返回单个 ChartGroup 对象。以下示例向图表工作表 1 上的图表组 1 中添加垂直线。

``` JavaScript
Application.Charts.Item(1).ChartGroups(1).HasDropLines = true
```

如果图表已被激活，可使用 ActiveChart ：

``` JavaScript
Application.Charts.Item(1).Activate()
Application.ActiveChart.ChartGroups(1).HasDropLines = true
```

因为当特定图表组所用的图表格式更改时，该图表组的索引号可能会更改，所以使用命名图表组快捷方法之一来返回特定的图表组会更加容易。 PieGroups 方法返回图表中饼图图表组的集合， LineGroups 方法返回图表中折线图图表组的集合，依此类推。这些方法中的每一个都可以与索引号配合使用以返回单个 ChartGroup 对象，或不指定索引号而返回 ChartGroups 集合。

## 方法

| 序号 | 名称                      | 说明                   |
|:----:|:--------------------------|:-----------------------|
|  1   | [Item](#ChartGroups.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartGroups.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#ChartGroups.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#ChartGroups.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#ChartGroups.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### ChartGroups.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ChartGroups 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 必选      | Variant  | 对象的索引号。 |

**示例**

``` JavaScript
/*此示例为第一个图表工作表中的第一个图表组添加垂直线。*/
Application.Charts.Item(1).ChartGroups.Item(1).HasDropLines = true
```

## 成员属性

### ChartGroups.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ChartGroups 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
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

### ChartGroups.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 ChartGroups 对象的变量。

### ChartGroups.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartGroups 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ChartGroups.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartGroups 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartGroups](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# DownBars 对象

## 简介

代表图表组中的跌柱线。

## 说明

跌柱线将图表组中第一个系列的数据点与最后一个系列中相应的有较小值的数据点连接起来（从第一个系列向下生长）。只有至少包含两个系列的二维折线图才能有跌柱线。此对象不是集合。没有代表单个跌柱线的对象；或者打开图表组中所有数据点的涨跌柱线，或者将其全部关闭。

如果 HasUpDownBars 属性为 False ， DownBars 对象的绝大部分属性将被禁用。

使用 DownBars 属性可返回 DownBars 对象。下例打开工作表“Sheet5”上嵌入的第一个图表中第一个图表组的涨跌柱线，然后将涨柱线的颜色设置为蓝色，而将跌柱线设置为红色。

``` JavaScript
function test(){
let downbars = Application.Worksheets.Item("Sheet5").ChartObjects(1).Chart.ChartGroups(1)
    downbars.HasUpDownBars = true
    downbars.UpBars.Interior.Color = (0, 0, 255)
    downbars.DownBars.Interior.Color = (255, 0, 0)
}
```

## 方法

| 序号 | 名称                       | 说明       |
|:----:|:---------------------------|:-----------|
|  1   | [Delete](#DownBars.Delete) | 删除对象。 |
|  2   | [Select](#DownBars.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#DownBars.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#DownBars.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Format](#DownBars.Format)           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  4   | [Name](#DownBars.Name)               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  5   | [Parent](#DownBars.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### DownBars.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 DownBars 对象的变量。

**返回值**

Variant

### DownBars.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 DownBars 对象的变量。

**返回值**

Variant

## 成员属性

### DownBars.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 DownBars 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### DownBars.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DownBars 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### DownBars.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 DownBars 对象的变量。

### DownBars.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 DownBars 对象的变量。

### DownBars.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DownBars 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/DownBars](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FillFormat 对象

## 简介

代表形状的填充格式。

## 说明

形状可以有纯色、渐变、纹理、图案、图片或半透明填充。

FillFormat 对象的很多属性是只读的。要设置这些属性中每一个，必须使用相应的方法。

使用 Fill 属性可返回 FillFormat 对象。

``` JavaScript
/*下例向 myDocument 中添加矩形并且设置矩形填充的渐变和颜色。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shapes2 = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill
shapes2.ForeColor.RGB = (0, 128, 128)
shapes2.OneColorGradient(msoGradientHorizontal, 1, 1)
}
```

## 方法

| 序号 | 名称                                             | 说明                                                                                  |
|:----:|:-------------------------------------------------|:--------------------------------------------------------------------------------------|
|  1   | [OneColorGradient](#FillFormat.OneColorGradient) | 将指定填充设为单色渐变。                                                              |
|  2   | [Patterned](#FillFormat.Patterned)               | 将指定填充设为一种图案。                                                              |
|  3   | [PresetGradient](#FillFormat.PresetGradient)     | 将指定填充设为一个预设的渐变。                                                        |
|  4   | [PresetTextured](#FillFormat.PresetTextured)     | 将指定填充格式设为预设纹理。                                                          |
|  5   | [Solid](#FillFormat.Solid)                       | 将指定填充设为统一的颜色。使用此方法可将渐变、纹理、图案或背景填充转换为纯色填充。    |
|  6   | [TwoColorGradient](#FillFormat.TwoColorGradient) | 将指定填充设为双色渐变。                                                              |
|  7   | [UserPicture](#FillFormat.UserPicture)           | 用图像填充指定的形状。                                                                |
|  8   | [UserTextured](#FillFormat.UserTextured)         | 用平铺的小图像填充指定形状。如果要用一个大图像来填充该形状，请使用 UserPicture 方法。 |

## 属性

| 序号 | 名称                                                         | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FillFormat.Application)                       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [BackColor](#FillFormat.BackColor)                           | 返回或设置一个 ColorFormat 对象，它代表指定的填充背景色。                                                                                                                                                                       |
|  3   | [Creator](#FillFormat.Creator)                               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [ForeColor](#FillFormat.ForeColor)                           | 返回或设置一个 ColorFormat 对象，它代表指定的前景填充色或纯色。                                                                                                                                                                 |
|  5   | [GradientAngle](#FillFormat.GradientAngle)                   | 返回或设置指定填充格式的渐变填充角度。可读/写。                                                                                                                                                                                 |
|  6   | [GradientColorType](#FillFormat.GradientColorType)           | 返回指定填充的渐变颜色类型。 MsoGradientColorType ，只读。                                                                                                                                                                      |
|  7   | [GradientDegree](#FillFormat.GradientDegree)                 | 以浮点数值的方式返回单色阴影填充的渐变程度，数值大小介于 0.0（暗）到 1.0（亮）之间。 Single 型，只读。                                                                                                                          |
|  8   | [GradientStops](#FillFormat.GradientStops)                   | 返回渐变填充的终点。只读。                                                                                                                                                                                                      |
|  9   | [GradientStyle](#FillFormat.GradientStyle)                   | 返回指定填充的渐变样式。 MsoGradientStyle ，只读。                                                                                                                                                                              |
|  10  | [GradientVariant](#FillFormat.GradientVariant)               | 以整数值（1 至 4）形式返回指定填充的底纹变量。此属性的值对应于 “填充效果” 对话框中 “渐变” 选项卡的渐变变量（从左到右、从上到下编号）。 Long 型，只读。                                                                          |
|  11  | [Parent](#FillFormat.Parent)                                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  12  | [Pattern](#FillFormat.Pattern)                               | 返回或设置一个代表填充图案的 MsoPatternType 值。                                                                                                                                                                                |
|  13  | [PictureEffects](#FillFormat.PictureEffects)                 | 返回一个对象，该对象表示指定的填充格式的图片或纹理填充。只读                                                                                                                                                                    |
|  14  | [PresetGradientType](#FillFormat.PresetGradientType)         | 返回指定填充的预设渐变类型。 MsoPresetGradientType ，只读。                                                                                                                                                                     |
|  15  | [PresetTexture](#FillFormat.PresetTexture)                   | 返回指定填充的预设纹理。 MsoPresetTexture ，只读。                                                                                                                                                                              |
|  16  | [RotateWithObject](#FillFormat.RotateWithObject)             | 返回或设置是否随对象旋转。可读/写 MsoTriState 类型。                                                                                                                                                                            |
|  17  | [TextureAlignment](#FillFormat.TextureAlignment)             | 返回或设置指定 FillFormat 对象的文本对齐方式。可读/写。                                                                                                                                                                         |
|  18  | [TextureHorizontalScale](#FillFormat.TextureHorizontalScale) | 返回或设置 FillFormat 对象的横向文本缩放值。可读/写 Single 类型。                                                                                                                                                               |
|  19  | [TextureName](#FillFormat.TextureName)                       | 返回指定填充的自定义纹理文件名称。 String 型，只读。                                                                                                                                                                            |
|  20  | [TextureOffsetX](#FillFormat.TextureOffsetX)                 | 返回指定填充的偏移 X 值。可读/写 Single 类型。                                                                                                                                                                                  |
|  21  | [TextureOffsetY](#FillFormat.TextureOffsetY)                 | 返回指定填充的偏移 Y 值。可读/写 Single 类型。                                                                                                                                                                                  |
|  22  | [TextureTile](#FillFormat.TextureTile)                       | 返回指定填充的纹理平铺样式。可读/写 MsoTriState 类型。                                                                                                                                                                          |
|  23  | [TextureType](#FillFormat.TextureType)                       | 返回指定填充的纹理类型。 MsoTextureType ，只读。                                                                                                                                                                                |
|  24  | [TextureVerticalScale](#FillFormat.TextureVerticalScale)     | 返回指定填充的纹理垂直缩放。可读/写 Single 类型。                                                                                                                                                                               |
|  25  | [Transparency](#FillFormat.Transparency)                     | 返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。 Double 型，可读写。                                                                                                                                    |
|  26  | [Type](#FillFormat.Type)                                     | 返回代表填充类型的 MsoFillType 值。                                                                                                                                                                                             |
|  27  | [Visible](#FillFormat.Visible)                               | 返回或设置一个 MsoTriState 值，它确定对象是否可见。可读写。                                                                                                                                                                     |

## 成员方法

### FillFormat.OneColorGradient

将指定填充设为单色渐变。

**语法**

**express.OneColorGradient ( *Style* , *Variant* , *Degree* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型         | 说明                                                                                                                                                                          |
|-----------|-----------|------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*   | 必选      | MsoGradientStyle | 渐变样式。                                                                                                                                                                    |
| *Variant* | 必选      | Integer          | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 GradientStyle 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |
| *Degree*  | 必选      | Single           | 渐变程度。可以为 0.0（暗）到 1.0（亮）之间的值。                                                                                                                              |

### FillFormat.Patterned

将指定填充设为一种图案。

**语法**

**express.Patterned ( *Pattern* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型       | 说明         |
|-----------|-----------|----------------|--------------|
| *Pattern* | 必选      | MsoPatternType | 图案的类型。 |

### FillFormat.PresetGradient

将指定填充设为一个预设的渐变。

**语法**

**express.PresetGradient ( *Style* , *Variant* , *PresetGradientType* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型              | 说明                                                                                                                                                                          |
|----------------------|-----------|-----------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*              | 必选      | MsoGradientStyle      | 渐变样式。                                                                                                                                                                    |
| *Variant*            | 必选      | Integer               | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 GradientStyle 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |
| *PresetGradientType* | 必选      | MsoPresetGradientType | 预设渐变类型。                                                                                                                                                                |

### FillFormat.PresetTextured

将指定填充格式设为预设纹理。

**语法**

**express.PresetTextured ( *PresetTexture* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型         | 说明               |
|-----------------|-----------|------------------|--------------------|
| *PresetTexture* | 必选      | MsoPresetTexture | 要应用的纹理类型。 |

**示例**

``` JavaScript
/*此示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillTextured) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    if(c1f.TextureType == msoTexturePreset) {
        c2f.PresetTextured(c1f.PresetTexture)
    }
    else {
         c2f.UserTextured(c1f.TextureName)
    }
}
}
```

### FillFormat.Solid

将指定填充设为统一的颜色。使用此方法可将渐变、纹理、图案或背景填充转换为纯色填充。

**语法**

**express.Solid ()**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*此示例将 myDocument 中的所有填充转换为统一的红色填充。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
for(let s = 1; s <= myDocument.Shapes.Count; s++) {
    let s2 = myDocument.Shapes.Item(s).Fill
        s2.Solid()
        s2.ForeColor.RGB = (255, 0, 0)
}
}
```

### FillFormat.TwoColorGradient

将指定填充设为双色渐变。

**语法**

**express.TwoColorGradient ( *Style* , *Variant* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型         | 说明                                                                                                                                                                  |
|-----------|-----------|------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*   | 必选      | MsoGradientStyle | 渐变样式。                                                                                                                                                            |
| *Variant* | 必选      | Integer          | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 Style 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |

### FillFormat.UserPicture

用图像填充指定的形状。

**语法**

**express.UserPicture ( *PictureFile* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明         |
|---------------|-----------|----------|--------------|
| *PictureFile* | 必选      | String   | 图片文件名。 |

### FillFormat.UserTextured

用平铺的小图像填充指定形状。如果要用一个大图像来填充该形状，请使用 **UserPicture** 方法。

**语法**

**express.UserTextured ( *TextureFile* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明         |
|---------------|-----------|----------|--------------|
| *TextureFile* | 必选      | String   | 图片文件名。 |

**示例**

``` JavaScript
/*本示例为第二个图表设置填充格式。*/
Application.Charts.Item(2).ChartArea.Fill.UserTextured("brick.gif")
```

## 成员属性

### FillFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### FillFormat.BackColor

返回或设置一个 **ColorFormat** 对象，它代表指定的填充背景色。

**语法**

**express.BackColor**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FillFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### FillFormat.ForeColor

返回或设置一个 **ColorFormat** 对象，它代表指定的前景填充色或纯色。

**语法**

**express.ForeColor**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.GradientAngle

返回或设置指定填充格式的渐变填充角度。可读/写。

**语法**

**express.GradientAngle**

\*express是一个代表 FillFormat 对象的变量。

**说明**

可以在图表中各种元素（形状）的格式设置中指定渐变填充。例如，可使用 **“设置数据系列格式”** 对话框将 **“列”** 图表中各列的格式设置为渐变填充。在此情况下， **GradientAngle** 属性对应于 **“设置数据系列格式”** 对话框内 **“填充”** 类别中的 **“角度”** 框的设置。 **GradientAngle** 属性值的有效范围是 0 到 359.9。

**示例**

``` JavaScript
/*下面的代码示例将 Chart 1 中第 1 个系列的渐变填充角度设置为 45 度。*/
function test(){
Application.ActiveSheet.ChartObjects("Chart 1").Activate()
Application.ActiveChart.SeriesCollection(1).Format.Fill.GradientAngle = 45
}
```

### FillFormat.GradientColorType

返回指定填充的渐变颜色类型。 **MsoGradientColorType** ，只读。

**语法**

**express.GradientColorType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

**MsoGradientColorMixed** 是一个返回值，它仅指明指定区域中其他状态的组合。使用 **OneColorGradient** 、 **PresetGradient** 或 **TwoColorGradient** 方法可设置填充的渐变类型。

**示例**

``` JavaScript
/*此示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillGradient && c1f.GradientColorType == msoGradientOneColor) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    c2f.OneColorGradient(c1f.GradientStyle, c1f.GradientVariant, c1f.GradientDegree)
}
}
```

### FillFormat.GradientDegree

以浮点数值的方式返回单色阴影填充的渐变程度，数值大小介于 0.0（暗）到 1.0（亮）之间。 **Single** 型，只读。

**语法**

**express.GradientDegree**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性为只读。使用 **OneColorGradient** 方法可设置填充的渐变程度。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillGradient && c1f.GradientColorType == msoGradientOneColor) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    c2f.OneColorGradient(c1f.GradientStyle, c1f.GradientVariant, c1f.GradientDegree)
}
}
```

### FillFormat.GradientStops

返回渐变填充的终点。只读。

**语法**

**express.GradientStops**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.GradientStyle

返回指定填充的渐变样式。 **MsoGradientStyle** ，只读。

**语法**

**express.GradientStyle**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.GradientVariant

以整数值（1 至 4）形式返回指定填充的底纹变量。此属性的值对应于 **“填充效果”** 对话框中 **“渐变”** 选项卡的渐变变量（从左到右、从上到下编号）。 **Long** 型，只读。

**语法**

**express.GradientVariant**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性为只读。使用 **OneColorGradient** 或 **TwoColorGradient** 方法可设置填充的渐变变量。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillGradient && c1f.GradientColorType == msoGradientOneColor) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    c2f.OneColorGradient(c1f.GradientStyle, c1f.GradientVariant, c1f.GradientDegree)
}
}
```

### FillFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.Pattern

返回或设置一个代表填充图案的 **MsoPatternType** 值。

**语法**

**express.Pattern**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **Patterned** 方法可设置填充的图案类型。

### FillFormat.PictureEffects

返回一个对象，该对象表示指定的填充格式的图片或纹理填充。只读

**语法**

**express.PictureEffects**

\*express是一个代表 FillFormat 对象的变量。

**说明**

可以在图表中各种元素（形状）的格式设置中指定图片或纹理填充。例如，可使用 **“设置数据系列格式”** 对话框将 **“列”** 图表中各列的格式设置为图片或纹理填充。在此情况下， **PictureEffects** 属性返回的 **PictureEffects** 集合对应于 **“设置数据系列格式”** 对话框的 **“填充”** 类别中 **“图片或纹理填充”** 选项的关联设置。

### FillFormat.PresetGradientType

返回指定填充的预设渐变类型。 **MsoPresetGradientType** ，只读。

**语法**

**express.PresetGradientType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **PresetGradient** 方法可设置填充的预设渐变类型。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillGradient) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    c2f.PresetGradient(c1f.GradientStyle, c1f.GradientVariant, c1f.PresetGradientType)
}
}
```

### FillFormat.PresetTexture

返回指定填充的预设纹理。 **MsoPresetTexture** ，只读。

**语法**

**express.PresetTexture**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **PresetTextured** 方法可设置填充的预设纹理。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillTextured) {
    let c2f = Charts.Item(2).ChartArea.Fill
        c2f.Visible = true
        if(c1f.TextureType == msoTexturePreset) {
            c2f.PresetTextured(c1f.PresetTexture)
        }
        else {
            c2f.UserTextured(c1f.TextureName)
        }
}
}
```

### FillFormat.RotateWithObject

返回或设置是否随对象旋转。可读/写 **MsoTriState** 类型。

**语法**

**express.RotateWithObject**

\*express是一个代表 FillFormat 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| MsoTriState 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue**                                      |
| **msoFalse**                                      |
| **msoTriStateMixed**                              |
| **msoTriStateToggle**                             |
| **msoTrue**                                       |

### FillFormat.TextureAlignment

返回或设置指定 **FillFormat** 对象的文本对齐方式。可读/写。

**语法**

**express.TextureAlignment**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.TextureHorizontalScale

返回或设置 **FillFormat** 对象的横向文本缩放值。可读/写 **Single** 类型。

**语法**

**express.TextureHorizontalScale**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.TextureName

返回指定填充的自定义纹理文件名称。 **String** 型，只读。

**语法**

**express.TextureName**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **UserPicture** 或 **UserTextured** 方法可设置填充的纹理文件。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillTextured) {
    let c2f = Charts.Item(2).ChartArea.Fill
        c2f.Visible = true
        if(c1f.TextureType == msoTexturePreset) {
            c2f.PresetTextured(c1f.PresetTexture)
        }
        else {
            c2f.UserTextured(c1f.TextureName)
        }
}
}
```

### FillFormat.TextureOffsetX

返回指定填充的偏移 X 值。可读/写 **Single** 类型。

**语法**

**express.TextureOffsetX**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.TextureOffsetY

返回指定填充的偏移 Y 值。可读/写 **Single** 类型。

**语法**

**express.TextureOffsetY**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.TextureTile

返回指定填充的纹理平铺样式。可读/写 **MsoTriState** 类型。

**语法**

**express.TextureTile**

\*express是一个代表 FillFormat 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| MsoTriState 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue**                                      |
| **msoFalse**                                      |
| **msoTriStateMixed**                              |
| **msoTriStateToggle**                             |
| **msoTrue**                                       |

### FillFormat.TextureType

返回指定填充的纹理类型。 **MsoTextureType** ，只读。

**语法**

**express.TextureType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **UserTextured** 方法可设置填充的纹理类型。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillTextured) {
    let c2f = Charts.Item(2).ChartArea.Fill
        c2f.Visible = true
        if(c1f.TextureType == msoTexturePreset) {
            c2f.PresetTextured(c1f.PresetTexture)
        }
        else {
            c2f.UserTextured(c1f.TextureName)
        }
}
}
```

### FillFormat.TextureVerticalScale

返回指定填充的纹理垂直缩放。可读/写 **Single** 类型。

**语法**

**express.TextureVerticalScale**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.Transparency

返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。 **Double** 型，可读写。

**语法**

**express.Transparency**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

### FillFormat.Type

返回代表填充类型的 **MsoFillType** 值。

**语法**

**express.Type**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性也可返回 **xlAutomatic** 或 **xlNone** 。

在 ET 中不使用 **msoFillBackground** 常量。

### FillFormat.Visible

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 FillFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/FillFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Filter 对象

## 简介

代表单个列的筛选。

## 说明

Filter 对象是 Filters 集合的成员。 Filters 集合包含自动筛选区域中的所有筛选。

使用 Filters ( *index* )（其中 *index* 是筛选的标题或索引号）可返回单个 Filter 对象。下例将一个变量设置为工作表 Crew 上的筛选区域中第一列的筛选的 On 属性的值。

``` JavaScript
function test()
{
    let w = Application.Worksheets.Item("Crew")
    if(w.AutoFilterMode) {
        filterIsOn = w.AutoFilter.Filters.Item(1).On
    }
}
```

注意： Filter 对象的所有属性都是只读的。要设置这些属性，请手动应用自动筛选，或使用 Range 对象的 Aut oFilter 方法，如下例中所示。

``` JavaScript
function test() {
    let w = Application.Worksheets.Item("Crew")     
    w.Cells.AutoFilter(2, "Crucial", xlOr, "Important")
}
```

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Filter.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Filter.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                                                                                          |
|  3   | [Creator](#Filter.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Criteria1](#Filter.Criteria1)     | 返回筛选区域内指定列上的第一个筛选值。 Variant 类型，只读。                                                                                                                                                                     |
|  5   | [Criteria2](#Filter.Criteria2)     | 返回已筛选区域中指定列的第二个筛选值。 Variant 类型，只读。                                                                                                                                                                     |
|  6   | [On](#Filter.On)                   | 如果打开指定的筛选，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                                   |
|  7   | [Operator](#Filter.Operator)       | 返回一个 XlAutoFilterOperator 值，它代表与指定筛选所应用的两个条件相关的操作符。                                                                                                                                                |
|  8   | [Parent](#Filter.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员属性

### Filter.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Filter 对象的变量。

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

### Filter.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 Filter 对象的变量。

### Filter.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Filter 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Filter.Criteria1

返回筛选区域内指定列上的第一个筛选值。 **Variant** 类型，只读。

**语法**

**express.Criteria1**

\*express是一个代表 Filter 对象的变量。

**示例**

``` JavaScript
/* 下例将变量设置为工作表“Crew”上的筛选区域中第一列的筛选的 Criteria1 属性值*/
function test()
{
    let c1 = Application.Worksheets.Item("Crew")
    if(c1.AutoFilterMode) {
        let c2 = c1.AutoFilter.Filters.Item(1)
            if(c2.On) { 
                c1 = c2.Criteria1
            }
    }
}
```

### Filter.Criteria2

返回已筛选区域中指定列的第二个筛选值。 **Variant** 类型，只读。

**语法**

**express.Criteria2**

\*express是一个代表 Filter 对象的变量。

**说明**

使用筛选的 **Criteria2** 属性时，如果该筛选没有使用两个条件，则将出现错误。在试图使用 **Operator** 属性之前，需保证 **Filter** 对象的 **Criteria2** 属性不为零 (0)。

**示例**

``` JavaScript
/* 下面的示例将一个变量设置为 Criteria2 属性值。该属性是 Crew 工作表上已筛选区域中的第一列筛选的*/
function test()
{
    let c1 = Application.Worksheets.Item("Crew")
    if(c1.AutoFilterMode) {
        let c2 = c1.AutoFilter.Filters.Item(1)
            if(c2.On && c2.Operator) {
                let c3 = c2.Criteria2
            }
            else {
                let c3 = "Not set"
            }
    }
}
```

### Filter.On

如果打开指定的筛选，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.On**

\*express是一个代表 Filter 对象的变量。

**示例**

``` JavaScript
/* 下例将变量设置为工作表“Crew”上的筛选区域中第一列的筛选的 Criteria1 属性值*/
function test()
{
    let on1 = Application.Worksheets.Item("Crew")
    if(on1.AutoFilterMode) {
        let on2 = on1.AutoFilter.Filters.Item(1)
            if(on2.On) {
                let c1 = on2.Criteria1
            }
    }
}
```

### Filter.Operator

返回一个 **XlAutoFilterOperator** 值，它代表与指定筛选所应用的两个条件相关的操作符。

**语法**

**express.Operator**

\*express是一个代表 Filter 对象的变量。

### Filter.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Filter 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Filter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListColumn 对象

## 简介

代表表格中的一列。

## 说明

ListColumn 对象是 ListColumns 集合的成员。 ListColumns 集合包含表格中的所有列（ ListObject 对象）。

使用 ListObject 对象的 ListColumns 属性可返回一个 ListColumns 集合。

``` JavaScript
/*下例给活动工作簿的第一个工作表的默认 ListObject 对象添加新 ListColumn 对象。由于未指定位置，因此在最右边添加一个新列。*/
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Add()
}
```

## 方法

| 序号 | 名称                         | 说明                 |
|:----:|:-----------------------------|:---------------------|
|  1   | [Delete](#ListColumn.Delete) | 在列表中删除数据列。 |

## 属性

| 序号 | 名称                                               | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ListColumn.Application)             | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#ListColumn.Creator)                     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [DataBodyRange](#ListColumn.DataBodyRange)         | 返回一个 Range 对象，该对象为列的数据部分的大小。只读。                                                                                                                                                                         |
|  4   | [Index](#ListColumn.Index)                         | 返回一个 Long 值，它代表 ListColumn 对象在 ListColumns 集合中的索引号。                                                                                                                                                         |
|  5   | [Name](#ListColumn.Name)                           | 返回或设置一个 String 值，它代表列表列的名称。                                                                                                                                                                                  |
|  6   | [Parent](#ListColumn.Parent)                       | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  7   | [Range](#ListColumn.Range)                         | 返回一个 Range 对象，它代表在以上列表中的指定列表对象所应用的区域。                                                                                                                                                             |
|  8   | [Total](#ListColumn.Total)                         | 返回 ListColumn 对象的总计行。只读。                                                                                                                                                                                            |
|  9   | [TotalsCalculation](#ListColumn.TotalsCalculation) | 基于 XlTotalsCalculation 枚举的值，确定列表列的“总计”行中的计算类型。可读写。                                                                                                                                                   |
|  10  | [XPath](#ListColumn.XPath)                         | 返回一个 XPath 对象，它代表映射到指定 Range 对象的元素的 XPath。该区域的上下文确定操作是否成功，或返回空对象。只读。                                                                                                            |

## 成员方法

### ListColumn.Delete

在列表中删除数据列。

**语法**

**express.Delete ()**

\*express是一个代表 ListColumn 对象的变量。

**说明**

此方法不从工作表中删除列。如果列表链接到 Microsoft SharePoint Foundation 网站，则无法从服务器中删除该列，并产生一个错误。

## 成员属性

### ListColumn.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ListColumn 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### ListColumn.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ListColumn 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ListColumn.DataBodyRange

返回一个 **Range** 对象，该对象为列的数据部分的大小。只读。

**语法**

**express.DataBodyRange**

\*express是一个代表 ListColumn 对象的变量。

**说明**

返回的对象不包含标题单元格和总计单元格。

### ListColumn.Index

返回一个 **Long** 值，它代表 **ListColumn** 对象在 **ListColumns** 集合中的索引号。

**语法**

**express.Index**

\*express是一个代表 ListColumn 对象的变量。

### ListColumn.Name

返回或设置一个 **String** 值，它代表列表列的名称。

**语法**

**express.Name**

\*express是一个代表 ListColumn 对象的变量。

**说明**

该属性值还被用作列表列的显示名称。此名称必须在列表中唯一。

### ListColumn.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ListColumn 对象的变量。

### ListColumn.Range

返回一个 **Range** 对象，它代表在以上列表中的指定列表对象所应用的区域。

**语法**

**express.Range**

\*express是一个代表 ListColumn 对象的变量。

### ListColumn.Total

返回 **ListColumn** 对象的总计行。只读。

**语法**

**express.Total**

\*express是一个代表 ListColumn 对象的变量。

### ListColumn.TotalsCalculation

基于 **XlTotalsCalculation** 枚举的值，确定列表列的“总计”行中的计算类型。可读写。

**语法**

**express.TotalsCalculation**

\*express是一个代表 ListColumn 对象的变量。

**说明**

|                                                                     |
|---------------------------------------------------------------------|
| **XlTotalsCalculation** 可为以下 **XlTotalsCalculation** 常量之一。 |
| **xlTotalsCalculationNone**                                         |
| **xlTotalsCalculationSum**                                          |
| **xlTotalsCalculationAverage**                                      |
| **xlTotalsCalculationCount**                                        |
| **xlTotalsCalculationCountNums**                                    |
| **xlTotalsCalculationMin**                                          |
| **xlTotalsCalculationStdDev**                                       |
| **xlTotalsCalculationVar**                                          |
| **xlTotalsCalculationMax**                                          |

无须显示“总计”行即可设置此属性。此属性没有固定的“默认”值。随着对其他列的添加或删除，ET 会更改该属性的状态。

**示例**

``` JavaScript
Application.ActiveSheet.ListColumns.Item(1).TotalsCalculation=xlTotalsCalculationSum
```

### ListColumn.XPath

返回一个 **XPath** 对象，它代表映射到指定 **Range** 对象的元素的 XPath。该区域的上下文确定操作是否成功，或返回空对象。只读。

**语法**

**express.XPath**

\*express是一个代表 ListColumn 对象的变量。

**说明**

当 **XPath** 属性包含的区域满足以下条件时，该属性有效：

- 区域中仅有一个单元格。
- 如果区域中包含两个或更多个单元格，那么必须满足以下条件之一：
  1.  如果单元格包含 XPath 信息，区域内的所有单元格都必须包含 XPath 信息（也即，每个单元格与一个或多个数据映射关联），所有单元格还必须具有相同的 XPath 内容（也即，每个单元格向同一组数据映射提供数据）。
  2.  所有单元格都不包含 XPath 信息。
- 区域内没有不连续的区域。
  | 注释                                      |
  |-------------------------------------------|
  | 表格的标题和总计行被视为包含 XPath 信息。 |

任何不满足上述条件的区域将返回运行时错误。

如果区域选择有效，但未映射任何单元格，ET 将返回一个 **XPath** 对象，这样您就能通过访问 **SetValue** 方法来创建映射。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ListColumn](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Name 对象

## 简介

代表单元格区域的定义名。名称可以是内置名称（如“Database”、“Print_Area”和“Auto_Open”）或自定义名称。

## 说明

应用程序、工作簿和 Worksheet 对象

Name 对象是 Application 、 Workbook 和 Worksheet 对象的 Names 集合的成员。使用 Names ( *index* )（其中 *index* 是名称索引号或定义名称）可返回一个 Name 对象。

索引号表明名称在集合中的位置。名称按字母顺序从 a 到 z 放置，不区分大小写。

Range 对象

虽然 Range 对象可以有多个名称，但 Range 对象没有 Names 集合。将 Name 与一个 Range 对象一起使用可从名称列表（按字母顺序排序）中返回第一个名称。下例为指定给工作表一上单元格 A1:B1 的第一个名称设置 Visible 属性。

## 示例

本示例显示应用程序集合中第一个名称的单元格引用。

``` JavaScript
function test() {
    alert(Application.Application.Names.Item(1).RefersTo) 
} 
```

下例从活动工作簿中删除名称“mySortRange”。

``` JavaScript
function test() {
    Application.ActiveWorkbook.Names.Item("mySortRange").Delete()
} 
```

使用 Name 属性可返回或设置名称本身的文本。本示例更改活动工作簿中第一个 Name 对象的名称。

``` JavaScript
function test() {
    Application.Names.Item(1).Name = "stock_values"
} 
```

下例为指定给工作表一上单元格 A1:B1 的第一个名称设置 Visible 属性。

``` JavaScript
function test() {
    Application.Worksheets.Item(1).Range("a1:b1").Name.Visible = false
} 
```

## 方法

| 序号 | 名称                   | 说明       |
|:----:|:-----------------------|:-----------|
|  1   | [Delete](#Name.Delete) | 删除对象。 |

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                                                                                                        |
|:----:|:-------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Name.Application)     | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 Application 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 Application 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
|  2   | [Category](#Name.Category)           | 返回或设置指定名称在宏语言中的分类。该名称必须针对一个自定义函数或命令。 String 类型，可读写。                                                                                                                              |
|  3   | [Comment](#Name.Comment)             | 返回或设置与名称关联的批注。可读/写 String 类型。                                                                                                                                                                           |
|  4   | [Index](#Name.Index)                 | 返回 Number 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                              |
|  5   | [Name](#Name.Name)                   | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                |
|  6   | [RefersTo](#Name.RefersTo)           | 用宏语言以 A1 样式表示法返回或设置名称所引用的公式（以等号开头）。 String 类型，可读写。                                                                                                                                    |
|  7   | [RefersToR1C1](#Name.RefersToR1C1)   | 返回或设置指定名称所引用的公式。公式以等号开头，由宏语言和 R1C1-样式引用组成。 String 类型，可读写。                                                                                                                        |
|  8   | [RefersToRange](#Name.RefersToRange) | 返回由 Name 对象引用的 Range 对象。只读。                                                                                                                                                                                   |
|  9   | [Value](#Name.Value)                 | 返回或设置一个 String 值，它代表规定名称去引用的公式。                                                                                                                                                                      |
|  10  | [Visible](#Name.Visible)             | 返回或设置一个 Boolean 值，它确定对象是否可见。可读写。                                                                                                                                                                     |

## 成员方法

### Name.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Name 对象的变量。

## 成员属性

### Name.Application

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 Name 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息。 */
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "Microsoft Excel") {
        alert("This is a Microsoft Excel Application object.")
    }
    else {
        alert("This is not a Microsoft Excel Application object.")
    }
}
```

### Name.Category

返回或设置指定名称在宏语言中的分类。该名称必须针对一个自定义函数或命令。 **String** 类型，可读写。

**语法**

**express.Category**

\*express是一个代表 Name 对象的变量。

**示例**

``` JavaScript
/* 本示例假定已在 ET 4.0 宏表中创建了一个自定义函数或命令。本示例显示该名称在宏语言中的分类。假定该自定义函数或命令的名称在指定工作簿中唯一。*/
function test() {
    if (Application.ActiveWorkbook.Names.Item('test').MacroType != Application.Enum.xlNone) {
        alert("The category for this name is " + Application.ActiveWorkbook.Names.Item('test').Category)
    }
    else {
        alert("This name does not refer to a custom function or command.")
    }
}
```

### Name.Comment

返回或设置与名称关联的批注。可读/写 **String** 类型。

**语法**

**express.Comment**

\*express是一个代表 Name 对象的变量。

**说明**

批注文本不能超过 255 个字符。

### Name.Index

返回 **Number** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 Name 对象的变量。

### Name.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Name 对象的变量。

### Name.RefersTo

用宏语言以 A1 样式表示法返回或设置名称所引用的公式（以等号开头）。 **String** 类型，可读写。

**语法**

**express.RefersTo**

\*express是一个代表 Name 对象的变量。

**示例**

``` JavaScript
/* 本示例创建活动工作簿中所有名称的列表，并用宏语言以 A1 样式表示法表示这些名称所引用的公式。该列表将出现在本示例新建的工作表中。 */
function test() {
    let newSheet = Application.Worksheets.Add()
    let i = 1
    for (let nm = 1; nm <= Application.ActiveWorkbook.Names.Count; nm++) {
        newSheet.Cells.Item(i, 1).Value2 = Application.ActiveWorkbook.Names.Item(nm).Name
        newSheet.Cells.Item(i, 2).Value2 = "'" + Application.ActiveWorkbook.Names.Item(nm).RefersTo
        i = i + 1
    }
    newSheet.Columns.Item("A:B").AutoFit()
}
```

### Name.RefersToR1C1

返回或设置指定名称所引用的公式。公式以等号开头，由宏语言和 R1C1-样式引用组成。 **String** 类型，可读写。

**语法**

**express.RefersToR1C1**

\*express是一个代表 Name 对象的变量。

**示例**

``` JavaScript
/* 本示例新建一个工作表，并将当前工作簿所有名称的列表插入到新工作表中，名称列表中包括其公式（由 R1C1-样式引用和宏语言组成）。 */
function test() {
    let newSheet = Application.ActiveWorkbook.Worksheets.Add()
    let i = 1
    for (let nm = 1; nm <= Application.ActiveWorkbook.Names.Count; nm++) {
        newSheet.Cells.Item(i, 1).Value2 = Application.ActiveWorkbook.Names.Item(nm).Name
        newSheet.Cells.Item(i, 2).Value2 = "'" + Application.ActiveWorkbook.Names.Item(nm).RefersToR1C1
        i = i + 1
    }
}
```

### Name.RefersToRange

返回由 **Name** 对象引用的 **Range** 对象。只读。

**语法**

**express.RefersToRange**

\*express是一个代表 Name 对象的变量。

**说明**

如果 **Name** 对象并不引用区域（例如，如果该对象引用的是一个常量或公式），则该属性无效。

要更改名称所引用的区域，请使用 **RefersTo** 属性。

**示例**

``` JavaScript
/* 本示例显示当前工作表打印区域中的行数和列数。 */
/* 注释:确保在当前工作簿的活动工作表上建立打印区域。 */
function test() {
    let p = Application.Sheets.Item(Application.ActiveSheet.Name).Names.Item("Print_Area").RefersToRange.Value2
    alert("Print_Area: " + p.length + " rows, " + p[0].length + " columns")
}
```

### Name.Value

返回或设置一个 **String** 值，它代表规定名称去引用的公式。

**语法**

**express.Value**

\*express是一个代表 Name 对象的变量。

### Name.Visible

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 Name 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Name](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# OLEDBConnection 对象

## 简介

代表 OLE DB 连接。

## 说明

OLE DB 连接可存储在 ET 工作簿中。当 Micrososft ET 打开此工作簿时， ET 会创建一个称为 OLEDBConnection 对象的 OLE DB 连接的内存副本。 OLEDBConnection 对象包含与连接相关的信息，例如，要连接到的服务器的名称和要在该服务器上打开的对象的名称。（可选）OLEDBConnection 对象也可以包括身份验证凭据信息或要传递到服务器并执行的命令（例如，要由 SQL Server 执行的 SELECT 语句）。

## 方法

| 序号 | 名称                                              | 说明                                             |
|:----:|:--------------------------------------------------|:-------------------------------------------------|
|  1   | [CancelRefresh](#OLEDBConnection.CancelRefresh)   | 对指定的 OLE DB 连接取消正在进行的所有刷新操作。 |
|  2   | [MakeConnection](#OLEDBConnection.MakeConnection) | 为指定的 OLE DB 连接建立连接。                   |
|  3   | [Reconnect](#OLEDBConnection.Reconnect)           | 先放弃再重新连接指定的连接。                     |
|  4   | [Refresh](#OLEDBConnection.Refresh)               | 刷新 OLE DB 连接。                               |
|  5   | [SaveAsODC](#OLEDBConnection.SaveAsODC)           | 将 OLE DB 连接保存为 WPS Office 数据连接文件。   |

## 属性

| 序号 | 名称                                                                | 说明                                                                                                                                                               |
|:----:|:--------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ADOConnection](#OLEDBConnection.ADOConnection)                     | 如果将数据透视表缓存连接到 OLE DB 数据源，则返回一个 ADO Connection 对象。只读。                                                                                   |
|  2   | [AlwaysUseConnectionFile](#OLEDBConnection.AlwaysUseConnectionFile) | 如果连接文件始终用于建立与数据源的连接，则为 True 。可读/写 Boolean 类型。                                                                                         |
|  3   | [Application](#OLEDBConnection.Application)                         | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  4   | [BackgroundQuery](#OLEDBConnection.BackgroundQuery)                 | 如果 OLE DB 连接的查询是异步执行（在后台执行）的，则为 True 。可读/写 Boolean 类型。                                                                               |
|  5   | [CalculatedMembers](#OLEDBConnection.CalculatedMembers)             | 返回指定连接的 CalculatedMembers 集合。只读                                                                                                                        |
|  6   | [CommandText](#OLEDBConnection.CommandText)                         | 返回或设置指定数据源的命令字符串。可读/写 Variant 类型。                                                                                                           |
|  7   | [CommandType](#OLEDBConnection.CommandType)                         | 返回或设置 XlCmdType 常量之一。可读/写 XlCmdType 。                                                                                                                |
|  8   | [Connection](#OLEDBConnection.Connection)                           | 返回或设置一个包含 OLE DB 设置的字符串，这些设置使 ET 可以连接 OLE DB 数据源。可读/写 Variant 类型。                                                               |
|  9   | [Creator](#OLEDBConnection.Creator)                                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  10  | [EnableRefresh](#OLEDBConnection.EnableRefresh)                     | 如果用户可刷新连接，则为 True 。默认值是 True 。可读/写 Boolean 类型。                                                                                             |
|  11  | [IsConnected](#OLEDBConnection.IsConnected)                         | 如果 MaintainConnection 属性为 True ，则返回 True ，如果它当前未与其源相连，则返回 False 。只读 Boolean 类型。                                                     |
|  12  | [LocalConnection](#OLEDBConnection.LocalConnection)                 | 返回或设置脱机多维数据集文件的连接字符串。可读/写 String 类型。                                                                                                    |
|  13  | [LocaleID](#OLEDBConnection.LocaleID)                               | 返回或设置指定的连接的区域设置标识符。可读/写。                                                                                                                    |
|  14  | [MaintainConnection](#OLEDBConnection.MaintainConnection)           | 如果从刷新操作开始直至关闭工作簿，都一直保留与指定数据源的连接，则返回 True 。可读/写 Boolean 类型。                                                               |
|  15  | [MaxDrillthroughRecords](#OLEDBConnection.MaxDrillthroughRecords)   | 返回或设置要检索的记录的最大数目。可读/写 Long 类型。                                                                                                              |
|  16  | [OLAP](#OLEDBConnection.OLAP)                                       | 如果 OLE DB 连接与联机分析处理 (OLAP) 服务器相连，则返回 True ，只读 Boolean 类型。                                                                                |
|  17  | [Parent](#OLEDBConnection.Parent)                                   | 返回指定对象的父对象。只读。                                                                                                                                       |
|  18  | [RefreshDate](#OLEDBConnection.RefreshDate)                         | 返回上次刷新 OLE DB 连接的日期。只读 Date 类型。                                                                                                                   |
|  19  | [Refreshing](#OLEDBConnection.Refreshing)                           | 如果正在进行指定 OLE DB 连接的后台 OLE DB 查询，则为 True 。可读/写 Boolean 类型。                                                                                 |
|  20  | [RefreshOnFileOpen](#OLEDBConnection.RefreshOnFileOpen)             | 如果每次打开工作簿都自动更新连接，则为 True 。默认值为 False 。可读/写 Boolean 类型。                                                                              |
|  21  | [RefreshPeriod](#OLEDBConnection.RefreshPeriod)                     | 返回或设置两次刷新之间的时间间隔（单位为分钟）。可读/写 Long 类型。                                                                                                |
|  22  | [RetrieveInOfficeUILang](#OLEDBConnection.RetrieveInOfficeUILang)   | 如果要以 Office 用户界面显示语言（如果可用）检索数据和错误，则为 True 。可读/写 Boolean 类型。                                                                     |
|  23  | [RobustConnect](#OLEDBConnection.RobustConnect)                     | 返回或设置 OLE DB 连接与其数据源连接的方式。可读/写 XlRobustConnect 。                                                                                             |
|  24  | [SavePassword](#OLEDBConnection.SavePassword)                       | 如果将 OLE DB 连接字符串中的密码信息保存在连接字符串中，则为 True 。如果删除密码，则为 False 。可读/写 Boolean 类型。                                              |
|  25  | [ServerCredentialsMethod](#OLEDBConnection.ServerCredentialsMethod) | 返回或设置应该用于服务器验证的凭据的类型。可读/写 XlCredentialsMethod 。                                                                                           |
|  26  | [ServerFillColor](#OLEDBConnection.ServerFillColor)                 | 如果在使用指定连接时从服务器中检索 OLAP 服务器的填充色格式，则为 True 。可读/写 Boolean 类型。                                                                     |
|  27  | [ServerFontStyle](#OLEDBConnection.ServerFontStyle)                 | 如果在使用指定连接时从服务器中检索 OLAP 服务器的字体样式格式，则为 True 。可读/写 Boolean 类型。                                                                   |
|  28  | [ServerNumberFormat](#OLEDBConnection.ServerNumberFormat)           | 如果在使用指定连接时从服务器中检索 OLAP 服务器的数字格式，则为 True 。可读/写 Boolean 类型。                                                                       |
|  29  | [ServerSSOApplicationID](#OLEDBConnection.ServerSSOApplicationID)   | 返回或设置单一登录 (SSO) 应用程序标识符，该标识符用于在 SSO 数据库中查找凭据。可读/写 String 类型。                                                                |
|  30  | [ServerTextColor](#OLEDBConnection.ServerTextColor)                 | 如果在使用指定连接时从服务器中检索 OLAP 服务器的文本颜色格式，则为 True 。可读/写 Boolean 类型。                                                                   |
|  31  | [SourceConnectionFile](#OLEDBConnection.SourceConnectionFile)       | 返回或设置一个 String 类型的值，该值指示用于创建连接的 WPS Office 数据连接文件或类似文件。可读/写。                                                                |
|  32  | [SourceDataFile](#OLEDBConnection.SourceDataFile)                   | 返回或设置一个 String 类型的值，该值指示 OLE DB 连接的源数据文件。可读/写。                                                                                        |
|  33  | [UseLocalConnection](#OLEDBConnection.UseLocalConnection)           | 如果 LocalConnection 属性用于指定使 ET 能够连接到数据源的字符串，则为 True 。如果使用 Connection 属性指定的连接字符串，则为 False 。可读/写 Boolean 类型。         |

## 成员方法

### OLEDBConnection.CancelRefresh

对指定的 OLE DB 连接取消正在进行的所有刷新操作。

**语法**

**express.CancelRefresh ()**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

使用 **Refreshing** 属性可以确定当前是否正在进行刷新操作。

### OLEDBConnection.MakeConnection

为指定的 OLE DB 连接建立连接。

**语法**

**express.MakeConnection ()**

\*express是一个代表 OLEDBConnection 对象的变量。

**返回值**

无

**说明**

**MakeConnection** 方法可在断开连接和用户想要重新建立连接时使用。

如果断开连接，则各种对象和方法可能都会返回运行时错误。使用该方法可确保在执行其他对象或方法之前已建立一个连接。

### OLEDBConnection.Reconnect

先放弃再重新连接指定的连接。

**语法**

**express.Reconnect ()**

\*express是一个代表 OLEDBConnection 对象的变量。

**示例**

下面的代码示例将使指定的连接重新连接。

``` JavaScript
ActiveWorkbook.Connections(1).Reconnect()
```

### OLEDBConnection.Refresh

刷新 OLE DB 连接。

**语法**

**express.Refresh ()**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

创建到 OLE DB 数据源的连接时，ET 使用 **Connection** 属性指定的连接字符串。如果指定的连接字符串缺少必需的值，则显示对话框提示用户输入必需的信息。如果 **DisplayAlerts** 属性为 **False** ，则不显示对话框， **Refresh** 方法将失败，并引发“连接信息无效”异常。

ET 成功创建连接之后，将存储完整的连接字符串，这样，以后在同一编辑会话中调用 **Refresh** 方法时就不会显示提示。通过检查 **Connection** 属性的值，可以获得完整的连接字符串。

建立数据库连接后，将检查 SQL 查询的有效性。如果该查询无效， **Refresh** 方法将失败并引发“SQL 语法错误”异常。

如果查询成功完成或成功启动，则 **Refresh** 方法返回 **True** ；如果用户取消连接，则返回 **False** 。

要查看取回的行数是否超过了工作表中的可用行数，请检查 **FetchedRowOverflow** 属性。每次调用 **Refresh** 方法之前，该属性都将被初始化。

### OLEDBConnection.SaveAsODC

将 OLE DB 连接保存为 WPS Office 数据连接文件。

**语法**

**express.SaveAsODC ( *ODCFileName* , *Description* , *Keywords* )**

\*express是一个代表 OLEDBConnection 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                   |
|---------------|-----------|----------|----------------------------------------|
| *ODCFileName* | 必选      | String   | 保存文件的位置。                       |
| *Description* | 可选      | Variant  | 将保存在文件中的说明。                 |
| *Keywords*    | 可选      | Variant  | 可用于搜索该文件的以空格分隔的关键字。 |

**示例**

以下示例将连接保存为名为“ODCFile”的 ODC 文件。此示例假定 OLE DB 连接位于活动工作表上。

``` JavaScript
function UseSaveAsODC(){
    Application.ActiveWorkbook.OLEDBConnection.SaveAsODC ("ODCFile")
}
```

## 成员属性

### OLEDBConnection.ADOConnection

如果将数据透视表缓存连接到 OLE DB 数据源，则返回一个 ADO Connection 对象。只读。

**语法**

**express.ADOConnection**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

**ADOConnection** 属性表明 ET 与数据提供程序相连，允许用户在 ET 正在使用的同一会话的上下文中编写代码。

**ADOConnection** 属性只可用于数据源是 OLE DB 数据源的会话。如果没有 ADO 会话，查询将导致运行时错误。 **ADOConnection** 属性可用于任意包含 ADO 的基于 OLE DB 的缓存。ADO Connection 对象可与 ADO MD 一起使用，以查找有关缓存所基于的 OLAP 多维数据集的信息。

### OLEDBConnection.AlwaysUseConnectionFile

如果连接文件始终用于建立与数据源的连接，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.AlwaysUseConnectionFile**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

当此属性为 **True** 时，连接文件将始终用于建立与数据源的连接。如果在工作簿中嵌入的连接不同于外部连接文件，嵌入的连接将被忽略，外部连接文件将是唯一考虑的版本。

### OLEDBConnection.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### OLEDBConnection.BackgroundQuery

如果 OLE DB 连接的查询是异步执行（在后台执行）的，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.BackgroundQuery**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

对于 OLAP 数据源，此属性是只读的，并且始终返回 **False** 。

### OLEDBConnection.CalculatedMembers

返回指定连接的 **CalculatedMembers** 集合。只读

**语法**

**express.CalculatedMembers**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

使用 **OLEDBConnection** 集合的 **CalculatedMembers** 属性可在使用同一连接的多个数据透视表和多维数据集函数之间共享计算的成员。

### OLEDBConnection.CommandText

返回或设置指定数据源的命令字符串。可读/写 **Variant** 类型。

**语法**

**express.CommandText**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

应使用 **CommandText** 属性，而不应使用 **SQL** 属性，现在 SQL 属性的存在主要是为了与 ET 的早期版本兼容。如果您同时使用这两个属性，则 **CommandText** 属性的值优先。

**CommandType** 属性描述了 **CommandText** 属性的值。

### OLEDBConnection.CommandType

返回或设置 **XlCmdType** 常量之一。可读/写 **XlCmdType** 。

**语法**

**express.CommandType**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

返回或设置的常量描述了 **CommandText** 属性的值。默认值是 **xlCmdSQL** 。

### OLEDBConnection.Connection

返回或设置一个包含 OLE DB 设置的字符串，这些设置使 ET 可以连接 OLE DB 数据源。可读/写 **Variant** 类型。

**语法**

**express.Connection**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

设置 **Connection** 属性并不会立即初始化到数据源的连接。必须使用 **Refresh** 方法建立连接并检索数据。 在使用脱机多维数据集文件时，请将 **UseLocalConnection** 属性设置为 **True** 并使用 **LocalConnection** 属性，而不是 **Connection** 属性。

**示例**

``` JavaScript
/*本示例在活动工作表的 A3 单元格上创建一个基于 OLAP 提供程序的数据透视表缓存，然后基于该缓存创建一个数据透视表。 */
function test(){
　　　　let apadd = ActiveWorkbook.PivotCaches.Add(xlExternal)
　　　　apadd.Connection = "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National"
　　　　apadd.MaintainConnection = true
　　　　apadd.CreatePivotTable (Range("A3"), "PivotTable1")
　　　　
　　　　let pts = ActiveSheet.PivotTables.Item("PivotTable1")
　　　　pts.SmallGrid = false
　　　　pts.PivotCache.RefreshPeriod = 0

　　　　let cfs = pts.CubeFields.Item("[state]")
　　　　cfs.Orientation = xlColumnField
　　　　cfs.Position = 0
    
　　　　let cfs2 = pts.CubeFields.Item("[Measures].[Count Of au_id]")
　　　　cfs2.Orientation = xlDataField
　　　　cfs2.Position = 0
}
```

### OLEDBConnection.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### OLEDBConnection.EnableRefresh

如果用户可刷新连接，则为 **True** 。默认值是 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.EnableRefresh**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

如果 **EnableRefresh** 属性设置为 **False** ，则 **RefreshOnFileOpen** 属性被忽略。

### OLEDBConnection.IsConnected

如果 **MaintainConnection** 属性为 **True** ，则返回 **True** ，如果它当前未与其源相连，则返回 **False** 。只读 **Boolean** 类型。

**语法**

**express.IsConnected**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

**IsConnected** 属性不检查是否建立了连接。即使该属性返回 **True** ，如果连接不再有效，则向提供程序发送命令将导致错误

### OLEDBConnection.LocalConnection

返回或设置脱机多维数据集文件的连接字符串。可读/写 **String** 类型。

**语法**

**express.LocalConnection**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

对于非 OLAP 数据源， **LocalConnection** 属性的值为空字符串，并且 **UseLocalConnection** 属性设置为 **False** 。

设置 **LocalConnection** 属性并不会立即初始化到数据源的连接。必须首先使用 **Refresh** 方法建立连接并检索数据。

如果 **UseLocalConnection** 属性设置为 **True** ，则使用 **LocalConnection** 属性的值。如果 **UseLocalConnection** 属性设置为 **False** ，则 **Connection** 属性将为基于源而不是基于本地多维数据集文件的查询表指定连接字符串。

### OLEDBConnection.LocaleID

返回或设置指定的连接的区域设置标识符。可读/写。

**语法**

**express.LocaleID**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

在将 **LocaleID** 属性设置为新的区域设置之前，必须将 **OLEDBConnection** 对象的 **RetrieveInOfficeUILang** 属性设置为 **False** 。有关有效的区域设置 ID (LCID) 值的详细信息，请搜索 MSDN 网站查找“Locale IDs Assigned by Microsoft”（Microsoft 分配的区域设置 ID）。

### OLEDBConnection.MaintainConnection

如果从刷新操作开始直至关闭工作簿，都一直保留与指定数据源的连接，则返回 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.MaintainConnection**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

默认值是 **True** 。如果预计会频繁对服务器进行查询，则可以将此属性设置为 **True** ，这样能减少重新连接的时间，因而可提高性能。如果将此属性设置为 **False** ，则会使打开的连接关闭。

### OLEDBConnection.MaxDrillthroughRecords

返回或设置要检索的记录的最大数目。可读/写 **Long** 类型。

**语法**

**express.MaxDrillthroughRecords**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.OLAP

如果 OLE DB 连接与联机分析处理 (OLAP) 服务器相连，则返回 **True** ，只读 **Boolean** 类型。

**语法**

**express.OLAP**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.RefreshDate

返回上次刷新 OLE DB 连接的日期。只读 **Date** 类型。

**语法**

**express.RefreshDate**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

对于 OLAP 数据源，在每次查询后都会更新此属性。

### OLEDBConnection.Refreshing

如果正在进行指定 OLE DB 连接的后台 OLE DB 查询，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.Refreshing**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

使用 **CancelRefresh** 方法可取消后台查询。

### OLEDBConnection.RefreshOnFileOpen

如果每次打开工作簿都自动更新连接，则为 **True** 。默认值为 **False** 。可读/写 **Boolean** 类型。

**语法**

**express.RefreshOnFileOpen**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

在 Visual Basic 中使用 **Open** 方法打开工作簿时，不会自动刷新连接。打开工作簿后，可使用 **Refresh** 方法刷新数据。

### OLEDBConnection.RefreshPeriod

返回或设置两次刷新之间的时间间隔（单位为分钟）。可读/写 **Long** 类型。

**语法**

**express.RefreshPeriod**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

将时间段设置为 0（零）会禁用定时自动刷新，等同于将此属性设置为 **Null** 。 **RefreshPeriod** 属性的值可以是从 0 到 32767 的整数。

### OLEDBConnection.RetrieveInOfficeUILang

如果要以 Office 用户界面显示语言（如果可用）检索数据和错误，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.RetrieveInOfficeUILang**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

如果此属性设置为 **False** ，则改用连接字符串中的 LCID 值。如果未指定 LCID，则使用服务器中的默认 LCID。

### OLEDBConnection.RobustConnect

返回或设置 OLE DB 连接与其数据源连接的方式。可读/写 **XlRobustConnect** 。

**语法**

**express.RobustConnect**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

如果您使用可靠的连接，那么当 ET 无法使用工作簿连接信息建立连接时， ET 将检查工作簿连接是否具有外部连接文件的路径。如果有， ET 将打开外部连接文件并尝试使用外部连接文件中包含的信息进行连接。如果在使用外部连接文件之后能够建立连接，则 ET 将更新工作簿的连接对象。

在信息技术部门在一个中央位置维护和更新连接的情况下，这可以提供可靠的连接，每当上一版本的连接（缓存在工作簿中）失败时，允许用户的工作簿自动获取那些更新。

### OLEDBConnection.SavePassword

如果将 OLE DB 连接字符串中的密码信息保存在连接字符串中，则为 **True** 。如果删除密码，则为 **False** 。可读/写 **Boolean** 类型。

**语法**

**express.SavePassword**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

此属性由数据透视表和查询表在 ODBC 和 OLEDB 查询中使用。

### OLEDBConnection.ServerCredentialsMethod

返回或设置应该用于服务器验证的凭据的类型。可读/写 **XlCredentialsMethod** 。

**语法**

**express.ServerCredentialsMethod**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.ServerFillColor

如果在使用指定连接时从服务器中检索 OLAP 服务器的填充色格式，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.ServerFillColor**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.ServerFontStyle

如果在使用指定连接时从服务器中检索 OLAP 服务器的字体样式格式，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.ServerFontStyle**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.ServerNumberFormat

如果在使用指定连接时从服务器中检索 OLAP 服务器的数字格式，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.ServerNumberFormat**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.ServerSSOApplicationID

返回或设置单一登录 (SSO) 应用程序标识符，该标识符用于在 SSO 数据库中查找凭据。可读/写 **String** 类型。

**语法**

**express.ServerSSOApplicationID**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.ServerTextColor

如果在使用指定连接时从服务器中检索 OLAP 服务器的文本颜色格式，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.ServerTextColor**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.SourceConnectionFile

返回或设置一个 **String** 类型的值，该值指示用于创建连接的 WPS Office 数据连接文件或类似文件。可读/写。

**语法**

**express.SourceConnectionFile**

\*express是一个代表 OLEDBConnection 对象的变量。

### OLEDBConnection.SourceDataFile

返回或设置一个 **String** 类型的值，该值指示 OLE DB 连接的源数据文件。可读/写。

**语法**

**express.SourceDataFile**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

对于基于文件的数据源（如 Access）， **SourceDataFile** 属性包含源数据文件的完全限定路径。对于基于服务器的数据源（如 SQL Server），该属性为 null。如果通过编程方式更改 **Connection** 属性，则 **SourceDataFile** 属性设置为 null。

### OLEDBConnection.UseLocalConnection

如果 **LocalConnection** 属性用于指定使 ET 能够连接到数据源的字符串，则为 **True** 。如果使用 **Connection** 属性指定的连接字符串，则为 **False** 。可读/写 **Boolean** 类型。

**语法**

**express.UseLocalConnection**

\*express是一个代表 OLEDBConnection 对象的变量。

**说明**

本示例将第一个数据透视表缓存的连接字符串设置为引用脱机多维数据集文件

> WPS 开发文档： [WPS 基础接口/表格 API 参考/OLEDBConnection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Pages 对象

## 简介

文档页面的集合。使用 Pages 集合及相关对象和属性可通过编程方式定义工作簿的页面布局。

## 说明

使用 Pages 属性可返回 Pages 集合。下面的示例访问活动工作表的所有页面。

``` JavaScript
let objPage = Application.ActiveWindow.Panes.Item(1).Pages
```

使用 Item 方法可访问代表工作表中单个页面的单个 Page 对象。下面的示例访问活动工作表的第一页。

``` JavaScript
let objPage = Application.ActiveWindow.Panes.Item(1).Pages.Item(1)
```

## 方法

| 序号 | 名称                | 说明                                                     |
|:----:|:--------------------|:---------------------------------------------------------|
|  1   | [Item](#Pages.Item) | 返回一个 Page 对象，该对象代表工作簿中的页的集合。只读。 |

## 属性

| 序号 | 名称                  | 说明                                   |
|:----:|:----------------------|:---------------------------------------|
|  1   | [Count](#Pages.Count) | 返回集合中对象的数目。只读 Long 类型。 |

## 成员方法

### Pages.Item

返回一个 **Page** 对象，该对象代表工作簿中的页的集合。只读。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Pages 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明         |
|---------|-----------|----------|--------------|
| *Index* | 必选      | Variant  | 页的索引值。 |

## 成员属性

### Pages.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 Pages 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Pages](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# RecentFile 对象

## 简介

代表最近使用的文件列表中的某个文件。

## 说明

RecentFile 对象是 RecentFiles 集合的成员

## 方法

| 序号 | 名称                         | 说明                       |
|:----:|:-----------------------------|:---------------------------|
|  1   | [Delete](#RecentFile.Delete) | 删除对象。                 |
|  2   | [Open](#RecentFile.Open)     | 打开一个最近使用的工作簿。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#RecentFile.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#RecentFile.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Index](#RecentFile.Index)             | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                    |
|  4   | [Name](#RecentFile.Name)               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  5   | [Parent](#RecentFile.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [Path](#RecentFile.Path)               | 返回一个 String 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。                                                                                                                                                |

## 成员方法

### RecentFile.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 RecentFile 对象的变量。

### RecentFile.Open

打开一个最近使用的工作簿。

**语法**

**express.Open ()**

\*express是一个代表 RecentFile 对象的变量。

**返回值**

一个代表打开的工作簿的 Workbook 对象。

**示例**

本示例打开 Analysis.xls 工作簿，然后运行它的 Auto_Open 宏。

``` JavaScript
Application.RecentFiles.Item(2).Open()
```

## 成员属性

### RecentFile.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 RecentFile 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
　　　　let myObject = ActiveWorkbook
    if(myObject.Application.Value == "ET") {
　　　　    MsgBox("This is an ET Application object.")
　　　　}
　　　　else {
　　　　    MsgBox("This is not an ET Application object.")
　　　　}
}
```

### RecentFile.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 RecentFile 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### RecentFile.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 RecentFile 对象的变量。

### RecentFile.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 RecentFile 对象的变量。

### RecentFile.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 RecentFile 对象的变量。

### RecentFile.Path

返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。

**语法**

**express.Path**

\*express是一个代表 RecentFile 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/RecentFile](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Research 对象

## 简介

代表 Research 查询的控件。

## 说明

使用 Research 查询时，必须具有对应于实时数据源的现有 GUID。如果相应数据源不可用或不存在，将发生运行时错误。

## 方法

| 序号 | 名称                                             | 说明                                                          |
|:----:|:-------------------------------------------------|:--------------------------------------------------------------|
|  1   | [IsResearchService](#Research.IsResearchService) | 指示 *ServiceID* 参数中指定的 GUID 是否对应于当前配置的服务。 |
|  2   | [Query](#Research.Query)                         | 指定搜索查询。                                                |
|  3   | [SetLanguagePair](#Research.SetLanguagePair)     | 设置翻译服务的语言。                                          |

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                                               |
|:----:|:-------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Research.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#Research.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Parent](#Research.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员方法

### Research.IsResearchService

指示 *ServiceID* 参数中指定的 GUID 是否对应于当前配置的服务。

**语法**

**express.IsResearchService ( *ServiceID* )**

\*express是一个代表 Research 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                          |
|-------------|-----------|----------|-------------------------------|
| *ServiceID* | 必选      | String   | 指定一个标识搜索服务的 GUID。 |

**返回值**

Boolean

### Research.Query

指定搜索查询。

**语法**

**express.Query ( *ServiceID* , *QueryString* , *QueryLanguage* , *UseSelection* , *RequeryContextXML* , *NewQueryContextXML* , *LaunchQuery* )**

\*express是一个代表 Research 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                                    |
|----------------------|-----------|----------|---------------------------------------------------------------------------------------------------------|
| *ServiceID*          | 必选      | String   | 指定一个标识搜索服务的 GUID。                                                                           |
| *QueryString*        | 可选      | Variant  | 指定查询字符串。                                                                                        |
| *QueryLanguage*      | 可选      | Variant  | 指定查询字符串的查询语言。                                                                              |
| *UseSelection*       | 可选      | Variant  | 指定查询字符串的查询语言。                                                                              |
| *RequeryContextXML*  | 可选      | Variant  | 指定包含所请求内容的 xml 文件。                                                                         |
| *NewQueryContextXML* | 可选      | Variant  | 指定包含新查询内容的 XML 文件。                                                                         |
| *LaunchQuery*        | 可选      | Boolean  | 如果为 True，则启动查询。如果为 False，则显示“信息检索”任务窗格，该任务窗格用于搜索指定的信息检索服务。 |

**返回值**

Variant

### Research.SetLanguagePair

设置翻译服务的语言。

**语法**

**express.SetLanguagePair ( *LanguageFrom* , *LanguageTo* )**

\*express是一个代表 Research 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                 |
|----------------|-----------|----------|----------------------|
| *LanguageFrom* | 必选      | Long     | 指定翻译的源语言。   |
| *LanguageTo*   | 必选      | Long     | 指定翻译的目标语言。 |

**返回值**

Variant

## 成员属性

### Research.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Research 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### Research.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Research 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Research.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Research 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Research](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SeriesLines 对象

## 简介

代表图表组中的系列线。

## 说明

系列线连接每个系列中的数据。只有二维堆积条形图、二堆, 2-D 堆只柱形图、复合饼图或复合条饼图等图表可以有系列线。此对象不是集合。没有代表单个系列线的对象；您或者打开图表组中所有数据点的系列线，或者将其全部关闭。

如果 HasSeriesLines 属性是 False ， SeriesLines 对象的绝大部分属性都会被禁用。

## 方法

| 序号 | 名称                          | 说明       |
|:----:|:------------------------------|:-----------|
|  1   | [Delete](#SeriesLines.Delete) | 删除对象。 |
|  2   | [Select](#SeriesLines.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SeriesLines.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Border](#SeriesLines.Border)           | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  3   | [Creator](#SeriesLines.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Format](#SeriesLines.Format)           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  5   | [Name](#SeriesLines.Name)               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  6   | [Parent](#SeriesLines.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### SeriesLines.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 SeriesLines 对象的变量。

**返回值**

Variant

### SeriesLines.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 SeriesLines 对象的变量。

**返回值**

Variant

## 成员属性

### SeriesLines.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SeriesLines 对象的变量。

**示例**

``` JavaScript
//本示例显示一条有关创建 myObject 的应用程序的消息。
function test()
{
    if(Application.ActiveWorkbook.Application.Value == "ET") 
    {
        alert("This is an ET Application object.")
    }
    else
    {
        alert("This is not an ET Application object.")
    }
}
```

### SeriesLines.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 SeriesLines 对象的变量。

### SeriesLines.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 SeriesLines 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### SeriesLines.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 SeriesLines 对象的变量。

### SeriesLines.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 SeriesLines 对象的变量。

### SeriesLines.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SeriesLines 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/SeriesLines](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TableStyleElement 对象

## 简介

代表单个表格样式元素。

## 说明

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，标题行是表格的元素。表格样式可以规定标题行的填充色为红色。

表格中每个表格样式元素的格式设置可在适用于该元素的表格样式中指定。

## 方法

| 序号 | 名称                              | 说明               |
|:----:|:----------------------------------|:-------------------|
|  1   | [Clear](#TableStyleElement.Clear) | 清除此元素的格式。 |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                               |
|:----:|:----------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#TableStyleElement.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Borders](#TableStyleElement.Borders)         | 返回一个代表表样式元素的边框的 Borders 集合。只读。                                                                                                                |
|  3   | [Creator](#TableStyleElement.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Font](#TableStyleElement.Font)               | 返回一个 Font 对象，它代表指定对象的字体。 只读。                                                                                                                  |
|  5   | [HasFormat](#TableStyleElement.HasFormat)     | 返回表样式元素是否具有应用到指定元素的格式设置。只读 Boolean 类型。                                                                                                |
|  6   | [Interior](#TableStyleElement.Interior)       | 返回一个 Interior 对象，该对象代表指定对象的内部。 只读。                                                                                                          |
|  7   | [Parent](#TableStyleElement.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  8   | [StripeSize](#TableStyleElement.StripeSize)   | 返回或设置条带的大小。可读/写 Long 类型。                                                                                                                          |

## 成员方法

### TableStyleElement.Clear

清除此元素的格式。

**语法**

**express.Clear ()**

\*express是一个代表 TableStyleElement 对象的变量。

## 成员属性

### TableStyleElement.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TableStyleElement 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### TableStyleElement.Borders

返回一个代表表样式元素的边框的 **Borders** 集合。只读。

**语法**

**express.Borders**

\*express是一个代表 TableStyleElement 对象的变量。

**示例**

本示例将表格的上边框的颜色设置为红色。

``` JavaScript
function test() {
    let sty = Application.ActiveWorkbook.TableStyles.Item("Table Style 4").TableStyleElements.Item(Application.Enum.xlWholeTable).Borders.Item(Application.Enum.xlEdgeTop)
    sty.Color = 255
    sty.TintAndShade = 0
    sty.Weight = 2
    sty.LineStyle = 1
}
```

### TableStyleElement.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TableStyleElement 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL

### TableStyleElement.Font

返回一个 **Font** 对象，它代表指定对象的字体。 只读。

**语法**

**express.Font**

\*express是一个代表 TableStyleElement 对象的变量。

### TableStyleElement.HasFormat

返回表样式元素是否具有应用到指定元素的格式设置。只读 **Boolean** 类型。

**语法**

**express.HasFormat**

\*express是一个代表 TableStyleElement 对象的变量。

### TableStyleElement.Interior

返回一个 **Interior** 对象，该对象代表指定对象的内部。 只读。

**语法**

**express.Interior**

\*express是一个代表 TableStyleElement 对象的变量。

### TableStyleElement.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 TableStyleElement 对象的变量。

### TableStyleElement.StripeSize

返回或设置条带的大小。可读/写 **Long** 类型。

**语法**

**express.StripeSize**

\*express是一个代表 TableStyleElement 对象的变量。

**说明**

此属性不会应用于所有 **TableStyleElement** 对象。它只应用于 **xlColumnStripe1** 、 **xlColumnStripe2** 、 **xlRowStripe1** 和 **xlRowStripe2** 类型。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/TableStyleElement](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Validation 对象

## 简介

代表工作表区域的数据有效性规则。

## 说明

使用 Validation 属性可返回 Validation 对象。下例更改单元格 E5 的数据有效性规则。

``` JavaScript
Application.Range("e5").Validation.Modify(xlValidateList, xlValidAlertStop, null, "=$A$1:$A$10")
```

使用 Add 方法可将数据有效性添加到某个区域并创建一个新的 Validation 对象。此示例向单元格 E5 添加数据有效性检验。

``` JavaScript
let rng = Application.Range("e5").Validation
rng.Add(xlValidateWholeNumber, xlValidAlertInformation, "5", "10")
rng.InputTitle = "Integers"
rng.ErrorTitle = "Integers"
rng.InputMessage = "Enter an integer from five to ten"
rng.ErrorMessage = "You must enter a number from five to ten"
```

## 方法

| 序号 | 名称                         | 说明                             |
|:----:|:-----------------------------|:---------------------------------|
|  1   | [Add](#Validation.Add)       | 向指定区域内添加数据有效性验证。 |
|  2   | [Delete](#Validation.Delete) | 删除对象。                       |
|  3   | [Modify](#Validation.Modify) | 修改指定区域的数据有效性验证。   |

## 属性

| 序号 | 名称                                         | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AlertStyle](#Validation.AlertStyle)         | 返回有效性检验警告样式。 XlDVAlertStyle 类型，只读。                                                                                                                                                                            |
|  2   | [Application](#Validation.Application)       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  3   | [Creator](#Validation.Creator)               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [ErrorMessage](#Validation.ErrorMessage)     | 返回或设置数据有效性检验错误消息。 String 类型，可读写。                                                                                                                                                                        |
|  5   | [ErrorTitle](#Validation.ErrorTitle)         | 返回或设置数据有效性错误对话框的标题。 String 类型，可读写。                                                                                                                                                                    |
|  6   | [Formula1](#Validation.Formula1)             | 返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 String 型，只读。                                                                                                                      |
|  7   | [Formula2](#Validation.Formula2)             | 返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 Operator 属性为 xlBetween 或 xlNotBetween 的情况。可为常量值、字符串值、单元格引用或公式。 String 类型，只读。                               |
|  8   | [IgnoreBlank](#Validation.IgnoreBlank)       | 如果指定区域内的数据有效性检验允许空值，则该值为 True 。 Boolean 类型，可读写。                                                                                                                                                 |
|  9   | [IMEMode](#Validation.IMEMode)               | 返回或设置日文输入规则的说明。可以是下表列出的 XlIMEMode 常量之一。 Long 类型，可读写。                                                                                                                                         |
|  10  | [InCellDropdown](#Validation.InCellDropdown) | 如果数据有效性显示含有有效取值的下拉列表，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                           |
|  11  | [InputMessage](#Validation.InputMessage)     | 返回或设置数据有效性检验输入信息。 String 类型，可读写。                                                                                                                                                                        |
|  12  | [InputTitle](#Validation.InputTitle)         | 返回或设置数据有效性输入对话框的标题。 String 类型，可读写。                                                                                                                                                                    |
|  13  | [Operator](#Validation.Operator)             | 返回一个 Long 值，它代表数据有效性的运算符。                                                                                                                                                                                    |
|  14  | [Parent](#Validation.Parent)                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  15  | [ShowError](#Validation.ShowError)           | 如果用户输入无效数据时显示数据有效性检查错误消息，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                   |
|  16  | [ShowInput](#Validation.ShowInput)           | 如果用户在数据有效性检查区域内选定了某一单元格时，显示数据有效性检查输入消息，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                       |
|  17  | [Type](#Validation.Type)                     | 返回一个包含 XlDVType 常量的 Long 值，它代表区域的数据类型有效性验证。                                                                                                                                                          |
|  18  | [Value](#Validation.Value)                   | 返回一个 Boolean 值，它指明是否符合所有的有效性验证条件（即，该区域是否包含有效数据）。                                                                                                                                         |

## 成员方法

### Validation.Add

向指定区域内添加数据有效性验证。

**语法**

**express.Add ( *Type* , *AlertStyle* , *Operator* , *Formula1* , *Formula2* )**

\*express是一个代表 Validation 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                                                                |
|--------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Type*       | 必选      | XlDVType | 有效性验证类型。                                                                                                                                                    |
| *AlertStyle* | 可选      | Variant  | 有效性验证警告的样式。可为以下 XlDVAlertStyle 常量之一：xlValidAlertInformation、xlValidAlertStop 或 xlValidAlertWarning。                                          |
| *Operator*   | 可选      | Variant  | 数据有效性验证运算符。可为以下 XlFormatConditionOperator 常量之一：xlBetween、xlEqual、xlGreater、xlGreaterEqual、xlLess、xlLessEqual、xlNotBetween 或 xlNotEqual。 |
| *Formula1*   | 可选      | Variant  | 数据有效性验证等式中的第一部分。                                                                                                                                    |
| *Formula2*   | 可选      | Variant  | 当 Operator 为 xlBetween 或 xlNotBetween 时，数据有效性验证等式的第二部分（其他情况下，此参数被忽略）。                                                             |

**说明**

**Add** 方法所要求的参数依有效性验证的类型而定，如下表所示。

| 有效性验证类型                                                                                                             | 参数                                                                                                                                             |
|----------------------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------|
| **xlValidateCustom**                                                                                                       | **Formula1** 必需，忽略 **Formula2** 。 **Formula1** 必须包含一个表达式，数据项有效时该表达式的值为 **True** ，数据项无效时，该值为 **False** 。 |
| **xlInputOnly**                                                                                                            | 使用 **AlertStyle** 、 **Formula1** 或 **Formula2** 。                                                                                           |
| **xlValidateList**                                                                                                         | **Formula1** 必需，忽略 **Formula2** 。 **Formula1** 必须包含以逗号分隔的值列表，或对该列表的工作表引用。                                        |
| **xlValidateWholeNumber** 、 **xlValidateDate** 、 **xlValidateDecimal** 、 **xlValidateTextLength** 或 **xlValidateTime** | 必须指定 **Formula1** 或 **Formula2** 之一，或两者均指定。                                                                                       |

**示例**

``` JavaScript
//本示例向单元格 E5 添加数据有效性验证。
function test()
{
    let rng = Application.Range("e5").Validation
    rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")
    rng.InputTitle = "Integers"
    rng.ErrorTitle = "Integers"
    rng.InputMessage = "Enter an integer from five to ten"
    rng.ErrorMessage = "You must enter a number from five to ten"
}
```

### Validation.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Validation 对象的变量。

### Validation.Modify

修改指定区域的数据有效性验证。

**语法**

**express.Modify ( *Type* , *AlertStyle* , *Operator* , *Formula1* , *Formula2* )**

\*express是一个代表 Validation 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                          |
|--------------|-----------|----------|-----------------------------------------------------------------------------------------------|
| *Type*       | 可选      | Variant  | 一个代表有效性验证类型的 XlDVType 值。                                                        |
| *AlertStyle* | 可选      | Variant  | 一个代表有效性验证警告样式的 XlDVAlertStyle 值。                                              |
| *Operator*   | 可选      | Variant  | 一个代表数据有效性验证运算符的 XlFormatConditionOperator 值。                                 |
| *Formula1*   | 可选      | Variant  | 数据有效性验证等式中的第一部分。                                                              |
| *Formula2*   | 可选      | Variant  | 当 Operator 为 xlBetween 或 xlNotBetween 时，数据有效性的第二部分；其他情况下，此参数被忽略。 |

**说明**

**Modify** 方法所要求的参数依有效性验证的类型而定，如下表所示。

| 有效性验证类型                                                                                                             | 参数                                                                                                                                             |
|----------------------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------|
| **xlInputOnly**                                                                                                            | 不使用 **AlertStyle** 、 **Formula1** 和 **Formula2** 。                                                                                         |
| **xlValidateCustom**                                                                                                       | **Formula1** 必需，忽略 **Formula2** 。 **Formula1** 必须包含一个表达式，数据项有效时该表达式的值为 **True** ，数据项无效时，该值为 **False** 。 |
| **xlValidateList**                                                                                                         | **Formula1** 必需，忽略 **Formula2** 。 **Formula1** 必须包含一个以逗号分隔的值列表，或对该列表的工作表引用。                                    |
| **xlValidateDate** 、 **xlValidateDecimal** 、 **xlValidateTextLength** 、 **xlValidateTime** 或 **xlValidateWholeNumber** | 必须指定 **Formula1** 或 **Formula2** ，或两者均指定。                                                                                           |

**示例**

``` JavaScript
//本示例更改单元格 E5 的数据有效性验证。
Application.Range("e5").Validation.Modify(xlValidateList, xlValidAlertStop, xlBetween, "=$A$1:$A$10")
```

## 成员属性

### Validation.AlertStyle

返回有效性检验警告样式。 **XlDVAlertStyle** 类型，只读。

**语法**

**express.AlertStyle**

\*express是一个代表 Validation 对象的变量。

**说明**

使用 **Add** 方法设置某一区域的警告样式。如果该区域已有数据有效性检验，则可用 **Modify** 方法更改警告样式。

**示例**

``` JavaScript
//本示例显示单元格 E5 的警告样式。
alert(Range("e5").Validation.AlertStyle)
```

### Validation.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Validation 对象的变量。

**示例**

``` JavaScript
//本示例显示一条有关创建 myObject 的应用程序的消息。
function test()
{
    let myObject = Application.ActiveWorkbook
    if(myObject.Value == "ET") 
    {
        alert("This is an ET Application object.")
    }
    else 
    {
        alert("This is not an ET Application object.")
    }
}
```

### Validation.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Validation 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Validation.ErrorMessage

返回或设置数据有效性检验错误消息。 **String** 类型，可读写。

**语法**

**express.ErrorMessage**

\*express是一个代表 Validation 对象的变量。

**示例**

``` JavaScript
//本示例向单元格 E5 添加数据有效性检验，并指定输入信息和错误消息。
function test()
{
    let rng = Application.Range("e5").Validation
    rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")
    rng.InputTitle = "Integers"
    rng.ErrorTitle = "Integers"
    rng.InputMessage = "Enter an integer from five to ten"
    rng.ErrorMessage = "You must enter a number from five to ten"
}
```

### Validation.ErrorTitle

返回或设置数据有效性错误对话框的标题。 **String** 类型，可读写。

**语法**

**express.ErrorTitle**

\*express是一个代表 Validation 对象的变量。

**示例**

``` JavaScript
//本示例向单元格 E5 添加数据有效性检验，并指定输入信息和错误消息。
function test()
{
    let rng = Application.Range("e5").Validation
    rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")
    rng.InputTitle = "Integers"
    rng.ErrorTitle = "Integers"
    rng.InputMessage = "Enter an integer from five to ten"
    rng.ErrorMessage = "You must enter a number from five to ten"
}
```

### Validation.Formula1

返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 **String** 型，只读。

**语法**

**express.Formula1**

\*express是一个代表 Validation 对象的变量。

### Validation.Formula2

返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 **Operator** 属性为 **xlBetween** 或 **xlNotBetween** 的情况。可为常量值、字符串值、单元格引用或公式。 **String** 类型，只读。

**语法**

**express.Formula2**

\*express是一个代表 Validation 对象的变量。

### Validation.IgnoreBlank

如果指定区域内的数据有效性检验允许空值，则该值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IgnoreBlank**

\*express是一个代表 Validation 对象的变量。

**说明**

当 **IgnoreBlank** 属性为 **True** 时，如果单元格为空白，或者 **MinVal** 或 **MaxVal** 属性所引用的单元格为空白，仍认为该单元格数据有效。

**示例**

``` JavaScript
//本示例使单元格 E5 的数据有效性检验为允许有空值。
Application.Range("e5").Validation.IgnoreBlank = true
```

### Validation.IMEMode

返回或设置日文输入规则的说明。可以是下表列出的 **XlIMEMode** 常量之一。 **Long** 类型，可读写。

**语法**

**express.IMEMode**

\*express是一个代表 Validation 对象的变量。

**说明**

| 常量                      | 说明             |
|---------------------------|------------------|
| **xlIMEModeAlpha**        | 半角字母数字     |
| **xlIMEModeAlphaFull**    | 全角字母数字     |
| **xlIMEModeDisable**      | 禁用             |
| **xlIMEModeHiragana**     | 平假名           |
| **xlIMEModeKatakana**     | 片假名           |
| **xlIMEModeKatakanaHalf** | 片假名（半角）   |
| **xlIMEModeNoControl**    | 无控制           |
| **xlIMEModeOff**          | 关闭（英文模式） |
| **xlIMEModeOn**           | 开启             |

只有安装和选择了日文语言支持时，才可设置该属性。

**示例**

``` JavaScript
//本示例设置单元格 E5 的数据输入规则。
function test()
{
    let rng = Application.Range("E5").Validation
    rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")
    rng.InputTitle = " seisuuchi seisuuchi seisuuchi"
    rng.ErrorTitle = " seisuuchi seisuuchi seisuuchi"
    rng.InputMessage = "5  kara kara 10  否 seisuu seisuu o nyuuryoku nyuuryoku shite shite kudasai kudasai kudasai kudasai ."
    rng.ErrorMessage = " nyuuryoku nyuuryoku dekirunowa dekirunowa dekirunowa dekirunowa dekirunowa 5  kara kara 10  madeno madeno madeno ne desu desu ."
    rng.IMEMode = xlIMEModeAlpha
}
```

### Validation.InCellDropdown

如果数据有效性显示含有有效取值的下拉列表，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.InCellDropdown**

\*express是一个代表 Validation 对象的变量。

**说明**

如果数据有效性检验类型不是 **xlValidateList** ，将忽略该属性。

可用 **Validation** 对象的 **Add** 方法或 **Modify** 方法的 *Minimum* 参数来指定包含有效数据的区域。

**示例**

``` JavaScript
//本示例向单元格 E5 添加数据有效性检验。单元格区域 A1:A10 包含单元格的有效取值，且该单元格显示含有这些有效取值的下拉列表。
function test()
{
    let rng = Application.Range("e5").Validation
    rng.Add(xlValidateList, xlValidAlertStop, xlBetween, "=$A$1:$A$10")
    rng.InCellDropdown = true
}
```

### Validation.InputMessage

返回或设置数据有效性检验输入信息。 **String** 类型，可读写。

**语法**

**express.InputMessage**

\*express是一个代表 Validation 对象的变量。

**示例**

``` JavaScript
//本示例向单元格 E5 添加数据有效性检验，并指定输入信息和错误消息。
function test()
{
    let rng = Application.Range("e5").Validation
    rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")
    rng.InputTitle = "Integers"
    rng.ErrorTitle = "Integers"
    rng.InputMessage = "Enter an integer from five to ten"
    rng.ErrorMessage = "You must enter a number from five to ten"
}
```

### Validation.InputTitle

返回或设置数据有效性输入对话框的标题。 **String** 类型，可读写。

**语法**

**express.InputTitle**

\*express是一个代表 Validation 对象的变量。

**示例**

``` JavaScript
//本示例向单元格 E5 添加数据有效性检验，并指定输入信息和错误消息。
function test()
{
    let rng = Application.Range("e5").Validation
    rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")
    rng.InputTitle = "Integers"
    rng.ErrorTitle = "Integers"
    rng.InputMessage = "Enter an integer from five to ten"
    rng.ErrorMessage = "You must enter a number from five to ten"
}
```

### Validation.Operator

返回一个 **Long** 值，它代表数据有效性的运算符。

**语法**

**express.Operator**

\*express是一个代表 Validation 对象的变量。

### Validation.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Validation 对象的变量。

### Validation.ShowError

如果用户输入无效数据时显示数据有效性检查错误消息，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowError**

\*express是一个代表 Validation 对象的变量。

**示例**

``` JavaScript
//本示例向第一张工作表上的单元格 A10 中添加数据有效性检查。输入的值必须处于 5 到 10 之间；如果用户输入了无效数据，将显示错误消息，但不显示输入消息。
function test()
{
    let rng = Application.Worksheets.Item(1).Range("A10").Validation
    rng.Add(1, xlValidAlertStop, xlBetween, "5", "10")
    rng.ErrorMessage = "value must be between 5 and 10"
    rng.ShowInput = false
    rng.ShowError = true
}
```

### Validation.ShowInput

如果用户在数据有效性检查区域内选定了某一单元格时，显示数据有效性检查输入消息，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowInput**

\*express是一个代表 Validation 对象的变量。

**示例**

``` JavaScript
//本示例向第一张工作表上的单元格 A10 中添加数据有效性检查。输入的值必须处于 5 到 10 之间；如果用户输入了无效数据，将显示错误消息，但不显示输入消息。
function test()
{
    let rng = Application.Worksheets.Item(1).Range("A10").Validation
    rng.Add(1, xlValidAlertStop, xlBetween, "5", "10")
    rng.ErrorMessage = "value must be between 5 and 10"
    rng.ShowInput = false
    rng.ShowError = true
}
```

### Validation.Type

返回一个包含 **XlDVType** 常量的 **Long** 值，它代表区域的数据类型有效性验证。

**语法**

**express.Type**

\*express是一个代表 Validation 对象的变量。

### Validation.Value

返回一个 **Boolean** 值，它指明是否符合所有的有效性验证条件（即，该区域是否包含有效数据）。

**语法**

**express.Value**

\*express是一个代表 Validation 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Validation](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# WebOptions 对象

## 简介

包含工作簿级的属性，当以网页保存文档或打开网页时，ET 将使用这些属性。

## 说明

您可以在应用程序（全局）级或在工作簿级返回或设置属性。（请注意，根据保存工作簿时的属性值，属性值可能会因工作簿而异。工作簿级的属性设置会重写应用程序级的属性设置。应用程序级属性包含在 DefaultWebOptions 对象中。

示例

使用 WebOptions 属性可返回 WebOptions 对象。下例实现的功能是：检查是否允许 PNG（便携网络图形）作为图像格式，然后设置相应的 ` strImageFileType ` 变量。

``` JavaScript
function test() {
    let strImageFileType
    let objAppWebOptions = Application.Workbooks.Item(1).WebOptions
    if (objAppWebOptions.AllowPNG == true) {
        strImageFileType = "PNG"
    }
    else {
        strImageFileType = "JPG"
    }
}
```

## 方法

| 序号 | 名称                                                         | 说明                                                               |
|:----:|:-------------------------------------------------------------|:-------------------------------------------------------------------|
|  1   | [UseDefaultFolderSuffix](#WebOptions.UseDefaultFolderSuffix) | 根据所选或安装的语言支持，将含指定文档的文件夹后缀设置为默认后缀。 |

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.AllowPNG">AllowPNG</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果以网页保存文档时，允许将 PNG（可移植网络图形）作为图像格式使用，则为 True 。如果不允许将 PNG 作为输出格式使用，则为 False 。默认值是 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.DownloadComponents">DownloadComponents</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果必要的“Microsoft Office Web 组件”在您用 Web 浏览器查看已保存的文档时已经下载，但只有在尚未安装这些组件时，该值为 True 。如果没有下载这些组件，则为 False 。默认值是 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.Encoding">Encoding</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>当查看已保存的文档时，返回或设置 Web 浏览器使用的文档编码（代码页或字符集）。默认值为系统代码页。 <span>MsoEncoding</span> 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.FolderSuffix">FolderSuffix</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回文件夹后缀，当您将文档保存为网页、使用长文件名，以及选择将支持文件单独保存在某个文件夹中（即，如果 <span>UseLongFileNames</span> 和 <span>OrganizeInFolder</span> 属性设置为 True ）时，ET 使用该后缀。 String 型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.LocationOfComponents">LocationOfComponents</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置中央 URL（对于 Intranet 或 Web）或路径（对于本地或网络而言），授权用户可以在查看保存的文档时，从这些位置下载“Microsoft Office Web 组件”。默认值是 Microsoft Office 的本地或网络安装路径。 String 型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.OrganizeInFolder">OrganizeInFolder</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则在将指定的文档保存为网页时，所有的支持文件（如背景纹理和图形）将组织在一个单独的文件夹中；如果为 False ，则支持文件和网页将保存在同一文件夹中。默认值是 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.PixelsPerInch">PixelsPerInch</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置网页中图形图像的密度（每英寸的像素数）及表格单元格的密度。密度设置范围通常为 19 到 480 之间，对于普通的屏幕大小，通常的设置一般为 72、96 和 120。默认设置是 96。 Long 型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.RelyOnCSS">RelyOnCSS</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在 Web 浏览器中查看保存的文档时，字体格式使用的是级联样式表 (CSS)，则为 True 。ET 会创建一个级联样式表文件，然后依据 <span>OrganizeInFolder</span> 属性的值，将该文件保存到指定文件夹或与网页相同的文件夹中。如果使用了 HTML &lt;FONT&gt; 标记和级联样式表，则为 False 。默认值是 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.RelyOnVML">RelyOnVML</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，将文档保存为网页时，图形对象不能生成图像文件。如果为 False ，则可生成图像文件。默认值是 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.ScreenSize">ScreenSize</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置在 Web 浏览器中查看已保存文档时应使用的理想的最小屏幕大小（以像素为单位，宽度乘以高度）。可以是 <span>MsoScreenSize</span> 常量之一。默认常量是 msoScreenSize800x600 。 MsoScreenSize 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.TargetBrowser">TargetBrowser</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 <span>MsoTargetBrowser</span> 常量，它指明浏览器的版本。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#WebOptions.UseLongFileNames">UseLongFileNames</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果该属性值为 True ，则将文档保存为网页时使用长文件名。如果该属性值为 False ，则不使用长文件名，而使用 DOS 文件名格式（8.3）。默认值是 True 。 Boolean 类型，可读写。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### WebOptions.UseDefaultFolderSuffix

根据所选或安装的语言支持，将含指定文档的文件夹后缀设置为默认后缀。

**语法**

**express.UseDefaultFolderSuffix ()**

\*express是一个代表 WebOptions 对象的变量。

**说明**

将文档保存为网页时，ET 会使用文件夹后缀，同时还会使用长文件名，并选择是否将支持文件保存在单独的文件夹中（即，是否将 **UseLongFileNames** 和 **OrganizeInFolder** 属性设置为 **True** ）。

此后缀位于文档名后面的文件夹名中。例如，如果文档名为“Book1”，语言为“英语”，则文件夹名为“Book1_files”。可用的文件夹后缀列在 **FolderSuffix** 属性主题中。

**示例**

本示例将第一个工作簿的文件夹后缀设置为默认后缀。

``` JavaScript
function test() {
    Application.Workbooks.Item(1).WebOptions.UseDefaultFolderSuffix()
}
```

## 成员属性

### WebOptions.AllowPNG

如果以网页保存文档时，允许将 PNG（可移植网络图形）作为图像格式使用，则为 **True** 。如果不允许将 PNG 作为输出格式使用，则为 **False** 。默认值是 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.AllowPNG**

\*express是一个代表 WebOptions 对象的变量。

**说明**

如果将图像保存为 PNG 格式而非其他文件格式，则可以提高图像的质量或减小图像文件的大小，因而也就可以减少下载所需的时间，前提是所使用的 Web 浏览器支持 PNG 格式。

**示例**

此示例使用 PNG 格式作为第一个工作簿的输出格式。

``` JavaScript
function test() {
    Application.Workbooks.Item(1).WebOptions.AllowPNG = true
}
```

或者，新建文档的应用程序可以将 PNG 作为全局默认设置。

``` JavaScript
function test() {
    Application.DefaultWebOptions.AllowPNG = true
}
```

### WebOptions.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 WebOptions 对象的变量。

**示例**

本示例显示一条有关创建 ` myObject ` 的应用程序的消息。

``` JavaScript
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        slert("This is an ET Application object.")
    }
    else {
        slert("This is not an ET Application object.")
    }
}
```

### WebOptions.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 WebOptions 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### WebOptions.DownloadComponents

如果必要的“Microsoft Office Web 组件”在您用 Web 浏览器查看已保存的文档时已经下载，但只有在尚未安装这些组件时，该值为 **True** 。如果没有下载这些组件，则为 **False** 。默认值是 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.DownloadComponents**

\*express是一个代表 WebOptions 对象的变量。

**说明**

可将 **LocationOfComponents** 属性设置为主 URL（在 Intranet 或网站上）或路径（本地或网络），授权用户在查看已保存的文档时，可从该路径所指向的位置下载组件。该路径必须是有效路径，且指向的位置必须包含必要的组件，同时用户还必须拥有有效的 Microsoft Office 许可。

Office Web 组件可增强保存为网页的文档的交互性。如果在没有安装这些组件的计算机上使用 Web 浏览器查看网页，则该页上的交互部分将表现为静态的（不能交互）。

**示例**

如果还没有安装“Office Web 组件”，此示例允许“Office Web 组件”与指定的网页一起下载。

``` JavaScript
function test() {
    Application.DefaultWebOptions.DownloadComponents = true
    Application.DefaultWebOptions.LocationOfComponents = Application.Path + Application.PathSeparator + "foo"
}
```

### WebOptions.Encoding

table { word-break:break-all; }

当查看已保存的文档时，返回或设置 Web 浏览器使用的文档编码（代码页或字符集）。默认值为系统代码页。 **MsoEncoding** 类型，可读写。

**语法**

**express.Encoding**

\*express是一个代表 WebOptions 对象的变量。

**说明**

不能使用任何带有 **AutoDetect** 后缀的常量。这些常量由 **ReloadAs** 方法使用。

### WebOptions.FolderSuffix

返回文件夹后缀，当您将文档保存为网页、使用长文件名，以及选择将支持文件单独保存在某个文件夹中（即，如果 **UseLongFileNames** 和 **OrganizeInFolder** 属性设置为 **True** ）时，ET 使用该后缀。 **String** 型，只读。

**语法**

**express.FolderSuffix**

\*express是一个代表 WebOptions 对象的变量。

**说明**

新建的文档将使用 **DefaultWebOptions** 对象的 **FolderSuffix** 属性所返回的后缀。如果文档曾在其他语言版本的 ET 中进行过编辑，则 **WebOptions** 对象的 **FolderSuffix** 属性的值可能会与 **DefaultWebOptions** 对象中该属性的值不同。可以使用 **UseDefaultFolderSuffix** 方法将后缀更改为当前在 Microsoft Office 中所使用的语言。

默认情况下，支持文件夹名由网页的名称加上一个下划线 (\_)、一个句点 (.) 或连字号 (-)，再加上单词“files”（将文件保存为网页的 ET 版本的语言形式）组成。例如，如果使用荷兰语版本的 ET 将名为“Page1”的文件保存为网页，则支持文件夹的默认名称为“Page1_bestanden”。

下表列出 Office 的各个语言版本及相应的 **LanguageID** 属性值和文件夹后缀。对于表中没有列出的语言，其后缀为“.files”。

| 语言                 | LanguageID | 文件夹后缀    |
|----------------------|------------|---------------|
| 阿拉伯语             | 1025       | .files        |
| 巴斯克语             | 1069       | \_fitxategiak |
| 巴西葡萄牙语         | 1046       | \_arquivos    |
| 保加利亚语           | 1026       | .files        |
| 加泰罗尼亚语         | 1027       | \_fitxers     |
| 中文 - 简体          | 2052       | .files        |
| 中文 - 繁体          | 1028       | .files        |
| 克罗地亚语           | 1050       | \_datoteke    |
| 捷克语               | 1029       | \_soubory     |
| 丹麦语               | 1030       | -filer        |
| 荷兰语               | 1043       | \_bestanden   |
| 英语                 | 1033       | \_files       |
| 爱沙尼亚语           | 1061       | \_failid      |
| 芬兰语               | 1035       | \_tiedostot   |
| 法语                 | 1036       | \_fichiers    |
| 德语                 | 1031       | -Dateien      |
| 希腊语               | 1032       | .files        |
| 希伯来语             | 1037       | .files        |
| 匈牙利语             | 1038       | \_elemei      |
| 意大利语             | 1040       | -file         |
| 日语                 | 1041       | .files        |
| 朝鲜语               | 1042       | .files        |
| 拉脱维亚语           | 1062       | \_fails       |
| 立陶宛语             | 1063       | \_bylos       |
| 挪威语               | 1044       | -filer        |
| 波兰语               | 1045       | \_pliki       |
| 葡萄牙语             | 2070       | \_ficheiros   |
| 罗马尼亚语           | 1048       | .files        |
| 俄语                 | 1049       | .files        |
| 塞尔维亚语(西里尔文) | 3098       | .files        |
| 塞尔维亚语（拉丁文） | 2074       | \_fajlovi     |
| 斯洛伐克语           | 1051       | .files        |
| 斯洛文尼亚语         | 1060       | \_datoteke    |
| 西班牙语             | 3082       | \_archivos    |
| 瑞典语               | 1053       | -filer        |
| 泰语                 | 1054       | .files        |
| 土耳其语             | 1055       | \_dosyalar    |
| 乌克兰语             | 1058       | .files        |
| 越南语               | 1066       | .files        |

**示例**

此示例返回第一个工作簿所使用的文件夹后缀。该后缀将返回到字符串变量 ` strFolderSuffix ` 中。

``` JavaScript
function test() {
    let strFolderSuffix = Application.Workbooks.Item(1).WebOptions.FolderSuffix
}
```

### WebOptions.LocationOfComponents

返回或设置中央 URL（对于 Intranet 或 Web）或路径（对于本地或网络而言），授权用户可以在查看保存的文档时，从这些位置下载“Microsoft Office Web 组件”。默认值是 Microsoft Office 的本地或网络安装路径。 **String** 型，可读写。

**语法**

**express.LocationOfComponents**

\*express是一个代表 WebOptions 对象的变量。

**说明**

如果 **DownloadComponents** 属性设置为 **True** ，Office Web 组件尚未安装，路径是有效路径且指向包含必要组件的位置，同时用户还具有有效的 Microsoft Office 许可，则会从指定网页自动下载这些组件。

### WebOptions.OrganizeInFolder

如果为 **True** ，则在将指定的文档保存为网页时，所有的支持文件（如背景纹理和图形）将组织在一个单独的文件夹中；如果为 **False** ，则支持文件和网页将保存在同一文件夹中。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.OrganizeInFolder**

\*express是一个代表 WebOptions 对象的变量。

**说明**

在保存网页的文件夹中创建新文件夹，并依据文档命名。如果使用长文件名，则在文件夹名称后加后缀。 **FolderSuffix** 属性返回已选择或已安装的语言支持的文件夹后缀，或返回默认的文件夹后缀。

如果在上次保存某个文档时， **OrganizeInFolder** 属性的值与现在不同，那么现在保存该文档时，ET 会相应地将支持文件自动移进或移出该文件夹。

如果并未使用长文件名（即，如果 **UseLongFileNames** 属性设置为 **False** ），则 ET 会自动将任何支持文件保存到其他文件夹中。这些文件不能保存到与网页相同的文件夹中。

**示例**

此示例指定在将文档保存为网页时，所有的支持文件和网页保存在同一文件夹中。

``` JavaScript
function test() {
    Application.DefaultWebOptions.OrganizeInFolder = false
}
```

### WebOptions.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 WebOptions 对象的变量。

### WebOptions.PixelsPerInch

返回或设置网页中图形图像的密度（每英寸的像素数）及表格单元格的密度。密度设置范围通常为 19 到 480 之间，对于普通的屏幕大小，通常的设置一般为 72、96 和 120。默认设置是 96。 **Long** 型，可读写。

**语法**

**express.PixelsPerInch**

\*express是一个代表 WebOptions 对象的变量。

**说明**

无论何时在 Web 浏览器中查看所保存文档，此属性都可确定指定的网页中相对于文本尺寸的图像以及单元格的大小。图像或单元格的实际最终尺寸等于原始尺寸（以英寸为单位）乘以每英寸中的像素数。

使用 **ScreenSize** 属性可为目标 Web 浏览器设置最佳屏幕大小。

**示例**

此示例根据浏览器的目标屏幕大小来设置像素密度。对于 800x600 像素的屏幕，其密度为 72 像素/英寸。对于 1024x768 像素的屏幕，其密度为 96 像素/英寸。对于所有其他情况，则使用的密度为 120 像素/英寸。

``` JavaScript
function test() {
    let web = Application.ActiveWorkbook.WebOptions
    switch (web.ScreenSize) {
        case msoScreenSize800x600:
            web.PixelsPerInch = 72
            break
        case msoScreenSize1024x768:
            web.PixelsPerInch = 96
            break
        default:
            web.PixelsPerInch = 120
    }
}
```

### WebOptions.RelyOnCSS

如果在 Web 浏览器中查看保存的文档时，字体格式使用的是级联样式表 (CSS)，则为 **True** 。ET 会创建一个级联样式表文件，然后依据 **OrganizeInFolder** 属性的值，将该文件保存到指定文件夹或与网页相同的文件夹中。如果使用了 HTML \<FONT\> 标记和级联样式表，则为 **False** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.RelyOnCSS**

\*express是一个代表 WebOptions 对象的变量。

**说明**

如果 Web 浏览器支持级联样式表，则应将此属性设置为 **True** ，这样，就可以为该网页上提供更精确的布局和格式控制，网页也会更接近原来的文档（如同在 ET 中显示一样）。

**示例**

此示例启用级联样式表作为应用程序的全局默认值。

``` JavaScript
function test() {
    Application.ActiveWorkbook.WebOptions.RelyOnCSS = true
}
```

### WebOptions.RelyOnVML

如果为 **True** ，将文档保存为网页时，图形对象不能生成图像文件。如果为 **False** ，则可生成图像文件。默认值是 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.RelyOnVML**

\*express是一个代表 WebOptions 对象的变量。

**说明**

如果您的 Web 浏览器支持 Vector Markup Language (VML)，也可不从图形对象生成图像，这样可减少文件大小。例如，Microsoft Internet Explorer 5 支持此项功能，因此如果要将其作为目标浏览器，应该将 **RelyOnVML** 属性设为 **True** 。如果浏览器不支持 VML，则在查看用此属性保存的网页时，将不会显示图像。

例如，如果网页中使用了已生成的图像文件，并且文档保存的位置与该网页在 Web 服务器上的最终位置不同，则不应生成图像。

**示例**

此示例指定在将工作表保存到网页上时，生成图像。

``` JavaScript
function test() {
    Application.Workbooks.Item(1).WebOptions.RelyOnVML = false
}
```

### WebOptions.ScreenSize

返回或设置在 Web 浏览器中查看已保存文档时应使用的理想的最小屏幕大小（以像素为单位，宽度乘以高度）。可以是 **MsoScreenSize** 常量之一。默认常量是 **msoScreenSize800x600** 。 **MsoScreenSize** 类型，可读写。

**语法**

**express.ScreenSize**

\*express是一个代表 WebOptions 对象的变量。

### WebOptions.TargetBrowser

返回或设置一个 **MsoTargetBrowser** 常量，它指明浏览器的版本。可读写。

**语法**

**express.TargetBrowser**

\*express是一个代表 WebOptions 对象的变量。

**说明**

|                                                                                               |
|-----------------------------------------------------------------------------------------------|
| **MsoTargetBrowser** 可为以下这些 **MsoTargetBrowser** 常量之一。                             |
| **msoTargetBrowserIE4** 。Microsoft Internet Explorer 4.0 或更高版本。                        |
| **msoTargetBrowserIE5** 。Microsoft Internet Explorer 5.0 或更高版本。                        |
| **msoTargetBrowserIE6** 。Microsoft Internet Explorer 6.0 或更高版本。                        |
| **msoTargetBrowserV3** 。Microsoft Internet Explorer 3.0、Netscape Navigator 3.0 或更高版本。 |
| **msoTargetBrowserV4** 。Microsoft Internet Explorer 4.0、Netscape Navigator 4.0 或更高版本。 |

**示例**

在本示例中，ET 确定 Web 选项的浏览器版本是否为 IE5，并通知用户。

``` JavaScript
function test() {
    let wkbOne = Application.Workbooks.Item(1)

    //Determine if IE5 is the target browser.
    if (wkbOne.WebOptions.TargetBrowser == Application.Enum.msoTargetBrowserIE5) {
        alert("The target browser is IE5 or later.")
    } else {
        alert("The target browser is not IE5 or later.")
    }
}
```

### WebOptions.UseLongFileNames

如果该属性值为 **True** ，则将文档保存为网页时使用长文件名。如果该属性值为 **False** ，则不使用长文件名，而使用 DOS 文件名格式（8.3）。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UseLongFileNames**

\*express是一个代表 WebOptions 对象的变量。

**说明**

如果不使用长文件名且文档还具有支持文件，ET 会自动将那些文件组织到单独的文件夹中。另外，使用 **OrganizeInFolder** 属性可确定支持文件是否组织在单独的文件夹中。

**示例**

此示例禁止将长文件名用作应用程序的全局默认设置。

``` JavaScript
function test() {
    Application.DefaultWebOptions.UseLongFileNames = false
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/WebOptions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Windows 对象

## 简介

ET 中所有 Window 对象的集合。

## 说明

Application 对象的 Windows 集合包含应用程序中的所有窗口，而 Workbook 对象的 Windows 集合只包含指定工作簿中的窗口。

## 示例

使用 Windows 属性可返回 Windows 集合。下例层叠当前在 ET 中显示的所有窗口。

``` JavaScript
Application.Windows.Arrange(xlCascade)
```

使用 NewWindow 方法可返回一个新窗口并将它添加到集合。下例为活动工作簿创建一个新窗口。

``` JavaScript
Application.ActiveWorkbook.NewWindow()
```

使用 Windows ( *index* )（其中 *index* 是窗口名称或索引号）可返回一个 Window 对象。下例最大化活动窗口。

注意，活动窗口总是 ` Windows(1) ` 。

``` JavaScript
Application.Windows.Item(1).WindowState = xlMaximized
```

## 方法

| 序号 | 名称                  | 说明                   |
|:----:|:----------------------|:-----------------------|
|  1   | [Item](#Windows.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Windows.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Windows.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |

## 成员方法

### Windows.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Windows 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
/*此示例将当前窗口最大化。*/
Application.Windows.Item(1).WindowState = xlMaximized
```

## 成员属性

### Windows.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Windows 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    if (Application.ActiveWorkbook.Application.Value == "ET" ) {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### Windows.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Windows 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Windows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Application 对象

## 简介

代表 WPS 应用程序。 Application 对象包含可返回顶级对象的属性和方法。例如， [ActiveDocument](#Application.ActiveDocument) 属性返回 Document 对象。

## 说明

整个wps应用程序的api对象结构是一个树状结构，而Application对象是树状结构的根对象，同时，Application对象也提供了对于应用程序相关的各种访问接口，例如：

- 应用程序的设置选项、环境、版本号等相关信息
- 一些常见的属性，如ActiveWindow、ActiveDocument、Name等

以下示例显示 WPS 用户名。

``` JavaScript
alert(Application.UserName);
```

``` JavaScript
/* 添加一个文档，并且打印该文档名 */
function test() {
    Application.Documents.Add();
    alert(Application.ActiveDocument.Name);
}
```

WPS宏编辑器环境下，许多返回最常见用户界面对象（如活动文档， ActiveDocument 属性）的属性和方法可在没有 Application 对象识别符的情况下使用。例如，您可以编写 ActiveDocument.PrintOut，而不是编写 Application.ActiveDocument.PrintOut。可以在没有 Application 对象识别符的情况下使用的属性和方法被视为“全局”属性和方法。

## 方法

| 序号 | 名称                                                            | 说明                                                                                                                                                                                                      |
|:----:|:----------------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Activate](#Application.Activate)                               | 激活指定的对象。                                                                                                                                                                                          |
|  2   | [AddAddress](#Application.AddAddress)                           | 本方法在通讯簿中新增一项。每项都有一个或多个标签 ID 值。                                                                                                                                                  |
|  3   | [AutomaticChange](#Application.AutomaticChange)                 | 当“WPS 助手”建议进行更改时，将执行 “自动套用格式” 操作。如果未激活任何“自动套用格式”操作，该方法将导致出错。                                                                                              |
|  4   | [BuildKeyCode](#Application.BuildKeyCode)                       | 为指定的组合键返回唯一的代码。                                                                                                                                                                            |
|  5   | [CentimetersToPoints](#Application.CentimetersToPoints)         | 将度量单位由厘米转换为磅数（1 厘米=28.35 磅）。所返回的转换结果为 Single 类型。                                                                                                                           |
|  6   | [ChangeFileOpenDirectory](#Application.ChangeFileOpenDirectory) | 设置 WPS 在其中搜索文档的文件夹。                                                                                                                                                                         |
|  7   | [CheckGrammar](#Application.CheckGrammar)                       | 检查字符串的语法错误。返回一个 Boolean 类型的值，以表示字符串中是否包含语法错误。如果字符串未包含语法错误，则该值为 True 。                                                                               |
|  8   | [CheckSpelling](#Application.CheckSpelling)                     | 检查字符串的拼写错误。返回一个 Boolean 类型的值，以显示字符串是否包含拼写错误。如果字符串未包含拼写错误，则该值为 True 。                                                                                 |
|  9   | [CleanString](#Application.CleanString)                         | 从指定字符串中删除非打印字符（字符代码为 1-29）及 WPS 的特殊字符，或将它们替换为空格（字符代码为 32）。返回的结果为 String 类型。                                                                         |
|  10  | [CompareDocuments](#Application.CompareDocuments)               | 比较两个文档，并返回一个 Document 对象，该对象表示包含两个文档之间的差别的文档，该文档用修订进行标记。                                                                                                    |
|  11  | [DDEExecute](#Application.DDEExecute)                           | 通过指定的 DDE（即动态数据交换）通道，向应用程序发送一条或一组命令。                                                                                                                                      |
|  12  | [DDEInitiate](#Application.DDEInitiate)                         | 打开通向其他应用程序的 DDE（动态数据交换）通道，并返回通道序号。                                                                                                                                          |
|  13  | [DDEPoke](#Application.DDEPoke)                                 | 通过打开的 DDE（动态数据交换）通道向应用程序发送数据。                                                                                                                                                    |
|  14  | [DDERequest](#Application.DDERequest)                           | 通过打开的 DDE（动态数据交换）通道向接收应用程序查询信息，并以 String 类型返回该信息。                                                                                                                    |
|  15  | [DDETerminate](#Application.DDETerminate)                       | 关闭指定的通向另一应用程序的 DDE（动态数据交换）通道。                                                                                                                                                    |
|  16  | [DDETerminateAll](#Application.DDETerminateAll)                 | 关闭所有由 WPS 打开的动态数据交换 (DDE) 通道。                                                                                                                                                            |
|  17  | [DefaultWebOptions](#Application.DefaultWebOptions)             | 返回 DefaultWebOptions 对象，该对象包含 WPS 每当将文档保存为网页或打开网页时使用的应用程序级全局属性。                                                                                                    |
|  18  | [FileDialog](#Application.FileDialog)                           | 返回一个 FileDialog 对象，该对象代表文件对话框的单个实例。                                                                                                                                                |
|  19  | [FindKey](#Application.FindKey)                                 | 返回 KeyBinding 对象，该对象代表指定的组合键。只读。                                                                                                                                                      |
|  20  | [GetAddress](#Application.GetAddress)                           | 返回默认通讯簿中的地址。                                                                                                                                                                                  |
|  21  | [GetDefaultTheme](#Application.GetDefaultTheme)                 | 返回代表默认 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">主题</span> 名称的 String ，以及 WPS 用于新文档、电子邮件或网页的主题格式选项。 |
|  22  | [GetSpellingSuggestions](#Application.GetSpellingSuggestions)   | 返回一个 SpellingSuggestions 集合，该集合代表指定单词的建议拼写替换单词。                                                                                                                                 |
|  23  | [GoBack](#Application.GoBack)                                   | 将插入点在活动文档中最近进行过编辑操作的三个位置间移动，其作用等同于按 Shift+F5。                                                                                                                         |
|  24  | [GoForward](#Application.GoForward)                             | 将插入点在活动文档中进行编辑的最后三个位置之间向前移动。                                                                                                                                                  |
|  25  | [Help](#Application.Help)                                       | 显示安装的帮助信息。                                                                                                                                                                                      |
|  26  | [HelpTool](#Application.HelpTool)                               | 指针从箭头改为问号，表明这是关于将要单击的命令或屏幕元素的快捷帮助。如果单击文本， WPS 将显示描述当前段落及字符格式的对话框。按 Esc 可将指针恢复为原始状态。                                              |
|  27  | [InchesToPoints](#Application.InchesToPoints)                   | 将度量单位从英寸转换为磅（1 英寸 = 72 磅）。以 Single 类型返回转换结果。                                                                                                                                  |
|  28  | [International](#Application.International)                     | 返回当前的国家/地区和国际设置信息。 Variant 类型，只读。                                                                                                                                                  |
|  29  | [KeysBoundTo](#Application.KeysBoundTo)                         | 返回一个 KeysBoundTo 对象，该对象代表指定项的所有组合键。                                                                                                                                                 |
|  30  | [Keyboard](#Application.Keyboard)                               | 返回或设置键盘的语言和布局设置。                                                                                                                                                                          |
|  31  | [KeyboardBidi](#Application.KeyboardBidi)                       | 将键盘语言设置为从右向左的语言，并将文字的输入方向设置为从右向左。                                                                                                                                        |
|  32  | [KeyboardLatin](#Application.KeyboardLatin)                     | 将键盘语言设置为从左向右的语言，并将文字的输入方向设置为从左向右。                                                                                                                                        |
|  33  | [KeyString](#Application.KeyString)                             | 为指定的键返回组合键字符串（例如，Ctrl+Shift+A）。                                                                                                                                                        |
|  34  | [LinesToPoints](#Application.LinesToPoints)                     | 将度量单位从行转换为磅（1 行 = 12 磅）。以 Single 类型返回转换结果。                                                                                                                                      |
|  35  | [ListCommands](#Application.ListCommands)                       | 新建文档，然后插入带有相关快捷键和菜单设定的 WPS 命令表。                                                                                                                                                 |
|  36  | [LoadMasterList](#Application.LoadMasterList)                   | 加载书目源文件。                                                                                                                                                                                          |
|  37  | [LookupNameProperties](#Application.LookupNameProperties)       | 在全局通讯簿列表中查找一个姓名，并显示 “属性” 对话框，其中包含了有关指定姓名的信息。                                                                                                                      |
|  38  | [MergeDocuments](#Application.MergeDocuments)                   | 比较两个文档，并返回一个 Document 对象，该对象表示包含两个文档之间的差别的文档，该文档用修订进行标记。                                                                                                    |
|  39  | [MillimetersToPoints](#Application.MillimetersToPoints)         | 把度量单位从毫米转换为磅值（1 毫米 = 2.85 磅）。以 Single 类型返回转换结果。                                                                                                                              |
|  40  | [Move](#Application.Move)                                       | 设置任务窗口或活动文档窗口的位置。                                                                                                                                                                        |
|  41  | [NewWindow](#Application.NewWindow)                             | 打开一个新的窗口作为指定窗口，并显示原文档。返回一个 Window 对象。                                                                                                                                        |
|  42  | [NextLetter](#Application.NextLetter)                           | 您已经就仅在 Macintosh 上使用的 Visual Basic 关键字请求帮助。有关 Application 对象的 NextLetter 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。                                     |
|  43  | [OnTime](#Application.OnTime)                                   | 启动在指定的时间运行宏的后台计时器。                                                                                                                                                                      |
|  44  | [OrganizerCopy](#Application.OrganizerCopy)                     | 将指定的“自动图文集”词条、工具栏、样式或宏方案项从源文档或模板中复制到目标文档或模板上。                                                                                                                  |
|  45  | [OrganizerDelete](#Application.OrganizerDelete)                 | 从文档或模板中删除指定的样式、“自动图文集”词条、工具栏或宏方案项。                                                                                                                                        |
|  46  | [OrganizerRename](#Application.OrganizerRename)                 | 重新命名文档或模板中指定的样式、“自动图文集”词条、工具栏或宏方案项。                                                                                                                                      |
|  47  | [PicasToPoints](#Application.PicasToPoints)                     | 将度量单位从十二点活字转换为磅值（1 十二点活字 = 12 磅）。以 Single 类型返回转换结果。                                                                                                                    |
|  48  | [PixelsToPoints](#Application.PixelsToPoints)                   | 将长度值的单位由像素转换为磅。以 Single 类型返回转换结果。                                                                                                                                                |
|  49  | [PointsToCentimeters](#Application.PointsToCentimeters)         | 将度量单位由磅转换为厘米（1 厘米 = 28.35 磅）。以 Single 类型返回转换结果。                                                                                                                               |
|  50  | [PointsToInches](#Application.PointsToInches)                   | 将磅转换为英寸（1 英寸 = 72 磅）。以 Single 类型返回转换结果。                                                                                                                                            |
|  51  | [PointsToLines](#Application.PointsToLines)                     | 将磅转换为行（1 行 = 12 磅）。以 Single 类型返回转换结果。                                                                                                                                                |
|  52  | [PointsToMillimeters](#Application.PointsToMillimeters)         | 将度量单位由磅转换为毫米（1 毫米 = 2.835 磅）。以 Single 类型返回转换结果。                                                                                                                               |
|  53  | [PointsToPicas](#Application.PointsToPicas)                     | 将度量单位由磅转换为十二点活字（1 十二点活字 = 12 磅）。以 Single 类型返回转换结果。                                                                                                                      |
|  54  | [PointsToPixels](#Application.PointsToPixels)                   | 将长度值的单位由磅转换为像素。以 Single 类型返回转换结果。                                                                                                                                                |
|  55  | [PrintOut](#Application.PrintOut)                               | 打印指定文档的全部或部分内容。                                                                                                                                                                            |
|  56  | [ProductCode](#Application.ProductCode)                         | 返回一个 String 类型的 WPS 全球唯一标识符 (GUID)。                                                                                                                                                        |
|  57  | [PutFocusInMailHeader](#Application.PutFocusInMailHeader)       | 如果活动窗口中的文档是电子邮件文档，则将插入点置于邮件标题的 “收件人” 行。                                                                                                                                |
|  58  | [Quit](#Application.Quit)                                       | 退出 WPS ，并可选择保存或传送打开的文档。                                                                                                                                                                 |
|  59  | [Repeat](#Application.Repeat)                                   | 一次或多次重复最近的编辑操作。如果成功地重复了命令，则返回 True 。                                                                                                                                        |
|  60  | [ResetIgnoreAll](#Application.ResetIgnoreAll)                   | 在拼写检查时，取消以前的忽略词汇表。                                                                                                                                                                      |
|  61  | [Resize](#Application.Resize)                                   | 调整 WPS 应用程序或某一任务的窗口大小。                                                                                                                                                                   |
|  62  | [Run](#Application.Run)                                         | 运行JS函数                                                                                                                                                                                                |
|  63  | [ScreenRefresh](#Application.ScreenRefresh)                     | 使用视频内存缓冲区的当前信息更新显示器的显示。                                                                                                                                                            |
|  64  | [SetDefaultTheme](#Application.SetDefaultTheme)                 | 为 WPS 设置用于新文档、电子邮件或网页的默认 主题 。                                                                                                                                                       |
|  65  | [ShowClipboard](#Application.ShowClipboard)                     | 显示 “剪贴板” 任务窗格。                                                                                                                                                                                  |
|  66  | [ShowMe](#Application.ShowMe)                                   | 当存在更多可用信息时，显示“WPS 助手”或“帮助”窗口。                                                                                                                                                        |
|  67  | [SubstituteFont](#Application.SubstituteFont)                   | 设置字体映射选项。                                                                                                                                                                                        |
|  68  | [SynonymInfo](#Application.SynonymInfo)                         | 返回一个 SynonymInfo 对象，该对象包含来自同义词库的、与指定字词或短语的同义词、反义词或相关词语及表达相关的信息。                                                                                         |
|  69  | [ToggleKeyboard](#Application.ToggleKeyboard)                   | 在从左向右的语言和从右向左语言之间切换键盘语言设置。                                                                                                                                                      |

## 属性

| 序号 | 名称                                                                                | 说明                                                                                                                                                                                                                                                               |
|:----:|:------------------------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ActiveDocument](#Application.ActiveDocument)                                       | 返回一个 Document 对象，该对象代表活动文档。如果没有打开的文档，就会导致出错。只读。                                                                                                                                                                               |
|  2   | [ActiveEncryptionSession](#Application.ActiveEncryptionSession%20)                  | 返回一个 Long 类型的值，该值代表与活动文档相关的加密会话。只读。                                                                                                                                                                                                   |
|  3   | [ActivePrinter](#Application.ActivePrinter)                                         | 返回或设置活动打印机名称。可读/写 String 类型。                                                                                                                                                                                                                    |
|  4   | [ActiveProtectedViewWindow](#Application.ActiveProtectedViewWindow)                 | 返回代表活动的受保护的视图窗口的 ProtectedViewWindow 对象。只读。                                                                                                                                                                                                  |
|  5   | [ActiveWindow](#Application.ActiveWindow)                                           | 返回 Window 对象，该对象代表活动窗口（焦点所在的窗口）。如果没有打开的窗口，则会导致出错。只读。                                                                                                                                                                   |
|  6   | [AddIns](#Application.AddIns)                                                       | 返回一个 AddIns 集合，该集合代表所有有效加载项，而不考虑当前是否已加载它们。只读。                                                                                                                                                                                 |
|  7   | [AnswerWizard](#Application.AnswerWizard)                                           | 返回一个 AnswerWizard 对象，该对象包含联机“帮助”搜索引擎所用的文件。                                                                                                                                                                                               |
|  8   | [Application](#Application.Application)                                             | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                                                                                                                                                                                               |
|  9   | [ArbitraryXMLSupportAvailable](#Application.ArbitraryXMLSupportAvailable)           | 返回一个 Boolean 类型的值，代表 WPS 是否接受自定义的 XML 架构。值为 True 表明 WPS 接受自定义的 XML 架构。                                                                                                                                                          |
|  10  | [Assistance](#Application.Assistance)                                               | 返回一个代表 WPS Office帮助查看器的 Assistance 对象。只读。                                                                                                                                                                                                        |
|  11  | [Assistant](#Application.Assistant)                                                 | 返回一个 Assistant 对象，该对象代表“WPS Office 助手”。                                                                                                                                                                                                             |
|  12  | [AutoCaptions](#Application.AutoCaptions)                                           | 返回一个 AutoCaptions 集合。当诸如表格或图片之类的项目插入到文档中时，此集合代表自动添加的题注。只读。                                                                                                                                                             |
|  13  | [AutoCorrect](#Application.AutoCorrect)                                             | 返回一个 AutoCorrect 对象，该对象包含当前“自动更正”的选项、词条和例外项。只读。                                                                                                                                                                                    |
|  14  | [AutoCorrectEmail](#Application.AutoCorrectEmail)                                   | 返回一个 AutoCorrect 对象，该对象代表对电子邮件消息进行的自动更正。                                                                                                                                                                                                |
|  15  | [AutomationSecurity](#Application.AutomationSecurity)                               | 返回或设置一个 MsoAutomationSecurity 常量，该常量代表当用编程方法打开文件时 WPS 所使用的安全设置。                                                                                                                                                                 |
|  16  | [BackgroundPrintingStatus](#Application.BackgroundPrintingStatus)                   | 返回后台打印队列中的打印任务数目。 Long 类型，只读。                                                                                                                                                                                                               |
|  17  | [BackgroundSavingStatus](#Application.BackgroundSavingStatus)                       | 返回后台保存队列中的文件数。 Long 类型，只读。                                                                                                                                                                                                                     |
|  18  | [Bibliography](#Application.Bibliography)                                           | 返回一个 Bibliography 对象，该对象代表 WPS 中存储的书目参考源。只读。                                                                                                                                                                                              |
|  19  | [BrowseExtraFileTypes](#Application.BrowseExtraFileTypes)                           | 如果将该属性设置为“text/html”，则用 WPS （而不是默认的 Internet 浏览器）可以打开超链接的 HTML 文件。 String 类型，可读写。                                                                                                                                         |
|  20  | [Browser](#Application.Browser)                                                     | 返回一个 Browser 对象，该对象代表垂直滚动条上的 “选择浏览对象” 工具。只读。                                                                                                                                                                                        |
|  21  | [Build](#Application.Build)                                                         | 返回 WPS 应用程序的版本号及编译序号。 String 类型，只读。                                                                                                                                                                                                          |
|  22  | [CapsLock](#Application.CapsLock)                                                   | 如果已打开 Caps Lock 键，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                                                                 |
|  23  | [Caption](#Application.Caption)                                                     | 返回或设置出现在应用程序窗口的标题栏中的文本。 String 类型，可读写。                                                                                                                                                                                               |
|  24  | [CaptionLabels](#Application.CaptionLabels)                                         | 返回一个 CaptionLabels 集合，该集合包括全部有效题注标签。只读。                                                                                                                                                                                                    |
|  25  | [CheckLanguage](#Application.CheckLanguage)                                         | 如果 WPS 在键入时自动检测使用的语言，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                   |
|  26  | [COMAddIns](#Application.COMAddIns)                                                 | 返回对 COMAddIns 集合的引用，该集合代表 WPS 中当前加载的所有 组件对象模型 (COM) 加载项?（COM 加载项：通过添加自定义命令和指定的功能来扩展 WPS Office程序的功能的补充程序。COM 加载项可在一个或多个 Office 程序中运行。COM 加载项使用文件扩展名 .dll 或 .exe。） 。 |
|  27  | [CommandBars](#Application.CommandBars)                                             | 返回一个 CommandBars 集合，该集合代表 WPS 中的菜单栏以及所有工具栏。                                                                                                                                                                                               |
|  28  | [Creator](#Application.Creator)                                                     | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。                                                                                                                                                                                       |
|  29  | [CustomDictionaries](#Application.CustomDictionaries)                               | 返回一个 Dictionaries 对象，该对象代表活动自定义词典的集合。只读。                                                                                                                                                                                                 |
|  30  | [CustomizationContext](#Application.CustomizationContext)                           | 返回或设置一个 Template 或 Document 对象，该对象代表存储菜单栏、工具栏和键绑定的更改的模板或文档。可读/写。                                                                                                                                                        |
|  31  | [DefaultLegalBlackline](#Application.DefaultLegalBlackline)                         | 如果为 True ，则 WPS 使用 “比较并合并文档” 对话框中的 “精确比较” 选项比较和合并文档。 Boolean 类型，可读写。                                                                                                                                                       |
|  32  | [DefaultSaveFormat](#Application.DefaultSaveFormat)                                 | 返回或设置将在 “另存为” 对话框上的 “保存类型” 框中显示的默认格式。 String 类型，可读写。                                                                                                                                                                           |
|  33  | [DefaultTableSeparator](#Application.DefaultTableSeparator)                         | 返回或设置一个字符；在将文本转换为表格时，该字符用来将文本分隔为单元格。 String 类型，可读写。                                                                                                                                                                     |
|  34  | [Dialogs](#Application.Dialogs)                                                     | 返回一个 Dialogs 集合，该集合代表 WPS 中所有内置对话框。只读。                                                                                                                                                                                                     |
|  35  | [DisplayAlerts](#Application.DisplayAlerts)                                         | 返回或设置运行宏时产生的一些警告和消息的处理方式。 WdAlertLevel 类型，可读写。                                                                                                                                                                                     |
|  36  | [DisplayAutoCompleteTips](#Application.DisplayAutoCompleteTips)                     | 如果 WPS 显示有关所键入整个单词、日期或词组的建议的提示，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                               |
|  37  | [DisplayDocumentInformationPanel](#Application.DisplayDocumentInformationPanel)     | 返回或设置 Boolean 值，该值表示是否显示文档属性面板。可读/写。                                                                                                                                                                                                     |
|  38  | [DisplayRecentFiles](#Application.DisplayRecentFiles)                               | 如果在 “文件” 菜单中显示最近使用的文件名，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                              |
|  39  | [DisplayScreenTips](#Application.DisplayScreenTips)                                 | 如果为 True ，则批注、脚注、尾注和超链接以提示形式显示。突出显示那些标记有批注的文本。 Boolean 类型，可读写。                                                                                                                                                      |
|  40  | [DisplayScrollBars](#Application.DisplayScrollBars)                                 | 如果为 True ，则 WPS 在至少一个文档窗口中显示滚动条；如果为 False ，则在任何窗口中均不显示滚动条。 Boolean 类型，可读写。                                                                                                                                          |
|  41  | [Documents](#Application.Documents)                                                 | 返回一个 Documents 集合，该集合代表所有打开的文档。只读。                                                                                                                                                                                                          |
|  42  | [DontResetInsertionPointProperties](#Application.DontResetInsertionPointProperties) | 返回或设置一个 Boolean 类型的值，该值代表 WPS 在运行其他代码后是否保持位于插入点位置的文本的格式属性。可读写。                                                                                                                                                     |
|  43  | [EmailOptions](#Application.EmailOptions)                                           | 返回一个 EmailOptions 对象，该对象代表创作电子邮件的全局首选项。只读。                                                                                                                                                                                             |
|  44  | [EmailTemplate](#Application.EmailTemplate)                                         | 返回或设置一个 String 类型的值，该值代表发送电子邮件时使用的文档模板。可读写。                                                                                                                                                                                     |
|  45  | [EnableCancelKey](#Application.EnableCancelKey)                                     | 返回或设置 WPS 处理 Ctrl+Break 用户中断的方式。 WdEnableCancelKey 类型，可读写。                                                                                                                                                                                   |
|  46  | [FeatureInstall](#Application.FeatureInstall)                                       | 返回或设置 WPS 将如何处理对所需功能尚未安装的方法和属性的调用。 MsoReatureInstall 类型，可读写。                                                                                                                                                                   |
|  47  | [FileConverters](#Application.FileConverters)                                       | 返回一个 FileConverters 集合，该集合代表可用于 WPS 的所有文件转换器。只读。                                                                                                                                                                                        |
|  48  | [FileValidation](#Application.FileValidation)                                       | 返回或设置 WPS 在打开文件之前验证这些文件的方式。可读/写 MsoFileValidationMode。                                                                                                                                                                                   |
|  49  | [FindKey](#Application.FindKey)                                                     | 返回 KeyBinding 对象，该对象代表指定的组合键。只读。                                                                                                                                                                                                               |
|  50  | [FocusInMailHeader](#Application.FocusInMailHeader)                                 | 如果插入点位于电子邮件标题域中（如“收件人”域），则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                                          |
|  51  | [FontNames](#Application.FontNames)                                                 | 返回一个 FontNames 对象，该对象包含所有有效字体的名称。只读。                                                                                                                                                                                                      |
|  52  | [HangulHanjaDictionaries](#Application.HangulHanjaDictionaries)                     | 返回一个 HangulHanjaConversionDictionaries 集合，该集合代表所有活动的自定义转换词典。                                                                                                                                                                              |
|  53  | [Height](#Application.Height)                                                       | 返回或设置活动文档窗口的高度。 Long 类型，可读写。                                                                                                                                                                                                                 |
|  54  | [IsObjectValid](#Application.IsObjectValid)                                         | 如果引用某个对象的指定变量有效，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                                                          |
|  55  | [IsSandboxed](#Application.IsSandboxed)                                             | 如果应用程序窗口为受保护的视图窗口，则该属性值为 True 。只读。                                                                                                                                                                                                     |
|  56  | [KeyBindings](#Application.KeyBindings)                                             | 返回一个 KeyBindings 集合，该集合代表自定义的按键指定方案，包含了键代码、键类别和命令。                                                                                                                                                                            |
|  57  | [LandscapeFontNames](#Application.LandscapeFontNames)                               | 返回一个 FontNames 对象，该对象包括了所有有效的横向字体的名称。                                                                                                                                                                                                    |
|  58  | [Language](#Application.Language)                                                   | 返回一个 MsoLanguageID 常量，该常量代表为 WPS 用户界面选定的语言。                                                                                                                                                                                                 |
|  59  | [Languages](#Application.Languages)                                                 | 返回一个 Languages 集合，该集合代表列在 “语言” 对话框（在 “工具” 菜单上，单击 “语言” ，再单击 “设置语言” ）中的校对语言。                                                                                                                                          |
|  60  | [LanguageSettings](#Application.LanguageSettings)                                   | 返回一个 LanguageSettings 对象，该对象包含 WPS 中语言设置的信息。                                                                                                                                                                                                  |
|  61  | [Left](#Application.Left)                                                           | 返回或设置一个 Long 类型的值，该值代表活动文档的水平位置，以磅为单位。可读写。                                                                                                                                                                                     |
|  62  | [ListGalleries](#Application.ListGalleries)                                         | 返回一个 ListGalleries 集合，该集合代表三个列表模板库。                                                                                                                                                                                                            |
|  63  | [MacroContainer](#Application.MacroContainer)                                       | 返回 Template 或 Document 对象，该对象代表其中存储模块（该模块中包含着运行的过程）的模板或文档。                                                                                                                                                                   |
|  64  | [MailingLabel](#Application.MailingLabel)                                           | 返回一个 MailingLabel 对象，该对象代表一个邮件标签。                                                                                                                                                                                                               |
|  65  | [MailMessage](#Application.MailMessage)                                             | 返回一个 MailMessage 对象，该对象代表活动的电子邮件。                                                                                                                                                                                                              |
|  66  | [MailSystem](#Application.MailSystem)                                               | 返回主机上安装的邮件系统（或系统）。 WdMailSystem 类型，只读。                                                                                                                                                                                                     |
|  67  | [MAPIAvailable](#Application.MAPIAvailable)                                         | 如果已经安装了 MAPI，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                                                                     |
|  68  | [MathCoprocessorAvailable](#Application.MathCoprocessorAvailable)                   | 如果安装了数学协处理器并且该处理器适用于 WPS ，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                                           |
|  69  | [MouseAvailable](#Application.MouseAvailable)                                       | 如果系统有可用的鼠标，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                                                                    |
|  70  | [Name](#Application.Name)                                                           | 返回指定对象的名称。 String 类型，只读。                                                                                                                                                                                                                           |
|  71  | [NewDocument](#Application.NewDocument)                                             | 返回一个 NewFile 对象，该对象代表在“新建文档”任务窗格上所列的一个文档。                                                                                                                                                                                            |
|  72  | [NormalTemplate](#Application.NormalTemplate)                                       | 返回一个 Template 对象，该对象代表 Normal 模板。                                                                                                                                                                                                                   |
|  73  | [NumLock](#Application.NumLock)                                                     | 此属性返回 Num Lock 键的状态。如果为 True ，则数字键盘可用于输入数字；如果为 False ，则该键盘用于移动插入点。 Boolean 类型，只读。                                                                                                                                 |
|  74  | [OMathAutoCorrect](#Application.OMathAutoCorrect)                                   | 返回一个 OMathAutoCorrect 对象，该对象代表公式的自动更正条目。只读。                                                                                                                                                                                               |
|  75  | [OpenAttachmentsInFullScreen](#Application.OpenAttachmentsInFullScreen)             | 返回或设置 Boolean 值，该值表示 WPS 是否在读取模式下打开电子邮件附件。可读写。                                                                                                                                                                                     |
|  76  | [Options](#Application.Options)                                                     | 返回一个 Options 对象，该对象代表 WPS 中的应用程序设置。                                                                                                                                                                                                           |
|  77  | [Parent](#Application.Parent)                                                       | 返回一个 Object 类型的值，该值代表指定 Application 对象的父对象。                                                                                                                                                                                                  |
|  78  | [Path](#Application.Path)                                                           | 返回指定对象的磁盘或 Web 路径。只读 String 类型。                                                                                                                                                                                                                  |
|  79  | [PathSeparator](#Application.PathSeparator)                                         | 返回用于分隔文件夹名称的字符。在 Windows 中，该属性返回一个反斜线 (\\。只读 String 类型。                                                                                                                                                                          |
|  80  | [PickerDialog](#Application.PickerDialog)                                           | 返回 PickerDialog 对象，以提供在对话框中选择用户或数据的功能。只读。                                                                                                                                                                                               |
|  81  | [PortraitFontNames](#Application.PortraitFontNames)                                 | 返回一个 FontNames 对象，该对象包括所有有效纵向字体的名称。                                                                                                                                                                                                        |
|  82  | [PrintPreview](#Application.PrintPreview)                                           | 如果当前视图是打印预览，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                |
|  83  | [ProtectedViewWindows](#Application.ProtectedViewWindows)                           | 返回代表所有受保护的视图窗口的 ProtectedViewWindows 集合。只读。                                                                                                                                                                                                   |
|  84  | [RecentFiles](#Application.RecentFiles)                                             | 返回一个 RecentFiles 集合，该集合代表最近存取过的文件。                                                                                                                                                                                                            |
|  85  | [RestrictLinkedStyles](#Application.RestrictLinkedStyles)                           | 返回或设置 Boolean 值，该值表示 WPS 是否允许链接样式。可读写。                                                                                                                                                                                                     |
|  86  | [ScreenUpdating](#Application.ScreenUpdating)                                       | 如果启用屏幕更新，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                      |
|  87  | [Selection](#Application.Selection)                                                 | 返回 Selection 对象，该对象代表一选定的范围或插入点。只读。                                                                                                                                                                                                        |
|  88  | [ShowStartupDialog](#Application.ShowStartupDialog)                                 | 如果为 True ，则在启动 WPS 时显示 “任务窗格” 。 Boolean 类型，可读写。                                                                                                                                                                                             |
|  89  | [ShowStylePreviews](#Application.ShowStylePreviews)                                 | 返回或设置 Boolean 值，该值表示 WPS 是否在 “样式” 对话框中显示样式的格式预览。可读/写。                                                                                                                                                                            |
|  90  | [ShowVisualBasicEditor](#Application.ShowVisualBasicEditor)                         | 如果显示“Visual Basic 编辑器”窗口，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                     |
|  91  | [ShowWindowsInTaskbar](#Application.ShowWindowsInTaskbar)                           | 如果为 True ，则在任务栏中显示打开的文档，默认值为 Single Document Interface (SDI)。如果为 False ，则只在 “窗口” 菜单中列出打开的文档，提供 Multiple Document Interface (MDI) 的外观。 Boolean 类型，可读写。                                                      |
|  92  | [SmartArtColors](#Application.SmartArtColors)                                       | 返回 SmartArtColors 对象，该对象代表应用程序中当前加载的颜色样式集。只读。                                                                                                                                                                                         |
|  93  | [SmartArtLayouts](#Application.SmartArtLayouts)                                     | 返回 SmartArtLayouts 对象，该对象代表应用程序中当前加载的 SmartArt 布局集。只读。                                                                                                                                                                                  |
|  94  | [SmartArtQuickStyles](#Application.SmartArtQuickStyles)                             | 返回 SmartArtQuickStyles 对象，该对象代表应用程序中当前加载的 SmartArt 样式集。只读。                                                                                                                                                                              |
|  95  | [SmartTagRecognizers](#Application.SmartTagRecognizers)                             | 返回应用程序的一个 SmartTagRecognizers 集合。                                                                                                                                                                                                                      |
|  96  | [SmartTagTypes](#Application.SmartTagTypes)                                         | 返回一个 SmartTagTypes 集合，代表 WPS 中所安装的智能标记组件的智能标记类型。                                                                                                                                                                                       |
|  97  | [SpecialMode](#Application.SpecialMode)                                             | 如果 WPS 处于某个特殊模式（如 CopyText 模式或 MoveText 模式）下，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                         |
|  98  | [StartupPath](#Application.StartupPath)                                             | 返回或设置启动文件夹的完整路径，不包括最后的分隔符。 String 类型，可读写。                                                                                                                                                                                         |
|  99  | [System](#Application.System)                                                       | 返回一个 System 对象，该对象可用于返回与系统相关的信息，并执行与系统相关的任务。                                                                                                                                                                                   |
| 100  | [TaskPanes](#Application.TaskPanes)                                                 | 返回一个 TaskPanes 集合，该集合代表 WPS 中最常执行的任务。                                                                                                                                                                                                         |
| 101  | [Tasks](#Application.Tasks)                                                         | 返回一个 Tasks 集合，该集合代表所有正在运行的应用程序。                                                                                                                                                                                                            |
| 102  | [Templates](#Application.Templates)                                                 | 返回一个 Templates 集合，该集合代表所有可用的模板，包括共用模板和附加到打开文档的模板。                                                                                                                                                                            |
| 103  | [Top](#Application.Top)                                                             | 返回或设置活动文档的垂直位置。 Long 类型，可读写。                                                                                                                                                                                                                 |
| 104  | [UndoRecord](#Application.UndoRecord)                                               | 返回 UndoRecord 对象，该对象提供撤消堆栈的自定义入口点。只读。                                                                                                                                                                                                     |
| 105  | [UsableHeight](#Application.UsableHeight)                                           | 返回 WPS 文档窗口高度可设置的最大值（以磅为单位）。 Long 类型，只读。                                                                                                                                                                                              |
| 106  | [UsableWidth](#Application.UsableWidth)                                             | 返回 WPS 文档窗口可设置的最大宽度（以磅为单位）。 Long 类型，只读。                                                                                                                                                                                                |
| 107  | [UserAddress](#Application.UserAddress)                                             | 该属性返回或设置用户的通讯地址。 String 类型，可读写。                                                                                                                                                                                                             |
| 108  | [UserControl](#Application.UserControl)                                             | 如果文档或应用程序是由用户创建或打开的，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                                                                                  |
| 109  | [UserName](#Application.UserName)                                                   | 该属性返回或设置用户姓名， WPS 将其用于信封和文档的“作者”属性。 String 类型，可读写。                                                                                                                                                                              |
| 110  | [VBE](#Application.VBE)                                                             | 返回一个 VBE 对象，该对象代表“Visual Basic 编辑器”。                                                                                                                                                                                                               |
| 111  | [Version](#Application.Version)                                                     | 返回 WPS 版本号。 String 类型，只读。                                                                                                                                                                                                                              |
| 112  | [Visible](#Application.Visible)                                                     | 如果指定对象可见，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                      |
| 113  | [Width](#Application.Width)                                                         | 返回或设置应用程序窗口的宽度（以磅为单位）。 Long 类型，可读写。                                                                                                                                                                                                   |
| 114  | [Windows](#Application.Windows)                                                     | 返回一个 Windows 集合，该集合代表所有文档窗口。只读。                                                                                                                                                                                                              |
| 115  | [WindowState](#Application.WindowState)                                             | 返回或设置指定文档窗口或任务窗口的状态。 WdWindowState 类型，可读写。                                                                                                                                                                                              |
| 116  | [WordBasic](#Application.WordBasic)                                                 | 返回一个自动化对象 (Word.Basic)，其中包含适用于 WPS 6.0 和 WPS for Windows 中所有 WordBasic 语句和函数的方法。只读。                                                                                                                                               |
| 117  | [XMLNamespaces](#Application.XMLNamespaces)                                         | 返回一个 XMLNamespaces 集合，代表“架构库”中的 XML 架构。                                                                                                                                                                                                           |

## 成员方法

### Application.Activate

激活指定的对象。

**语法**

**express.Activate ()**

\*express是一个代表 Application 对象的变量。

### Application.AddAddress

本方法在通讯簿中新增一项。每项都有一个或多个标签 ID 值。

**语法**

**express.AddAddress ( *TagID* , *Value* )**

\*express是一个代表 Application 对象的变量。

**参数**

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th>名称</th>
<th>必选/可选</th>
<th>数据类型</th>
<th>说明</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td class="mainsection"><em>TagID</em></td>
<td class="mainsection">必选</td>
<td class="mainsection">String array</td>
<td class="mainsection" style="word-break: break-all">新地址项的标签 ID 值。数组中的每个元素可包含下表所列的字符串之一。只有显示名称是必需的，其余项都是可选的。
<table>
<thead>
<tr class="header">
<th>标签 ID</th>
<th>说明</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td>PR_DISPLAY_NAME</td>
<td>“通讯簿” 对话框中显示的姓名</td>
</tr>
<tr class="even">
<td>PR_DISPLAY_NAME_PREFIX</td>
<td>称谓（例如，“Ms.”或“Dr.”）</td>
</tr>
<tr class="odd">
<td>PR_GIVEN_NAME</td>
<td>名</td>
</tr>
<tr class="even">
<td>PR_SURNAME</td>
<td>姓</td>
</tr>
<tr class="odd">
<td>PR_STREET_ADDRESS</td>
<td>街道地址</td>
</tr>
<tr class="even">
<td>PR_LOCALITY</td>
<td>市/县或地方</td>
</tr>
<tr class="odd">
<td>PR_STATE_OR_PROVINCE</td>
<td>省/市/自治区</td>
</tr>
<tr class="even">
<td>PR_POSTAL_CODE</td>
<td>邮政编码</td>
</tr>
<tr class="odd">
<td>PR_COUNTRY</td>
<td>国家/地区</td>
</tr>
<tr class="even">
<td>PR_TITLE</td>
<td>职务</td>
</tr>
<tr class="odd">
<td>PR_COMPANY_NAME</td>
<td>公司名称</td>
</tr>
<tr class="even">
<td>PR_DEPARTMENT_NAME</td>
<td>公司内部门名称</td>
</tr>
<tr class="odd">
<td>PR_OFFICE_LOCATION</td>
<td>办公室位置</td>
</tr>
<tr class="even">
<td>PR_PRIMARY_TELEPHONE_NUMBER</td>
<td>主要电话号码</td>
</tr>
<tr class="odd">
<td>PR_PRIMARY_FAX_NUMBER</td>
<td>主要传真号码</td>
</tr>
<tr class="even">
<td>PR_OFFICE_TELEPHONE_NUMBER</td>
<td>办公室电话号码</td>
</tr>
<tr class="odd">
<td>PR_OFFICE2_TELEPHONE_NUMBER</td>
<td>第二个办公室电话号码</td>
</tr>
<tr class="even">
<td>PR_HOME_TELEPHONE_NUMBER</td>
<td>住宅电话号码</td>
</tr>
<tr class="odd">
<td>PR_CELLULAR_TELEPHONE_NUMBER</td>
<td>移动电话号码</td>
</tr>
<tr class="even">
<td>PR_BEEPER_TELEPHONE_NUMBER</td>
<td>BP 机号码</td>
</tr>
<tr class="odd">
<td>PR_COMMENT</td>
<td>地址项的 “注释” 选项卡包含的文字</td>
</tr>
<tr class="even">
<td>PR_EMAIL_ADDRESS</td>
<td>电子邮件地址</td>
</tr>
<tr class="odd">
<td>PR_ADDRTYPE</td>
<td>电子邮件地址类型</td>
</tr>
<tr class="even">
<td>PR_OTHER_TELEPHONE_NUMBER</td>
<td>其他电话号码（区别于住宅或办公室电话号码）</td>
</tr>
<tr class="odd">
<td>PR_BUSINESS_FAX_NUMBER</td>
<td>商务传真号码</td>
</tr>
<tr class="even">
<td>PR_HOME_FAX_NUMBER</td>
<td>住宅传真号码</td>
</tr>
<tr class="odd">
<td>PR_RADIO_TELEPHONE_NUMBER</td>
<td>无线电话号码</td>
</tr>
<tr class="even">
<td>PR_INITIALS</td>
<td>缩写</td>
</tr>
<tr class="odd">
<td>PR_LOCATION</td>
<td>地点，格式为楼号/房间号（例如，7/3007 代表 7 号楼 3007 房间）</td>
</tr>
<tr class="even">
<td>PR_CAR_TELEPHONE_NUMBER</td>
<td>汽车电话号码</td>
</tr>
</tbody>
</table></td>
</tr>
<tr class="even">
<td class="mainsection"><em>Value</em></td>
<td class="mainsection">必选</td>
<td class="mainsection">String array</td>
<td class="mainsection" style="word-break: break-all">新地址项的值。每个元素对应于 TagID 数组中的一个元素。有关详细信息，请参阅以下示例。</td>
</tr>
</tbody>
</table>

**示例**

``` JavaScript
/*本示例在通讯簿中添加一项。*/
function test() {
    let tagIDArray = [0, 1, 2, 3]
    let valueArray = [0, 1, 2, 3]

    tagIDArray[0] = "PR_DISPLAY_NAME"
    tagIDArray[1] = "PR_GIVEN_NAME"
    tagIDArray[2] = "PR_SURNAME"
    tagIDArray[3] = "PR_COMMENT"
    valueArray[0] = "Kim Buhler"
    valueArray[1] = "Kim"
    valueArray[2] = "Buhler"
    valueArray[3] = "This is a comment"

    Application.AddAddress(tagIDArray, valueArray)
}
```

### Application.AutomaticChange

当“WPS 助手”建议进行更改时，将执行 **“自动套用格式”** 操作。如果未激活任何“自动套用格式”操作，该方法将导致出错。

**语法**

**express.AutomaticChange ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将完成“WPS 助手”建议进行的“自动套用格式”操作（如果存在一个待执行的操作）。*/
Application.AutomaticChange()
```

### Application.BuildKeyCode

为指定的组合键返回唯一的代码。

**语法**

**express.BuildKeyCode ( *Arg1* , *Arg2 - Arg4* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                          |
|---------------|-----------|----------|-------------------------------|
| *Arg1*        | 必选      | WdKey    | 使用 WdKey 常量之一指定的键。 |
| *Arg2 - Arg4* | 可选      | WdKey    | 使用 WdKey 常量之一指定的键。 |

**示例**

``` JavaScript
/*本示例为 Organizer 命令指定组合键 Alt+F1。*/
Application.CustomizationContext = NormalTemplate
Application.KeyBindings.Add(wdKeyCategoryCommand,"Organizer",BuildKeyCode(wdKeyAlt,wdKeyF1))

/*本示例从 Normal 模板中删除 Alt+F1 按键指定方案。*/
Application.CustomizationContext = NormalTemplate
Application.FindKey(BuildKeyCode(wdKeyAlt,wdKeyF1)).Clear()
```

### Application.CentimetersToPoints

将度量单位由厘米转换为磅数（1 厘米=28.35 磅）。所返回的转换结果为 **Single** 类型。

**语法**

**express.CentimetersToPoints ( *Centimeters* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                 |
|---------------|-----------|----------|----------------------|
| *Centimeters* | 必选      | Single   | 要转换为磅的厘米值。 |

**示例**

``` JavaScript
/*本示例用于为选定内容的所有段落添加居中制表位。此制表位位于距页面左边距 1.5 厘米处。*/
Application.Selection.Paragraphs.TabStops.Add(Application.CentimetersToPoints(1.5),Application.Enum.wdAlignTabCenter)
```

### Application.ChangeFileOpenDirectory

设置 WPS 在其中搜索文档的文件夹。

**语法**

**express.ChangeFileOpenDirectory ( *Path* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                      |
|--------|-----------|----------|-------------------------------------------|
| *Path* | 必选      | String   | 文件夹路径， WPS 将在该文件夹中搜索文档。 |

**说明**

下次显示 **“文件”** 菜单的 **“打开”** 对话框时，将列出指定文件夹的内容。用户在 **“打开”** 对话框中更改该文件夹或当前 WPS 会话结束前， WPS 将一直在指定的文件夹中搜索文档。使用 **DefaultFilePath** 属性更改每个 WPS 会话中的文档默认文件夹。

**示例**

``` JavaScript
/*本示例更改文件夹，以使 WPS 在该文件夹中搜索文档，然后打开“Test.doc”文档。*/
function test() {
    Application.ChangeFileOpenDirectory("C:\\Documents")
    Application.Documents.Open("Test.doc")
}
```

### Application.CheckGrammar

检查字符串的语法错误。返回一个 **Boolean** 类型的值，以表示字符串中是否包含语法错误。如果字符串未包含语法错误，则该值为 **True** 。

**语法**

**express.CheckGrammar ( *String* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                         |
|----------|-----------|----------|------------------------------|
| *String* | 必选      | String   | 要对其进行语法检查的字符串。 |

**示例**

``` JavaScript
/*本示例显示选定部分的语法检查结果。*/
function test() {
    let strPass = Application.CheckGrammar(Application.Selection.Text)
    alert("Selection is grammatically correct: " + strPass)
}
```

### Application.CheckSpelling

检查字符串的拼写错误。返回一个 **Boolean** 类型的值，以显示字符串是否包含拼写错误。如果字符串未包含拼写错误，则该值为 **True** 。

**语法**

**express.CheckSpelling ( *Word* , *CustomDictionary* , *IgnoreUppercase* , *MainDictionary* , *CustomDictionary2* , *CustomDictionary3* , *CustomDictionary4* , *CustomDictionary5* , *CustomDictionary6* , *CustomDictionary7* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------------------|-----------|----------|------------------------------------------------------------------------------------------|
| *Word*              | 必选      | String   | 要对其进行拼写检查的文本。                                                               |
| *CustomDictionary*  | 可选      | Variant  | 可以是返回 Dictionary 对象的表达式，也可以是自定义词典的文件名。                         |
| *IgnoreUppercase*   | 可选      | Variant  | 如果忽略大小写，则该属性值为 True。如果省略该参数，则使用 IgnoreUppercase 属性的当前值。 |
| *MainDictionary*    | 可选      | Variant  | 可以是返回 Dictionary 对象的表达式，也可以是主词典的文件名。                             |
| *CustomDictionary2* | 可选      | Variant  | 可以是返回 Dictionary 对象的表达式，也可以是主词典的文件名。                             |
| *CustomDictionary3* | 可选      | Variant  | 可以是返回 Dictionary 对象的表达式，也可以是主词典的文件名。                             |
| *CustomDictionary4* | 可选      | Variant  | 可以是返回 Dictionary 对象的表达式，也可以是主词典的文件名。                             |
| *CustomDictionary5* | 可选      | Variant  | 可以是返回 Dictionary 对象的表达式，也可以是主词典的文件名。                             |
| *CustomDictionary6* | 可选      | Variant  | 可以是返回 Dictionary 对象的表达式，也可以是主词典的文件名。                             |
| *CustomDictionary7* | 可选      | Variant  | 可以是返回 Dictionary 对象的表达式，也可以是主词典的文件名。                             |

### Application.CleanString

从指定字符串中删除非打印字符（字符代码为 1-29）及 WPS 的特殊字符，或将它们替换为空格（字符代码为 32）。返回的结果为 **String** 类型。

**语法**

**express.CleanString ( *String* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明       |
|----------|-----------|----------|------------|
| *String* | 必选      | String   | 源字符串。 |

**说明**

以下字符将按此表所示进行转换。

| 字符代码          | 说明                                                                  |
|-------------------|-----------------------------------------------------------------------|
| 7（蜂鸣）         | 如果前导字符代码不是 13（段落），则将其删除并转换为字符 9（制表符）。 |
| 10（换行）        | 如果前导字符代码不是 13，则转换为字符 13（段落），然后将其删除。      |
| 13（段落）        | 不改变。                                                              |
| 31（可选连字符）  | 删除。                                                                |
| 160（不间断空格） | 转换为字符 32（空格）。                                               |
| 172（可选连字符） | 删除。                                                                |
| 176（不间断空格） | 转换为字符 32（空格）。                                               |
| 182（段落标记）   | 删除。                                                                |
| 183（项目符号）   | 转换为字符 32（空格）。                                               |

**示例**

``` JavaScript
/*本示例删除选定文本的非打印字符，并将结果插入新文档中。*/
function test() {
    let strClean = Application.CleanString(Application.Selection.Text)
    let docNew = Application.Documents.Add()
    docNew.Content.InsertAfter(strClean)
}
```

### Application.CompareDocuments

比较两个文档，并返回一个 **Document** 对象，该对象表示包含两个文档之间的差别的文档，该文档用修订进行标记。

**语法**

**express.CompareDocuments ( *OriginalDocument* , *RevisedDocument* , *Destination* , *Granularity* , *CompareFormatting* , *CompareCaseChanges* , *CompareWhitespace* , *CompareTables* , *CompareHeaders* , *CompareFootnotes* , *CompareTextboxes* , *CompareFields* , *CompareComments* , *RevisedAuthor* , *IgnoreAllComparisonWarnings* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称                          | 必选/可选 | 数据类型             | 说明                                                                                                    |
|-------------------------------|-----------|----------------------|---------------------------------------------------------------------------------------------------------|
| *OriginalDocument*            | 必选      | Document             | 指定原始文档的路径和文件名。                                                                            |
| *RevisedDocument*             | 必选      | Document             | 指定与原始文档进行比较的修订后文档的路径和文件名。                                                      |
| *Destination*                 | 可选      | WdCompareDestination | 指定是创建新文件还是在原始文档或修订后文档中标记两个文档之间的差别。默认值为 wdCompareDestinationNew 。 |
| *Granularity*                 | 可选      | WdGranularity        | 指定是按字符还是按字词进行修订。默认值为 wdGranularityWordLevel 。                                      |
| *CompareFormatting*           | 可选      | Boolean              | 指定是否标记两个文档之间的格式差别。默认值为 True 。                                                    |
| *CompareCaseChanges*          | 可选      |                      | 指定是否标记两个文档之间的大小写差别。默认值为 True 。                                                  |
| *CompareWhitespace*           | 可选      | Boolean              | 指定是否标记两个文档之间的空白区域（如段落或间距）差别。默认值为 True 。                                |
| *CompareTables*               | 可选      | Boolean              | 指定是否比较两个文档之间的表中所包含数据的差别。默认值为 True 。                                        |
| *CompareHeaders*              | 可选      | Boolean              | 指定是否比较两个文档之间的页眉和页脚的差别。默认值为 True 。                                            |
| *CompareFootnotes*            | 可选      | Boolean              | 指定是否比较两个文档之间的脚注和尾注的差别。默认值为 True 。                                            |
| *CompareTextboxes*            | 可选      | Boolean              | 指定是否比较两个文档之间的文本框中所包含数据的差别。默认值为 True 。                                    |
| *CompareFields*               | 可选      | Boolean              | 指定是否比较两个文档之间的域的差别。默认值为 True 。                                                    |
| *CompareComments*             | 可选      | Boolean              | 指定是否比较两个文档之间的批注的差别。默认值为 True 。                                                  |
| *RevisedAuthor*               | 可选      | String               | 指定在比较两个文档时做出更改的人员的姓名。                                                              |
| *IgnoreAllComparisonWarnings* | 可选      | Boolean              | 指定在比较两个文档时是否忽视警告消息。                                                                  |

### Application.DDEExecute

通过指定的 DDE（即动态数据交换）通道，向应用程序发送一条或一组命令。

**语法**

**express.DDEExecute ( *Channel* , *Command* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                 |
|-----------|-----------|----------|------------------------------------------------------------------------------------------------------|
| *Channel* | 必选      | Long     | DDEInitiate 方法返回的通道号。                                                                       |
| *Command* | 必选      | String   | 由接收应用程序（DDE 服务器）识别的一条或一组命令。如果接收应用程序不能执行指定的命令，则会出现错误。 |

**说明**

**安全性** ??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

**示例**

``` JavaScript
/*本示例在 ET 中创建新的工作表。创建新工作表的 XLM 宏指令是 New(1)。*/
function test() {
    let lngChannel = Application.DDEInitiate("ET","System")
    Application.DDEExecute(lngChannel,"[New(1)]")
    Application.DDETerminate(lngChannel)
}
```

### Application.DDEInitiate

打开通向其他应用程序的 DDE（动态数据交换）通道，并返回通道序号。

**语法**

**express.DDEInitiate ( *App* , *Topic* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                       |
|---------|-----------|----------|----------------------------------------------------------------------------|
| *App*   | 必选      | String   | 应用程序名。                                                               |
| *Topic* | 必选      | String   | DDE 主题名称（如某打开文档的名称），打开通道所指向的应用程序将识别该名称。 |

**说明**

**安全性** ??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

如果成功， **DDEInitiate** 方法将返回打开通道的序号。所有后续 DDE 函数使用该序号来指定通道。

**示例**

``` JavaScript
/*本示例用 System 主题启动 DDE 会话，并打开ET 工作簿 Sales.xls。然后，本示例终止 DDE 通道，启动通向 Sales.xls 的通道，并在单元格 R1C1 中插入文本。*/
function test() {
    let lngChannel = Application.DDEInitiate("ET", "System")
    Application.DDEExecute(lngChannel, "[OPEN(" + "C:\Sales.xls)]")
    Application.DDETerminate(lngChannel)
    lngChannel = Application.DDEInitiate("ET", "Sales.xls")
    Application.DDEPoke(lngChannel, "R1C1", "1996 Sales")
    Application.DDETerminate(lngChannel)
}
```

### Application.DDEPoke

通过打开的 DDE（动态数据交换）通道向应用程序发送数据。

**语法**

**express.DDEPoke ( *Channel* , *Item* , *Data* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                            |
|-----------|-----------|----------|-------------------------------------------------|
| *Channel* | 必选      | Long     | DDEInitiate 方法返回的通道号。                  |
| *Item*    | 必选      | String   | DDE 主题中的项，指定的数据将发送给该项。        |
| *Data*    | 必选      | String   | 需要发送至接收应用程序（即 DDE 服务器）的数据。 |

**说明**

**安全性** ??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

如果使用 **DDEPoke** 方法不成功，将导致出错。

**示例**

``` JavaScript
/*本示例打开 ET 工作表 Sales.xls，然后在 R1C1 单元格内插入“1996 Sales”。*/
function test() {
    let lngChannel = Application.DDEInitiate("ET","System")
    Application.DDEExecute(lngChannel,"[OPEN(C:\Sales.xls)]")
    Application.DDETerminate(lngChannel)
    lngChannel = Application.DDEInitiate("ET", "Sales.xls")
    Application.DDEPoke(lngChannel,"R1C1", "1996 Sales")
    Application.DDETerminate(lngChannel)
}
```

### Application.DDERequest

通过打开的 DDE（动态数据交换）通道向接收应用程序查询信息，并以 **String** 类型返回该信息。

**语法**

**express.DDERequest ( *Channel* , *Item* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                           |
|-----------|-----------|----------|--------------------------------|
| *Channel* | 必选      | Long     | DDEInitiate 方法返回的通道号。 |
| *Item*    | 必选      | String   | 需要进行查询的项。             |

**说明**

**安全性** ??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

请求服务器应用程序的主题信息时，必须指定包含所请求内容的主题中的项目。例如，在ET 中，单元格是合法项目，可以通过“R1C1”格式或已命名的引用来引用单元格。

ET 和其他支持 DDE 的应用程序可识别名为“System”的主题。下表对 System 主题中的三个标准项目加以说明。请注意：使用 SysItems 项目可得到 System 主题中的其他项目列表。

| System 主题中的项目 | 效果                                  |
|---------------------|---------------------------------------|
| SysItems            | 返回 System 主题中的所有项目的列表。  |
| Topics              | 返回所有有效主题的列表。              |
| Formats             | 返回 WPS 支持的所有剪贴板格式的列表。 |

**示例**

``` JavaScript
/*本示例打开 ET 工作簿 Book1.xls，然后检索单元格 R1C1 的内容。*/
function test() {
    let lngChannel = DDEInitiate("ET", "System")
    Application.DDEExecute(lngChannel,"[OPEN(C:\Documents\Book1.xls)]")
    Application.DDETerminate(lngChannel)
    lngChannel = Application.DDEInitiate("ET","Book1.xls")
    alert(Application.DDERequest(lngChannel,"R1C1"))
    Application.DDETerminateAll()
}
```

### Application.DDETerminate

关闭指定的通向另一应用程序的 DDE（动态数据交换）通道。

**语法**

**express.DDETerminate ( *Channel* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                           |
|-----------|-----------|----------|--------------------------------|
| *Channel* | 必选      | Long     | DDEInitiate 方法返回的通道号。 |

**说明**

**安全性** ??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

**示例**

``` JavaScript
/*本示例在 ET 中创建一张新工作簿，然后终止 DDE 会话。*/
function test() {
    let lngChannel = DDEInitiate("ET", "System")
    Application.DDEExecute(lngChannel,"[New(1)]")
    Application.DDETerminate(lngChannel)
}
```

### Application.DDETerminateAll

关闭所有由 WPS 打开的动态数据交换 (DDE) 通道。

**语法**

**express.DDETerminateAll ()**

\*express是一个代表 Application 对象的变量。

**说明**

此方法不关闭由客户端应用程序打开的到 WPS 的通道。使用此方法与对每一个打开的通道使用 **DDETerminate** 方法等效。

**安全性** ??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

如果中断打开 DDE 通道的宏，可能会无意中使一个通道处于打开状态。宏结束时打开的通道不会自行关闭，并且每一个打开的通道都会占用系统资源。因此，在调试打开一个或多个 DDE 通道的宏时，最好使用此方法关闭 DDE 通道。

**示例**

``` JavaScript
/*本示例打开ET 工作簿 Book1.xls，在单元格 R2C3 中插入文本，然后保存此工作簿，再关闭所有的 DDE 通道。*/
function test() {
    let lngChannel = DDEInitiate("ET", "System")
    Application.DDEExecute(lngChannel,"[OPEN(C:\Documents\Book1.xls)]")
    Application.DDETerminate(lngChannel)
    lngChannel = Application.DDEInitiate("ET","Book1.xls")
    Application.DDEPoke(lngChannel, "R2C3", "Hello World")
    Application.DDEExecute(lngChannel,"[Save]")
    Application.DDETerminateAll()
}
```

### Application.DefaultWebOptions

返回 **DefaultWebOptions** 对象，该对象包含 WPS 每当将文档保存为网页或打开网页时使用的应用程序级全局属性。

**语法**

**express.DefaultWebOptions ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例检查文档编码的默认设置是否为“Western”，然后根据检查结果设置 strDocEncoding 字符串。*/
function test() {
    let strDocEncoding

    if(Application.DefaultWebOptions.Encoding == msoEncodingWestern) {
        strDocEncoding = "Western"
    } else {
        strDocEncoding = "Other"
    }
}
```

### Application.FileDialog

返回一个 **FileDialog** 对象，该对象代表文件对话框的单个实例。

**语法**

**express.FileDialog ( *FileDialogType* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型          | 说明         |
|------------------|-----------|-------------------|--------------|
| *FileDialogType* | 必选      | MsoFileDialogType | 对话框类型。 |

**示例**

``` JavaScript
/*该示例显示“另存为”对话框。*/
function ShowSaveAsDialog() {
    let dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
    dlgSaveAs.Show()
}
```

``` JavaScript
/*本示例显示“打开”对话框并允许用户选择打开多个文件。*/
function ShowFileDialog() {
    let dlgOpen = Application.FileDialog(msoFileDialogOpen)
    dlgOpen.AllowMultiSelect = true
    dlgOpen.Show()
}
```

### Application.FindKey

返回 **KeyBinding** 对象，该对象代表指定的组合键。只读。

**语法**

**express.FindKey ( *KeyCode* , *KeyCode2* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                              |
|------------|-----------|----------|-----------------------------------|
| *KeyCode*  | 必选      | Long     | 用 WdKey 常量之一指定的一个键。   |
| *KeyCode2* | 可选      | Variant  | 用 WdKey 常量之一指定的第二个键。 |

**说明**

可用 **BuildKeyCode** 方法创建 *KeyCode* 或 *KeyCode2* 参数。

**示例**

``` JavaScript
/*本示例取消活动文档模板中的 Alt+Shift+F12。要返回包含不止两个键的 KeyBinding 对象，请用 BuildKeyCode 方法，如本例所示。*/
function test() {
    let CustomizationContext = Application.ActiveDocument.AttachedTemplate
    Application.FindKey(Application.BuildKeyCode(Application.Enum.wdKeyAlt,Application.Enum.wdKeyShift,Application.Enum.wdKeyF12).Disable()
}
```

``` JavaScript
/*本示例显示 F1 指向的命令。*/
function test() {
    CustomizationContext = NormalTemplate
    alert((Application.FindKey(Application.Enum.wdKeyF1).Command))
}
```

### Application.GetAddress

返回默认通讯簿中的地址。

**语法**

**express.GetAddress ( *Name* , *AddressProperties* , *UseAutoText* , *DisplaySelectDialog* , *SelectDialog* , *CheckNamesDialog* , *RecentAddressesChoice* , *UpdateRecentAddresses* )**

\*express是一个代表 Application 对象的变量。

**参数**

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th>名称</th>
<th>必选/可选</th>
<th>数据类型</th>
<th>说明</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td class="mainsection"><em>Name</em></td>
<td class="mainsection">可选</td>
<td class="mainsection">Variant</td>
<td class="mainsection" style="word-break: break-all">收件人姓名，与通讯簿中“查找姓名”对话框中所指定的相同。</td>
</tr>
<tr class="even">
<td class="mainsection"><em>AddressProperties</em></td>
<td class="mainsection">可选</td>
<td class="mainsection">Variant</td>
<td class="mainsection" style="word-break: break-all">如果 UseAutoText 为 True，则该参数代表定义通讯簿属性系列的“自动图文集”词条的名称。 如果 UseAutoText 为 False 或被省略，则该参数定义自定义布局。有效通讯簿属性名称或属性名称集合用尖括号（“&lt;”和“&gt;”）括起来并用空格或段落标记隔开（例如，“ ”&amp; vbCr &amp;“ ”）。 如果省略 AddressProperties 参数，则使用名为“AddressLayout”的默认“自动图文集”词条。如果还没有定义“AddressLayout”，则使用下面的地址布局定义：" " &amp; vbCr &amp; " " &amp; vbCr &amp; " " &amp; ", " &amp; " " &amp; " " &amp; " " &amp; vbCr &amp; " "。 要获取有效的通讯簿属性名列表，请参阅 AddAddress 方法。</td>
</tr>
<tr class="odd">
<td class="mainsection"><em>UseAutoText</em></td>
<td class="mainsection">可选</td>
<td class="mainsection">Variant</td>
<td class="mainsection" style="word-break: break-all">如果 AddressProperties 指定了定义通讯簿属性系列的“自动图文集”词条的名称，则为 True ；如果它指定了一个自定义布局，则为 False。</td>
</tr>
<tr class="even">
<td class="mainsection"><em>DisplaySelectDialog</em></td>
<td class="mainsection">可选</td>
<td class="mainsection">Variant</td>
<td class="mainsection" style="word-break: break-all">指定是否显示“选择姓名”对话框，如下表所示。
<table>
<thead>
<tr class="header">
<th>值</th>
<th>说明</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td>0（零）</td>
<td>不显示 “选择姓名” 对话框。</td>
</tr>
<tr class="even">
<td>1 或省略</td>
<td>显示 “选择姓名” 对话框。</td>
</tr>
<tr class="odd">
<td>2</td>
<td>不显示 “选择姓名” 对话框，并且不搜索指定的姓名。该方法返回的地址将是以前选择的地址。</td>
</tr>
</tbody>
</table></td>
</tr>
<tr class="odd">
<td class="mainsection"><em>SelectDialog</em></td>
<td class="mainsection">可选</td>
<td class="mainsection">Variant</td>
<td class="mainsection" style="word-break: break-all">指定“选择姓名”对话框的显示方式（即以何种模式显示），如下表所示。
<table>
<thead>
<tr class="header">
<th>值</th>
<th>显示模式</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td>0（零）或省略</td>
<td>浏览模式</td>
</tr>
<tr class="even">
<td>1</td>
<td>紧凑模式，只显示 “收件人:” 框</td>
</tr>
<tr class="odd">
<td>2</td>
<td>紧凑模式， “收件人:” 和 “抄送:” 框都显示</td>
</tr>
</tbody>
</table></td>
</tr>
<tr class="even">
<td class="mainsection"><em>CheckNamesDialog</em></td>
<td class="mainsection">可选</td>
<td class="mainsection">Variant</td>
<td class="mainsection" style="word-break: break-all">如果该属性值为 True ，则在 Name 参数的值不够具体时显示“检查姓名”对话框。</td>
</tr>
<tr class="odd">
<td class="mainsection"><em>RecentAddressesChoice</em></td>
<td class="mainsection">可选</td>
<td class="mainsection">Variant</td>
<td class="mainsection" style="word-break: break-all">如果该属性值为 True ，则使用最近使用的回信地址列表。</td>
</tr>
<tr class="even">
<td class="mainsection"><em>UpdateRecentAddresses</em></td>
<td class="mainsection">可选</td>
<td class="mainsection">Variant</td>
<td class="mainsection" style="word-break: break-all">如果该属性值为 True ，则向最近使用的地址列表中添加一个地址；如果该属性值为 False，则不添加地址。如果 SelectDialog 设置为 1 或 2，则忽略该参数。</td>
</tr>
</tbody>
</table>

**示例**

``` JavaScript
/*本示例将 John Smith 的地址赋给变量 strAddress，将插入点移到文档的开头，并插入地址。插入的地址包括默认的通讯簿属性。*/
function test() {
    let strAddress = Application.GetAddress("John Smith", null, null, null, null, true)
    Application.ActiveDocument.Range(0, 0).InsertAfter(strAddress)
}
```

### Application.GetDefaultTheme

返回代表默认 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">主题</span> 名称的 **String** ，以及 WPS 用于新文档、电子邮件或网页的主题格式选项。

**语法**

**express.GetDefaultTheme ( *DocumentType* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型         | 说明                                 |
|----------------|-----------|------------------|--------------------------------------|
| *DocumentType* | 必选      | WdDocumentMedium | 要为其检索默认主题名的新文档的类型。 |

**说明**

也可使用 **ThemeName** 属性为新电子邮件返回并设置默认主题。

**示例**

``` JavaScript
/*本示例显示 WPS 用于新网页的主题名。*/
alert(Application.GetDefaultTheme(wdWebPage))
```

### Application.GetSpellingSuggestions

返回一个 **SpellingSuggestions** 集合，该集合代表指定单词的建议拼写替换单词。

**语法**

**express.GetSpellingSuggestions ( *Word* , *IgnoreUppercase* , *SuggestionMode* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型 | 说明                                                                                                       |
|-------------------|-----------|----------|------------------------------------------------------------------------------------------------------------|
| *Word*            | 必选      | String   | 将接受拼写检查的单词。                                                                                     |
| *IgnoreUppercase* | 可选      | Variant  | 如果为 True IgnoreUppercase                                                                                |
| *SuggestionMode*  | 可选      | Variant  | 指定 WPS 提出拼写建议的方式。可以是下列 WdSpellingWordType wdAnagram wdSpellword wdWildcard WdSpellword 。 |

**说明**

如果此单词拼写正确，则 **SpellingSuggestions** 对象的 **Count** 属性返回 0（零）。

### Application.GoBack

将插入点在活动文档中最近进行过编辑操作的三个位置间移动，其作用等同于按 Shift+F5。

**语法**

**express.GoBack ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例打开最近使用过的文件，并将插入点移至文档中最后进行编辑操作的位置。*/
function test() {
    Application.RecentFiles.Item(1).Open()
    Application.GoBack()
}
```

### Application.GoForward

将插入点在活动文档中进行编辑的最后三个位置之间向前移动。

**语法**

**express.GoForward ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将插入点移至曾经进行过编辑的下一个位置。*/
Application.GoForward()
```

### Application.Help

显示安装的帮助信息。

**语法**

**express.Help ( *HelpType* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                             |
|------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *HelpType* | 可选      | Variant  | 联机帮助主题或窗口。可以是下列 WdHelpType 常量之一： wdHelp 、 wdHelpAbout 、 wdHelpActiveWindow 、 wdHelpContents 、 wdHelpHWP 、 wdHelpIchitaro 、 wdHelpIndex 、 wdHelpPE2 、 wdHelpPSSHelp 、 wdHelpSearch 和 wdHelpUsingHelp （由于选择或安装的语言不同，此处列出的某些常量可能无法使用）。 |

**示例**

``` JavaScript
/*本示例显示“帮助主题”对话框。*/
Application.Help(Application.Enum.wdHelp)
```

### Application.HelpTool

指针从箭头改为问号，表明这是关于将要单击的命令或屏幕元素的快捷帮助。如果单击文本， WPS 将显示描述当前段落及字符格式的对话框。按 Esc 可将指针恢复为原始状态。

**语法**

**express.HelpTool ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将鼠标指针从箭头变为问号。*/
Application.HelpTool()
```

### Application.InchesToPoints

将度量单位从英寸转换为磅（1 英寸 = 72 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.InchesToPoints ( *Inches* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                   |
|----------|-----------|----------|------------------------|
| *Inches* | 必选      | Single   | 要转换为磅值的英寸值。 |

**示例**

``` JavaScript
/*本示例将选定段落的段前间距设置为 0.25 英寸。*/
Application.Selection.ParagraphFormat.SpaceBefore =Application. InchesToPoints(0.25)
```

### Application.International

返回当前的国家/地区和国际设置信息。 **Variant** 类型，只读。

**语法**

**express.International ( *Index* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型             | 说明                           |
|---------|-----------|----------------------|--------------------------------|
| *Index* | 必选      | WdInternationalIndex | 当前的国家/地区和/或国际设置。 |

**示例**

``` JavaScript
/*本示例在状态栏中显示货币格式。*/
StatusBar = "Currency Format: " + Application.International(Application.Enum.wdCurrencyCode)
```

### Application.KeysBoundTo

返回一个 **KeysBoundTo** 对象，该对象代表指定项的所有组合键。

**语法**

**express.KeysBoundTo ( *KeyCategory* , *Command* , *CommandParameter* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型      | 说明                                                                                                          |
|--------------------|-----------|---------------|---------------------------------------------------------------------------------------------------------------|
| *KeyCategory*      | 必选      | WdKeyCategory | 组合键的类别。                                                                                                |
| *Command*          | 必选      | String        | 命令的名称。                                                                                                  |
| *CommandParameter* | 可选      | Variant       | 由 Command 指定的命令所需的附加文字（如果有）。有关详细信息，请参阅 KeyBindings 对象的 Add 方法的“说明”部分。 |

**示例**

``` JavaScript
/*本示例显示活动文档附加的模板中为 FileOpen 命令指定的所有组合键。*/
function test() {
    let kbLoop = Application.KeysBoundTo(Application.Enum.wdKeyCategoryCommand, "FileOpen")
    CustomizationContext = Application.ActiveDocument.AttachedTemplate

    for (let i = 1; i <= kbLoop.Count; i++) {
        strOutput = strOutput + kbLoop.Item(i).KeyString + "\r"
    }
    alert(strOutput)
}
```

``` JavaScript
/*本示例删除“Normal”模板中“Macro1”的所有按键指定方案。*/
function test() {
    let kb = KeysBoundTo(Application.Enum.wdKeyCategoryMacro, "Macro1")
    CustomizationContext = NormalTemplate
    for (let i = 1; i <= kb.Count; i++) {
        kb.Item(i).Disable()
    }
}
```

### Application.Keyboard

返回或设置键盘的语言和布局设置。

**语法**

**express.Keyboard ( *LangId* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                  |
|----------|-----------|----------|-------------------------------------------------------------------------------------------------------|
| *LangId* | 可选      | Long     | {description}WPS 要针对其设置键盘的语言和布局组合。如果省略该参数，则该方法返回当前的语言和布局设置。 |

**说明**

Microsoft Windows 用一种名为“输入语言句柄”的变量类型来跟踪键盘的语言和布局设置，常把这种变量称为“HKL”。句柄的低端字是语言 ID，句柄的高端字是键盘布局。

**示例**

``` JavaScript
/*本示例将当前的键盘语言和布局设置赋给一个变量。*/
let lngKeyboard = Application.Keyboard()
```

### Application.KeyboardBidi

将键盘语言设置为从右向左的语言，并将文字的输入方向设置为从右向左。

**语法**

**express.KeyboardBidi ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*将键盘语言设置为从右向左的语言，并将文字的输入方向设置为从右向左。*/
Application.KeyboardBidi()
```

### Application.KeyboardLatin

将键盘语言设置为从左向右的语言，并将文字的输入方向设置为从左向右。

**语法**

**express.KeyboardLatin ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例为从左向右的语言输入配置键盘。*/
Application.KeyboardLatin()
```

### Application.KeyString

为指定的键返回组合键字符串（例如，Ctrl+Shift+A）。

**语法**

**express.KeyString ( *KeyCode* , *KeyCode2* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                              |
|------------|-----------|----------|-----------------------------------|
| *KeyCode*  | 必选      | Long     | 用某个 WdKey 常量指定的键。       |
| *KeyCode2* | 可选      | Variant  | 用某个 WdKey 常量指定的另一个键。 |

**说明**

可以使用 **BuildKeyCode** 方法创建 *KeyCode* 或 *KeyCode2* 参数。

**示例**

``` JavaScript
/*本示例为以下三个 <b>wdKeyA</b> 常量显示组合键字符串 (Ctrl+Shift+A)：<b>WdKey</b>、<b>wdKeyControl</b> 和
<b>wdKeyShift</b>。*/
function test() {
    Application.CustomizationContext = Application.ActiveDocument.AttachedTemplate
    alert(KeyString(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyA)))
}
```

### Application.LinesToPoints

将度量单位从行转换为磅（1 行 = 12 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.LinesToPoints ( *Lines* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Lines* | 必选      | Single   | 要转换为磅值的行值。 |

**示例**

``` JavaScript
/*本示例将选定段落的行距设置为 3 倍行距。*/
function test() {
    Application.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
    Application.Selection.ParagraphFormat.LineSpacing = LinesToPoints(3)
}
```

### Application.ListCommands

新建文档，然后插入带有相关快捷键和菜单设定的 WPS 命令表。

**语法**

**express.ListCommands ( *ListAllCommands* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型 | 说明                                                                                                                               |
|-------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------|
| *ListAllCommands* | 必选      | Boolean  | 如果该属性值为 True ，则包括所有的 WPS 命令及其设定（不论是自定义的还是内置的）。如果该属性值为 False ，则只包括自定义设定的命令。 |

**示例**

``` JavaScript
/*本示例将新建文档，该文档将列出所有的 WPS 命令及其快捷键和菜单设定，然后打印并关闭新文档，不保存所进行的修改。*/
function test() {
    Application.ListCommands(true)
    Application.ActiveDocument.PrintOut()
    Application.ActiveDocument.Close(wdDoNotSaveChanges)
}
```

### Application.LoadMasterList

加载书目源文件。

**语法**

**express.LoadMasterList ( *FileName* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                       |
|------------|-----------|----------|----------------------------|
| *FileName* | 必选      | String   | 书目源文件的路径和文件名。 |

### Application.LookupNameProperties

在全局通讯簿列表中查找一个姓名，并显示 **“属性”** 对话框，其中包含了有关指定姓名的信息。

**语法**

**express.LookupNameProperties ( *Name* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| *Name* | 必选      | String   | 全局通讯簿中的一个姓名 |

**说明**

如果该方法找到了多个匹配姓名，则显示 **“检查姓名”** 对话框。

**示例**

``` JavaScript
/*本示例在通讯簿中查找“Don Funk”这一姓名，并显示“Don Funk”的“属性”对话框。*/
Application.LookupNameProperties("Don Funk")
```

### Application.MergeDocuments

比较两个文档，并返回一个 **Document** 对象，该对象表示包含两个文档之间的差别的文档，该文档用修订进行标记。

**语法**

**express.MergeDocuments ( *OriginalDocument* , *RevisedDocument* , *Destination* , *Granularity* , *CompareFormatting* , *CompareCaseChanges* , *CompareWhitespace* , *CompareTables* , *CompareHeaders* , *CompareFootnotes* , *CompareFields* , *CompareComments* , *OriginalAuthor* , *RevisedAuthor* , *FormatFrom* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型             | 说明                                                                                                    |
|----------------------|-----------|----------------------|---------------------------------------------------------------------------------------------------------|
| *OriginalDocument*   | 必选      | Document             | 指定原始文档的路径和文件名。                                                                            |
| *RevisedDocument*    | 必选      | Document             | 指定与原始文档进行比较的修订后文档的路径和文件名。                                                      |
| *Destination*        | 可选      | WdCompareDestination | 指定是创建新文件还是在原始文档或修订后文档中标记两个文档之间的差别。默认值为 wdCompareDestinationNew 。 |
| *Granularity*        | 可选      | WdGranularity        | 指定是按字符还是按字词进行修订。默认值为 wdGranularityWordLevel 。                                      |
| *CompareFormatting*  | 可选      | Boolean              | 指定是否标记两个文档之间的格式差别。默认值为 True 。                                                    |
| *CompareCaseChanges* | 可选      | Boolean              | 指定是否标记两个文档之间的大小写差别。默认值为 True 。                                                  |
| *CompareWhitespace*  | 可选      | Boolean              | 指定是否标记两个文档之间的空白区域（如段落或间距）差别。默认值为 True 。                                |
| *CompareTables*      | 可选      | Boolean              | 指定是否比较两个文档之间的表中所包含数据的差别。默认值为 True 。                                        |
| *CompareHeaders*     | 可选      | Boolean              | 指定是否比较两个文档之间的页眉和页脚的差别。默认值为 True 。                                            |
| *CompareFootnotes*   | 可选      | Boolean              | 指定是否比较两个文档之间的脚注和尾注的差别。默认值为 True 。                                            |
| *CompareFields*      | 可选      | Boolean              | 指定是否比较两个文档之间的域的差别。默认值为 True 。                                                    |
| *CompareComments*    | 可选      | Boolean              | 指定是否比较两个文档之间的批注的差别。默认值为 True 。                                                  |
| *OriginalAuthor*     | 可选      | String               | 指定原始文档作者的姓名。                                                                                |
| *RevisedAuthor*      | 可选      | String               | 指定合并两个文档之后进行非属性化更改的人员的姓名。                                                      |
| *FormatFrom*         | 可选      | WdMergeFormatFrom    | 指定要保留其格式的文档。                                                                                |

### Application.MillimetersToPoints

把度量单位从毫米转换为磅值（1 毫米 = 2.85 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.MillimetersToPoints ( *Millimeters* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                   |
|---------------|-----------|----------|------------------------|
| *Millimeters* | 必选      | Single   | 要转换为磅值的毫米值。 |

**示例**

``` JavaScript
/*本示例将活动文档的断字区设置为 8.8 毫米。*/
Application.ActiveDocument.HyphenationZone = Application.MillimetersToPoints(8.8)
```

### Application.Move

设置任务窗口或活动文档窗口的位置。

**语法**

**express.Move ( *Left* , *Top* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                     |
|--------|-----------|----------|--------------------------|
| *Left* | 必选      | Long     | 指定窗口的水平屏幕位置。 |
| *Top*  | 必选      | Long     | 指定窗口的垂直屏幕位置。 |

**示例**

``` JavaScript
/*本示例启动“计算器”应用程序 (Calc.exe)，并使用 Move 方法重新调整应用程序窗口的位置。*/
function test() {
    Application.Tasks("Calculator");
    Application.WindowState = Application.Enum.wdWindowStateNormal
    Application.Move(50,50)
}
```

### Application.NewWindow

打开一个新的窗口作为指定窗口，并显示原文档。返回一个 ****Window**** 对象。

**语法**

**express.NewWindow ()**

\*express是一个代表 Application 对象的变量。

**说明**

对同一文档打开多个窗口时，窗口标题将出现一个冒号 (:) 和一个数字。如果将 **NewWindow** 方法与 **Application** 对象配合使用，则为活动窗口打开一个新的窗口。下列两条指令在功能上是等效的。

``` JavaScript
let myWindow = Application.ActiveDocument.ActiveWindow.NewWindow()
myWindow = Application.NewWindow()
```

**示例**

``` JavaScript
/* 新增一个窗口，并且打印新增前后窗口数量 */
function test() {
    alert(Application.Windows.Count);
    Application.NewWindow();
    alert(Application.Windows.Count);
}
```

### Application.NextLetter

您已经就仅在 Macintosh 上使用的 Visual Basic 关键字请求帮助。有关 **Application** 对象的 **NextLetter** 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。

**语法**

**express.NextLetter ()**

\*express是一个代表 Application 对象的变量。

### Application.OnTime

启动在指定的时间运行宏的后台计时器。

**语法**

**express.OnTime ( *When* , *Name* , *Tolerance* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                    |
|-------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *When*      | 必选      | Variant  | 运行宏的时间。                                                                                                                                                                                                                                                          |
| *Name*      | 必选      | String   | 要运行的宏的名称。                                                                                                                                                                                                                                                      |
| *Tolerance* | 可选      | Variant  | 表示未在 When 指定的时间运行的宏最多经过多长时间（单位为秒）被取消。宏不一定总在指定时间运行。例如，如果正在进行排序操作，或者正在显示某一对话框，则宏将延迟到 WPS 完成当前任务后才运行。如果该参数值为 0（零）或被省略，则不管距 When 指定的时间逾期多久，宏都将运行。 |

**说明**

*When* 参数可以是指定时间的字符串（如 ` "4:30 pm" ` 或 ` "16:30" ` ），也可以是诸如 **TimeValue** 或 **TimeSerial** 等函数返回的一系列数字（如 ` TimeValue("2:30 pm") ` 或 ` TimeSerial(14, 30, 00) ` ）。也可以包括日期（如 ` "6/30 4:15 pm" ` 或 ` TimeValue("6/30 4:15 pm") ` ）。

对于 *Name* 参数，应使用完整的宏路径来确保运行正确的宏（如 ` "Project.Module1.Macro1" ` ）。对于要运行的宏，文档或模板必须同时满足以下条件：既在 **OnTime** 指令运行时可用，又在到达 *When* 指定的时间时可用。因此，最好将宏存储在 Normal.dot 或其他能够自动加载的共用模板上。

使用 **Now** 函数以及 **TimeValue** 或 **TimeSerial** 函数的返回值的和可以设置计时器，使宏在语句运行后的指定时间运行。例如，使用 ` Now+TimeValue("00:05:30") ` 可以使宏在语句运行后再过 5 分 30 秒运行。

WPS 只能保留一个由 **OnTime** 设置的后台计时器。如果在现有计时器运行前启动另一个计时器，则将取消现有计时器。

**示例**

``` JavaScript
/*本示例在 3:55 P.M 运行当前模块中名为“Macro1”的宏。*/
Application.OnTime("15:55:00","Macro1")
```

``` JavaScript
/*本示例在运行 15 秒后运行名为“Macro1”的宏。*/
function test()
{
let now = new Date()
let now_tmp = now.getTime()
//获取15秒后的时间
now.setTime(now_tmp + 1000 * 15)
Application.OnTime(now.toLocaleString(),"Macro1")
}
```

``` JavaScript
/*本示例在 1:30 P.M 运行名为“Start”的宏。宏的名称包括了项目名称和模块名称。*/
Application.OnTime("1:30 pm", "Start")
```

### Application.OrganizerCopy

将指定的“自动图文集”词条、工具栏、样式或宏方案项从源文档或模板中复制到目标文档或模板上。

**语法**

**express.OrganizerCopy ( *Source* , *Destination* , *Name* , *Object* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型          | 说明                                               |
|---------------|-----------|-------------------|----------------------------------------------------|
| *Source*      | 必选      | String            | 含有要复制的条目的文档或模板的名称。               |
| *Destination* | 必选      | String            | 要向其复制条目的文档或模板的名称。                 |
| *Name*        | 必选      | String            | 要复制的“自动图文集”词条、工具栏、样式或宏的名称。 |
| *Object*      | 必选      | WdOrganizerObject | 要复制的项目的类型。                               |

**示例**

``` JavaScript
/*本示例将与活动文档相关的模板中的所有“自动图文集”词条复制到 Normal 模板中。*/
function test() {
    let atEntry = Application.ActiveDocument.AttachedTemplate.AutoTextEntries
    for (let i = 1; i <= atEntry.Count; i++) {
        Application.OrganizerCopy(Application.ActiveDocument.AttachedTemplate.FullName, NormalTemplate.FullName, atEntry.Item(i).Name, wdOrganizerObjectAutoText)
    }
}
```

``` JavaScript
/*如果活动文档中含有名为“SubText”的样式，本示例将该样式复制到 C:\Templates\Template1.dot 中。*/
function test() {
    let styleLoop = Application.ActiveDocument.Styles
    for (let i = 1; i <= styleLoop.Count; i++) {
        if (styleLoop.Item(i) == "SubText") {
            Application.OrganizerCopy(Application.ActiveDocument.Name, "C:\\Templates\\Template1.dot", "SubText", wdOrganizerObjectStyles)
        }
    }
}
```

### Application.OrganizerDelete

从文档或模板中删除指定的样式、“自动图文集”词条、工具栏或宏方案项。

**语法**

**express.OrganizerDelete ( *Source* , *Name* , *Object* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型          | 说明                                               |
|----------|-----------|-------------------|----------------------------------------------------|
| *Source* | 必选      | String            | 含有需要删除的条目的文档或模板的名称。             |
| *Name*   | 必选      | String            | 要删除的样式、“自动图文集”词条、工具栏或宏的名称。 |
| *Object* | 必选      | WdOrganizerObject | 要复制的项目的类型。                               |

**示例**

``` JavaScript
/*本示例将从 Normal 模板中删除名为“Custom 1”的工具栏。*/
function test() {
    for (let i = 1; i <= Application.CommandBars.Count; i++) {
        if (Application.CommandBars.Item(i).Name == "Custom 1") {
            Application.OrganizerDelete(NormalTemplate.Name, "Custom 1", wdOrganizerObjectCommandBars)
        }
    }
}
```

``` JavaScript
/*本示例提示用户删除活动文档的相关模板中的每一个“自动图文集”词条。如果用户单击“确定”按钮，则将删除“自动图文集”词条。*/
function test() {
    let intResponse
    let temp = Application.ActiveDocument.AttachedTemplate
    let atEntry = Application.ActiveDocument.AttachedTemplate.AutoTextEntries

    for (let i = 1; i <= atEntry.Count; i++) {
        intResponse = alert("Do you want to delete the " + atEntry.Item(i).Name + " AutoText entry?", jsYesNoCancel)
        if (intResponse == jsResultYes) {
            Application.OrganizerDelete(temp.Path + "\\" + temp.Name, atEntry.Item(i).Name, wdOrganizerObjectAutoText)
        }
        if (intResponse == jsResultCancel) {
            break
        }
    }
}
```

### Application.OrganizerRename

重新命名文档或模板中指定的样式、“自动图文集”词条、工具栏或宏方案项。

**语法**

**express.OrganizerRename ( *Source* , *Name* , *NewName* , *Object* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型          | 说明                                                 |
|-----------|-----------|-------------------|------------------------------------------------------|
| *Source*  | 必选      | String            | 含有需要重新命名的项目的文档或模板的名称。           |
| *Name*    | 必选      | String            | 要重命名的样式、“自动图文集”词条、工具栏或宏的名称。 |
| *NewName* | 必选      | String            | 项目的新名称。                                       |
| *Object*  | 必选      | WdOrganizerObject | 要复制的项目的类型。                                 |

**示例**

``` JavaScript
/*本示例将活动文档中名为“SubText”的样式改为“SubText2”。*/
function test() {
    let styleLoop = Application.ActiveDocument.Styles
    for (let i = 1; i <= styleLoop.Count; i++) {
        if (styleLoop.Item(i).NameLocal == "SubText") {
            Application.OrganizerRename(Application.ActiveDocument.Name, "SubText", "SubText2", wdOrganizerObjectStyles)
        }
    }
}
```

``` JavaScript
/*本示例将附加模板中名为“Module1”的宏改为“Macros1”。*/
function test()
{
    let dotTemp = Application.ActiveDocument.AttachedTemplate.Name
    Application.OrganizerRename(dotTemp,"Module1","Macros1",wdOrganizerObjectProjectItems)
}
```

### Application.PicasToPoints

将度量单位从十二点活字转换为磅值（1 十二点活字 = 12 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PicasToPoints ( *Picas* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                         |
|---------|-----------|----------|------------------------------|
| *Picas* | 必选      | Single   | 要转换为磅值的十二点活字数。 |

**返回值**

Single

**示例**

``` JavaScript
本示例将为活动文档添加行号，并将行号和文档文本之间的距离设置为 4 个十二点活字。
function test()
{
ActiveDocument.PageSetup.LineNumbering.Active = true
ActiveDocument.PageSetup.LineNumbering.DistanceFromText = PicasToPoints(4)
}
```

``` JavaScript
/*本示例将选定段落的首行缩进设置为 3 个十二点活字。*/
Application.Selection.ParagraphFormat.FirstLineIndent = PicasToPoints(3)
```

### Application.PixelsToPoints

将长度值的单位由像素转换为磅。以 **Single** 类型返回转换结果。

**语法**

**express.PixelsToPoints ( *Pixels* , *fVertical* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                        |
|-------------|-----------|----------|-----------------------------------------------------------------------------|
| *Pixels*    | 必选      | Single   | 要转换为磅值的像素值。                                                      |
| *fVertical* | 可选      | Variant  | 如果该属性值为 True，则转换垂直像素；如果该属性值为 False，则转换水平像素。 |

**返回值**

Single

**示例**

``` JavaScript
/*本示例将显示一个以像素度量的对象在以磅为单位时的高度和宽度。*/
alert("320x240 pixels is equivalent to " + PixelsToPoints(320, false) + "x" + PixelsToPoints(240, true) + " points on this display.")
```

### Application.PointsToCentimeters

将度量单位由磅转换为厘米（1 厘米 = 28.35 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToCentimeters ( *Points* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                 |
|----------|-----------|----------|----------------------|
| *Points* | 必选      | Single   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

``` JavaScript
/*本示例将 30 磅的度量数值转换为相应的厘米数。*/
alert(PointsToCentimeters(30) + " centimeters")
```

``` JavaScript
/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/
function test()
{
let intUnit
let sngData
                                    
function ConvertPoints(intUnit,sngData){
    switch(intUnit) {
    case 1:
        ConvertPoints = PointsToCentimeters(sngData)
    case 2:
        ConvertPoints = PointsToInches(sngData)
    case 3:
        ConvertPoints = PointsToLines(sngData)
    case 4:
        ConvertPoints = PointsToMillimeters(sngData)
    case 5:
        ConvertPoints = PointsToPicas(sngData)
    default:
        alert("无效的过程调用")
    }
}
}
```

### Application.PointsToInches

将磅转换为英寸（1 英寸 = 72 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToInches ( *Points* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                 |
|----------|-----------|----------|----------------------|
| *Points* | 必选      | Single   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

``` JavaScript
/*本示例将活动文档上边距转换为英寸，并在消息框中显示结果。*/
alert(PointsToInches(Applicatiion.ActiveDocument.Sections.Item(1).PageSetup.TopMargin))
```

``` JavaScript
/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/
function test()
{
let intUnit
let sngData
                                    
function ConvertPoints(intUnit,sngData){
    switch(intUnit) {
    case 1:
        ConvertPoints = PointsToCentimeters(sngData)
    case 2:
        ConvertPoints = PointsToInches(sngData)
    case 3:
        ConvertPoints = PointsToLines(sngData)
    case 4:
        ConvertPoints = PointsToMillimeters(sngData)
    case 5:
        ConvertPoints = PointsToPicas(sngData)
    default:
        alert("无效的过程调用")
    }
}
}
```

### Application.PointsToLines

将磅转换为行（1 行 = 12 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToLines ( *Points* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                 |
|----------|-----------|----------|----------------------|
| *Points* | 必选      | Single   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

``` JavaScript
/*本示例将选定内容中第一段的行间距的值由磅转换为行。*/
alert(PointsToLines(Application.Selection.Paragraphs.Item(1).LineSpacing) + " lines")
```

``` JavaScript
/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/
function test()
{
let intUnit
let sngData
                                    
function ConvertPoints(intUnit,sngData){
    switch(intUnit) {
    case 1:
        ConvertPoints = PointsToCentimeters(sngData)
    case 2:
        ConvertPoints = PointsToInches(sngData)
    case 3:
        ConvertPoints = PointsToLines(sngData)
    case 4:
        ConvertPoints = PointsToMillimeters(sngData)
    case 5:
        ConvertPoints = PointsToPicas(sngData)
    default:
        alert("无效的过程调用")
    }
}
}
```

### Application.PointsToMillimeters

将度量单位由磅转换为毫米（1 毫米 = 2.835 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToMillimeters ( *Points* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                 |
|----------|-----------|----------|----------------------|
| *Points* | 必选      | Single   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

``` JavaScript
/*本示例将 72 磅转换为相应的毫米数。*/
alert(PointsToMillimeters(72) + " millimeters")
```

``` JavaScript
/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/
function test()
{
let intUnit
let sngData
                                    
function ConvertPoints(intUnit,sngData){
    switch(intUnit) {
    case 1:
        ConvertPoints = PointsToCentimeters(sngData)
    case 2:
        ConvertPoints = PointsToInches(sngData)
    case 3:
        ConvertPoints = PointsToLines(sngData)
    case 4:
        ConvertPoints = PointsToMillimeters(sngData)
    case 5:
        ConvertPoints = PointsToPicas(sngData)
    default:
        alert("无效的过程调用")
    }
}
}
```

### Application.PointsToPicas

将度量单位由磅转换为十二点活字（1 十二点活字 = 12 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToPicas ( *Points* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                 |
|----------|-----------|----------|----------------------|
| *Points* | 必选      | Single   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

``` JavaScript
/*本示例将 36 磅转换为相应的十二点活字数。*/
alert(Application.PointsToPicas(36) + " picas")
```

``` JavaScript
/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/
function test()
{
let intUnit
let sngData
                                    
function ConvertPoints(intUnit,sngData){
    switch(intUnit) {
    case 1:
        ConvertPoints = PointsToCentimeters(sngData)
    case 2:
        ConvertPoints = PointsToInches(sngData)
    case 3:
        ConvertPoints = PointsToLines(sngData)
    case 4:
        ConvertPoints = PointsToMillimeters(sngData)
    case 5:
        ConvertPoints = PointsToPicas(sngData)
    default:
        alert("无效的过程调用")
    }
}
}
```

### Application.PointsToPixels

将长度值的单位由磅转换为像素。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToPixels ( *Points* , *fVertical* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                            |
|-------------|-----------|----------|-------------------------------------------------------------------------------------------------|
| *Points*    | 必选      | Single   | 要转换为像素值的磅值。                                                                          |
| *fVertical* | 可选      | Variant  | 如果该属性值为 True，则返回的结果是垂直像素值；如果该属性值为 False，则返回的结果是水平像素值。 |

**返回值**

Single

**示例**

``` JavaScript
/*本示例以像素为单位显示以磅为单位的对象的高度和宽度。*/
alert("180x120 points is equivalent to " + Application.PointsToPixels(180, false) + "x" + Application.PointsToPixels(120, true) + " pixels on this display.")
```

### Application.PrintOut

打印指定文档的全部或部分内容。

**语法**

**express.PrintOut ( *Background* , *Append* , *Range* , *OutputFileNameOutputFileName* , *From* , *To* , *Item* , *Copies* , *Pages* , *PageType* , *PrintToFile* , *Collate* , *FileName* , *ActivePrinterMacGX* , *ManualDuplexPrint* , *PrintZoomColumn* , *PrintZoomRow* , *PrintZoomPaperWidth* , *PrintZoomPaperHeight* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称                           | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                 |
|--------------------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Background*                   | 可选      | Variant  | 如果将该属性设置为 True，则 WPS 在打印文档时继续运行宏。                                                                                                                                                                                                                                             |
| *Append*                       | 可选      | Variant  | 如果将该属性设置为 True，则将指定文档附加到 OutputFileName 参数所指定的文件名。如果将该属性设置为 False，则覆盖 OutputFileName 参数所指定文件的内容。                                                                                                                                                |
| *Range*                        | 可选      | Variant  | 页码范围。可以是任意 WdPrintOutRange 常量。                                                                                                                                                                                                                                                          |
| *OutputFileNameOutputFileName* | 可选      | Variant  | 如果 PrintToFile 为 True，则该参数指定输出文件的路径和文件名。                                                                                                                                                                                                                                       |
| *From*                         | 可选      | Variant  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定起始页码。                                                                                                                                                                                                                                            |
| *To*                           | 可选      | Variant  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定结束页码。                                                                                                                                                                                                                                            |
| *Item*                         | 可选      | Variant  | 要打印的项目。可以是任意 WdPrintOutItem 常量。                                                                                                                                                                                                                                                       |
| *Copies*                       | 可选      | Variant  | 要打印的份数。                                                                                                                                                                                                                                                                                       |
| *Pages*                        | 可选      | Variant  | 要打印的页码和页码范围，中间用逗号分开。例如，“2, 6-10”表示打印第 2 页以及第 6 至第 10 页。                                                                                                                                                                                                          |
| *PageType*                     | 可选      | Variant  | 要打印的页面类型。可以是任意 WdPrintOutPages 常量。                                                                                                                                                                                                                                                  |
| *PrintToFile*                  | 可选      | Variant  | 如果该属性值为 True，则将打印指令发送到文件。请确保使用 OutputFileName 指定文件名。                                                                                                                                                                                                                  |
| *Collate*                      | 可选      | Variant  | 在打印文档的多份副本时，如果该属性值为 True，则完成打印所有页面后再打印下一份副本。                                                                                                                                                                                                                  |
| *FileName*                     | 可选      | Variant  | 要打印的文档的路径和文件名。如果省略该参数， WPS 将打印活动文档（仅应用于 Application 对象）。                                                                                                                                                                                                       |
| *ActivePrinterMacGX*           | 可选      | Variant  | 该参数仅适用于 WPS OfficeMacintosh Edition。有关该参数的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。                                                                                                                                                                            |
| *ManualDuplexPrint*            | 可选      | Variant  | 如果该属性值为 True，则在无双面打印组件的打印机上打印双面文档。如果该属性值为 True，则忽略 PrintBackground 和 PrintReverse 属性。使用 PrintOddPagesInAscendingOrder 和 PrintEvenPagesInAscendingOrder 属性可在手动双面打印时控制输出。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *PrintZoomColumn*              | 可选      | Variant  | 表示 WPS 在一页纸上水平放置的页数。可以是 1、2、3 或 4 页。与 PrintZoomRow 参数一同使用可在单张纸上打印多页文档。                                                                                                                                                                                    |
| *PrintZoomRow*                 | 可选      | Variant  | 表示 WPS 在一页纸上垂直放置的页数。可以是 1、2 或 4 页。与 PrintZoomColumn 参数一同使用可在单张纸上打印多页文档。                                                                                                                                                                                    |
| *PrintZoomPaperWidth*          | 可选      | Variant  | WPS 将打印页面缩放到的宽度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。                                                                                                                                                                                                                             |
| *PrintZoomPaperHeight*         | 可选      | Variant  | WPS 将打印页面缩放到的高度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。                                                                                                                                                                                                                             |

**示例**

``` JavaScript
/*本示例打印活动文档的当前页面。*/
Application.ActiveDocument.PrintOut(null,null,wdPrintCurrentPage)
```

``` JavaScript
/*本示例打印活动窗口中文档的前三页。*/
Application.ActiveDocument.ActiveWindow.PrintOut(null,null,wdPrintFromTo,null,"1","3")
```

``` JavaScript
/*本示例打印活动文档中的备注。*/
function test()
{
if(Application.ActiveDocument.Comments.Count >= 1) {
    Application.ActiveDocument.PrintOut(null,null,null,null,null,null,wdPrintComments)
}
}
```

``` JavaScript
/*本示例将打印活动文档，每张纸上打印六页文档。*/
Application.ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,3,2)
```

``` JavaScript
/*本示例按实际尺寸的 75% 打印活动文档。*/
Application.ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,0.75 * (8.5 * 1440),0.75 * (11 * 1440))
```

### Application.ProductCode

返回一个 **String** 类型的 WPS 全球唯一标识符 (GUID)。

**语法**

**express.ProductCode ()**

\*express是一个代表 Application 对象的变量。

**返回值**

String

**示例**

``` JavaScript
/*本示例将显示 WPS 的全球唯一标识符 (GUID)。*/
alert(Application.ProductCode())
```

### Application.PutFocusInMailHeader

如果活动窗口中的文档是电子邮件文档，则将插入点置于邮件标题的 **“收件人”** 行。

**语法**

**express.PutFocusInMailHeader ()**

\*express是一个代表 Application 对象的变量。

**说明**

为了获得最佳结果，可使用具有 **PutFocusInMailHeader** 属性的 **EnvelopeVisible** 方法。 **EnvelopeVisible** 属性设置为 **True** 时， **PutFocusInMailHeader** 方法会将插入点置于邮件标题中。

**示例**

``` JavaScript
/*以下示例显示活动文档的邮件标题，然后将插入点置于邮件标题的“收件人”行。*/
function test()
{
    Application.ActiveDocument.ActiveWindow.EnvelopeVisible = true
    Application.PutFocusInMailHeader()
}
```

### Application.Quit

退出 WPS ，并可选择保存或传送打开的文档。

**语法**

**express.Quit ( *SaveChanges* , *OriginalFormat* , *RouteDocument* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                   |
|------------------|-----------|----------|----------------------------------------------------------------------------------------|
| *SaveChanges*    | 可选      | Variant  | 指定 WPS 在关闭前是否保存更改过的文档。可以是 WdSaveOptions 常量之一。                 |
| *OriginalFormat* | 可选      | Variant  | 指定 WPS 保存其原格式不是 WPS 文档格式的文档的方式。可以是 WdOriginalFormat 常量之一。 |
| *RouteDocument*  | 可选      | Variant  | 如果为 True ，则将文档传送给下一个收件人。如果文档没有附加传送名单，则忽略该参数。     |

**示例**

``` JavaScript
/* 本示例关闭 WPS，并提示用户保存自上次保存后已更改过的每篇文档。 */
Application.Quit(wdPromptToSaveChanges)
```

### Application.Repeat

一次或多次重复最近的编辑操作。如果成功地重复了命令，则返回 **True** 。

**语法**

**express.Repeat ( *Times* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                         |
|---------|-----------|----------|------------------------------|
| *Times* | 可选      | Variant  | 为需要重复最近的命令的次数。 |

**返回值**

Boolean

**说明**

使用这种方法等同于使用 **“编辑”** 菜单上的 “重复”命令。

**示例**

``` JavaScript
/*本示例插入文本“Hello”及两个段落（重复一次第二个键入操作）。*/
function test() {
    Application.Selection.TypeText("Hello")
    Application.Selection.TypeParagraph()
    Application.Repeat()
}
```

``` JavaScript
/*本示例重复三次最后的命令（如果该命令可以重复的话）。*/
function test() {
    try {
        if (Application.Repeat(3) == true) {
            StatusBar = "Action repeated"
        }
    }
    catch (exception) {
        Application.console.log(Application.exception.message)
    }
}
```

### Application.ResetIgnoreAll

在拼写检查时，取消以前的忽略词汇表。

**语法**

**express.ResetIgnoreAll ()**

\*express是一个代表 Application 对象的变量。

**说明**

运行该方法后，以前被忽略的词现在同其他所有词一同受到检查。为使 **ResetIgnoreAll** 方法工作，必须首先将 **SpellingChecked** 属性设置为 **False** 。

**示例**

``` JavaScript
/*本示例取消在上一次拼写检查中忽略的词汇列表，然后对活动文档重新进行拼写检查。*/
function test() {
    Application.ResetIgnoreAll()
    Application.ActiveDocument.SpellingChecked = false
    Application.ActiveDocument.CheckSpelling()
}
```

### Application.Resize

调整 WPS 应用程序或某一任务的窗口大小。

**语法**

**express.Resize ( *Width* , *Height* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                            |
|----------|-----------|----------|---------------------------------|
| *Width*  | 必选      | Long     | 窗口的宽度，以磅值为单位。      |
| *Height* | 必选      | Long     | Long 窗口的高度，以磅值为单位。 |

**说明**

如果该窗口被最大化或最小化，将出现错误。使用 **Width** 或 **Height** 属性为窗口单独设置宽度和高度。

**示例**

``` JavaScript
/*本示例将 ET 应用程序窗口调整为 6 英寸宽、4 英寸高。*/
function test() {
    if (Application.Tasks.Exists("ET") == true) {
        Application.Tasks("ET").WindowState = wdWindowStateNormal
        Application.Tasks("ET").Resize(Application.InchesToPoints(6), Application.InchesToPoints(4))
    }
}
```

``` JavaScript
/*本示例将 WPS 应用程序窗口调整为 7 英寸宽、6 英寸高。*/
function test() {
    Application.WindowState = wdWindowStateNormal
    Application.Resize(InchesToPoints(7), InchesToPoints(6))
}
```

### Application.Run

运行JS函数

**语法**

**express.Run ( *MacroName* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                     |
|-------------|-----------|----------|--------------------------|
| *MacroName* | 必选      | String   | 函数的名称，可以带参数。 |

**说明**

*MacroName* 参数可以是任意模板、模块和宏名的组合。

**示例**

``` JavaScript
/*本示例运行test函数。*/
Application.Run(test())
```

### Application.ScreenRefresh

使用视频内存缓冲区的当前信息更新显示器的显示。

**语法**

**express.ScreenRefresh ()**

\*express是一个代表 Application 对象的变量。

**说明**

可以在使用 **ScreenUpdating** 属性禁用屏幕更新后使用该方法。 **ScreenRefresh** 只对某一指令打开屏幕更新，然后立即关闭。直到再次使用 **ScreenUpdating** 属性打开屏幕更新后，后续指令才更新屏幕。

**示例**

``` JavaScript
/*本示例关闭屏幕更新，打开文档 Test.doc，插入文本，刷新屏幕，然后关闭该文档（保存所做的更改）。*/
function test() {
    let rngTemp
    Application.ScreenUpdating = false
    Application.Documents.Open("C:\\DOCS\\TEST.DOC")

    rngTemp = Application.ActiveDocument.Range(0, 0)

    rngTemp.InsertBefore("new")
    Application.ScreenRefresh()
    Application.ActiveDocument.Close(wdSaveChanges)
    Application.ScreenUpdating = true
}
```

### Application.SetDefaultTheme

为 WPS 设置用于新文档、电子邮件或网页的默认 主题 。

**语法**

**express.SetDefaultTheme ( *Name* , *DocumentType* )**

\*express是一个代表 Application 对象的变量。

**参数**

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th>名称</th>
<th>必选/可选</th>
<th>数据类型</th>
<th>说明</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td class="mainsection"><em>Name</em></td>
<td class="mainsection">必选</td>
<td class="mainsection">String</td>
<td class="mainsection" style="word-break: break-all">由要指定为默认主题的主题名称加上任何要应用的主题格式选项组成。该字符串格式为“themennn”，其中 theme 和 nnn 的定义如下：
主题 :包含所需主题的数据的文件夹名称（主题数据文件夹的默认位置是 C:\Program Files\Kingsoft\WPS Office Professional\office6\document theme）。主题必须使用该文件夹名，不能使用“主题”对话框（“格式”菜单中的“主题”命令）中显示的名称。
nnn :三个数字的字符串，表示要激活的主题格式选项（1 为激活，0 为停用）。每个数字分别对应“主题”对话框（“格式”菜单中的“主题”命令）中的“鲜艳颜色”、“动态图形”和“背景图像”复选框。如果省略该字符串，则 nnn 的默认值为“011”（“动态图形”和“背景图像”均被激活）。</td>
</tr>
<tr class="even">
<td class="mainsection"><em>DocumentType</em></td>
<td class="mainsection">必选</td>
<td class="mainsection">WdDocumentMedium</td>
<td class="mainsection" style="word-break: break-all">为其指定默认主题的新文档的类型。</td>
</tr>
</tbody>
</table>

**说明**

如果设置了一个默认主题，则此主题将不用于启动 WPS 时自动创建的空白文档。此后所创建的任何文档都将具有该默认主题。

也可使用 **ThemeName** 属性为新电子邮件返回并设置默认主题。

**示例**

``` JavaScript
/*本示例指定 WPS 在所有的新电子邮件中使用“Blueprint”主题。*/
Application.SetDefaultTheme("blueprnt", wdEmailMessage)
```

``` JavaScript
/*本示例指定 WPS 在所有新网页中的“动态图形”中使用“Expedition”主题。*/
Application.SetDefaultTheme("expeditn 010", wdWebPage)
```

### Application.ShowClipboard

显示 **“剪贴板”** 任务窗格。

**语法**

**express.ShowClipboard ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*以下示例显示“剪贴板”任务窗格。*/
Application.ShowClipboard()
```

### Application.ShowMe

当存在更多可用信息时，显示“WPS 助手”或“帮助”窗口。

**语法**

**express.ShowMe ()**

\*express是一个代表 Application 对象的变量。

**说明**

如果没有进一步的有效信息，该方法将显示一则信息以提示不存在相关“帮助”主题。

**示例**

``` JavaScript
/*如果可能的话，本示例将完成一个“操作向导详细说明”操作。*/
Application.ShowMe()
```

### Application.SubstituteFont

设置字体映射选项。

**语法**

**express.SubstituteFont ( *UnavailableFont* , *SubstituteFont* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型 | 说明                                                                         |
|-------------------|-----------|----------|------------------------------------------------------------------------------|
| *UnavailableFont* | 必选      | String   | 在本计算机上无效的字体名称（因而您希望映射为其他字体以进行显示和打印输出）。 |
| *SubstituteFont*  | 必选      | String   | 用于替换无效字体并在本计算机上有效的字体的名称。                             |

**说明**

可以在 **“字体替换”** 对话框中查找字体映射选项。

**示例**

``` JavaScript
/*本示例将 Courier 字体替换为 CustomFont1 字体。*/
Application.SubstituteFont("CustomFont1","Courier")
```

### Application.SynonymInfo

返回一个 **SynonymInfo** 对象，该对象包含来自同义词库的、与指定字词或短语的同义词、反义词或相关词语及表达相关的信息。

**语法**

**express.SynonymInfo ( *Word* , *LanguageID* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                                 |
|--------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------|
| *Word*       | 必选      | String   | 指定的字词或短语。                                                                                                                   |
| *LanguageID* | 可选      | Variant  | 用于同义词库的语言。可以是 WdLanguageID 常量之一。（但是根据您选择或安装的语言支持不同（例如美国英语），部分列出的常量可能无法使用） |

**示例**

``` JavaScript
/*本示例返回美国英语中单词“big”的反义词列表。*/
function test() {
    let Alist = Application.SynonymInfo("big", wdEnglishUS).AntonymList
    for (let i = 1; i <= Alist.length - 1; i++) {
        alert(Alist[i])
    }
}
```

### Application.ToggleKeyboard

在从左向右的语言和从右向左语言之间切换键盘语言设置。

**语法**

**express.ToggleKeyboard ()**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将键盘语言设置由从右向左的语言切换为从左向右的语言。*/
function test() {
    let x = alert("Switch the keyboard language setting")
     Application.ToggleKeyboard()
}
```

## 成员属性

### Application.ActiveDocument

返回一个 Document 对象，该对象代表活动文档。如果没有打开的文档，就会导致出错。只读。

**语法**

**express.ActiveDocument**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例打印当前激活的文档名*/
// 当前有打开文档
function test() {
    if(Application.Documents.Count >= 1) { 
        alert(ActiveDocument.Name)
    }else {
        alert("No documents are open")
    }
}

/*本示例将选定内容折叠为插入点，然后创建一个区域，包括选定内容的后面五个字符*/
function test() {
    Application.Selection.Collapse(Application.Enum.wdCollapseStart)
    let rng = Application.ActiveDocument.Range(Application.Selection.Start, Application.Selection.Start + 5)
}
```

### Application.ActiveEncryptionSession

返回一个 **Long** 类型的值，该值代表与活动文档相关的加密会话。只读。

**语法**

**express.ActiveEncryptionSession**

\*express是一个代表 Application 对象的变量。

**说明**

加密提供程序机制在单独的进程中管理每个文件，因此每个文件均与单独的加密会话相关。

| 注释                              |
|-----------------------------------|
| 此属性仅用于文档执行自定义加密时. |

### Application.ActivePrinter

返回或设置活动打印机名称。可读/写 **String** 类型。

**语法**

**express.ActivePrinter**

\*express是一个代表 Application 对象的变量。

**说明**

使用 **ActivePrinter** 属性设置打印机会更改默认打印机。有关详细信息，请参阅 设置 ActivePrinter 会更改系统默认打印机。

**示例**

``` JavaScript
/*本示例显示当前打印机的名称。*/
alert( "The name of the active printer is " + ActivePrinter)
```

``` JavaScript
/*本示例将网络 HP LaserJet IIISi 打印机设置为活动打印机。*/
Application.ActivePrinter = "HP LaserJet IIISi on \\printers\laser"
```

``` JavaScript
/*本示例将 LPT1 端口的本地 HP LaserJet 4 打印机设置为活动打印机。*/
Application.ActivePrinter = "HP LaserJet 4 local on LPT1:"
```

### Application.ActiveProtectedViewWindow

返回代表活动的受保护的视图窗口的 ProtectedViewWindow 对象。只读。

**语法**

**express.ActiveProtectedViewWindow**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示活动的受保护的视图窗口的标题文本。*/
alert(Application.ActiveProtectedViewWindow.Caption)
```

### Application.ActiveWindow

返回 **Window** 对象，该对象代表活动窗口（焦点所在的窗口）。如果没有打开的窗口，则会导致出错。只读。

**语法**

**express.ActiveWindow**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
//本示例显示活动窗口的标题文本。

Application.ActiveWindow.Caption
```

### Application.AddIns

返回一个 **AddIns** 集合，该集合代表所有有效加载项，而不考虑当前是否已加载它们。只读。

**语法**

**express.AddIns**

\*express是一个代表 Application 对象的变量。

**说明**

此 **AddIns** 集合包含全局模板和 **“工具”** 菜单的 **“模板和加载项”** 对话框中所列的 WPS 加载项库（即 WLL）。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*此示例返回加载项的总数。*/
let intAddIns = Application.AddIns.Count
```

``` JavaScript
/*此示例显示 Addins 集合中每一加载项的名称。*/
function test() {
    for (i = 1; i <= Application.AddIns.Count; i++) {
        alert(Application.AddIns.Item(i).Name)
    }
}
```

### Application.AnswerWizard

返回一个 **AnswerWizard** 对象，该对象包含联机“帮助”搜索引擎所用的文件。

**语法**

**express.AnswerWizard**

\*express是一个代表 Application 对象的变量。

**说明**

“WPS Office 助手”和“应答向导”在 WPS Office 2015 版中已弃用。

**示例**

``` JavaScript
/*本示例重新设置“操作向导”文件列表。*/
function test()
{
    function AnswerWizardReset()
        Application.AnswerWizard.ResetFileList()
    }
}
```

### Application.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例使 WPS 显示滚动条、屏幕提示和状态栏。*/
function test() {
    let app = Application
    app.DisplayScrollBars = true
    app.DisplayScreenTips = true
    app.DisplayStatusBar = true
}
```

### Application.ArbitraryXMLSupportAvailable

返回一个 **Boolean** 类型的值，代表 WPS 是否接受自定义的 XML 架构。值为 **True** 表明 WPS 接受自定义的 XML 架构。

**语法**

**express.ArbitraryXMLSupportAvailable**

\*express是一个代表 Application 对象的变量。

**说明**

WPS Office 2003 使用 WPS XML 架构提供 XML 支持，但不提供对自定义 XML 架构的支持。仅在 WPS Office 2003 或更高版本提供对自定义 XML 架构的支持。使用 **ArbitraryXMLSupportAvailable** 属性可确定安装了哪种版本。

**示例**

``` JavaScript
/*如果安装的 WPS 版本不支持自定义的 XML 架构，则以下代码将显示一条消息。*/
function test()
{
    if(Application.ArbitraryXMLSupportAvailable == false) {
        alert("Custom XML schemas are not supported in this version of WPS.")
    }
}
```

### Application.Assistance

返回一个代表 WPS Office帮助查看器的 **Assistance** 对象。只读。

**语法**

**express.Assistance**

\*express是一个代表 Application 对象的变量。

**说明**

**Assistance** 对象使开发人员可以在 Office 帮助查看器中显示自定义帮助以及随 Office 一起安装的帮助。

### Application.Assistant

返回一个 **Assistant** 对象，该对象代表“WPS Office 助手”。

**语法**

**express.Assistant**

\*express是一个代表 Application 对象的变量。

**说明**

WPS Office 助手和 AnswerWizard 在 2013 版的 WPS Office 中的作用已经弱化。

**示例**

``` JavaScript
/*本示例显示 WPS 助手。*/
Application.Assistant.Visible = true
```

``` JavaScript
/*本示例显示 WPS 助手，并将它移至屏幕的左上角。*/
function test()
{
assistant.Visible = true
assistant.Move(100,100)
}
```

``` JavaScript
/*本示例显示 WPS 助手和在气球中的自定义消息。*/
function test()
{
Assistant.Visible = true
let bln = Assistant.NewBalloon
bln.Mode = msoModeAutoDown
bln.Text = "Hello"
bln.Button = msoButtonSetNone
bln.Show()
}
```

### Application.AutoCaptions

返回一个 **AutoCaptions** 集合。当诸如表格或图片之类的项目插入到文档中时，此集合代表自动添加的题注。只读。

**语法**

**express.AutoCaptions**

\*express是一个代表 Application 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示在插入到文档时自动添加题注的各个项目的名称。*/
function test() {
    for (let i = 1; i <= Application.AutoCaptions.Count; i++) {
        if (Application.AutoCaptions.Item(i).AutoInsert == true) {
            alert(Application.AutoCaptions.Item(i).name)
        }
    }
}
```

### Application.AutoCorrect

返回一个 **AutoCorrect** 对象，该对象包含当前“自动更正”的选项、词条和例外项。只读。

**语法**

**express.AutoCorrect**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例添加一个“自动更正”替换词条。运行此代码后，在文档中键入的所有“sr”都会自动替换为“Stella Richards”。*/
Application.AutoCorrect.Entries.Add("sr","Stella Richards")
```

``` JavaScript
/*本示例将删除指定的原有“自动更正”词条。*/
function test() {
    let strInput
    let aceLoop
    let blnMatch
    let intConfirm

    blnMatch = false

    strInput = alert("Enter the AutoCorrect entry to delete.")

    for (let i = 1; i <= Application.AutoCorrect.Entries.Count; i++) {
        aceLoop = Application.AutoCorrect.Entries.Item(i)
        if (aceLoop.Name == strInput) {
            blnMatch == true
            intConfirm = alert("Are you sure you want to delete " + aceLoop.Name, jsYesNo)
            if (intConfirm == jsResultYes) {
                aceLoop.Delete()
            }
        }
    }

    if (blnMatch != true) {
        alert("There was no AutoCorrect entry: " + strInput)
    }
}
```

### Application.AutoCorrectEmail

返回一个 **AutoCorrect** 对象，该对象代表对电子邮件消息进行的自动更正。

**语法**

**express.AutoCorrectEmail**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例添加电子邮件消息的“自动更正”词条。运行该代码后，在电子邮件消息中键入的所有“allways”、“hte”和“hwen”将被分别替换为“always”、“the”和“when”。*/
function AutoCorrectEMailAddress(){
    var Email = Application.AutoCorrectEmail
    Email.Entries.Add("allways","always")
    Email.Entries.Add("hte","the")
    Email.Entries.Add("hwen","when")
}
```

### Application.AutomationSecurity

返回或设置一个 **MsoAutomationSecurity** 常量，该常量代表当用编程方法打开文件时 WPS 所使用的安全设置。

**语法**

**express.AutomationSecurity**

\*express是一个代表 Application 对象的变量。

**说明**

**AutomationSecurity** 属性的默认设置为 **msoAutomationSecurityLow** 。因此，若要避免更改用户的安全设置或中断依赖于默认设置的解决方案，在程序打开文件后，应谨慎地将此属性设置回其初始设置。

将 **ScreenUpdating** 设置为 **False** 不会影响警告提醒和安全警告。 **DisplayAlerts** 设置不会应用于安全警告。例如，如果用户将 **DisplayAlerts** 设置为等于 **False** ，将 **AutomationSecurity** 设置为 **msoAutomationSecurityByUI** ，同时用户处于“中等”安全级别，则在运行宏时会显示安全警告。这使宏可以捕获文件打开错误，而即使文件成功打开，仍然会显示安全警告。

**示例**

``` JavaScript
/*本示例更改设置以禁用宏，显示“打开”对话框，然后将 AutomationSecurity 属性设置回其初始设置。*/
function Security(){
    let lngAutomation = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    let dialog = Application.FileDialog(msoFileDialogOpen)
    dialog.Show()
    dialog.Execute()
    Application.AutomationSecurity = lngAutomation
}
```

### Application.BackgroundPrintingStatus

返回后台打印队列中的打印任务数目。 **Long** 类型，只读。

**语法**

**express.BackgroundPrintingStatus**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例返回当前排队等候后台打印的 WPS 打印任务数目。*/
function test() {
    let lngStatus
    if (Application.Options.PrintBackground == true) {
        lngStatus = Application.BackgroundPrintingStatus
    }
}
```

``` JavaScript
/*如果打印任务数大于 0（零），本示例将在状态栏中显示消息。*/
function test() {
    if (Application.BackgroundPrintingStatus > 0) {
        Application.StatusBar = Application.BackgroundPrintingStatus + " print jobs are queued up"
    }
}
```

### Application.BackgroundSavingStatus

返回后台保存队列中的文件数。 **Long** 类型，只读。

**语法**

**express.BackgroundSavingStatus**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例在状态栏中显示当前正在保存的文档数。*/
function test() {
    Application.Options.BackgroundSave = true
    Application.Documents.Add()
    Application.ActiveDocument.SaveAs()
    while (Application.BackgroundSavingStatus != 0) {
        Application.StatusBar = "Documents remaining to save: " + Application.BackgroundSavingStatus
    }
}
```

### Application.Bibliography

返回一个 **Bibliography** 对象，该对象代表 WPS 中存储的书目参考源。只读。

**语法**

**express.Bibliography**

\*express是一个代表 Application 对象的变量。

### Application.BrowseExtraFileTypes

如果将该属性设置为“text/html”，则用 WPS （而不是默认的 Internet 浏览器）可以打开超链接的 HTML 文件。 **String** 类型，可读写。

**语法**

**express.BrowseExtraFileTypes**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例使 WPS（而不是默认的 Internet 浏览器）打开超链接的 HTML 文件。*/
function test()
{
Application.BrowseExtraFileTypes = "text/html"
}
```

### Application.Browser

返回一个 **Browser** 对象，该对象代表垂直滚动条上的 **“选择浏览对象”** 工具。只读。

**语法**

**express.Browser**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例跳转到活动文档的下一个脚注引用标记。*/
function test() {
    let Borwser = Application.Browser
    Borwser.Target = wdBrowseFootnote
    Borwser.Next()
}
```

``` JavaScript
/*本示例跳转到活动文档的下一个域。将从所选内容的开始处到下一个域之间的文本设置为加粗格式。*/
function test() {
    Application.Selection.ExtendMode = true
    let ExtendMode = Application.Browser
    ExtendMode.Target = wdBrowseField
    ExtendMode.Next()
    Application.Selection.Font.Bold = true
    Application.Selection.ExtendMode = false
    Application.Selection.Collapse(wdCollapseEnd)
}
```

### Application.Build

返回 WPS 应用程序的版本号及编译序号。 **String** 类型，只读。

**语法**

**express.Build**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例用于显示 WPS 的版本号和编译序号。*/
alert(Application.Build+"WPS  Version")
```

### Application.CapsLock

如果已打开 Caps Lock 键，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.CapsLock**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例检索 Caps Lock 键的当前状态。*/
function test()
{
    let blnCapsLock
    blnCapsLock = Application.CapsLock
}
```

### Application.Caption

返回或设置出现在应用程序窗口的标题栏中的文本。 **String** 类型，可读写。

**语法**

**express.Caption**

\*express是一个代表 Application 对象的变量。

**说明**

要将应用程序窗口的题注改为默认文本，请将该属性设置为空字符串（""）。

**示例**

``` JavaScript
/*本示例重新设置应用程序窗口的题注。*/
Application.Caption = ""
```

``` JavaScript
/*本示例改变 WPS 应用程序窗口的题注，使其包括用户名。*/
Application.Caption = Application.UserName + "'s copy of Word"
```

### Application.CaptionLabels

返回一个 **CaptionLabels** 集合，该集合包括全部有效题注标签。只读。

**语法**

**express.CaptionLabels**

\*express是一个代表 Application 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例为表格题注设置编号样式。*/
Application.CaptionLabels(wdCaptionTable).NumberStyle = Application.Enum.wdCaptionNumberStyleLowercaseRoman
```

``` JavaScript
/*本示例先添加一个名为“Photo”的新题注标签，然后插入一个照片题注。*/
function test() {
    Application.CaptionLabels.Add("Photo")
    Application.Selection.InsertParagraphAfter()
    Application.Selection.InsertCaption("Photo")
}
```

### Application.CheckLanguage

如果 WPS 在键入时自动检测使用的语言，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CheckLanguage**

\*express是一个代表 Application 对象的变量。

**说明**

如果未将 WPS 设为可用于多语言编辑，则 **CheckLanguage** 属性总返回 **False** 。

**示例**

``` JavaScript
/*本示例将检查是否已激活自动语言检测功能。*/
function test() {
    If(Application.CheckLanguage == true) {
        alert("Automatic language detection is activated!")
    }
}
```

### Application.COMAddIns

返回对 **COMAddIns** 集合的引用，该集合代表 WPS 中当前加载的所有 组件对象模型 (COM) 加载项?（COM 加载项：通过添加自定义命令和指定的功能来扩展 WPS Office程序的功能的补充程序。COM 加载项可在一个或多个 Office 程序中运行。COM 加载项使用文件扩展名 .dll 或 .exe。） 。

**语法**

**express.COMAddIns**

\*express是一个代表 Application 对象的变量。

**说明**

COM 加载项在 **“COM 加载项”** 对话框中列出。可以使用 **“自定义”** 对话框（ **“工具”** 菜单），向 **“工具”** 菜单添加 **“COM 加载项”** 命令。

### Application.CommandBars

返回一个 **CommandBars** 集合，该集合代表 WPS 中的菜单栏以及所有工具栏。

**语法**

**express.CommandBars**

\*express是一个代表 Application 对象的变量。

**说明**

在访问 **CommandBars** 集合之前，可用 **CustomizationContext** 属性设置模板或文档上下文。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例增大所有命令栏按钮并激活工具提示。*/
function test() {
    let Com = Application.CommandBars
    Com.LargeButtons = true
    Com.DisplayTooltips = true
}
```

``` JavaScript
/*本示例在应用程序窗口底部显示“绘图”工具栏。*/
function test() {
    let Com = Application.CommandBars.Item("Drawing")
    Com.Visible = true
    Com.Position = msoBarBottom
}
```

``` JavaScript
/*本示例将“版本”命令按钮添至“常用”工具栏。*/
function test() {
    CustomizationContext = NormalTemplate
    Application.CommandBars.Item("Standard").Controls.Add(msoControlButton, 2522, null, 4)
}
```

### Application.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Application 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如， WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>值也可用常量 <strong>wdCreatorCode</strong> 表示。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

### Application.CustomDictionaries

返回一个 **Dictionaries** 对象，该对象代表活动自定义词典的集合。只读。

**语法**

**express.CustomDictionaries**

\*express是一个代表 Application 对象的变量。

**说明**

在 **“自定义词典”** 对话框中选中复选框的词典即为活动自定义词典。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例在该集合中添加一个新的空白自定义词典。然后在消息框中显示该自定义词典的路径和文件名。*/
function test() {
    let dicHome = Application.CustomDictionaries.Add("Home.dic")
    alert(Application.CustomDictionaries.Item(1).Path + Application.PathSeparator + Application.CustomDictionaries.Item(1).Name)
}
```

``` JavaScript
/*本示例清除所有的自定义词典以便没有激活状态的自定义词典。但并非删除自定义词典文件。*/
Application.CustomDictionaries.ClearAll()
```

``` JavaScript
/*本示例显示该集合中所有自定义词典的名称。*/
function test() {
    for (let i = 1; i <= Application.CustomDictionaries.Count; i++) {
        alert(Application.CustomDictionaries.Item(i).Name)
    }
}
```

### Application.CustomizationContext

返回或设置一个 **Template** 或 **Document** 对象，该对象代表存储菜单栏、工具栏和键绑定的更改的模板或文档。可读/写。

**语法**

**express.CustomizationContext**

\*express是一个代表 Application 对象的变量。

**说明**

对应于 **“工具”** 菜单 **“自定义”** 对话框 **“命令”** 选项卡上 **“保存于”** 框中的值。

**示例**

``` JavaScript
/*本示例将 Alt+Ctrl+W 添至 FileClose 命令，然后将该键盘自定义方案保存于“Normal”模板中。*/
function test() {
    CustomizationContext = NormalTemplate
    KeyBindings.Add(wdKeyCategoryCommand, "FileClose", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyW))
}
```

``` JavaScript
/*本示例向“常用”工具栏中添加“文件版本”按钮。该命令栏自定义方案保存于活动文档所用的模板中。*/
function test() {
    CustomizationContext = Application.ActiveDocument.AttachedTemplate
    Application.CommandBars("Standard").Controls.Add(msoControlButton, 2522, null, 8)
}
```

### Application.DefaultLegalBlackline

如果为 **True** ，则 WPS 使用 **“比较并合并文档”** 对话框中的 **“精确比较”** 选项比较和合并文档。 **Boolean** 类型，可读写。

**语法**

**express.DefaultLegalBlackline**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例启用 WPS 的“精确比较”选项比较和合并法律文档。*/
function CreateLegalBlackline() {
    Application.DefaultLegalBlackline = true
}
```

### Application.DefaultSaveFormat

返回或设置将在 **“另存为”** 对话框上的 **“保存类型”** 框中显示的默认格式。 **String** 类型，可读写。

**语法**

**express.DefaultSaveFormat**

\*express是一个代表 Application 对象的变量。

**说明**

本属性所用字符串是文件转换程序类名。 WPS 内部格式的类名列于下表。

| wps格式          | 文件转化程序类名 |
|------------------|------------------|
| wps文档          | ""               |
| 文档模板         | "Dot"            |
| 纯文本           | "Text"           |
| 带换行符的纯文本 | "CRText"         |
| MS-DOS 文本      | "9Text"          |
| RTF格式          | "Rtf"            |
| Unicode文本      | "Unicode"        |

使用 **FileConverter** 对象的 **ClassName** 属性可以确定外部文件转换器的类名。

**示例**

``` JavaScript
/*本示例将默认的保存格式设置为 WPS 文档格式。*/
Application.DefaultSaveFormat = ""
```

``` JavaScript
/*此示例返回 WPS 用来保存文档的当前设置。*/
alert(Application.DefaultSaveFormat)
```

### Application.DefaultTableSeparator

返回或设置一个字符；在将文本转换为表格时，该字符用来将文本分隔为单元格。 **String** 类型，可读写。

**语法**

**express.DefaultTableSeparator**

\*express是一个代表 Application 对象的变量。

**说明**

如果省略 **ConvertToTable** 方法、 **Range** 或 **Selection** 对象中的 *Separator* 参数，则使用 **DefaultTableSeparator** 属性的值。

**示例**

``` JavaScript
/*本示例更改默认的表格分隔字符。*/
Application.DefaultTableSeparator = "%"
```

### Application.Dialogs

返回一个 ****Dialogs**** 集合，该集合代表 WPS 中所有内置对话框。只读。

**语法**

**express.Dialogs**

\*express是一个代表 Application 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 显示“打开文件”对话框 */
Application.Dialogs.Item(Application.Enumerate.wdDialogFileOpen).Show()
```

### Application.DisplayAlerts

返回或设置运行宏时产生的一些警告和消息的处理方式。 **WdAlertLevel** 类型，可读写。

**语法**

**express.DisplayAlerts**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在运行宏时显示所有警告和消息框。*/
Application.DisplayAlerts = wdAlertsAll
```

``` JavaScript
/*本示例返回 DisplayAlerts 属性的当前设置。*/
let lngTemp = Application.DisplayAlerts
```

### Application.DisplayAutoCompleteTips

如果 WPS 显示有关所键入整个单词、日期或词组的建议的提示，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayAutoCompleteTips**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在您键入文字时，显示有关所键入的整个单词、日期或词组的建议的提示。*/
Application.DisplayAutoCompleteTips = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动图文集”选项卡的“显示有关‘自动图文集’和日期的‘记忆式键入’提示”选项的状态。*/
let blnTemp = Application.DisplayAutoCompleteTips
```

### Application.DisplayDocumentInformationPanel

返回或设置 **Boolean** 值，该值表示是否显示文档属性面板。可读/写。

**语法**

**express.DisplayDocumentInformationPanel**

\*express是一个代表 Application 对象的变量。

### Application.DisplayRecentFiles

如果在 **“文件”** 菜单中显示最近使用的文件名，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayRecentFiles**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例使 WPS 在“文件”菜单上显示至多 6 个文件名。*/
function test() {
    Application.DisplayRecentFiles = true
    Application.RecentFiles.Maximum = 6
}
```

``` JavaScript
/*本示例用于从“文件”菜单中删除最近使用的文件列表。*/
Application.DisplayRecentFiles = false
```

### Application.DisplayScreenTips

如果为 **True** ，则批注、脚注、尾注和超链接以提示形式显示。突出显示那些标记有批注的文本。 **Boolean** 类型，可读写。

**语法**

**express.DisplayScreenTips**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例使 WPS 将批注、脚注和尾注显示为提示。而且突出显示那些标记有批注的文本。*/
Application.DisplayScreenTips = true
```

``` JavaScript
/*本示例返回“选项”对话框中“视图”选项卡上的“显示”区域中“屏幕提示”复选框当前的状态。*/
let temp = Application.DisplayScreenTips
```

### Application.DisplayScrollBars

如果为 **True** ，则 WPS 在至少一个文档窗口中显示滚动条；如果为 **False** ，则在任何窗口中均不显示滚动条。 **Boolean** 类型，可读写。

**语法**

**express.DisplayScrollBars**

\*express是一个代表 Application 对象的变量。

**说明**

将 **DisplayScrollBars** 属性设置为 **True** ，可在所有窗口中显示垂直滚动条和水平滚动条。将该属性设置为 **False** 将关闭所有窗口内的滚动条。

使用 **DisplayHorizontalScrollBar** 和 **DisplayVerticalScrollBar** 属性在指定的窗口中显示单独的滚动条。

**示例**

``` JavaScript
/*本示例在所有窗口中显示垂直滚动条和水平滚动条。*/
Application.DisplayScrollBars = true
```

``` JavaScript
/*如果当前在任何窗口中显示有滚动条，则本示例返回 True。*/
let blnTemp = Application.DisplayScrollBars
```

### Application.Documents

返回一个 ****Documents**** 集合，该集合代表所有打开的文档。只读。

**语法**

**express.Documents**

\*express是一个代表 Application 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 新增一篇文档，并显示“另存为”对话框 */
Application.Documents.Add().Save()
```

### Application.DontResetInsertionPointProperties

返回或设置一个 **Boolean** 类型的值，该值代表 WPS 在运行其他代码后是否保持位于插入点位置的文本的格式属性。可读写。

**语法**

**express.DontResetInsertionPointProperties**

\*express是一个代表 Application 对象的变量。

**说明**

有时，在运行其他 示例代码 (VBA) 代码后， WPS 会丢失插入点的格式。如果发生这样的情况，则会给依赖屏幕读取器应用程序的用户造成困难。当用户的辅助应用程序执行看起来不相关的任务时，他们会丢失格式。当运行包含 WPS 对象模型中的属性或方法的其他代码时，此属性可防止 WPS 丢失或更改位于插入点位置的文本的格式。

**要点** ??除非明确需要使用此属性来纠正解决方案功能，否则不要使用此属性。

### Application.EmailOptions

返回一个 **EmailOptions** 对象，该对象代表创作电子邮件的全局首选项。只读。

**语法**

**express.EmailOptions**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 标记电子邮件中的批注。*/
Application.EmailOptions.MarkComments = true
```

### Application.EmailTemplate

返回或设置一个 **String** 类型的值，该值代表发送电子邮件时使用的文档模板。可读写。

**语法**

**express.EmailTemplate**

\*express是一个代表 Application 对象的变量。

**说明**

如果在 Microsoft Outlook 中指定 WPS 为电子邮件编辑器，可使用 **EmailTemplate** 属性。

**示例**

``` JavaScript
/*本示例设置 WPS 对所有的新电子邮件使用名为“Email”的模板。本示例假定有一个名为“Email”的模板保存于默认的模板位置。*/
function MessageTemplate() {
    Application.EmailTemplate = "Email"
}
```

### Application.EnableCancelKey

返回或设置 WPS 处理 Ctrl+Break 用户中断的方式。 **WdEnableCancelKey** 类型，可读写。

**语法**

**express.EnableCancelKey**

\*express是一个代表 Application 对象的变量。

**说明**

使用此属性时应非常小心。如果使用 **wdCancelDisabled** ，则无法中断失控的循环或其他非自我中断的代码。另外，当代码停止运行时，除非明确地重新设置其值，否则 **EnableCancelKey** 属性不会重新设置为 **wdCancelInterrupt** ，该属性在 WPS 会话期间内仍为 **wdCancelDisabled** 。

**示例**

``` JavaScript
/*本示例使 Ctrl+Break 不能中断计数器循环。*/
function test() {
    Application.EnableCancelKey = wdCancelDisabled
    for (let intWait = 1; intWait <= 10000; intWait++) {
        StatusBar = intWait
    }
    Application.EnableCancelKey = wdCancelInterrupt
}
```

### Application.FeatureInstall

返回或设置 WPS 将如何处理对所需功能尚未安装的方法和属性的调用。 **MsoReatureInstall** 类型，可读写。

**语法**

**express.FeatureInstall**

\*express是一个代表 Application 对象的变量。

**说明**

当某个功能正在安装时，可以使用 **msoFeatureInstallOnDemandWithUI** 常量来防止用户认为应用程序没有响应。如果希望只有开发者才能安装新功能，则使用 **msoFeatureInstallNone** 常量。

如果将 **DisplayAlerts** 属性设为 **False** ，则即使将 **FeatureInstall** 属性设为 **msoFeatureInstallOnDemand** ，也不会提示用户安装新功能。如果将 **DisplayAlerts** 属性设置为 **True** ，同时将 **FeatureInstall** 属性设置为 **msoFeatureInstallOnDemand** ，则会显示安装进度指示器。

### Application.FileConverters

返回一个 **FileConverters** 集合，该集合代表可用于 WPS 的所有文件转换器。只读。

**语法**

**express.FileConverters**

\*express是一个代表 Application 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示一条消息，以表明 FileConverters 集合的第三个转换器能否用来保存文件。*/
function test() {
    if (Application.FileConverters.Item(3).CanSave == true) {
        alert(Application.FileConverters.Item(3).FormatName + " can save files")
    }
    else {
        alert(Application.FileConverters.Item(3).FormatName + " cannot save files")
    }
}
```

``` JavaScript
/*本示例显示最后一个文件转换器的名称。*/
function test() {
    let fcTemp = Application.FileConverters.Item(Application.FileConverters.Count)
    alert("The file name extensions for " + fcTemp.FormatName + " files are: " + fcTemp.Extensions)
}
```

### Application.FileValidation

返回或设置 WPS 在打开文件之前验证这些文件的方式。可读/写 MsoFileValidationMode。

**语法**

**express.FileValidation**

\*express是一个代表 Application 对象的变量。

**说明**

未通过验证的文件将在 受保护的视图窗口 中打开。 **FileValidation** 属性仅针对每个会话。如果设置了 **FileValidation** 属性，该设置将在应用程序处于打开状态的整个会话中保持有效。

### Application.FindKey

返回 **KeyBinding** 对象，该对象代表指定的组合键。只读。

**语法**

**express.FindKey**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 显示F1指向的命令 */
alert(Application.FindKey(Application.Enum.wdKeyF1).Command)
```

### Application.FocusInMailHeader

如果插入点位于电子邮件标题域中（如“收件人”域），则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.FocusInMailHeader**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*如果插入点位于电子邮件标题域中，则本示例在状态栏显示一条消息。*/
function test() {
    If(Application.FocusInMailHeader == true){
        StatusBar = "Selection is in message header"
    }
}
```

### Application.FontNames

返回一个 **FontNames** 对象，该对象包含所有有效字体的名称。只读。

**语法**

**express.FontNames**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例用于显示 FontNames 集合中的字体名称。*/
function test() {
    let intResponse

    for (let i = 1; i <= Application.FontNames.Count; i++) {
        intResponse = alert(Application.FontNames.Item(i))
        if (intResponse == jsResultCancel) {
            break
        }
    }
}
```

### Application.HangulHanjaDictionaries

返回一个 **HangulHanjaConversionDictionaries** 集合，该集合代表所有活动的自定义转换词典。

**语法**

**express.HangulHanjaDictionaries**

\*express是一个代表 Application 对象的变量。

**说明**

在 **“自定义词典”** 对话框中，活动的自定义转换词典使用复选框进行标记（在 **“工具”** 菜单上，单击 **“选项”** ，单击 **“拼写和语法”** 选项卡，然后单击 **“自定义词典”** 按钮可显示该对话框）。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例在该集合中添加一个新的空白自定义词典。然后在消息框中显示该自定义词典的路径和文件名。*/
function test() {
    let myHome = Application.HangulHanjaDictionaries.Add("Home.hhd")
    alert(myHome.Path + Application.PathSeparator + myHome.Name)
}
```

``` JavaScript
/*本示例取消所有自定义词典的激活状态，但不删除自定义词典文件。*/
Application.HangulHanjaDictionaries.ClearAll()
```

``` JavaScript
/*本示例显示该集合中所有自定义词典的名称。*/
function test() {
    for (let i = 1; i <= Application.HangulHanjaDictionaries.Count; i++) {
        alert(Application.HangulHanjaDictionaries.Item(i).Name)
    }
}
```

### Application.Height

返回或设置活动文档窗口的高度。 **Long** 类型，可读写。

**语法**

**express.Height**

\*express是一个代表 Application 对象的变量。

### Application.IsObjectValid

如果引用某个对象的指定变量有效，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.IsObjectValid**

\*express是一个代表 Application 对象的变量。

**说明**

如果该变量引用的对象已被删除，则 **IsObjectValid** 属性返回 **False** 。

**示例**

本示例向活动文档添加一个表格，并将该表格赋给 ` aTable ` 变量。然后，本示例删除文档中的第一个表格。如果 aTable 引用的表格不是文档中的第一个表格（也就是说， ` aTable ` 仍是一个有效对象），则本示例将同时删除该表格中的所有边框。

``` JavaScript
function test() {
    let aTable = Application.ActiveDocument.Tables.Add(Selection.Range, 2, 3)
    Application.ActiveDocument.Tables.Item(1).Delete()
    if (IsObjectValid(aTable) == true) {
        aTable.Borders.Enable = false
    }
}
```

### Application.IsSandboxed

如果应用程序窗口为受保护的视图窗口，则该属性值为 **True** 。只读。

**语法**

**express.IsSandboxed**

\*express是一个代表 Application 对象的变量。

**说明**

使用 **IsSandboxed** 属性可确定文档是否在受保护的视图窗口内打开。

**示例**

``` JavaScript
/*下面的代码示例显示指定的文档是否在受保护的视图窗口中打开。*/
function CheckIfSandboxed(doc) {
    alert(doc.Application.IsSandboxed)  
}
 
```

### Application.KeyBindings

返回一个 **KeyBindings** 集合，该集合代表自定义的按键指定方案，包含了键代码、键类别和命令。

**语法**

**express.KeyBindings**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将 Ctrl+Alt+W 指定给 FileClose 命令。这个键盘自定义方案保存在“Normal”模板中。*/
function test() {
    CustomizationContext = NormalTemplate
    Application.KeyBindings.Add(wdKeyCategoryCommand, "FileClose", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyW))
}
```

``` JavaScript
/*本示例为 KeyBindings 集合中的每项插入命令名称和组合键字符串。*/
function test() {
    CustomizationContext = NormalTemplate

    for (let i = 1; i <= Application.KeyBindings.Count; i++) {
        Application.Selection.InsertAfter(Application.KeyBindings.Item(i).Command + "\t" + KeyBindings.Item(i).KeyString + "\r")
        Application.Selection.Collapse(wdCollapseEnd)
    }
}
```

### Application.LandscapeFontNames

返回一个 **FontNames** 对象，该对象包括了所有有效的横向字体的名称。

**语法**

**express.LandscapeFontNames**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例创建在 FontNames 对象中命名的横向字体名称新文档的排序列表。*/
function ListLandscapeFonts() {
    let docNew = Application.Documents.Add()
    docNew.Content.InsertAfter("Landscape Fonts" + "\n")

    for (let i = 1; i <= LandscapeFontNames.Count; i++) {
        docNew.Content.InsertAfter(LandscapeFontNames.Item(i) + "\n")
    }

    docNew.Range(docNew.Paragraphs.Item(2).Range.Start, docNew.Paragraphs.Item(docNew.Paragraphs.Count).Range.End).Select()

    Selection.Sort()
}
```

### Application.Language

返回一个 **MsoLanguageID** 常量，该常量代表为 WPS 用户界面选定的语言。

**语法**

**express.Language**

\*express是一个代表 Application 对象的变量。

**说明**

该属性的值与下列表达式的返回值相同：

``` JavaScript
Application.LanguageSettings.LanguageID(msoLanguageIDUI)
```

**示例**

``` JavaScript
/*本示例显示一条消息，表明 WPS 用户界面所选语言是否是英语（美国）。*/
function LangSetting(){
    if(Application.Language == msoLanguageIDEnglishUS){
        alert("The user interface language is U.S. English.")
    }
    else{
        alert("The user interface language is not U.S. English.")
    }
}
```

### Application.Languages

返回一个 **Languages** 集合，该集合代表列在 **“语言”** 对话框（在 **“工具”** 菜单上，单击 **“语言”** ，再单击 **“设置语言”** ）中的校对语言。

**语法**

**express.Languages**

\*express是一个代表 Application 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例返回活动拼写词典的完整路径和文件名。*/
function test() {
    let dicSpell = Application.Languages.Item(Selection.LanguageID).ActiveSpellingDictionary
    alert(dicSpell.Path + Application.PathSeparator + dicSpell.Name)
}
```

``` JavaScript
/*本示例用 aLang() 数组来存储校对语言的名称。*/
function test() {
    let aLang = []
    intCount = 0

    for (let i = 1; i <= Languages.Count; i++) {
        aLang[intCount] = Languages.Item(i).NameLocal
        intCount++
    }
}
```

### Application.LanguageSettings

返回一个 **LanguageSettings** 对象，该对象包含 WPS 中语言设置的信息。

**语法**

**express.LanguageSettings**

\*express是一个代表 Application 对象的变量。

### Application.Left

返回或设置一个 **Long** 类型的值，该值代表活动文档的水平位置，以磅为单位。可读写。

**语法**

**express.Left**

\*express是一个代表 Application 对象的变量。

### Application.ListGalleries

返回一个 **ListGalleries** 集合，该集合代表三个列表模板库。

**语法**

**express.ListGalleries**

\*express是一个代表 Application 对象的变量。

**说明**

每个模板库（项目符号、编号和多级符号）都对应于 **“项目符号和编号”** 对话框（ **“格式”** 菜单）上的一个选项卡。有关返回集合中单个成员的信息，请参阅 从集合返回一个对象。

**示例**

``` JavaScript
/*本示例将变量 mylsttmp 设置为“项目符号和编号”对话框的“多级符号”选项卡中的第二种列表模板。然后，该示例将此模板应用到活动文档中的第一个列表。*/
function test() {
    let mylsttmp = Application.ListGalleries.Item(wdOutlineNumberGallery).ListTemplates.Item(2)
    Application.ActiveDocument.Lists.Item(1).ApplyListTemplate(mylsttmp)
}
```

``` JavaScript
/*本示例在 ListGalleries 集合中循环搜索，并将每个列表模板库中的模板重新设置为内置模板。*/
function test() {
    for (let i = 1; i <= Application.ListGalleries.Count; i++) {
        for (let j = 1; j <= 7; j++) {
            Application.ListGalleries.Item(i).Reset(j)
        }
    }
}
```

### Application.MacroContainer

返回 **Template** 或 **Document** 对象，该对象代表其中存储模块（该模块中包含着运行的过程）的模板或文档。

**语法**

**express.MacroContainer**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将显示文档或模板名称，其中包含正在运行的过程。*/
function test() {
    let cntnr = Application.MacroContainer
    alert(cntnr.Name)
}
```

### Application.MailingLabel

返回一个 **MailingLabel** 对象，该对象代表一个邮件标签。

**语法**

**express.MailingLabel**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例为指定地址创建一个新的 Avery 2160 小型标签文档。*/
function test()
{
    let addr = "Dave Edson" + "\r" + "123 Skye St." + "\r" + "Our Town, WA  98004"
    Application.MailingLabel.CreateNewDocument("2160 mini", addr, null, false)
}
```

### Application.MailMessage

返回一个 **MailMessage** 对象，该对象代表活动的电子邮件。

**语法**

**express.MailMessage**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例为活动电子邮件显示“选择姓名”对话框。*/
Application.MailMessage.DisplaySelectNamesDialog()
```

### Application.MailSystem

返回主机上安装的邮件系统（或系统）。 **WdMailSystem** 类型，只读。

**语法**

**express.MailSystem**

\*express是一个代表 Application 对象的变量。

**说明**

某些 **WdMailSystem** 常量仅可用于 WPS OfficeMacintosh Edition。有关这些常量的其他信息，请参阅 WPS OfficeMacintosh Edition 中的语言参考的帮助文件。

**示例**

``` JavaScript
/*本示例显示一条消息，指示邮件系统是否安装在此计算机上。*/
function test() {
    let ms = Application.MailSystem
    if (ms != wdNoMailSystem) {
        alert("This computer has a mail system installed.")
    }
    else {
        alert("This computer has no mail system installed.")
    }
}
```

### Application.MAPIAvailable

如果已经安装了 MAPI，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.MAPIAvailable**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*如果已安装了 MAPI，则本示例显示一条消息。*/
function test() {
    if (Application.MAPIAvailable == true) {
        alert("MAPI is available")
    }
}
```

### Application.MathCoprocessorAvailable

如果安装了数学协处理器并且该处理器适用于 WPS ，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.MathCoprocessorAvailable**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条消息以表明是否安装了数学协处理器并且是否适用于 WPS。*/
function test()
{
if(Application.MathCoprocessorAvailable == true){
    alert("A math coprocessor is available.")
}
else{
    alert("A math coprocessor is not installed.")
}
}
```

### Application.MouseAvailable

如果系统有可用的鼠标，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.MouseAvailable**

\*express是一个代表 Application 对象的变量。

**示例**

------------------------------------------------------------------------

``` JavaScript
/*本示例显示一条消息，表明没有可用的鼠标。*/
function test()
{
if(Application.MouseAvailable == false){
    alert("Make sure your mouse is plugged in.")
}
else{
    alert("Mouse is available")
}
}
```

### Application.Name

返回指定对象的名称。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 Application 对象的变量。

### Application.NewDocument

返回一个 NewFile 对象，该对象代表在“新建文档”任务窗格上所列的一个文档。

**语法**

**express.NewDocument**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例在“新建文档”任务窗格的“根据现有文档新建”部分中创建一个文档列表项目。*/
function CreateNewDocument() {
    Application.NewDocument.Add("C:\\NewFile.doc", msoNewfromExistingFile, "New File", msoCreateNewFile)
}
```

### Application.NormalTemplate

返回一个 **Template** 对象，该对象代表 Normal 模板。

**语法**

**express.NormalTemplate**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*如果 AutoTextEntries 集合中包含一个名为“Test”的“自动图文集”词条，则以下示例从 Normal 模板插入该词条。*/
function test() {
    for (let i = 1; i <= Application.NormalTemplate.AutoTextEntries.Count; i++) {
        let entry = Application.NormalTemplate.AutoTextEntries.Item(i)
        if (entry.Name == "Test") {
            entry.Insert(Selection.Range)
        }
    }
}
```

``` JavaScript
/*如果修改了 Normal 模板，本示例将保存该模板。*/
function test() {
    if (Application.NormalTemplate.Saved == false) {
        Application.NormalTemplate.Save()
    }
}
```

### Application.NumLock

此属性返回 Num Lock 键的状态。如果为 **True** ，则数字键盘可用于输入数字；如果为 **False** ，则该键盘用于移动插入点。 **Boolean** 类型，只读。

**语法**

**express.NumLock**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例返回 Num Lock 键的当前状态。*/
let theState = Application.NumLock
```

### Application.OMathAutoCorrect

返回一个 **OMathAutoCorrect** 对象，该对象代表公式的自动更正条目。只读。

**语法**

**express.OMathAutoCorrect**

\*express是一个代表 Application 对象的变量。

### Application.OpenAttachmentsInFullScreen

返回或设置 **Boolean** 值，该值表示 WPS 是否在读取模式下打开电子邮件附件。可读写。

**语法**

**express.OpenAttachmentsInFullScreen**

\*express是一个代表 Application 对象的变量。

**说明**

此属性对应于 **“ WPS 选项”** 对话框中的 **“在读取模式下打开电子邮件附件”** 复选框。

### Application.Options

返回一个 **Options** 对象，该对象代表 WPS 中的应用程序设置。

**语法**

**express.Options**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 本示例打印 Sales.doc 文档及其批注和域结果。 */
function test() {
    Application.Options.PrintFieldCodes = false
    Application.Options.PrintComments = true
    Application.Documents("Sales.doc").PrintOut()
}
```

### Application.Parent

返回一个 **Object** 类型的值，该值代表指定 **Application** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Application 对象的变量。

### Application.Path

返回指定对象的磁盘或 Web 路径。只读 **String** 类型。

**语法**

**express.Path**

\*express是一个代表 Application 对象的变量。

**说明**

此路径不包括尾部字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符。使用 **Document** 对象的 **Name** 属性可返回不带路径的文件名，而用 **FullName** 属性可同时返回文件名和路径。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>可以用 <strong>PathSeparator</strong> 属性建立网站地址，即使此地址中会包含正斜线 (/)，而 <strong>PathSeparator</strong> 属性默认使用反斜线 (\)。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

**示例**

``` JavaScript
/*本示例显示活动文档的路径和文件名。*/
MsgBox(ActiveDocument.Path + Application.PathSeparator + ActiveDocument.Name)
```

``` JavaScript
/*本示例将当前文件夹路径改为活动文档附加的模板的路径。*/
ChDir ActiveDocument.AttachedTemplate.Path
```

``` JavaScript
/*本示例显示 AddIns 集合中第一个加载项的路径。*/
function test()
{
if(AddIns.Count >= 1 ) {
    MsgBox(AddIns.Item(1).Path)
}
}
```

### Application.PathSeparator

返回用于分隔文件夹名称的字符。在 Windows 中，该属性返回一个反斜线 (\\。只读 **String** 类型。

**语法**

**express.PathSeparator**

\*express是一个代表 Application 对象的变量。

**说明**

可以用 **PathSeparator** 属性建立网站地址，尽管在地址中可能包含正斜线 (/)。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td><strong><span>FullName</span></strong> 属性以单个字符串形式返回路径和文件名，其中包含路径分隔符。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

**示例**

``` JavaScript
/*本示例显示活动文档的路径和文件名。*/
alert(Application.ActiveDocument.Path + Application.PathSeparator + Application.ActiveDocument.Name)
 
```

``` JavaScript
/*如果第一个加载项是模板，本示例卸载该模板并打开它。*/
function test() {
    if (Application.Addins.Item(1).Compiled == false) {
        Application.Addins.Item(1).Installed = false
        Application.Documents.Open(Application.AddIns.Item(1).Path + Application.PathSeparator + Application.AddIns.Item(1).Name)
    }
}
```

### Application.PickerDialog

返回 PickerDialog 对象，以提供在对话框中选择用户或数据的功能。只读。

**语法**

**express.PickerDialog**

\*express是一个代表 Application 对象的变量。

**说明**

**PickerDialog** 对象可用于 WPS Office类型库。有关详细信息，请参阅下列对象及其成员：

- PickerDialog

- PickerField

- PickerFields

- PickerProperties

- PickerProperty

- PickerResult

- PickerResults

### Application.PortraitFontNames

返回一个 **FontNames** 对象，该对象包括所有有效纵向字体的名称。

**语法**

**express.PortraitFontNames**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例在插入点插入一组纵向字体。*/
function test() {
    let sel = Application.Selection
    for (let i = 1; i <= Application.PortraitFontNames.Count; i++) {
        sel.Collapse(wdCollapseEnd)
        sel.InsertAfter(PortraitFontNames.Item(i))
        sel.InsertParagraphAfter()
        sel.Collapse(wdCollapseEnd)
    }
}
```

### Application.PrintPreview

如果当前视图是打印预览，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintPreview**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将视图切换为打印预览。*/
Application.PrintPreview = true
```

``` JavaScript
/*本示例将活动窗口从打印预览切换至普通视图。*/
function test() {
    Application.PrintPreview = false
    Application.ActiveDocument.ActiveWindow.View.Type = wdNormalView
}
```

### Application.ProtectedViewWindows

返回代表所有受保护的视图窗口的 ProtectedViewWindows 集合。只读。

**语法**

**express.ProtectedViewWindows**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示打开的受保护的视图窗口数量。*/
alert("There are " + Application.ProtectedViewWindows.Count + " protected view windows open.")
```

### Application.RecentFiles

返回一个 **RecentFiles** 集合，该集合代表最近存取过的文件。

**语法**

**express.RecentFiles**

\*express是一个代表 Application 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例打开 RecentFiles 集合中的第一项（即“文件”菜单上列出的第一个文档名）。*/
function test() {
    if (Application.RecentFiles.Count >= 1) {
        Application.RecentFiles.Item(1).Open()
    }
}
```

``` JavaScript
/*本示例显示 RecentFiles 集合中的每个文件的名称。*/
function test() {
    for (let i = 1; i <= Application.RecentFiles.Count; i++) {
        alert(Applicaiton.RecentFiles.Item(i).Name)
    }
}
```

### Application.RestrictLinkedStyles

返回或设置 **Boolean** 值，该值表示 WPS 是否允许链接样式。可读写。

**语法**

**express.RestrictLinkedStyles**

\*express是一个代表 Application 对象的变量。

**说明**

链接样式是可以作为字符样式或段落样式应用的样式。此属性对应于 **“样式”** 对话框中的 **“禁用链接样式”** 复选框。

### Application.ScreenUpdating

如果启用屏幕更新，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ScreenUpdating**

\*express是一个代表 Application 对象的变量。

**说明**

当运行某一过程时， **ScreenUpdating** 属性控制显示器的主要变化情况。如果关闭屏幕更新，屏幕依然显示工具栏，并且 WPS 仍允许此过程用状态栏提示、输入框、对话框或信息框等工具来显示和检索信息。关闭屏幕更新可提高某些过程的执行速度。如果此过程已结束，或由于某些错误而停止，请务必将 **ScreenUpdating** 属性设为 **True** 。

**示例**

``` JavaScript
/*本示例将关闭屏幕更新，并添加一篇新文档。将 500 行文本添加至文档中。宏每隔 50 行选定一行并刷新屏幕。*/
function test() {
    Application.ScreenUpdating = false
    Application.Documents.Add()

    for (let x = 1; x <= 500; x++) {
        Application.ActiveDocument.Content.InsertAfter("This is line " + x + ".")
        Application.ActiveDocument.Content.InsertParagraphAfter()

        if (x % 50 == 0) {
            Application.ActiveDocument.Paragraphs.Item(x).Range.Select()
            Application.ScreenRefresh()
        }
    }

    Application.ScreenUpdating = true
}
```

### Application.Selection

返回 **Selection** 对象，该对象代表一选定的范围或插入点。只读。

**语法**

**express.Selection**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 打印选区文本 */
function test() {
    if(Application.Selection.Type == wdSelectionNormal) {
        alert(Application.Selection.Text);
    }
}
```

### Application.ShowStartupDialog

如果为 **True** ，则在启动 WPS 时显示 **“任务窗格”** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowStartupDialog**

\*express是一个代表 Application 对象的变量。

**说明**

**ShowStartupDialog** 属性是一个全局选项，新的设置仅在重新启动 WPS 后生效。使用 **CommandBars** 集合的 **Visible** 属性可以在不重新启动 WPS 的情况下显示或隐藏任务窗格。

**示例**

``` JavaScript
/*本示例关闭“任务窗格”，使其不在启动 WPS 时显示。仅在用户下一次启动 WPS 时才生效。*/
function HideStartUpDlg(){
    Application.ShowStartupDialog = false
}
```

### Application.ShowStylePreviews

返回或设置 **Boolean** 值，该值表示 WPS 是否在 **“样式”** 对话框中显示样式的格式预览。可读/写。

**语法**

**express.ShowStylePreviews**

\*express是一个代表 Application 对象的变量。

**说明**

此属性对应于 **“样式”** 对话框中的 **“显示预览”** 复选框。

### Application.ShowVisualBasicEditor

如果显示“Visual Basic 编辑器”窗口，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowVisualBasicEditor**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例显示“Visual Basic 编辑器”窗口。*/
Application.ShowVisualBasicEditor = true
```

### Application.ShowWindowsInTaskbar

如果为 **True** ，则在任务栏中显示打开的文档，默认值为 Single Document Interface (SDI)。如果为 **False** ，则只在 **“窗口”** 菜单中列出打开的文档，提供 Multiple Document Interface (MDI) 的外观。 **Boolean** 类型，可读写。

**语法**

**express.ShowWindowsInTaskbar**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例切换界面，以只在“窗口”菜单上列出打开的文档。*/
function SDIToMDI(){
    Application.ShowWindowsInTaskbar = false
}
```

### Application.SmartArtColors

返回 SmartArtColors 对象，该对象代表应用程序中当前加载的颜色样式集。只读。

**语法**

**express.SmartArtColors**

\*express是一个代表 Application 对象的变量。

**说明**

**SmartArtColors** 属性代表的颜色集与 **“更改颜色”** 按钮上的可用颜色样式相对应，该按钮位于 WPS 2015 中 **“SmartArt 工具”** 上下文选项卡上的 **“设计”** 选项卡中。

**示例**

``` JavaScript
/*下面的代码示例向活动文档中添加 SmartArt 图形，然后将该 SmartArt 图形颜色设置为“深色 2 轮廓”。*/
function test() {
    let myShape
    let mySmartArt
    myShape = Application.ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts.Item(1), 50, 50, 200, 200)
    mySmartArt = myShape.SmartArt
    mySmartArt.Color = Application.SmartArtColors.Item(2)
}
```

### Application.SmartArtLayouts

返回 SmartArtLayouts 对象，该对象代表应用程序中当前加载的 SmartArt 布局集。只读。

**语法**

**express.SmartArtLayouts**

\*express是一个代表 Application 对象的变量。

**说明**

**SmartArtLayouts** 属性代表的布局集与 **“布局”** 组中的可用布局相对应，该组位于 WPS 2015 中 **“SmartArt 工具”** 上下文选项卡上的 **“设计”** 选项卡中。

**示例**

``` JavaScript
/*下面的代码示例向活动文档中添加 SmartArt 图形，然后将该 SmartArt 图形布局设置为“分组列表”。*/
function test() {
    let myShape
    let mySmartArt
    myShape = Application.ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts.Item(1), 50, 50, 200, 200, null)
    mySmartArt = myShape.SmartArt
    mySmartArt.Layout = Application.SmartArtLayouts.Item(15)
}
```

### Application.SmartArtQuickStyles

返回 SmartArtQuickStyles 对象，该对象代表应用程序中当前加载的 SmartArt 样式集。只读。

**语法**

**express.SmartArtQuickStyles**

\*express是一个代表 Application 对象的变量。

**说明**

**SmartArtQuickStyles** 属性代表的样式集与 **“样式”** 组中的可用样式相对应，该组位于 WPS 2015 中 **“SmartArt 工具”** 上下文选项卡上的 **“设计”** 选项卡中。

**示例**

``` JavaScript
/*下面的代码示例向活动文档中添加 SmartArt 图形，然后将该 SmartArt 图形样式设置为“抛光”。*/
function test() {
    let myShape
    let mySmartArt
    myShape = Application.ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts.Item(1), 50, 50, 200, 200, null)
    mySmartArt = myShape.SmartArt
    mySmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(6)
}
```

### Application.SmartTagRecognizers

返回应用程序的一个 **SmartTagRecognizers** 集合。

**语法**

**express.SmartTagRecognizers**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*下列示例访问 WPS 所安装的智能标记识别器列表中的第一个智能标记识别器。*/
Application.SmartTagRecognizers.Item(1)
```

### Application.SmartTagTypes

返回一个 **SmartTagTypes** 集合，代表 WPS 中所安装的智能标记组件的智能标记类型。

**语法**

**express.SmartTagTypes**

\*express是一个代表 Application 对象的变量。

**说明**

智能标记类型是智能标记组件中的单个项目。智能标记组件可以包含多个智能标记类型。例如，安装在英文系统上的 Address (English) 智能标记列表默认情况下包含 name 智能标记类型、street 智能标记类型和 city 智能标记类型，等等。 **SmartTagTypes** 集合包含用户计算机上安装的所有组件的所有智能标记类型。

**示例**

``` JavaScript
/*下列示例循环遍历 SmartTagTypes 集合。如果 SmartTagType 是 Address 智能标记，则重新加载该智能标记的识别器和处理程序。*/
function GetSmartTagsTypes(){
    let strSmartTagType
    let objSmartTagType = Application.SmartTagTypes
    strSmartTagType = "urn:schemas-microsoft-com" + ":office:smarttags#address"
    for(let i=1; i<=objSmartTagType.Count; i++){
        if(objSmartTagType == strSmartTagType){
            objSmartTagType.Item(i).smartTagActions.ReloadActions()
            objSmartTagType.Item(i).SmartTagRecognizers.ReloadRecognizers()
        }
    }
}
```

### Application.SpecialMode

如果 WPS 处于某个特殊模式（如 CopyText 模式或 MoveText 模式）下，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.SpecialMode**

\*express是一个代表 Application 对象的变量。

**说明**

在选定文本后，按 F2 或 Shift+F2， WPS 将进入特殊的复制或移动模式。

**示例**

``` JavaScript
/*本示例检查 WPS 是否处于某个特殊模式下。如果是，则在删除和粘贴当前所选内容之前，激活 Esc。*/
function test() {
    if (Application.SpecialMode == true) {
        SendKeys("ESC")
        Selection.Cut()
        Selection.EndKey(wdStory)
        Selection.Paste()
    }
}
```

### Application.StartupPath

返回或设置启动文件夹的完整路径，不包括最后的分隔符。 **String** 类型，可读写。

**语法**

**express.StartupPath**

\*express是一个代表 Application 对象的变量。

**说明**

启动 WPS 时，将自动装载启动文件夹中的模板和加载项。

**示例**

``` JavaScript
/*本示例显示启动文件夹的完整路径。*/
alert(Application.StartupPath)
```

``` JavaScript
/*本示例使用户可更改启动文件夹的路径。*/
function test()
{
    let x = alert("Do you want to change the startup path?")
    let newStartup = alert("Type a startup path")
    Application.StartupPath = newStartup
}
```

### Application.System

返回一个 **System** 对象，该对象可用于返回与系统相关的信息，并执行与系统相关的任务。

**语法**

**express.System**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例返回有关系统的信息。*/
function test()
{
let processor = System.ProcessorType
let enviro = System.OperatingSystem
}
```

``` JavaScript
/*本示例建立到网络驱动器的连接。*/
Application.System.Connect("\\Project\\Info")
```

### Application.TaskPanes

返回一个 **TaskPanes** 集合，该集合代表 WPS 中最常执行的任务。

**语法**

**express.TaskPanes**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*下列示例显示包含文档中格式信息的任务栏。*/
function showFormatting() {
    Application.TaskPanes.Item(wdTaskPaneFormatting).Visible = true
}
```

### Application.Tasks

返回一个 **Tasks** 集合，该集合代表所有正在运行的应用程序。

**语法**

**express.Tasks**

\*express是一个代表 Application 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例查看 ET 是否正在运行。如果该任务正在运行，则本示例激活 ET；否则将显示一个消息框。*/
function test() {
    if (Application.Tasks.Exists("ET") == true) {
        Application.Tasks("ET").Activate()
        Application.Tasks("ET").WindowState = wdWindowStateMaximize
    } else {
        alert("ET is not currently running.")
    }
}
```

### Application.Templates

返回一个 **Templates** 集合，该集合代表所有可用的模板，包括共用模板和附加到打开文档的模板。

**语法**

**express.Templates**

\*express是一个代表 Application 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示 Templates 集合中每个模板的名称。*/
function test() {
    let Count = 1
    for (let i = 1; i <= Application.Templates.Count; i++) {
        alert(Application.Templates.Item(i).Name + " is template number " + Count)
        Count = Count + 1
    }
}
```

``` JavaScript
/*本示例中，如果模板 1 是一个共用模板，则其路径存储在 thePath 中。ChDir 语句用于将具有存储在 thePath 中的路径的文件夹设为当前文件夹。进行更改后，将显示“打开”对话框。*/
function test() {
    if (Application.Templates.Item(1).Type == wdGlobalTemplate) {
        let thePath = Application.Templates.Item(1).Path
        if (thePath != "") {
            ChDir(thePath)
        }
        Dialogs(wdDialogFileOpen).Show()
    }
}
```

### Application.Top

返回或设置活动文档的垂直位置。 **Long** 类型，可读写。

**语法**

**express.Top**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 应用程序窗口置于距屏幕顶部 100 磅的位置。*/
function test()
{
    Application.WindowState = wdWindowStateNormal
    Application.Top = 100
}
```

### Application.UndoRecord

返回 UndoRecord 对象，该对象提供撤消堆栈的自定义入口点。只读。

**语法**

**express.UndoRecord**

\*express是一个代表 Application 对象的变量。

**说明**

使用 **UndoRecord** 对象可在 WPS 2015 撤消堆栈中创建和修改自定义撤消记录。

**示例**

``` JavaScript
/*下面的代码示例将实例化 UndoRecord 对象。*/
let objUndo = Application.UndoRecord
```

### Application.UsableHeight

返回 WPS 文档窗口高度可设置的最大值（以磅为单位）。 **Long** 类型，只读。

**语法**

**express.UsableHeight**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档窗口的尺寸设置为屏幕最大区域的四分之一。*/
function test() {
    let Win = Application.ActiveDocument.ActiveWindow
    Win.WindowState = wdWindowStateNormal
    Win.Top = 5
    Win.Left = 5
    Win.Height = (Application.UsableHeight * 0.5)
    Win.Width = (Application.UsableWidth * 0.5)
}
```

``` JavaScript
/*本示例显示活动文档窗口中工作区的大小。*/
function test() {
    let Win = Application.ActiveDocument.ActiveWindow
    alert("Working area height = " + Win.UsableHeight + "\n" + "Working area width = " + Win.UsableWidth)
}
```

### Application.UsableWidth

返回 WPS 文档窗口可设置的最大宽度（以磅为单位）。 **Long** 类型，只读。

**语法**

**express.UsableWidth**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档窗口的尺寸设置为屏幕最大区域的四分之一。*/
function test()
{
let Win = ActiveDocument.ActiveWindow
Win.WindowState = wdWindowStateNormal
Win.Top = 5
Win.Left = 5
Win.Height = (Application.UsableHeight*0.5)
Win.Width = (Application.UsableWidth*0.5)
}
```

``` JavaScript
/*本示例显示活动文档窗口中工作区的大小。*/
function test() {
    let Win = Application.ActiveDocument.ActiveWindow
    alert("Working area height = " + Win.UsableHeight + "\n" + "Working area width = " + Win.UsableWidth)
}
```

### Application.UserAddress

该属性返回或设置用户的通讯地址。 **String** 类型，可读写。

**语法**

**express.UserAddress**

\*express是一个代表 Application 对象的变量。

**说明**

使用该通讯地址作为信封上的寄信人地址。

**示例**

``` JavaScript
/*本示例设置用户的寄信人地址。Chr 函数用于返回一个换行符。*/
Application.UserAddress = "4200 Third Street NE" + "\n" + "Anytown, WA  98999"
```

``` JavaScript
/*本示例返回“工具”菜单的“选项”对话框中“用户信息”选项卡的“通讯地址”框中的地址。*/
alert(Application.UserAddress)
```

### Application.UserControl

如果文档或应用程序是由用户创建或打开的，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.UserControl**

\*express是一个代表 Application 对象的变量。

**说明**

如果应用程序是由另一 WPS Office应用程序使用 **Open** 方法、 **CreateObject** 或 **GetObject** 方法以编程方式创建或打开的，则 **UserControl** 属性返回 **False** 。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>如果 WPS 对用户可见，或者从代码模块中调用 <strong>Application</strong> 对象的 <strong>UserControl</strong> 属性，则该属性将始终返回 <strong>True</strong> 。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

**示例**

### Application.UserName

该属性返回或设置用户姓名， WPS 将其用于信封和文档的“作者”属性。 **String** 类型，可读写。

**语法**

**express.UserName**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例设置用户姓名。*/
Application.UserName = "Andrew Fuller"
```

``` JavaScript
/*本示例返回“工具”菜单上“选项”对话框中“用户信息”选项卡上“姓名”框中的姓名。*/
MsgBox(Application.UserName)
```

### Application.VBE

返回一个 VBE 对象，该对象代表“Visual Basic 编辑器”。

**语法**

**express.VBE**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*下面的示例将显示活动项目中可用的引用个数。*/
alert ("References = "+ Application..ActiveVBProject.References.Count)
```

### Application.Version

返回 WPS 版本号。 **String** 类型，只读。

**语法**

**express.Version**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 本示例在信息框中显示 WPS 的版本号。 */
alert("The version of  WPS is " + Application.Version)
```

### Application.Visible

如果指定对象可见，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Visible**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏 WPS 。*/
Application.Visible = false
```

### Application.Width

返回或设置应用程序窗口的宽度（以磅为单位）。 **Long** 类型，可读写。

**语法**

**express.Width**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 应用程序窗口的宽度和高度。*/
function test() {
    Application.WindowState = wdWindowStateNormal
    Application.Width = 500
    Application.Height = 400
}
```

### Application.Windows

返回一个 **Windows** 集合，该集合代表所有文档窗口。只读。

**语法**

**express.Windows**

\*express是一个代表 Application 对象的变量。

**说明**

此属性返回的集合中既包括可见窗口，也包括隐藏窗口。

**示例**

``` JavaScript
/* 获取第一个窗口的题注 */
Application.Windows.Item(1).Caption
```

### Application.WindowState

返回或设置指定文档窗口或任务窗口的状态。 **WdWindowState** 类型，可读写。

**语法**

**express.WindowState**

\*express是一个代表 Application 对象的变量。

**说明**

**wdWindowStateNormal** 常量表示没有最大化或最小化的窗口。无法设置不活动窗口的状态。在设置窗口状态之前可用 **Activate** 方法激活窗口。

### Application.WordBasic

返回一个自动化对象 (Word.Basic)，其中包含适用于 WPS 6.0 和 WPS for Windows 中所有 WordBasic 语句和函数的方法。只读。

**语法**

**express.WordBasic**

\*express是一个代表 Application 对象的变量。

**说明**

在 WPS 2000 及其后续版本中，当打开一个包含 WordBasic 宏的 WPS 6.0 或 WPS for Windows 模板时，宏将自动转换为 VB 模块。宏内部的每个 WordBasic 语句和函数都转换为相应的 Word.Basic 方法。

有关 WordBasic 语句和函数的信息，请参阅 WPS 6.0 或 WPS for Windows 中的“WordBasic 帮助”。有关将 WordBasic 转换为 Visual Basic 的信息，请参阅 将 WordBasic 宏转换为 Visual Basic。有关常规信息，请参阅 WordBasic 和 Visual Basic 在概念上的区别。

**示例**

### Application.XMLNamespaces

返回一个 **XMLNamespaces** 集合，代表“架构库”中的 XML 架构。

**语法**

**express.XMLNamespaces**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/*下列示例返回“架构库”中的第一个架构。*/
let objSchema = Application.XMLNamespaces.Item(1)
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Borders 对象

## 简介

Border 对象的集合，这些对象代表对象的边框。

## 方法

| 序号 | 名称                                                                    | 说明                                       |
|:----:|:------------------------------------------------------------------------|:-------------------------------------------|
|  1   | [ApplyPageBordersToAllSections](#Borders.ApplyPageBordersToAllSections) | 将指定的页面边框格式应用于文档中的所有节。 |
|  2   | [Item](#Borders.Item)                                                   | 返回指定区域或选定内容中的一个边框。       |

## 属性

| 序号 | 名称                                                            | 说明                                                                                                               |
|:----:|:----------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------|
|  1   | [AlwaysInFront](#Borders.AlwaysInFront)                         | 如果在文档文本之前显示页面边框，则该属性值为 True 。 Boolean 类型，可读写。                                        |
|  2   | [Application](#Borders.Application)                             | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                                               |
|  3   | [Count](#Borders.Count)                                         | 返回 Borders 集合中的项目数。 Long 类型，只读。                                                                    |
|  4   | [Creator](#Borders.Creator)                                     | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。                                       |
|  5   | [DistanceFrom](#Borders.DistanceFrom)                           | 此属性返回或设置数值，该值指明指定的页面边框是从页面边缘还是从环绕的文本开始度量。 WdBorderDistanceFrom ，可读写。 |
|  6   | [DistanceFromBottom](#Borders.DistanceFromBottom)               | 返回或设置文本与下边框的距离（以磅为单位）。 Long 类型，可读写。                                                   |
|  7   | [DistanceFromLeft](#Borders.DistanceFromLeft)                   | 返回或设置文本与左边框的距离（以磅为单位）。 Long 类型，可读写。                                                   |
|  8   | [DistanceFromRight](#Borders.DistanceFromRight)                 | 返回或设置文本的右边与右边框的距离（以磅为单位）。 Long 类型，可读写。                                             |
|  9   | [DistanceFromTop](#Borders.DistanceFromTop)                     | 返回或设置文本与上边框的距离（以磅为单位）。 Long 类型，可读写。                                                   |
|  10  | [Enable](#Borders.Enable)                                       | 返回或设置指定对象的边框格式。 Long 类型，可读写。                                                                 |
|  11  | [EnableFirstPageInSection](#Borders.EnableFirstPageInSection)   | 如果可为节内第一页设置页面边框，则该属性值为 True 。 Boolean 类型，可读写。                                        |
|  12  | [EnableOtherPagesInSection](#Borders.EnableOtherPagesInSection) | 如果可为本节中除首页之外的所有页面设置页面边框，则该属性值为 True 。 Boolean 类型，可读写。                        |
|  13  | [HasHorizontal](#Borders.HasHorizontal)                         | 如果可将水平方向的边框应用于对象，则该属性值为 True 。 Boolean 类型，只读。                                        |
|  14  | [HasVertical](#Borders.HasVertical)                             | 如果可将垂直边框应用于指定对象，则该属性值为 True 。 Boolean 类型，只读。                                          |
|  15  | [InsideColor](#Borders.InsideColor)                             | 返回或设置内部边框的 24 位颜色。可读写。                                                                           |
|  16  | [InsideColorIndex](#Borders.InsideColorIndex)                   | 返回或设置内部边框的颜色。 WdColorIndex 类型，可读写。                                                             |
|  17  | [InsideLineStyle](#Borders.InsideLineStyle)                     | 返回或设置指定对象的内部边框。                                                                                     |
|  18  | [InsideLineWidth](#Borders.InsideLineWidth)                     | 返回或设置对象内部边框的线条宽度。                                                                                 |
|  19  | [JoinBorders](#Borders.JoinBorders)                             | 如果删除段落和表格边缘的垂直边框，以使水平边框与页面边框相连，则该属性值为 True 。 Boolean 类型，可读写。          |
|  20  | [OutsideColor](#Borders.OutsideColor)                           | 返回或设置外部边框的 24 位颜色。                                                                                   |
|  21  | [OutsideColorIndex](#Borders.OutsideColorIndex)                 | 返回或设置外部边框的颜色。 WdColorIndex 类型，可读写。                                                             |
|  22  | [OutsideLineStyle](#Borders.OutsideLineStyle)                   | 返回或设置指定对象的外部边框。                                                                                     |
|  23  | [OutsideLineWidth](#Borders.OutsideLineWidth)                   | 返回或设置对象外部边框的线条宽度。可读写。                                                                         |
|  24  | [Parent](#Borders.Parent)                                       | 返回一个 Object 类型的值，该值代表指定 Borders 集合中的父对象。                                                    |
|  25  | [Shadow](#Borders.Shadow)                                       | 如果为 True ，则指定的边框设置为阴影格式。 Boolean 类型，可读写。                                                  |
|  26  | [SurroundFooter](#Borders.SurroundFooter)                       | 如果页面边框环绕文档的页脚，则该属性值为 True 。 Boolean 类型，可读写。                                            |
|  27  | [SurroundHeader](#Borders.SurroundHeader)                       | 如果页面边框环绕文档页眉，则该属性值为 True 。 Boolean 类型，可读写。                                              |

## 成员方法

### Borders.ApplyPageBordersToAllSections

将指定的页面边框格式应用于文档中的所有节。

**语法**

**express.ApplyPageBordersToAllSections ()**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档的所有节添加单线型页面边框*/
function test() {
    let sec = Application.ActiveDocument.Sections.Item(1)
    for(let i=1;i<=sec.Borders.Count;i++){
    sec.Borders.Item(i).LineStyle = wdLineStyleSingle
    sec.Borders.Item(i).LineWidth = wdLineWidth050pt
    }
    sec.Borders.ApplyPageBordersToAllSections()
}
```

### Borders.Item

返回指定区域或选定内容中的一个边框。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Borders 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明         |
|---------|-----------|--------------|--------------|
| *Index* | 必选      | WdBorderType | 返回的边框。 |

**示例**

``` JavaScript
/*本示例在活动文档的第一段之上插入双边框*/
function BorderItem() {
    Application.ActiveDocument.Paragraphs.Item(1).Borders.Item(wdBorderTop).LineStyle = wdLineStyleDouble
}
```

## 成员属性

### Borders.AlwaysInFront

如果在文档文本之前显示页面边框，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AlwaysInFront**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档中第一节的文档文本之前添加艺术型页面边框*/
function test() {
    let Sec = Application.ActiveDocument.Sections.Item(1)
    Sec.Borders.AlwaysInFront = true
    for(let i = 1; i <= Sec.Borders.Count; i++) {
        let Bor =Sec.Borders.Item(i)
        Bor.ArtStyle = wdArtPeople
        Bor.ArtWidth = 15
    }
}
```

### Borders.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Borders 对象的变量。

### Borders.Count

返回 **Borders** 集合中的项目数。 **Long** 类型，只读。

**语法**

**express.Count**

\*express是一个代表 Borders 对象的变量。

### Borders.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Borders 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

### Borders.DistanceFrom

此属性返回或设置数值，该值指明指定的页面边框是从页面边缘还是从环绕的文本开始度量。 **WdBorderDistanceFrom** ，可读写。

**语法**

**express.DistanceFrom**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档中第一节的每个页面添加单线型边框，然后设置每个边框到页面边缘的距离*/
function test() {
    let Sec = Application.ActiveDocument.Sections.Item(1)
    for(let i = 1 ; i <= Sec.Borders.Count; i++) {
        Sec.Borders.Item(i).LineStyle = wdLineStyleSingle
        Sec.Borders.Item(i).LineWidth = wdLineWidth050pt
    }
    let Bor = Sec.Borders
    Bor.DistanceFrom = wdBorderDistanceFromPageEdge
    Bor.DistanceFromTop = 20
    Bor.DistanceFromLeft = 22
    Bor.DistanceFromBottom = 20
    Bor.DistanceFromRight = 22
}

/*本示例为所选内容的第一节的每个页面添加边框，然后将文本与页面边框的距离设置为 6 磅*/
function test() {
    let Sec = Application.Selection.Sections.Item(1)
    for(let i = 1; i <= Sec.Borders.Count; i++) {
        Sec.Borders.Item(i).ArtStyle = wdArtSeattle
        Sec.Borders.Item(i).ArtWidth = 22
    }
    let Bor = Sec.Borders
    Bor.DistanceFrom = wdBorderDistanceFromText
    Bor.DistanceFromTop = 6
    Bor.DistanceFromLeft = 6
    Bor.DistanceFromBottom = 6
    Bor.DistanceFromRight = 6
}
```

### Borders.DistanceFromBottom

返回或设置文本与下边框的距离（以磅为单位）。 **Long** 类型，可读写。

**语法**

**express.DistanceFromBottom**

\*express是一个代表 Borders 对象的变量。

**说明**

将该属性与页边框结合使用，可以设置文本与页面下边框之间的距离，或页面下边缘与下边框之间的距离。这个距离是根据 **DistanceFrom** 属性的值测量的。

**示例**

``` JavaScript
/*本示例为活动文档的第一个段落添加边框，并将文本与下边框的距离设置为 6 磅。*/
function test() {
    let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders
    Bor.Enable = true
    Bor.DistanceFromBottom = 6
}

/*本示例为 Sales.doc 中的每张表格添加边框。然后将文本与上下边框的距离设置为 3 磅，将文本与左右边框的距离设置为 6 磅*/
function test() {
    for(let i = 1; i <= Application.Documents.Item("Sales.doc").Tables.Count; i++) {
        let Bor = Application.Documents.Item("Sales.doc").Tables.Item(i).Borders
        Bor.OutsideLineStyle = wdLineStyleSingle
        Bor.OutsideLineWidth = wdLineWidth150pt
        Bor.DistanceFromBottom = 3
        Bor.DistanceFromTop = 3
        Bor.DistanceFromLeft = 6
        Bor.DistanceFromRight = 6
    }
}
```

### Borders.DistanceFromLeft

返回或设置文本与左边框的距离（以磅为单位）。 **Long** 类型，可读写。

**语法**

**express.DistanceFromLeft**

\*express是一个代表 Borders 对象的变量。

**说明**

将该属性与页边框结合使用，可设置文本与页面左边框之间的距离，或页面左边缘与页面左边框之间的距离。这个距离是根据 **DistanceFrom** 属性的值测量的。

**示例**

``` JavaScript
/*本示例为活动文档中的每个图文框添加边框，并且将图文框与边框的距离设置为 5 磅。*/
function test() {
        for(let i = 1; i <= Application.ActiveDocument.Frames.Count; i++) {
        let Bor = Application.ActiveDocument.Frames.Item(i).Borders
        Bor.Enable = true
        Bor.DistanceFromLeft = 5
        Bor.DistanceFromRight = 5
        Bor.DistanceFromTop = 5
        Bor.DistanceFromBottom = 5
    }
}
/*本示例为活动文档的第一个段落添加边框，并将文本与左边框的距离设置为 3 磅*/
function test() {
    let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders
    Bor.Enable = true
    Bor.DistanceFromLeft = 3
}
```

### Borders.DistanceFromRight

返回或设置文本的右边与右边框的距离（以磅为单位）。 **Long** 类型，可读写。

**语法**

**express.DistanceFromRight**

\*express是一个代表 Borders 对象的变量。

**说明**

将该属性与页边框结合使用，可设置文本与右边框之间的距离，或页面右边缘与页面右边框之间的距离。这个距离是根据 **DistanceFrom** 属性的值测量的。

**示例**

``` JavaScript
/*本示例为所选内容的每个段落添加边框，并将文本与右边框的距离设置为 3 磅。*/
function test() {
    let Bor = Application.Selection.Paragraphs.Borders
    Bor.Enable = true
    Bor.DistanceFromRight = 3
}

/*本示例为活动文档中第一节的每页添加单线型页面边框。然后将左右边框与相应的页面边缘的距离设置为 22 磅。*/
function test() {
    let Sec = Application.ActiveDocument.Sections.Item(1)
    for(let i = 1; i <= Sec.Borders.Count; i++) {
        Sec.Borders.Item(i).LineStyle = wdLineStyleSingle
        Sec.Borders.Item(i).LineWidth = wdLineWidth050pt
    }
    let Bor = Sec.Borders
    Bor.DistanceFrom = wdBorderDistanceFromPageEdge
    Bor.DistanceFromLeft = 22
    Bor.DistanceFromRight = 22
}
```

### Borders.DistanceFromTop

返回或设置文本与上边框的距离（以磅为单位）。 **Long** 类型，可读写。

**语法**

**express.DistanceFromTop**

\*express是一个代表 Borders 对象的变量。

**说明**

将该属性与页边框结合使用，可设置文本与上边框之间的距离，或页面上边缘与上边框之间的距离。这个距离是根据 **DistanceFrom** 属性的值测量的。

**示例**

``` JavaScript
/*本示例为所选内容中的每个段落添加边框，并将文本与上边框的距离设置为 3 磅。*/
function test() {
    Application.Selection.Borders.Enable = true
    Application.Selection.Borders.DistanceFromTop = 3
}

/*本示例为所选内容中第一节的每页添加页面边框，并将文本与页面边框的距离设置为 6 磅。*/
function test() {
    let Sec = Application.Selection.Sections.Item(1)
    for(let i = 1; i <= Sec.Borders.Count; i++) {
        Sec.Borders.Item(i).ArtStyle = wdArtSeattle
        Sec.Borders.Item(i).ArtWidth = 22
    }
    let Bor = Sec.Borders
    Bor.DistanceFrom = wdBorderDistanceFromText
    Bor.DistanceFromTop = 6
    Bor.DistanceFromLeft = 6
    Bor.DistanceFromBottom = 6
    Bor.DistanceFromRight = 6
}
```

### Borders.Enable

返回或设置指定对象的边框格式。 **Long** 类型，可读写。

**语法**

**express.Enable**

\*express是一个代表 Borders 对象的变量。

**说明**

如果指定对象的全部或部分边框应用了边框格式，则 **Enable** 属性返回 **True** 或 **wdUndefined** 。可将该属性值设置为 **True** 、 **False** 或 **WdLineStyle** 常量。

**Enable** 属性应用于指定对象的所有边框。如果为 **True** ，则样式设置为默认样式，线条宽度设置为默认线条宽度。可以通过 **DefaultBorderLineWidth** 和 **DefaultBorderLineStyle** 属性来设置默认样式和默认线条宽度。

正如下例所示，将 **Enable** 属性的值设置为 **False** ，即可删除对象中的所有边框。

``` JavaScript
Application.ActiveDocument.Tables.Item(1).Borders.Enable = false
```

如果要删除或应用单一边框，可先用 **Borders** ( *index* ) 返回单一边框，然后设置 **LineStyle** 属性（其中 **index** 为 **WdBorderType** 常量）。下例删除了 ` `**`Selection`**` ` 中的下边框。

``` JavaScript
Applicaiton.Selection.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleNone
```

### Borders.EnableFirstPageInSection

如果可为节内第一页设置页面边框，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableFirstPageInSection**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/*本示例为选定内容的第一节的首页添加页面边框。*/
function test() {
    let Sec = Application.Selection.Sections.Item(1)
    Sec.Borders.EnableFirstPageInSection = true
    Sec.Borders.EnableOtherPagesInSection = false
    for(let i = 1; i <= Sec.Borders.Count; i++) {
        Sec.Borders.Item(i).ArtStyle = wdArtPeople
        Sec.Borders.Item(i).ArtWidth = 15
    }   
}
```

### Borders.EnableOtherPagesInSection

如果可为本节中除首页之外的所有页面设置页面边框，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableOtherPagesInSection**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/*本示例为选定内容的第一节除首页外的所有页面添加边框。*/
function test() {
    let Sec = Application.Selection.Sections.Item(1)
    Sec.Borders.EnableFirstPageInSection = false
    Sec.Borders.EnableOtherPagesInSection = true
    for(let i = 1; i <= Sec.Borders.Count; i++) {
        Sec.Borders.Item(i).ArtStyle = wdArtBabyRattle
        Sec.Borders.Item(i).ArtWidth = 22
    }
}
```

### Borders.HasHorizontal

如果可将水平方向的边框应用于对象，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.HasHorizontal**

\*express是一个代表 Borders 对象的变量。

**说明**

水平边框可以应用于包含两行或多行单元格的表格区域，或包含两个或多个段落的区域。

**示例**

``` JavaScript
/*如果选定内容支持水平方向的边框，则本示例为其设置单线型的水平边框。*/
function test() {
    if(Application.Selection.Borders.HasHorizontal == true) {
        Application.Selection.Borders.Item(wdBorderHorizontal).LineStyle = wdLineStyleSingle
    }
}
```

### Borders.HasVertical

如果可将垂直边框应用于指定对象，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.HasVertical**

\*express是一个代表 Borders 对象的变量。

**说明**

垂直边框可以应用于包含两列或多列单元格的表格区域。

**示例**

``` JavaScript
/*如果选定内容支持垂直边框，则本示例为其设置单线型的垂直边框。*/
function test() {
    if(Application.Selection.Borders.HasVertical == true) {
        Application.Selection.Borders.Item(wdBorderVertical).LineStyle = wdLineStyleSingle
    }
}
```

### Borders.InsideColor

返回或设置内部边框的 24 位颜色。可读写。

**语法**

**express.InsideColor**

\*express是一个代表 Borders 对象的变量。

**说明**

该属性可以是任何有效的 **WdColor** 常量或 Visual Basic 的 **RGB** 函数返回的值。如果将 **InsideLineStyle** 属性设置为 **wdLineStyleNone** 或 **False** ，则本属性的设置不产生任何效果。

**示例**

``` JavaScript
/*本示例在活动文档的前四个段落之间添加深红色边框。*/
function test() {
    let myDoc = Application.ActiveDocument
    let myRange = myDoc.Range(myDoc.Paragraphs.Item(1).Range.Start, myDoc.Paragraphs.Item(4).Range.End)
    let Bor = myRange.Borders
    Bor.InsideLineStyle = wdLineStyleSingle
    Bor.InsideLineWidth = wdLineWidth150pt
    Bor.InsideColor = wdDarkRed
}
```

### Borders.InsideColorIndex

返回或设置内部边框的颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.InsideColorIndex**

\*express是一个代表 Borders 对象的变量。

**说明**

如果将 **InsideLineStyle** 属性设置为 **wdLineStyleNone** 或 **False** ，则本属性的设置不产生任何效果。

**示例**

``` JavaScript
/*本示例在活动文档前四段之间添加红色边框。*/
function test() {
    let docActive = Application.ActiveDocument
    let rngTemp = docActive.Range(docActive.Paragraphs.Item(1).Range.Start, docActive.Paragraphs.Item(4).Range.End)
    let Bor = rngTemp.Borders
    Bor.InsideLineStyle = wdLineStyleSingle
    Bor.InsideLineWidth = wdLineWidth150pt
    Bor.InsideColorIndex = wdRed
}
```

### Borders.InsideLineStyle

返回或设置指定对象的内部边框。

**语法**

**express.InsideLineStyle**

\*express是一个代表 Borders 对象的变量。

**说明**

如果指定对象应用了多种边框，则该属性返回 **wdUndefined** ；否则，返回 **False** 或 **WdLineStyle** 常量。可将该属性值设置为 **True** 、 **False** 或 **WdLineStyle** 常量。

如果为 **True** ，则样式设置为默认样式，线条宽度设置为默认线条宽度。可用 **DefaultBorderLineWidth** 和 **DefaultBorderLineStyle** 属性设置默认样式和线条宽度。

可用下列任何一种指令删除活动文档中第一个表格的内部边框。

``` JavaScript
Application.ActiveDocument.Tables.Item(1).Borders.InsideLineStyle = wdLineStyleNone
Application.ActiveDocument.Tables.Item(1).Borders.InsideLineStyle = false
```

**示例**

``` JavaScript
/*本示例在活动文档的第一个表格中的行列之间添加边框。*/
function test() {
    if(Application.ActiveDocument.Tables.Count >= 1) {
        let tableTemp = ActiveDocument.Tables.Item(1)
        tableTemp.Borders.InsideLineStyle = true
    }
}

/*本示例在活动文档的前 4 个段落之间添加边框。*/
function test() {
    let docActive = Application.ActiveDocument
    let rngTemp = docActive .Range(docActive.Paragraphs.Item(1).Range.Start, docActive.Paragraphs.Item(4).Range.End)
    let Bor = rngTemp.Borders
    Bor.InsideLineStyle = wdLineStyleSingle
    Bor.InsideLineWidth = wdLineWidth150pt
}
```

### Borders.InsideLineWidth

返回或设置对象内部边框的线条宽度。

**语法**

**express.InsideLineWidth**

\*express是一个代表 Borders 对象的变量。

**说明**

如果对象的内部边框具有多个线条宽度，则该属性返回 **wdUndefined** ；否则，返回 **False** 或 **WdLineWidth** 常量。可将该属性值设置为 **True** 、 **False** 或下列 **WdLineWidth** 常量之一。

**示例**

``` JavaScript
/*本示例在活动文档的第一个表格中的行列之间添加边框。*/
function test() {
    if(Application.ActiveDocument.Tables.Count >= 1) {
        let tableTemp = Application.ActiveDocument.Tables.Item(1)
        tableTemp.Borders.InsideLineStyle = wdLineStyleDot
        tableTemp.Borders.InsideLineWidth = wdLineWidth050pt
    }
}

/*本示例在活动文档中的前 4 段之间添加点线型边框。*/
function test() {
    let docActive = Application.ActiveDocument
    let rngTemp = docActive.Range(docActive.Paragraphs.Item(1).Range.Start, docActive.Paragraphs.Item(4).Range.End)
    rngTemp.Borders.InsideLineStyle = wdLineStyleDot
    rngTemp.Borders.InsideLineWidth = wdLineWidth075pt
}
```

### Borders.JoinBorders

如果删除段落和表格边缘的垂直边框，以使水平边框与页面边框相连，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.JoinBorders**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/*本示例为所选内容的第一节中的每页添加边框。本示例还删除段落和表格边缘的垂直边框，以使水平边框能与页面边框相连。*/
function test() {
    let Sec = Application.Selection.Sections.Item(1)
    for(let i = 1; i <= Sec.Borders.Count; i++) {
        Sec.Borders.Item(i).ArtStyle = wdArtBalloonsHotAir
        Sec.Borders.Item(i).ArtWidth = 15
    }
    let Bor = Sec.Borders
    Bor.DistanceFromLeft = 2
    Bor.DistanceFromRight = 2
    Bor.DistanceFrom = wdBorderDistanceFromText
    Bor.JoinBorders = true
}
```

### Borders.OutsideColor

返回或设置外部边框的 24 位颜色。

**语法**

**express.OutsideColor**

\*express是一个代表 Borders 对象的变量。

**说明**

此属性可以是任何有效的 **WdColor** 常量，也可以是 Visual Basic 的 **RGB** 函数返回的值。

如果将 **OutsideLineStyle** 属性设置为 **wdLineStyleNone** 或 **False** ，则本属性的设置不会产生任何效果。

**示例**

``` JavaScript
/*本示例在活动文档的第一个表格的行列之间添加边框，然后设置内部和外部边框的颜色。*/
function test() {
    if(Application.ActiveDocument.Tables.Count >= 1) {
        let myTable = ActiveDocument.Tables.Item(1)
        let Bor = myTable.Borders
        Bor.InsideLineStyle = true
        Bor.InsideColor = wdColorBrightGreen
        Bor.OutsideColor = wdColorDarkTeal
    }
}

/*本示例在活动文档的第一段的周围添加深红色、0.75 磅宽的双边框。*/
function test() {
    let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders
    Bor.OutsideLineStyle = wdLineStyleDouble
    Bor.OutsideLineWidth = wdLineWidth075pt
    Bor.OutsideColor = wdColorDarkRed
}
```

### Borders.OutsideColorIndex

返回或设置外部边框的颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.OutsideColorIndex**

\*express是一个代表 Borders 对象的变量。

**说明**

如果将 **OutsideLineStyle** 属性设置为 **wdLineStyleNone** 或 **False** ，则该属性的设置不会产生任何效果。

**示例**

``` JavaScript
/*本示例在活动文档的第一个表格的行列之间添加边框，然后设置内部和外部边框的颜色。*/
function test() {
    if(Application.ActiveDocument.Tables.Count >= 1) {
        let myTable = Application.ActiveDocument.Tables.Item(1)
        let Bor = myTable.Borders
        Bor.InsideLineStyle = true
        Bor.InsideColorIndex = wdBrightGreen
        Bor.OutsideColorIndex = wdPink
    }
}

/*本示例为活动文档的第一段添加红色、0.75 磅的双边框。*/
function test() {
    let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders
    Bor.OutsideLineStyle = wdLineStyleDouble
    Bor.OutsideLineWidth = wdLineWidth075pt
    Bor.OutsideColorIndex = wdRed
}
```

### Borders.OutsideLineStyle

返回或设置指定对象的外部边框。

**语法**

**express.OutsideLineStyle**

\*express是一个代表 Borders 对象的变量。

**说明**

如果指定对象应用了多种边框，则该属性返回 **wdUndefined** ；否则，返回 **False** 或 **WdLineStyle** 常量。可将该属性值设置为 **True** 、 **False** 或 **WdLineStyle** 常量。

如果为 **True** ，则会将线条样式设置为默认线条样式，将线条宽度设置为默认线条宽度。可用 **DefaultBorderLineWidth** 和 **DefaultBorderLineStyle** 属性设置默认线条样式和宽度。

使用下列两条指令中的任何一条，都可删除活动文档中第一个表格的外部边框。

``` JavaScript
Application.ActiveDocument.Tables.Item(1).Borders.OutsideLineStyle = wdLineStyleNone
Application.ActiveDocument.Tables.Item(1).Borders.OutsideLineStyle = false
```

**示例**

``` JavaScript
/*本示例在活动文档的第一段的周围添加 0.75 磅宽的双边框。*/
function test() {
    let Bor = Application.ActiveDocument.Paragraphs.Item(1).Borders
    Bor.OutsideLineStyle = wdLineStyleDouble
    Bor.OutsideLineWidth = wdLineWidth075pt
}

/*本示例在活动文档的第一个表格周围添加边框。*/
function test() {
    if(Application.ActiveDocument.Tables.Count >= 1) {
        let myTable = Application.ActiveDocument.Tables.Item(1)
        myTable.Borders.OutsideLineStyle = wdLineStyleSingle
    }
}
```

### Borders.OutsideLineWidth

返回或设置对象外部边框的线条宽度。可读写。

**语法**

**express.OutsideLineWidth**

\*express是一个代表 Borders 对象的变量。

**说明**

如果对象的外部边框具有多个线条宽度，则该属性返回 **wdUndefined** ；否则，返回 **False** 或 **WdLineWidth** 常量。可将该属性值设置为 **True** 、 **False** 或 **WdLineWidth** 常量。

**示例**

``` JavaScript
/*本示例为活动文档的第一个表格添加波浪形边框。*/
function test() {
    if(Application.ActiveDocument.Tables.Count >= 1) {
        let Bor = Application.ActiveDocument.Tables.Item(1).Borders
        Bor.OutsideLineStyle = wdLineStyleSingleWavy
        Bor.OutsideLineWidth = wdLineWidth075pt
    }
}

/*本示例在活动文档的前 4 段四周添加点线型边框。*/
function test() {
    let myDoc = Application.ActiveDocument
    let myRange = myDoc.Range(myDoc.Paragraphs.Item(1).Range.Start, myDoc.Paragraphs.Item(4).Range.End)
    myRange.Borders.OutsideLineStyle = wdLineStyleDot
    myRange.Borders.OutsideLineWidth = wdLineWidth075pt
}
```

### Borders.Parent

返回一个 **Object** 类型的值，该值代表指定 **Borders** 集合中的父对象。

**语法**

**express.Parent**

\*express是一个代表 Borders 对象的变量。

### Borders.Shadow

如果为 **True** ，则指定的边框设置为阴影格式。 **Boolean** 类型，可读写。

**语法**

**express.Shadow**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/*本示例在新文档中演示两种不同的边框样式。*/
function test() {
    let myRange = Application.Documents.Add().Content
    myRange.InsertAfter("Demonstration of border with shadow.")
    myRange.InsertParagraphAfter()
    myRange.InsertParagraphAfter()
    myRange.InsertAfter("Demonstration of border without shadow.")
    ActiveDocument.Paragraphs.Item(1).Borders.Shadow = true
    ActiveDocument.Paragraphs.Item(3).Borders.Enable = true
}
```

### Borders.SurroundFooter

如果页面边框环绕文档的页脚，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SurroundFooter**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动文档中第一节的页面边框格式，使页面边框环绕该节中每个页面的页眉和页脚。*/
function test() {
    let Bor = Application.ActiveDocument.Sections.Item(1).Borders
    Bor.SurroundFooter = true
    Bor.SurroundHeader = true
}

/*本示例为第一节中的每一页添加艺术型页面边框。该页面边框不环绕页眉和页脚。*/
function test() {
    let Sec = Application.ActiveDocument.Sections.Item(1)
    Sec.Borders.SurroundFooter = false
    Sec.Borders.SurroundHeader = false
    for(let i = 1; i <= Sec.Borders.Count; i++) {
        Sec.Borders.Item(i).ArtStyle = wdArtPeople
        Sec.Borders.Item(i).ArtWidth = 15
    }
}
```

### Borders.SurroundHeader

如果页面边框环绕文档页眉，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SurroundHeader**

\*express是一个代表 Borders 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动文档中第一节的页面边框格式，使页面边框不环绕页眉和页脚。*/
function test() {
    let Bor = Application.ActiveDocument.Sections.Item(1).Borders
    Bor.SurroundFooter = false
    Bor.SurroundHeader = false
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Borders](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CoAuthors 对象

## 简介

文档中所有 CoAuthor 对象的集合。

## 说明

CoAuthors 集合包含文档中的所有共同作者（当前正在编辑文档的作者）。

``` JavaScript
/*下面的代码示例获取活动文档中的共同作者的数量。*/
function test()
{
let i
i = ActiveDocument.CoAuthoring.Authors.Count
MsgBox("The number of co-authors is " + i)
}
```

## 属性

| 序号 | 名称                                  | 说明                                                             |
|:----:|:--------------------------------------|:-----------------------------------------------------------------|
|  1   | [Application](#CoAuthors.Application) | 返回一个代表 WPS 应用程序的 Application 对象。只读。             |
|  2   | [Count](#CoAuthors.Count)             | 返回 CoAuthors 集合中的项目数。只读。                            |
|  3   | [Creator](#CoAuthors.Creator)         | 返回表示用于创建指定对象的应用程序的 32 位整数。只读 Long 类型。 |
|  4   | [Parent](#CoAuthors.Parent)           | 返回一个 Object 类型值，该值代表指定 CoAuthors 对象的父对象。    |

## 成员属性

### CoAuthors.Application

返回一个代表 WPS 应用程序的 Application 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CoAuthors 对象的变量。

### CoAuthors.Count

返回 CoAuthors 集合中的项目数。只读。

**语法**

**express.Count**

\*express是一个代表 CoAuthors 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示活动文档中的共同作者的数量。*/
MsgBox("The active document contains " + ActiveDocument.CoAuthoring.Authors.Count + " authors.")
```

### CoAuthors.Creator

返回表示用于创建指定对象的应用程序的 32 位整数。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CoAuthors 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表 **string** “WPS”。该属性主要设计用于 Apple Macintosh 平台，在该平台上，每个应用程序都有一个由四个字符组成的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的详细信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>该值也可用常量 <strong>wdCreatorCode</strong> 表示。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

### CoAuthors.Parent

返回一个 **Object** 类型值，该值代表指定 **CoAuthors** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 CoAuthors 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/CoAuthors](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Columns 对象

## 简介

Column 对象的集合，这些对象代表表格中的列。

## 方法

| 序号 | 名称                                        | 说明                                                               |
|:----:|:--------------------------------------------|:-------------------------------------------------------------------|
|  1   | [Add](#Columns.Add)                         | 返回一个 Column 对象，该对象代表添加到表中的列。                   |
|  2   | [AutoFit](#Columns.AutoFit)                 | 改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。 |
|  3   | [Delete](#Columns.Delete)                   | 删除指定列。                                                       |
|  4   | [DistributeWidth](#Columns.DistributeWidth) | 将指定列调整为等宽。                                               |
|  5   | [Item](#Columns.Item)                       | 返回集合中的单个 Column 对象。                                     |
|  6   | [Select](#Columns.Select)                   | 选择指定表格列。                                                   |
|  7   | [SetWidth](#Columns.SetWidth)               | 设置表格列的宽度。                                                 |

## 属性

| 序号 | 名称                                              | 说明                                                                                     |
|:----:|:--------------------------------------------------|:-----------------------------------------------------------------------------------------|
|  1   | [Application](#Columns.Application)               | 返回一个代表 WPS 应用程序的 Application 对象。                                           |
|  2   | [Borders](#Columns.Borders)                       | 返回一个 Borders 集合，该集合代表指定列的所有边框。                                      |
|  3   | [Count](#Columns.Count)                           | 返回一个 Number 类型的值，该值代表集合中列的数目。只读。                                 |
|  4   | [First](#Columns.First)                           | 返回一个 Column 对象，该对象代表在 Columns 集合中的第一个项目。                          |
|  5   | [Last](#Columns.Last)                             | 返回表示表格中最后一列的 Column 对象。                                                   |
|  6   | [NestingLevel](#Columns.NestingLevel)             | 返回指定列的嵌套层。只读 Number 类型。                                                   |
|  7   | [Parent](#Columns.Parent)                         | 返回一个 Object 类型值，该值代表指定 Columns 对象的父对象。                              |
|  8   | [PreferredWidth](#Columns.PreferredWidth)         | 返回或设置指定列的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 Single 类型，可读写。 |
|  9   | [PreferredWidthType](#Columns.PreferredWidthType) | 返回或设置用于指定单元格、列或表格的首选度量单位。可读/写 WdPreferredWidthType 类型。    |
|  10  | [Shading](#Columns.Shading)                       | 返回一个表示指定表格列的底纹格式的 Shading 对象。                                        |
|  11  | [Width](#Columns.Width)                           | 返回或设置指定列的宽度，以磅为单位。可读/写 Long 类型。                                  |

## 成员方法

### Columns.Add

返回一个 **Column** 对象，该对象代表添加到表中的列。

**语法**

**express.Add ( *BeforeColumn* )**

\*express是一个代表 Columns 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                       |
|----------------|-----------|----------|--------------------------------------------|
| *BeforeColumn* | 可选      | Variant  | 代表将会直接显示在新列右侧的 Column 对象。 |

**示例**

``` JavaScript
/* 本示例在活动文档中创建一个有两行两列的表格，然后在第一列之前添加一列。新列的宽度设为 1.5 英寸 */
function test() {
    let myTable = Application.ActiveDocument.Tables.Add(Application.Selection.Range, 2, 2)
    let newCol = myTable.Columns.Add(myTable.Columns.Item(1))
    newCol.SetWidth(Application.InchesToPoints(1.5), wdAdjustNone)
}
```

### Columns.AutoFit

改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。

**语法**

**express.AutoFit ()**

\*express是一个代表 Columns 对象的变量。

**说明**

如果表格的宽度已等于从左边界到右边界的距离，则此方法无效。

**示例**

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，然后调整第一列的宽度，使之与文本的宽度相称。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
tableNew.Cell(1,1).Range.InsertAfter("First cell")
tableNew.Columns.Item(1).AutoFit()
}
```

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，然后调整所有列的宽度，使之与文本的宽度相称。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
tableNew.Cell(1,1).Range.InsertAfter("First cell")
tableNew.Cell(1,2).Range.InsertAfter("This is cell (1,2)")
tableNew.Cell(1,3).Range.InsertAfter("(1,3)")
tableNew.Columns.AutoFit()
}
```

### Columns.Delete

删除指定列。

**语法**

**express.Delete ()**

\*express是一个代表 Columns 对象的变量。

### Columns.DistributeWidth

将指定列调整为等宽。

**语法**

**express.DistributeWidth ()**

\*express是一个代表 Columns 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中第一个表格的各列调整为等宽。*/
Application.ActiveDocument.Tables.Item(1).Columns.DistributeWidth()
```

``` JavaScript
/*本示例调整选定单元格的宽度。*/
function test() {
    if (Application.Selection.Cells.Count >= 2) {
        Application.Selection.Cells.DistributeWidth()
    }
}
```

### Columns.Item

返回集合中的单个 **Column** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Columns 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                           |
|---------|-----------|----------|----------------------------------------------------------------|
| *Index* | 可选      | Number   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

### Columns.Select

选择指定表格列。

**语法**

**express.Select ()**

\*express是一个代表 Columns 对象的变量。

**说明**

使用该方法后，可用 **Selection** 对象来处理选定项目。有关详细信息，请参阅 处理 Selection 对象。

### Columns.SetWidth

设置表格列的宽度。

**语法**

**express.SetWidth ( *ColumnWidth* , *RulerStyle* )**

\*express是一个代表 Columns 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型     | 说明                            |
|---------------|-----------|--------------|---------------------------------|
| *ColumnWidth* | 必选      | Single       | 指定列的宽度，以磅为单位。      |
| *RulerStyle*  | 必选      | WdRulerStyle | 控制 WPS 调整单元格宽度的方式。 |

**说明**

上述 **WdRulerStyle** 行为应用于左对齐的表格。 **WdRulerStyle** 行为应用于中对齐和右对齐的表格时可能会出现未知效果，因此应谨慎使用 **SetWidth** 方法。

## 成员属性

### Columns.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Columns 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Columns.Borders

返回一个 **Borders** 集合，该集合代表指定列的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Columns 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Columns.Count

返回一个 **Number** 类型的值，该值代表集合中列的数目。只读。

**语法**

**express.Count**

\*express是一个代表 Columns 对象的变量。

### Columns.First

返回一个 **Column** 对象，该对象代表在 **Columns** 集合中的第一个项目。

**语法**

**express.First**

\*express是一个代表 Columns 对象的变量。

### Columns.Last

返回表示表格中最后一列的 **Column** 对象。

**语法**

**express.Last**

\*express是一个代表 Columns 对象的变量。

### Columns.NestingLevel

返回指定列的嵌套层。只读 **Number** 类型。

**语法**

**express.NestingLevel**

\*express是一个代表 Columns 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

### Columns.Parent

返回一个 **Object** 类型值，该值代表指定 **Columns** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Columns 对象的变量。

### Columns.PreferredWidth

返回或设置指定列的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 **Single** 类型，可读写。

**语法**

**express.PreferredWidth**

\*express是一个代表 Columns 对象的变量。

**说明**

如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPoints** ，则 **PreferredWidth** 属性将返回或设置以磅为单位的宽度值；如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent** ，则 **PreferredWidth** 属性将返回或设置用窗口宽度百分比表示的宽度值。

### Columns.PreferredWidthType

返回或设置用于指定单元格、列或表格的首选度量单位。可读/写 **WdPreferredWidthType** 类型。

**语法**

**express.PreferredWidthType**

\*express是一个代表 Columns 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 以窗口宽度的百分比来接受宽度，然后将活动文档的第一个表格的所有列宽设置为窗口宽度的 50%。*/
function test() {
    let rng = Application.ActiveDocument.Tables.Item(1).Columns
    rng.PreferredWidthType = wdPreferredWidthPercent
    rng.PreferredWidth = 50
}
```

### Columns.Shading

返回一个表示指定表格列的底纹格式的 **Shading** 对象。

**语法**

**express.Shading**

\*express是一个代表 Columns 对象的变量。

**示例**

``` JavaScript
/*本示例对活动文档中第一个表格的所有列应用水平线纹理。*/
function test() {
    if (Application.ActiveDocument.Tables.Count >= 1) {
        Application.ActiveDocument.Tables.Item(1).Columns.Shading.Texture = wdTextureDiagonalDown
    }
}
```

### Columns.Width

返回或设置指定列的宽度，以磅为单位。可读/写 **Long** 类型。

**语法**

**express.Width**

\*express是一个代表 Columns 对象的变量。

**示例**

``` JavaScript
/*本示例在新文档中创建一张 5x5 的表格，然后将该表格中的所有列宽设为 1.5 英寸。*/
function test() {
    let objDoc = Application.ActiveDocument
    let objTable = objDoc.Tables.Add(Selection.Range, 5, 5)
    objTable.Columns.Width = InchesToPoints(1.5)
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Columns](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FileConverter 对象

## 简介

代表用于打开或保存文件的文件转换器。 FileConverter 对象是 FileConverters 集合的一个成员。 FileConverters 集合包含所有已安装的用于打开和保存文件的文件转换器。

## 说明

使用 FileConverters ( *Index* ) 可返回一个 FileConverter 对象，其中 *Index* 是类别名称或索引号。以下示例显示与 ET 工作表转换器相关联的扩展名。

``` JavaScript
MsgBox(FileConverters.Item("MSBiff").Extensions)
```

索引号代表文件转换器在 FileConverters 集合中的位置。以下示例显示第一个文件转换器的格式名称。

``` JavaScript
MsgBox(FileConverters.Item(1).FormatName)
```

不能新建一个文件转换器或向 FileConverters 集合添加一个文件转换器。 FileConverter 对象是在安装 WPS Office的过程中或通过安装补充文件转换器添加的。使用 CanSave 或 CanOpen 属性可确定 FileConverter 对象是否可用于打开或保存文档。

用于保存文档的文件转换器列在 “另存为” 对话框中。如果选中了 “选项” 对话框（ “工具” 菜单）中 “常规” 选项卡上的 “打开时确认转换” 复选框，则用于打开文档的文件转换器显示在一个对话框中。

## 属性

| 序号 | 名称                                      | 说明                                                                          |
|:----:|:------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#FileConverter.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                |
|  2   | [CanOpen](#FileConverter.CanOpen)         | 如果指定的文件转换器可用来打开文件，则该属性值为 True 。 Boolean 类型，只读。 |
|  3   | [CanSave](#FileConverter.CanSave)         | 如果指定的文件转换器可用来保存文件，则该属性值为 True 。 Boolean 类型，只读。 |
|  4   | [ClassName](#FileConverter.ClassName)     | 返回用来标识文件转换器的唯一名字。 String 类型，只读。                        |
|  5   | [Creator](#FileConverter.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。  |
|  6   | [Extensions](#FileConverter.Extensions)   | 返回与指定 FileConverter 对象关联的文件扩展名。 String 类型，只读。           |
|  7   | [FormatName](#FileConverter.FormatName)   | 返回指定文件转换器的名称。 String 类型，只读。                                |
|  8   | [Name](#FileConverter.Name)               | 返回指定对象的名称。 String 类型，只读。                                      |
|  9   | [OpenFormat](#FileConverter.OpenFormat)   | 返回指定文件转换器的文件格式。 Long 类型，只读。                              |
|  10  | [Parent](#FileConverter.Parent)           | 返回一个 Object 类型值，该值代表指定 FileConverter 对象的父对象。             |
|  11  | [Path](#FileConverter.Path)               | 返回指定对象的磁盘或 Web 路径。只读 String 类型。                             |
|  12  | [SaveFormat](#FileConverter.SaveFormat)   | 返回指定文档或文件转换器的文件格式。只读 Long 类型。                          |

## 成员属性

### FileConverter.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 FileConverter 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### FileConverter.CanOpen

如果指定的文件转换器可用来打开文件，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.CanOpen**

\*express是一个代表 FileConverter 对象的变量。

**说明**

如果 **CanSave** 属性返回 **True** ，则可以用指定的文件转换器保存（导出）文件。

**示例**

``` JavaScript
/*本示例判断能否用第一个文件转换器打开文件。*/
function test()
{
if(FileConverters.Item(1).CanOpen == true) {
    MsgBox(FileConverters.Item(1).FormatName + " can open files")
}
}
```

``` JavaScript
/*本示例判断能否用 WordPerfect6x 文件转换器打开文件。如果 CanOpen 属性返回 True，则打开名为“Test.wp”的文档。*/
function test()
{
if(FileConverters.Item("WordPerfect6x").CanOpen == true) {
    Documents.Open("C:\\Test.wp",null,null,null,null,null,null,null,null,FileConverters.Item("WordPerfect6x").OpenFormat)
}
}
```

### FileConverter.CanSave

如果指定的文件转换器可用来保存文件，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.CanSave**

\*express是一个代表 FileConverter 对象的变量。

**说明**

如果可以使用指定的文件转换器打开（导入）文件，则 **CanOpen** 属性返回 **True** 。

**示例**

``` JavaScript
/*本示例判断能否用 WordPerfect 转换器来保存文件。如果返回值为 True，则以 WordPerfect 6.x 格式保存活动文档。*/
function test()
{
if(Application.FileConverters.Item("WordPerfect6x").CanSave == true) {
    let lngSaveFormat = Application.FileConverters.Item("WordPerfect6x").SaveFormat
    ActiveDocument.SaveAs("C:\\Document.wp", lngSaveFormat)
}
}
```

``` JavaScript
/*本示例显示一条消息，以表明 FileConverters 集合的第三个转换器能否用来保存文件。*/
function test()
{
if(FileConverters.Item(3).CanSave == true) {
    MsgBox(FileConverters.Item(3).FormatName + " can save files")
}
else {
    MsgBox(FileConverters.Item(3).FormatName + " cannot save files")
}
}
```

### FileConverter.ClassName

返回用来标识文件转换器的唯一名字。 **String** 类型，只读。

**语法**

**express.ClassName**

\*express是一个代表 FileConverter 对象的变量。

**示例**

------------------------------------------------------------------------

``` JavaScript
/*本示例显示 FileConverters 集合中第一个转换器的类名和格式名。*/
MsgBox("ClassName= " + FileConverters.Item(1).ClassName + "\r" + "FormatName= " + FileConverters.Item(1).FormatName)
```

``` JavaScript
/*如果 HTML 文件转换器有效，本示例将 HTML 设为默认保存格式。*/
function test()
{
for(let i=1;i <= FileConverters.Count;i++) {
    if(FileConverters.Item(i).ClassName == "HTML") {
        Application.DefaultSaveFormat = "HTML"
    }
}
}
```

### FileConverter.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FileConverter 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### FileConverter.Extensions

返回与指定 **FileConverter** 对象关联的文件扩展名。 **String** 类型，只读。

**语法**

**express.Extensions**

\*express是一个代表 FileConverter 对象的变量。

**示例**

``` JavaScript
/*本示例显示第一个文件转换器的名称和文件扩展名。*/
function test()
{
let fcTemp = FileConverters.Item(1)
MsgBox("The file name extensions for " + fcTemp.FormatName + " files are: " + fcTemp.Extensions)
}
```

### FileConverter.FormatName

返回指定文件转换器的名称。 **String** 类型，只读。

**语法**

**express.FormatName**

\*express是一个代表 FileConverter 对象的变量。

**说明**

格式名称显示在 **“文件”** 菜单 **“另存为”** 对话框中的 **“保存类型”** 框中。

**示例**

``` JavaScript
/*本示例显示 FileConverters 集合中的第一个文件转换器的格式名称。*/
MsgBox(FileConverters.Item(1).FormatName)
```

``` JavaScript
/*本示例用 AvailableConv() 数组保存所有有效文件转换器的名称。*/
function test()
{
let AvailableConv = []
let intTemp = 0
for(let i=1;i <= FileConverters.Count;i++) {
    AvailableConv[intTemp] = FileConverters.Item(i).FormatName
    intTemp++
}
}
```

### FileConverter.Name

返回指定对象的名称。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 FileConverter 对象的变量。

### FileConverter.OpenFormat

返回指定文件转换器的文件格式。 **Long** 类型，只读。

**语法**

**express.OpenFormat**

\*express是一个代表 FileConverter 对象的变量。

**说明**

此属性可以是任何有效的 **WdOpenFormat** 常量，也可以是代表外部文件转换器的唯一数字。

**示例**

``` JavaScript
/*本示例显示可用来打开文档的转换器的唯一的格式值和格式名称。*/
function test()
{
for(let i=1;i <= FileConverters.Count;i++) {
    if(FileConverters.Item(i).CanOpen == true) {
        MsgBox(FileConverters.Item(i).OpenFormat + "\r" + FileConverters.Item(i).FormatName)
    }
}
}
```

``` JavaScript
/*本示例使用 WordPerfect 6x 文件转换器打开名为“Data.wp”的文件。*/
Documents.Open("C:\\Data.wp",null,null,null,null,null,null,null,null,FileConverters.Item("WordPerfect6x").OpenFormat)
```

### FileConverter.Parent

返回一个 **Object** 类型值，该值代表指定 **FileConverter** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 FileConverter 对象的变量。

### FileConverter.Path

返回指定对象的磁盘或 Web 路径。只读 **String** 类型。

**语法**

**express.Path**

\*express是一个代表 FileConverter 对象的变量。

**说明**

该路径不包括尾部字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符，使用 **Name** 属性可返回不带路径的文件名。您可以通过将 **Path** 、 **PathSeparator** 和 **Name** 属性连接在一起来创建文件转换器的完整名称。

### FileConverter.SaveFormat

返回指定文档或文件转换器的文件格式。只读 **Long** 类型。

**语法**

**express.SaveFormat**

\*express是一个代表 FileConverter 对象的变量。

**说明**

该属性返回一个指定外部文件转换器的唯一数字或一个 **WdSaveFormat** 常量。使用 **SaveAs** 方法的 *FileFormat* 参数的 **SaveFormat** 属性值可将文档保存为没有相应的 **WdSaveFormat** 常量的文件格式。

**示例**

``` JavaScript
/*以下示例新建一篇文档并在表格中列出可用于保存文档的转换器及其相应的 SaveFormat 值。*/
function FileConverterList() {
    //Create a new document and set a tab stop
    let docNew = Documents.Add()
    docNew.Paragraphs.Format.TabStops.Add(InchesToPoints.Item(3))
    //List all the converters in the FileConverters collection
    let Con = docNew.Content
    Con.InsertAfter("Name" + "\t" + "Number")
    Con.InsertParagraphAfter()
    for(let i=1;i <= FileConverters.Count;i++) {
        if(FileConverters.Item(i).CanSave == true) {
            Con.InsertAfter(FileConverters.Item(i).FormatName + "\t" + FileConverters.Item(i).SaveFormat)
            Con.InsertParagraphAfter()
        }
    }
    Con.ConvertToTable()
}
```

``` JavaScript
/*以下示例将活动文档保存为 WordPerfect 5.1 或 5.2 二级文件格式。*/
ActiveDocument.SaveAs(null,FileConverters.Item("WrdPrfctDat").SaveFormat)
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/FileConverter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# GroupShapes 对象

## 简介

代表组合形状中的单个形状。组合形状中包含的每个形状都由一个 Shape 对象表示。

## 方法

| 序号 | 名称                        | 说明                          |
|:----:|:----------------------------|:------------------------------|
|  1   | [Item](#GroupShapes.Item)   | 返回集合中的单个 Shape 对象。 |
|  2   | [Range](#GroupShapes.Range) | 返回一个 ShapeRange 对象。    |

## 属性

| 序号 | 名称                                    | 说明                                                                         |
|:----:|:----------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#GroupShapes.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#GroupShapes.Count)             | 返回一个 Long 类型的值，该值代表集合中图形的数量。只读。                     |
|  3   | [Creator](#GroupShapes.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#GroupShapes.Parent)           | 返回一个 Object 类型值，该值代表指定 GroupShapes 对象的父对象。              |

## 成员方法

### GroupShapes.Item

返回集合中的单个 **Shape** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 GroupShapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

### GroupShapes.Range

返回一个 **ShapeRange** 对象。

**语法**

**express.Range ( *Index* )**

\*express是一个代表 GroupShapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                         |
|---------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 指定要在指定区域内包含哪些形状。可以是指定 Shapes 集合内某个形状的索引号的整数，也可以是指定形状名称的字符串，或者是包含整数或字符串的数组。 |

**说明**

**ShapeRange** 对象不包含 **InlineShape** 对象。一个 **InlineShape** 对象相当于一个字符，并像字符一样位于文本区域中。 **Shape** 对象定位到文本区域（默认情况下为所选内容），但对象可以位于页面上的任何位置。 **Shape** 对象总是与定位的区域位于同一页上。

对 **Shape** 对象进行的大部分操作也适用于只包含一个形状的 **ShapeRange** 对象。有些操作如果用于包含多个形状的 **ShapeRange** 对象，则会产生错误。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个形状的填充前景色设置为紫色。*/
function test(){
    let fill = Application.ActiveDocument.Shapes.Range(1).Fill
    fill.ForeColor.RGB = (255, 0, 255)
    fill.Visible = msoTrue
}

/*以下示例为活动文档中变量形状应用阴影。*/
function test(strShpName){
    Application.ActiveDocument.Shapes.Range(strShpName).Shadow.Type = msoShadow6
}

/*要调用前一个子程序，请在标准代码模块中输入以下代码。*/
function test(){
    let shpArrow = Application.ActiveDocument.Shapes.AddShape(msoShapeLeftArrow,200,400,50,75)
    shpArrow.Name = "myShape"
    ShpRange2(shpArrow.Name)
}

/*以下示例选择活动文档的第一和第三个形状。*/
function test(){
    Application.ActiveDocument.Shapes.Range([1, 3]).Select()
}

/*以下示例选择并删除活动文档的第一个形状中的形状。本示例假定第一个形状是画布形状。*/
function CanvasShapeRange(){
    let rngCanvasShapes = Application.ActiveDocument.Shapes.Item(1).CanvasItems.Range(1)
        rngCanvasShapes.Select()
        rngCanvasShapes.Delete()
}
```

## 成员属性

### GroupShapes.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 GroupShapes 对象的变量。

### GroupShapes.Count

返回一个 **Long** 类型的值，该值代表集合中图形的数量。只读。

**语法**

**express.Count**

\*express是一个代表 GroupShapes 对象的变量。

### GroupShapes.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 GroupShapes 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### GroupShapes.Parent

返回一个 **Object** 类型值，该值代表指定 **GroupShapes** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 GroupShapes 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/GroupShapes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# InlineShapes 对象

## 简介

InlineShape 对象的集合，这些对象代表文档、范围或所选内容中的所有内嵌形状。

## 方法

| 序号 | 名称                                                                 | 说明                                                                                                                 |
|:----:|:---------------------------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------|
|  1   | [AddChart](#InlineShapes.AddChart)                                   | 将指定类型的图表作为内嵌形状插入到活动文档中，并且在ET 中打开包含 WPS 用于创建该图表的默认数据的工作表。             |
|  2   | [AddHorizontalLine](#InlineShapes.AddHorizontalLine)                 | 向当前文档添加一条基于图像文件的横线。                                                                               |
|  3   | [AddHorizontalLineStandard](#InlineShapes.AddHorizontalLineStandard) | 向当前文档添加一条横线。                                                                                             |
|  4   | [AddOLEControl](#InlineShapes.AddOLEControl)                         | 创建一个 ActiveX 控件（以前称为 OLE 控件）。返回代表新 ActiveX 控件的 InlineShape 对象。                             |
|  5   | [AddOLEObject](#InlineShapes.AddOLEObject)                           | 创建一个 OLE 对象。返回代表新 OLE 对象的 InlineShape 对象。                                                          |
|  6   | [AddPicture](#InlineShapes.AddPicture)                               | 在文档中添加一幅图片。返回一个代表该图片的 InlineShape 对象。                                                        |
|  7   | [AddPictureBullet](#InlineShapes.AddPictureBullet)                   | 向当前文档添加一个基于图像文件的图片项目符号。返回一个 InlineShape 对象。                                            |
|  8   | [AddSmartArt](#InlineShapes.AddSmartArt)                             | 将 SmartArt 图形作为内嵌形状插入活动文档。                                                                           |
|  9   | [Item](#InlineShapes.Item)                                           | 返回集合中的单个 InlineShape 对象。                                                                                  |
|  10  | [New](#InlineShapes.New)                                             | 插入一个空白的 WPS 图片对象，该对象是边长为一英寸的正方形，四周带有边框。此方法以 InlineShape 对象的形式返回新图形。 |

## 属性

| 序号 | 名称                                     | 说明                                                                         |
|:----:|:-----------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#InlineShapes.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#InlineShapes.Count)             | 返回一个 Long 类型的值，该值代表集合中内嵌图形的数目。只读。                 |
|  3   | [Creator](#InlineShapes.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#InlineShapes.Parent)           | 返回一个 Object 类型值，该值代表指定 InlineShapes 对象的父对象。             |

## 成员方法

### InlineShapes.AddChart

将指定类型的图表作为内嵌形状插入到活动文档中，并且在ET 中打开包含 WPS 用于创建该图表的默认数据的工作表。

**语法**

**express.AddChart ( *Type* , *Range* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型    | 说明                                                                                                                                                       |
|---------|-----------|-------------|------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Type*  | 必选      | XlChartType | 指定要创建的图表的类型。                                                                                                                                   |
| *Range* | 必选      | Variant     | 指定图表所绑定到的文本。如果指定了 Range，则会将图表放置在该区域中第一段的开头。如果省略该参数，则自动选择该区域，并且相对于页面的上边缘和左边缘放置图表。 |

**返回值**

InlineShape

**示例**

``` JavaScript
/*创建一个新的三维柱形图，然后将其添加到活动文档中。*/
ActiveDocument.InlineShapes.AddChart(xl3DColumn)
```

### InlineShapes.AddHorizontalLine

向当前文档添加一条基于图像文件的横线。

**语法**

**express.AddHorizontalLine ( *FileName* , *Range* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                  |
|------------|-----------|----------|---------------------------------------------------------------------------------------|
| *FileName* | 必选      | String   | 要用于横线的图像的文件名。                                                            |
| *Range*    | 可选      | Variant  | WPS 在其上方放置横线的区域。如果省略该参数，则 WPS 会将横线放置在当前选定内容的上方。 |

**说明**

要添加不基于现有图像文件的横线，可用 **AddHorizontalLineStandard** 方法。

**示例**

``` JavaScript
/*本示例将一条基于图像文件“ArtsyRule.gif”的横线添加到活动文档中当前选定内容的上方。*/
Selection.InlineShapes.AddHorizontalLine("C:\\Art files\\ArtsyRule.gif")
 
```

### InlineShapes.AddHorizontalLineStandard

向当前文档添加一条横线。

**语法**

**express.AddHorizontalLineStandard ( *Range* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                  |
|---------|-----------|----------|---------------------------------------------------------------------------------------|
| *Range* | 可选      | Variant  | WPS 在其上方放置横线的区域。如果省略该参数，则 WPS 会将横线放置在当前选定内容的上方。 |

**说明**

要添加基于现有图像文件的横线，可用 **AddHorizontalLine** 方法。

**示例**

``` JavaScript
/*本示例将一条横线添加到活动文档第五段的上方。*/
ActiveDocument.Paragraphs.Item(5).Range.InlineShapes.AddHorizontalLineStandard()
 
```

### InlineShapes.AddOLEControl

创建一个 ActiveX 控件（以前称为 OLE 控件）。返回代表新 ActiveX 控件的 **InlineShape** 对象。

**语法**

**express.AddOLEControl ( *ClassType* , *Range* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                   |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------|
| *ClassType* | 可选      | Variant  | 用于要创建的 ActiveX 控件的编程标识符。                                                                                                |
| *Range*     | 可选      | Variant  | 一个区域，ActiveX 控件将放置在该区域的文本中。如果该区域未折叠，则 ActiveX 控件将替换该区域。如果省略该参数，则自动放置 ActiveX 控件。 |

**说明**

在 WPS 中，将 ActiveX 控件表示为 **Shape** 对象或 **InlineShape** 对象。若要修改 ActiveX 控件的属性，可以对指定的图形或嵌入图形使用 **OLEFormat** 对象的 **Object** 属性。

关于 ActiveX 控件类型的可用信息，请参阅 OLE 编程标识符

### InlineShapes.AddOLEObject

创建一个 OLE 对象。返回代表新 OLE 对象的 **InlineShape** 对象。

**语法**

**express.AddOLEObject ( *ClassType* , *FileName* , *LinkToFile* , *DisplayAsIcon* , *IconFileName* , *IconIndex* , *IconLabel* , *Range* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                 |
|-----------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ClassType*     | 可选      | Variant  | 用于激活指定 OLE 对象的应用程序名称。                                                                                                                                                                                                                                                                |
| *FileName*      | 可选      | Variant  | 创建对象所用的文件。如果省略该参数，则使用当前的文件夹。必须指定对象的 ClassType 或 FileName 参数，但不能同时指定两者。                                                                                                                                                                              |
| *LinkToFile*    | 可选      | Variant  | 如果为 True，则将 OLE 对象链接到创建该对象所用的文件。如果为 False，则 OLE 对象成为其源文件的独立副本。如果为 ClassType 指定了值，则 LinkToFile 参数必须为 False。默认值为 False。                                                                                                                   |
| *DisplayAsIcon* | 可选      | Variant  | 如果为 True，则将 OLE 对象显示为图标。默认值为 False。                                                                                                                                                                                                                                               |
| *IconFileName*  | 可选      | Variant  | 要显示的图标所在的文件。                                                                                                                                                                                                                                                                             |
| *IconIndex*     | 可选      | Variant  | IconFileName 中图标的索引号。当选中“显示为图标”复选框时，在指定文件中图标的顺序对应于“插入”菜单“对象”对话框的“更改图标”对话框中图标出现的顺序。文件中的第一个图标的索引号为 0（零）。如果指定索引号的图标在 IconFileName 中不存在，则使用索引号为 1 的图标（文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选      | Variant  | 显示在图标下面的标签（题注）。                                                                                                                                                                                                                                                                       |
| *Range*         | 可选      | Variant  | 指定 OLE 对象放置在文本中的区域。除非区域是折叠的，否则 OLE 对象将替换该区域。如果省略该参数，则自动放置对象。                                                                                                                                                                                       |

**示例**

``` JavaScript
/*本示例在活动文档的第二段中添加一张新的 ET 工作表。*/
function test()
{
ActiveDocument.InlineShapes.AddOLEObject("Excel.Sheet", null, null, false,
    null, null, null, ActiveDocument.Paragraphs.Item(2).Range)
}
```

### InlineShapes.AddPicture

在文档中添加一幅图片。返回一个代表该图片的 InlineShape 对象。

**语法**

**express.AddPicture ( *FileName* , *LinkToFile* , *SaveWithDocument* , *Range* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                                         |
|--------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------|
| *FileName*         | 必选      | String   | 图片的路径和文件名。                                                                                         |
| *LinkToFile*       | 可选      | Variant  | 如果为 True，则将图片链接到创建它的文件；如果为 False，则使图片成为该文件的独立副本。默认值为 False。        |
| *SaveWithDocument* | 可选      | Variant  | 如果为 True，则将链接的图片与文档一起保存。默认值为 False。                                                  |
| *Range*            | 可选      | Variant  | 图片置于文本中的位置。如果该区域未折叠，那么图片将覆盖此区域，否则插入图片。如果省略此参数，则自动放置图片。 |

### InlineShapes.AddPictureBullet

向当前文档添加一个基于图像文件的图片项目符号。返回一个 **InlineShape** 对象。

**语法**

**express.AddPictureBullet ( *FileName* , *Range* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                               |
|------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName* | 必选      | String   | 要用于图片项目符号的图像的文件名称。                                                                                                               |
| *Range*    | 可选      | Variant  | WPS 向其添加图片项目符号的区域。 WPS 会将图片项目符号添加到该区域的每个段落。如果省略该参数，则 WPS 将图片项目符号添加到当前选定内容的每个段落中。 |

**返回值**

InlineShape

**示例**

``` JavaScript
/*本示例使用名为“RedBullet.gif”的图像文件向选定文本的每个段落添加一个图片项目符号。*/
Selection.InlineShapes.AddPictureBullet("C:\\Art files\\RedBullet.gif")
 
```

### InlineShapes.AddSmartArt

将 SmartArt 图形作为内嵌形状插入活动文档。

**语法**

**express.AddSmartArt ( *Layout* , *Range* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型           | 说明                                                                                                                                                                                   |
|----------|-----------|--------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Layout* | 必选      | \[SMARTARTLAYOUT\] | 指定 SmartArt 图形的布局的 SmartArtLayout 对象。                                                                                                                                       |
| *Range*  | 可选      | Variant            | 指定 SmartArt 图形所绑定到的文本。如果指定了 Range，则会将 SmartArt 图形置于该范围中第一段的开头。如果省略该参数，则自动选择该范围，并且相对于页面的上边缘和左边缘放置 SmartArt 图形。 |

**返回值**

InlineShape

### InlineShapes.Item

返回集合中的单个 **InlineShape** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

### InlineShapes.New

插入一个空白的 WPS 图片对象，该对象是边长为一英寸的正方形，四周带有边框。此方法以 **InlineShape** 对象的形式返回新图形。

**语法**

**express.New ( *Range* )**

\*express是一个代表 InlineShapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明           |
|---------|-----------|--------------|----------------|
| *Range* | 必选      | Range object | 新图形的位置。 |

**示例**

``` JavaScript
/*以下示例在活动文档中插入一个新的空白图片，并在该图片周围应用阴影边框。*/
function test()
{
let ishapeNew = ActiveDocument.InlineShapes.New(Selection.Range)
    ishapeNew.Borders.Shadow = true
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = false
}
```

## 成员属性

### InlineShapes.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 InlineShapes 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### InlineShapes.Count

返回一个 **Long** 类型的值，该值代表集合中内嵌图形的数目。只读。

**语法**

**express.Count**

\*express是一个代表 InlineShapes 对象的变量。

### InlineShapes.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 InlineShapes 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### InlineShapes.Parent

返回一个 **Object** 类型值，该值代表指定 **InlineShapes** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 InlineShapes 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/InlineShapes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListFormat 对象

## 简介

代表可应用于范围中各段落的列表格式属性。

## 说明

使用 ListFormat 属性可为范围返回 ListFormat 对象。以下示例对所选内容应用默认的项目符号列表格式。

``` JavaScript
Application.Selection.Range.ListFormat.ApplyBulletDefault()
```

应用列表格式的方便途径是使用 ApplyBulletDefault 、 ApplyNumberDefault 和 ApplyOutlineNumberDefault 方法，这些方法分别对应于 “项目符号和编号” 对话框中各选项卡上的第一种列表格式（不包括 “无” ）。

要应用一种非默认格式，请使用 ApplyListTemplate 方法，该方法允许您指定要应用的列表格式（列表模板）。

使用 List 或 ListTemplate 属性从指定范围中的第一段返回列表或列表模板。

可以使用 Range 属性和 ListFormat 对象访问可用于指定区域的列表格式属性和方法。以下示例对活动文档中的第二段应用默认项目符号列表格式。

``` JavaScript
Application.ActiveDocument.Paragraphs.Item(2).Range.ListFormat.ApplyBulletDefault()
```

但是，如果文档中已定义了列表，则可以使用 Lists 属性访问 List 对象。以下示例将在上例中创建的列表的格式更改为 “项目符号和编号” 对话框中 “编号” 选项卡上的第一种编号格式。

``` JavaScript
Application.ActiveDocument.Lists.Item(1).ApplyListTemplate(Application.ListGalleries.Item(2).ListTemplates.Item(1))
```

## 方法

| 序号 | 名称                                                 | 说明                                                                 |
|:----:|:-----------------------------------------------------|:---------------------------------------------------------------------|
|  1   | [ApplyBulletDefault](#ListFormat.ApplyBulletDefault) | 向指定的 ListFormat 对象的区域中的段落添加项目符号和格式。           |
|  2   | [ApplyListTemplate](#ListFormat.ApplyListTemplate)   | 为指定的 ListFormat 对象设置列表格式。                               |
|  3   | [ApplyNumberDefault](#ListFormat.ApplyNumberDefault) | 向指定的 ListFormat 对象的区域中的段落添加默认的编号格式。           |
|  4   | [ListIndent](#ListFormat.ListIndent)                 | 增加指定 ListFormat 对象的区域中的段落的列表级别，增加量为一个级别。 |
|  5   | [ListOutdent](#ListFormat.ListOutdent)               | 降低指定的 ListFormat 对象的区域中段落的列表级别，减量为一个级别。   |
|  6   | [RemoveNumbers](#ListFormat.RemoveNumbers)           | 删除指定列表中的编号或项目符号。                                     |

## 属性

| 序号 | 名称                                               | 说明                                                                                     |
|:----:|:---------------------------------------------------|:-----------------------------------------------------------------------------------------|
|  1   | [List](#ListFormat.List)                           | 返回一个 [List]() 对象，该对象代表在指定 ListFormat 对象中包含的第一个已设置格式的列表。 |
|  2   | [ListLevelNumber](#ListFormat.ListLevelNumber)     | 返回或设置指定的 ListFormat 对象中第一段的列表级别。可读写 Number 类型。                 |
|  3   | [ListPictureBullet](#ListFormat.ListPictureBullet) | 返回一个 InlineShape 对象，该对象代表图片项目符号列表中作为项目符号的图片。              |
|  4   | [ListType](#ListFormat.ListType)                   | 返回指定的 ListFormat 对象区域中所包含的列表类型。 WdListType 类型，只读。               |

## 成员方法

### ListFormat.ApplyBulletDefault

向指定的 **ListFormat** 对象的区域中的段落添加项目符号和格式。

**语法**

**express.ApplyBulletDefault ( *DefaultListBehavior* )**

\*express是一个代表 ListFormat 对象的变量。

**参数**

| 名称                  | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                                              |
|-----------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *DefaultListBehavior* | 可选      | Variant  | 设置一个值来指定 WPS 是否使用新的面向网站的格式，以使列表具有更好地显示效果。可以是下列常量之一：wdWord8ListBehavior（使用与 WPS 97 兼容的格式）或 wdWord9ListBehavior（使用面向网站的格式）。出于兼容原因，默认常量为 wdWord8ListBehavior。但在新建过程中，要建立缩进式多级项目列表，应当使用 wdWord9ListBehavior 来利用改进过的面向网站的格式。 |

**示例**

``` JavaScript
/* 本示例实现的功能是：为选定内容中的段落添加项目符号并设置格式。如果在选定内容中已经设置了项目符号，本示例将清除其原有的项目符号和格式 */
Application.Selection.Range.ListFormat.ApplyBulletDefault()
```

### ListFormat.ApplyListTemplate

为指定的 **ListFormat** 对象设置列表格式。

**语法**

**express.ApplyListTemplate ( *ListTemplate* , *ContinuePreviousList* , *ApplyTo* , *DefaultListBehavior* )**

\*express是一个代表 ListFormat 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型     | 说明                                                                                                                                                                                                                                                                                                                                                                  |
|------------------------|-----------|--------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ListTemplate*         | 必选      | ListTemplate | 要应用的列表模板。                                                                                                                                                                                                                                                                                                                                                    |
| *ContinuePreviousList* | 可选      | Variant      | 如果为 True，则接续前一列表的编号；如果为 False，则新建一个列表。                                                                                                                                                                                                                                                                                                     |
| *ApplyTo*              | 可选      | Variant      | 要应用列表模板的列表部分。可以是下列 WdListApplyTo 常量之一：wdListSelection、wdListWholeList 或 wdListThisPointForward。                                                                                                                                                                                                                                             |
| *DefaultListBehavior*  | 可选      | Variant      | 设置可指定 WPS 是否使用新的网页格式，以便列表具有更好的显示效果的值。其值可取下列 WdDefaultListBehavior 常量之一：wdWord8ListBehavior（使用与 WPS 97 兼容的格式）或 wdWord9ListBehavior（使用适于网页的格式）。考虑到兼容问题，默认常量为 wdWord8ListBehavior。但是在新建过程中若要建立缩进式多级列表，应当使用 wdWord9ListBehavior，以利用改进过的适用于网页的格式。 |

**示例**

``` JavaScript
/* 将第一段和第二段添加上编号 */
function test() {
    let actDoc = Application.ActiveDocument;
    let rng = actDoc.Range(actDoc.Paragraphs.Item(1).Range.Start, actDoc.Paragraphs.Item(2).Range.End);
    if (rng.ListFormat.ListType === wdListNoNumbering) {
        rng.ListFormat.ApplyListTemplate(Application.ListGalleries.Item(2).ListTemplates.Item(4))
    }
}
```

### ListFormat.ApplyNumberDefault

向指定的 **ListFormat** 对象的区域中的段落添加默认的编号格式。

**语法**

**express.ApplyNumberDefault ( *DefaultListBehavior* )**

\*express是一个代表 ListFormat 对象的变量。

**参数**

| 名称                  | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                  |
|-----------------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *DefaultListBehavior* | 可选      | Variant  | 设置一个值来指定 WPS 是否使用新的面向网站的格式，以使列表具有更好地显示效果。可以是下列常量之一：wdWord8ListBehavior或 wdWord9ListBehavior（使用面向网站的格式）。出于兼容原因，默认常量为 wdWord8ListBehavior。但在新建过程中，要建立缩进式多级项目列表，应当使用 wdWord9ListBehavior 来利用改进过的面向网站的格式。 |

**说明**

如果段落已设置了编号列表格式，该方法将清除其中已有的编号和格式。

### ListFormat.ListIndent

增加指定 **ListFormat** 对象的区域中的段落的列表级别，增加量为一个级别。

**语法**

**express.ListIndent ()**

\*express是一个代表 ListFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例设置文档 1 的第一个列表中的每一段依次缩进一个级别 */
Application.Documents.Item(1).Lists.Item(1).Range.ListFormat.ListIndent()
```

### ListFormat.ListOutdent

降低指定的 **ListFormat** 对象的区域中段落的列表级别，减量为一个级别。

**语法**

**express.ListOutdent ()**

\*express是一个代表 ListFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例依次降低活动文档中每一段落的缩进量，减量为一个级别 */
Application.Documents.Item(1).Lists.Item(1).Range.ListFormat.ListOutdent()
```

### ListFormat.RemoveNumbers

删除指定列表中的编号或项目符号。

**语法**

**express.RemoveNumbers ( *NumberType* )**

\*express是一个代表 ListFormat 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型     | 说明               |
|--------------|-----------|--------------|--------------------|
| *NumberType* | 可选      | WdNumberType | 要删除的编号类型。 |

## 成员属性

### ListFormat.List

返回一个 **[List]()** 对象，该对象代表在指定 **ListFormat** 对象中包含的第一个已设置格式的列表。

**语法**

**express.List**

\*express是一个代表 ListFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例返回选定内容中的第一个列表的段落数 */
Application.Selection.Range.ListFormat.List.ListParagraphs.Count
```

### ListFormat.ListLevelNumber

返回或设置指定的 **ListFormat** 对象中第一段的列表级别。可读写 **Number** 类型。

**语法**

**express.ListLevelNumber**

\*express是一个代表 ListFormat 对象的变量。

### ListFormat.ListPictureBullet

返回一个 InlineShape 对象，该对象代表图片项目符号列表中作为项目符号的图片。

**语法**

**express.ListPictureBullet**

\*express是一个代表 ListFormat 对象的变量。

### ListFormat.ListType

返回指定的 **ListFormat** 对象区域中所包含的列表类型。 **WdListType** 类型，只读。

**语法**

**express.ListType**

\*express是一个代表 ListFormat 对象的变量。

**说明**

**wdListListNumOnly** 常量引用 LISTNUM 域，可将该域添至段落的文本中。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ListFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Sections 对象

## 简介

所选内容、范围或文档中的 Section 对象的集合。

## 说明

表示指定文档的Section对象集合

## 方法

| 序号 | 名称                   | 说明                                                  |
|:----:|:-----------------------|:------------------------------------------------------|
|  1   | [Add](#Sections.Add)   | 返回一个 Section 对象，该对象代表添加到文档中的新节。 |
|  2   | [Item](#Sections.Item) | 返回集合中的单个 Section 对象。                       |

## 属性

| 序号 | 名称                                 | 说明                                                                         |
|:----:|:-------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Sections.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Sections.Count)             | 返回一个 Long 类型的值，该值代表集合中的节数。只读。                         |
|  3   | [Creator](#Sections.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [First](#Sections.First)             | 返回一个 Section 对象，该对象代表 Sections 集合中的第一个项目。              |
|  5   | [Last](#Sections.Last)               | 将 Sections 集合中的最后一项作为 Section 对象返回。                          |
|  6   | [PageSetup](#Sections.PageSetup)     | 返回一个 PageSetup 对象，该对象与指定文档、区域、节或选定内容相关联。        |
|  7   | [Parent](#Sections.Parent)           | 返回一个 Object 类型值，该值代表指定 Sections 对象的父对象。                 |

## 成员方法

### Sections.Add

返回一个 ****Section**** 对象，该对象代表添加到文档中的新节。

**语法**

**express.Add ( *Range* , *Start* )**

\*express是一个代表 Sections 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                               |
|---------|-----------|----------|----------------------------------------------------------------------------------------------------|
| *Range* | 可选      | Variant  | 要在其之前插入分节符的区域。如果省略该参数，则将分节符插至文档末尾。                               |
| *Start* | 可选      | Variant  | 要添加的分节符的类型。可以是 WdSectionStart 常量之一。如果省略该参数，则添加“下一页”类型的分节符。 |

**示例**

``` JavaScript
/*本示例在活动文档第三段之前添加一个“下一页”的分节符*/
function test()
{
    let myRange = Application.ActiveDocument.Paragraphs.Item(3).Range
    Application.ActiveDocument.Sections.Add(myRange)
}

/*本示例在选定内容中添加一个“连续”的分节符*/
function test()
{
    let myRange = Application.Selection.Range
    Application.ActiveDocument.Sections.Add(myRange, wdSectionContinuous)
}

/*本示例在活动文档的末尾添加一个“下一页”分节符*/
function test()
{
    Application.ActiveDocument.Sections.Add()
}
```

### Sections.Item

返回集合中的单个 **Section** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Sections 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

**示例**

``` JavaScript
/*返回当前活动文档第一节Section对象*/
Application.ActiveDocument.Sections.Item(1)
```

## 成员属性

### Sections.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Sections 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Sections.Count

返回一个 **Long** 类型的值，该值代表集合中的节数。只读。

**语法**

**express.Count**

\*express是一个代表 Sections 对象的变量。

**示例**

``` JavaScript
/*取当前活动文档有多少节*/
Application.ActiveDocument.Sections.Count
```

### Sections.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Sections 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Sections.First

返回一个 **Section** 对象，该对象代表 **Sections** 集合中的第一个项目。

**语法**

**express.First**

\*express是一个代表 Sections 对象的变量。

### Sections.Last

将 **Sections** 集合中的最后一项作为 **Section** 对象返回。

**语法**

**express.Last**

\*express是一个代表 Sections 对象的变量。

### Sections.PageSetup

返回一个 **PageSetup** 对象，该对象与指定文档、区域、节或选定内容相关联。

**语法**

**express.PageSetup**

\*express是一个代表 Sections 对象的变量。

**示例**

``` JavaScript
/*以下示例将 Summary.doc 中第一节的装订线设置为 36 磅（0.5 英寸）。*/
Documents.Item("Summary.doc").Sections.Item(1).PageSetup.Gutter = 36
```

### Sections.Parent

返回一个 **Object** 类型值，该值代表指定 **Sections** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Sections 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Sections](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XMLNamespace 对象

## 简介

代表架构库中的单个架构。

## 说明

在 WPS 中，可以通过 “模板和加载项” 对话框中的 “XML 架构” 选项卡来访问架构库。架构库表示安装在用户计算机上的架构，并且用户已将这些架构应用于 WPS 文档或者用户已使用 “架构库” 对话框将这些架构显式添加到架构库中。

使用 XMLNamespaces 集合的 Item 方法可返回一个 XMLNameSpace 对象。 Item 方法的索引值可以是 Long 类型，表明架构在架构库中的位置，也可以是 String 类型，代表使用 URI 属性（在架构中定义的 TargetNamespace 设置）返回的架构的名称。

以下示例将一个名为 SimpleSample 的架构附加到活动文档。

``` JavaScript
function ApplySampleSchema(){
    let objSchema = Application.XMLNamespaces
    
    for(let i =1; i <= objSchema.Count; i++){
        if(objSchema.Item(i).URI == "SimpleSample"){
            objSchema.Item(i).AttachToDocument (ActiveDocument)
            break
        }
    }
}
```

| 注释                                                                                                                                                            |
|-----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中。有关详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document SDK。 |

## 方法

| 序号 | 名称                                               | 说明                                       |
|:----:|:---------------------------------------------------|:-------------------------------------------|
|  1   | [AttachToDocument](#XMLNamespace.AttachToDocument) | 将 XML 架构附加到文档。                    |
|  2   | [Delete](#XMLNamespace.Delete)                     | 删除可用的 XML 架构列表中指定的 XML 架构。 |

## 属性

| 序号 | 名称                                               | 说明                                                                                                                                                        |
|:----:|:---------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Alias](#XMLNamespace.Alias)                       | 返回一个 String 类型的值，该值代表指定对象的显示名称。                                                                                                      |
|  2   | [Application](#XMLNamespace.Application)           | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                              |
|  3   | [Creator](#XMLNamespace.Creator)                   | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                |
|  4   | [DefaultTransform](#XMLNamespace.DefaultTransform) | 该属性返回一个 XSLTransform 对象，该对象表示对于特殊命名空间从 XML 架构中打开文档时要使用的默认 Extensible Stylesheet Language Transformation (XSLT) 文件。 |
|  5   | [Location](#XMLNamespace.Location)                 | 返回或设置 String 值，该值表示 XML 架构的命名空间在架构库中的物理位置。可读/写。                                                                            |
|  6   | [Parent](#XMLNamespace.Parent)                     | 返回一个 Object 类型值，该值代表指定 XMLNamespace 对象的父对象。                                                                                            |
|  7   | [URI](#XMLNamespace.URI)                           | 返回一个 String 类型的值，代表相关命名空间的“统一资源标识符”(URI)。                                                                                         |
|  8   | [XSLTransforms](#XMLNamespace.XSLTransforms)       | 返回一个 XSLTransforms 集合，代表指定用于架构的 Extensible Stylesheet Language Transformation (XSLT) 文件。                                                 |

## 成员方法

### XMLNamespace.AttachToDocument

将 XML 架构附加到文档。

**语法**

**express.AttachToDocument ( *Document* )**

\*express是一个代表 XMLNamespace 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                          |
|------------|-----------|----------|-------------------------------|
| *Document* | 必选      | Document | 将指定 XML 架构附加到的文档。 |

**示例**

以下示例将 SimpleSample 架构添加到 Schema Library，再附加到活动文档。

| 注释                                                                                                                                                            |
|-----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中。有关详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document SDK。 |

``` JavaScript
function test()
{
let objSchema = Application.XMLNamespaces.Add("c:\\schemas\\simplesample.xsd")
objSchema.AttachToDocument (ActiveDocument)
}
```

### XMLNamespace.Delete

删除可用的 XML 架构列表中指定的 XML 架构。

**语法**

**express.Delete ()**

\*express是一个代表 XMLNamespace 对象的变量。

## 成员属性

### XMLNamespace.Alias

返回一个 **String** 类型的值，该值代表指定对象的显示名称。

**语法**

**express.Alias**

\*express是一个代表 XMLNamespace 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动文档所附加的第一个架构的显示名称。*/
MsgBox(Application.XMLNamespaces.Item(1).Alias)
```

### XMLNamespace.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 XMLNamespace 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XMLNamespace.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XMLNamespace 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XMLNamespace.DefaultTransform

该属性返回一个 **XSLTransform** 对象，该对象表示对于特殊命名空间从 XML 架构中打开文档时要使用的默认 Extensible Stylesheet Language Transformation (XSLT) 文件。

**语法**

**express.DefaultTransform**

\*express是一个代表 XMLNamespace 对象的变量。

**示例**

``` JavaScript
/*以下示例返回架构库中第一个架构的默认 XSLT，WPS 将使用 XSLT 打开与该架构命名空间相关联的 XML 文件。本示例假设第一个架构有一个或多个已应用的 XSLT 文件。*/
let objXSLT = Application.XMLNamespaces.Item(1).DefaultTransform
```

### XMLNamespace.Location

返回或设置 **String** 值，该值表示 XML 架构的命名空间在架构库中的物理位置。可读/写。

**语法**

**express.Location**

\*express是一个代表 XMLNamespace 对象的变量。

### XMLNamespace.Parent

返回一个 **Object** 类型值，该值代表指定 **XMLNamespace** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XMLNamespace 对象的变量。

### XMLNamespace.URI

返回一个 **String** 类型的值，代表相关命名空间的“统一资源标识符”(URI)。

**语法**

**express.URI**

\*express是一个代表 XMLNamespace 对象的变量。

**示例**

``` JavaScript
/*下列示例显示“架构库”中第一个架构的 URI。*/
MsgBox(Application.XMLNamespaces.Item(1).URI)
```

### XMLNamespace.XSLTransforms

返回一个 XSLTransforms 集合，代表指定用于架构的 Extensible Stylesheet Language Transformation (XSLT) 文件。

**语法**

**express.XSLTransforms**

\*express是一个代表 XMLNamespace 对象的变量。

**说明**

下列示例将 simplesample.xsl 转换添加到 SimpleSample 架构的转换中。

| 注释                                                                                                                                                                                                                      |
|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中；但没有相应的 simplesample.xsl 文件。此代码仅为示例而创建。有关详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document SDK。 |

**示例**

``` JavaScript
function test()
{
let objSchema = Application.XMLNamespaces.Item("SimpleSample")
let objTransform = objSchema.XSLTransforms.Add("c:\\schemas\\simplesample.xsl")
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XMLNamespace](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ColorScheme 对象

## 简介

代表一种配色方案，该配色方案是八种颜色的组合，分别用于幻灯片、备注页或讲义中的不同元素，例如标题或背景（请注意，演示文稿中幻灯片、备注页及讲义的配色方案可以单独设置）。

## 说明

每种颜色由一个 RGBColor 对象代表。 ColorScheme 对象是 ColorSchemes 集合的成员。 ColorSchemes 集合包含演示文稿中所有的配色方案。

以下示例说明如何执行下列操作：

- 从演示文稿的所有配色方案的集合中返回 ColorScheme 对象
- 返回附加到特定幻灯片或母版的 ColorScheme 对象
- 从 ColorScheme 对象返回单个幻灯片元素的颜色

## 方法

| 序号 | 名称                          | 说明                                                     |
|:----:|:------------------------------|:---------------------------------------------------------|
|  1   | [Colors](#ColorScheme.Colors) | 返回一个 RGBColor 对象，该对象代表配色方案中的一种颜色。 |
|  2   | [Delete](#ColorScheme.Delete) | 删除指定的对象                                           |

## 属性

| 序号 | 名称                                    | 说明                                                    |
|:----:|:----------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#ColorScheme.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#ColorScheme.Count)             | 返回指定集合中的对象数目。                              |
|  3   | [Parent](#ColorScheme.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### ColorScheme.Colors

返回一个 **RGBColor** 对象，该对象代表配色方案中的一种颜色。

**语法**

**express.Colors ( *SchemeColor* )**

\*express是一个代表 ColorScheme 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型           | 说明                         |
|---------------|-----------|--------------------|------------------------------|
| *SchemeColor* | 必选      | PpColorSchemeIndex | 指定的配色方案中的一种颜色。 |

**返回值**

RGBColor

**示例**

``` JavaScript
/*本示例设置当前演示文稿中第一张和第三张幻灯片的标题颜色。*/
function test() {
    let mySlides = Application.ActivePresentation.Slides.Range([1, 3])
    mySlides.ColorScheme.Colors(ppTitle).RGB = 0, 255, 0
}
```

### ColorScheme.Delete

删除指定的对象

**语法**

**express.Delete ()**

\*express是一个代表 ColorScheme 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

## 成员属性

### ColorScheme.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 ColorScheme 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### ColorScheme.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 ColorScheme 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭第二个窗口。*/
function test(){
　　　　let w = Application.Windows
    for(let i = 2; i <= w.Count; i++) {
　　　　    w.Item(2).Close()
　　　　}
}
```

### ColorScheme.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ColorScheme 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/ColorScheme](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DocumentWindow 对象

## 简介

代表一个文档窗口。DocumentWindow对象是 DocumentWindows 集合的成员。 DocumentWindows 集合包含所有打开的文档窗口。

## 说明

使用 Presentation 属性可返回指定文档窗口中当前运行的演示文稿。 使用 Selection 属性可返回选定范围。 使用 SplitHorizontal 属性可返回普通视图中大纲窗格所占屏幕宽度的百分比。 使用 SplitVertical 属性可返回普通视图中幻灯片窗格所占屏幕高度的百分比。 使用 View 属性可返回指定文档窗口中的视图。

## 方法

| 序号 | 名称                                                           | 说明                                                                                                            |
|:----:|:---------------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------|
|  1   | [Activate](#DocumentWindow.Activate)                           | 激活指定的对象                                                                                                  |
|  2   | [Close](#DocumentWindow.Close)                                 | 关闭指定的文档窗口、演示文稿或打开的任意多边形。                                                                |
|  3   | [FitToPage](#DocumentWindow.FitToPage)                         | 调整指定文档窗口的大小，以适应当前显示的信息。                                                                  |
|  4   | [LargeScroll](#DocumentWindow.LargeScroll)                     | 滚动到指定的文档窗口中的页面。                                                                                  |
|  5   | [NewWindow](#DocumentWindow.NewWindow)                         | 打开一个包含指定窗口中所显示文档的新窗口。返回一个代表新窗口的 DocumentWindow 对象                              |
|  6   | [PointsToScreenPixelsX](#DocumentWindow.PointsToScreenPixelsX) | 将横向度量值的单位由磅值转换为像素。可用于返回文本框或形状的横向屏幕位置。以类型 Single 返回转换后的度量值。    |
|  7   | [PointsToScreenPixelsY](#DocumentWindow.PointsToScreenPixelsY) | 将纵向度量值的单位由磅值转换为像素。可用于返回文本框或形状的纵向屏幕位置。以类型 Single 返回转换后的度量值。    |
|  8   | [RangeFromPoint](#DocumentWindow.RangeFromPoint)               | 返回位于某一点（由屏幕上的位置坐标对指定）的 Shape 对象。如果在指定的坐标对上没有形状，则此方法将返回 Nothing。 |
|  9   | [ScrollIntoView](#DocumentWindow.ScrollIntoView)               | 滚动文档窗口，使指定矩形区域内的项目显示在文档窗口或窗格中。                                                    |
|  10  | [SmallScroll](#DocumentWindow.SmallScroll)                     | 按行和列滚动指定的文档窗口。                                                                                    |

## 属性

| 序号 | 名称                                         | 说明                                                                                                                                                                                            |
|:----:|:---------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Active](#DocumentWindow.Active)             | 返回指定的窗格或窗口是否处于激活的状态。只读MsoTriState 类型。                                                                                                                                  |
|  2   | [ActivePane](#DocumentWindow.ActivePane)     | 返回一个代表文档窗口中活动窗格的Pane对象。只读。                                                                                                                                                |
|  3   | [Application](#DocumentWindow.Application)   | 返回一个Application对象，该对象表示指定对象的创建者。                                                                                                                                           |
|  4   | [Caption](#DocumentWindow.Caption)           | 返回在文档窗口标题栏中显示的文本。String类型，只读。                                                                                                                                            |
|  5   | [Height](#DocumentWindow.Height)             | 以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。                                                                               |
|  6   | [Left](#DocumentWindow.Left)                 | 返回或设置一个 Single 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。 |
|  7   | [Panes](#DocumentWindow.Panes)               | 返回一个 Panes 集合，该集合代表文档窗口中的 窗格 （窗格：文档窗口的一部分，以垂直或水平条为界限并由此与其他部分分隔开。） 只读。                                                                |
|  8   | [Presentation](#DocumentWindow.Presentation) | 返回一个 Presentation 对象，该对象代表创建指定文档窗口或幻灯片放映窗口的演示文稿。只读。                                                                                                        |
|  9   | [Selection](#DocumentWindow.Selection)       | 返回一个 Selection 对象，该对象代表指定的文档窗口中的选定内容。只读。                                                                                                                           |
|  10  | [Top](#DocumentWindow.Top)                   | 返回或设置一个 Single 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。   |
|  11  | [View](#DocumentWindow.View)                 | 返回一个代表指定文档窗口中视图的View对象。只读。                                                                                                                                                |
|  12  | [ViewType](#DocumentWindow.ViewType)         | 返回或设置指定文档窗口中所包含的视图类型。可读/写。                                                                                                                                             |
|  13  | [Width](#DocumentWindow.Width)               | 以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。                                                                               |
|  14  | [WindowState](#DocumentWindow.WindowState)   | 返回或设置指定窗口的状态。可读/写。PpWindowState 类型。                                                                                                                                         |

## 成员方法

### DocumentWindow.Activate

激活指定的对象

**语法**

**express.Activate ()**

\*express是一个代表 DocumentWindow 对象的变量。

**示例**

``` JavaScript
/*本示例将激活第一个文档窗口。*/
Application.Windows.Item(1).Activate()
```

### DocumentWindow.Close

关闭指定的文档窗口、演示文稿或打开的任意多边形。

**语法**

**express.Close ()**

\*express是一个代表 DocumentWindow 对象的变量。

**说明**

使用此方法，WPP关闭打开的演示文档，并且不提示用户保存所做的工作。为防止工作丢失，您应在使用 SaveClose 方法之前使用 SaveAsSave 或 CloseSaveAs 方法。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口*/
function test(){
    let dClose = Application.Windows
    for ( let i = 2; i <= dClose.Count; i++ ) {
        dClose.Item(i).Close()
    }
}


/*本示例在不保存更改的情况下关闭 Pres1.ppt*/
function test(){
    let dClose1 = Application.Presentations.Item("pres1.ppt")
    dClose1.Saved = true
    dClose1.Close()
}
```

### DocumentWindow.FitToPage

调整指定文档窗口的大小，以适应当前显示的信息。

**语法**

**express.FitToPage ()**

\*express是一个代表 DocumentWindow 对象的变量。

**示例**

``` JavaScript
//以下示例将活动窗口中的视图设置为幻灯片视图，将显示比例设置为 25%，然后调整窗口大小以适应所显示的幻灯片。
function test() {
    let dWindow = Application.ActiveWindow
    dWindow.View.Zoom = 25
    dWindow.FitToPage()
}
```

### DocumentWindow.LargeScroll

滚动到指定的文档窗口中的页面。

**语法**

**express.LargeScroll ( *Down* , *Up* , *ToRight* , *ToLeft* )**

\*express是一个代表 DocumentWindow 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                   |
|-----------|-----------|----------|------------------------|
| *Down*    | 可选      | Long     | 指定要向下滚动的页数。 |
| *Up*      | 可选      | Long     | 指定要向上滚动的页数。 |
| *ToRight* | 可选      | Long     | 指定要向右滚动的页数。 |
| *ToLeft*  | 可选      | Long     | 指定要向左滚动的页数。 |

**说明**

如果未指定任何参数，则该方法向下滚动一页。如果同时指定了 **Down** 和 **Up** ，则效果是二者的叠加。例如，如果 **Down** 为 2 而 **Up** 为 4，则该方法向上滚动两页。同样，如果同时指定了 **Right** 和 **Left** ，则它们将共同起作用。

以上任何参数均可为负数。

**示例**

``` JavaScript
//本示例将活动窗口向下滚动三页。
Application.ActiveWindow.LargeScroll(3)
```

### DocumentWindow.NewWindow

打开一个包含指定窗口中所显示文档的新窗口。返回一个代表新窗口的 DocumentWindow 对象

**语法**

**express.NewWindow ()**

\*express是一个代表 DocumentWindow 对象的变量。

**示例**

``` JavaScript
/*本示例创建一个与当前窗口有相同内容的新窗口（同时激活新窗口）并返回到原窗口*/
function test(){
    let oldW = Application.ActiveWindow
    let newW = oldW.NewWindow()
    oldW.Activate()
}
```

### DocumentWindow.PointsToScreenPixelsX

将横向度量值的单位由磅值转换为像素。可用于返回文本框或形状的横向屏幕位置。以类型 **Single** 返回转换后的度量值。

**语法**

**express.PointsToScreenPixelsX ( *Points* )**

\*express是一个代表 DocumentWindow 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                             |
|----------|-----------|----------|--------------------------------------------------|
| *Points* | 必选      | Single   | 要转换为以像素为单位的横向度量值（以磅为单位）。 |

**示例**

``` JavaScript
//本示例将选中的文本框架边界框的宽度和高度从磅值转换到像素。并将值返回到 myXparm 和 myYparm。
function test() {
    let myXparm = Application.ActiveWindow.PointsToScreenPixelsX(Application.ActiveWindow.Selection.TextRange.BoundWidth)
    alert(myXparm)
    let myYparm = Application.ActiveWindow.PointsToScreenPixelsY(Application.ActiveWindow.Selection.TextRange.BoundHeight)
    alert(myYparm)
}
```

### DocumentWindow.PointsToScreenPixelsY

将纵向度量值的单位由磅值转换为像素。可用于返回文本框或形状的纵向屏幕位置。以类型 **Single** 返回转换后的度量值。

**语法**

**express.PointsToScreenPixelsY ( *Points* )**

\*express是一个代表 DocumentWindow 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                             |
|----------|-----------|----------|--------------------------------------------------|
| *Points* | 必选      | Single   | 要转换为以像素为单位的纵向度量值（以磅为单位）。 |

**示例**

``` JavaScript
//本示例将选取的文本框架边界框的宽度和高度从磅值转换到像素。并将值返回到 myXparm 和 myYparm。
function test() {
    let myXparm = Application.ActiveWindow.PointsToScreenPixelsX(Application.ActiveWindow.Selection.TextRange.BoundWidth)
    alert(myXparm)
    let myYparm = Application.ActiveWindow.PointsToScreenPixelsY(Application.ActiveWindow.Selection.TextRange.BoundHeight)
    alert(myYparm)
}
```

### DocumentWindow.RangeFromPoint

返回位于某一点（由屏幕上的位置坐标对指定）的 Shape 对象。如果在指定的坐标对上没有形状，则此方法将返回 Nothing。

**语法**

**express.RangeFromPoint ( *x* , *y* )**

\*express是一个代表 DocumentWindow 对象的变量。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                                           |
|------|-----------|----------|------------------------------------------------|
| *x*  | 必选      | Long     | 从屏幕左边缘到该点的水平距离（以像素为单位）。 |
| *y*  | 必选      | Long     | 从屏幕左边缘到该点的水平距离（以像素为单位）。 |

**示例**

``` JavaScript
//本示例使用坐标 (288, 100) 向第一张幻灯片中添加一个新的五角星。然后，再将坐标由单位磅转换为像素，并使用 RangeFromPoint 方法返回一个对该新对象的引用，最后更改此五角星的填充色。
function test() {
    Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShape5pointStar, 288, 100, 100, 72).Select()
    let myPointX = Application.ActiveWindow.PointsToScreenPixelsX(288)
    let myPointY = Application.ActiveWindow.PointsToScreenPixelsY(100)
    let myShape = Application.ActiveWindow.RangeFromPoint(myPointX, myPointY)
    myShape.Fill.ForeColor.RGB = (80, 160, 130)
}
```

### DocumentWindow.ScrollIntoView

滚动文档窗口，使指定矩形区域内的项目显示在文档窗口或窗格中。

**语法**

**express.ScrollIntoView ( *Left* , *Top* , *Width* , *Height* , *Start* )**

\*express是一个代表 DocumentWindow 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型    | 说明                                           |
|----------|-----------|-------------|------------------------------------------------|
| *Left*   | 必选      | Long        | 从文档窗口的左边缘到该矩形的水平距离（点数）。 |
| *Top*    | 必选      | Long        | 从文档窗口顶部到该矩形的垂直距离（点数）。     |
| *Width*  | 必选      | Long        | 矩形的宽度（点数）。                           |
| *Height* | 必选      | Long        | 矩形的高度（点数）。                           |
| *Start*  | 可选      | MsoTriState | 确定矩形相对于文档窗口的起始位置。             |

**说明**

如果矩形边界超过了文档窗口的大小，则 **Start** 参数可指定矩形的哪一个顶点将显示在屏幕上或具有初始焦点。该方法不能在大纲或幻灯片浏览视图中使用。

|                                                            |
|------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。              |
| **msoCTrue**                                               |
| **msoFalse** 矩形的右下角将显示于文档窗口的右下角。        |
| **msoTriStateMixed**                                       |
| **msoTriStateToggle**                                      |
| **msoTrue** 默认值。矩形的左上角将显示在文档窗口的左上角。 |

**示例**

``` JavaScript
//本示例将查看一个 100x200（单位：磅）的区域，该区域距幻灯片左边缘 50 磅，距幻灯片上边缘 20 磅。该矩形的左上角位于当前文档窗口的左上角。
Application.ActiveWindow.ScrollIntoView(50, 20, 100, 200)
```

### DocumentWindow.SmallScroll

按行和列滚动指定的文档窗口。

**语法**

**express.SmallScroll ( *Down* , *Up* , *ToRight* , *ToLeft* )**

\*express是一个代表 DocumentWindow 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                 |
|-----------|-----------|----------|----------------------|
| *Down*    | 可选      | Long     | 指定向下滚动的行数。 |
| *Up*      | 可选      | Long     | 指定向上滚动的行数。 |
| *ToRight* | 可选      | Long     | 指定向右滚动的列数。 |
| *ToLeft*  | 可选      | Long     | 指定向左滚动的列数。 |

**说明**

如果未指定任何参数，则该方法向下滚动一行。如果同时指定了 **Down** 和 **Up** ，则效果是二者的叠加。例如，如果 **Down** 为 2，而 **Up** 为 4，该方法将向上滚动两行。同样，如果同时指定了 **Right** 和 **Left** ，则它们将共同起作用。

以上任何参数均可为负数。

**示例**

``` JavaScript
//本示例将当前窗口向下滚动三行。
Application.ActiveWindow.SmallScroll(3)
```

## 成员属性

### DocumentWindow.Active

返回指定的窗格或窗口是否处于激活的状态。只读MsoTriState 类型。

**语法**

**express.Active**

\*express是一个代表 DocumentWindow 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。 msoCTrue msoFalse msoTriStateMixed msoTriStateToggle msoTrue：指定的窗格或窗口处于活动状态。

**示例**

``` JavaScript
/*以下示例检查演示文稿文件“test.ppt”是否位于当前窗口。如果不是，则将当前处于激活状态的演示文稿的名称保存在变量oldWin中，并激活“test.ppt”演示文稿*/
function test(){
    let dActive = Application.Presentations.Item("test.ppt").Windows.Item(1)
    if(!dActive.Active){
        let oldWin = Application.ActiveWindow
        dActive.Activate()
    }
}
```

### DocumentWindow.ActivePane

返回一个代表文档窗口中活动窗格的Pane对象。只读。

**语法**

**express.ActivePane**

\*express是一个代表 DocumentWindow 对象的变量。

**示例**

``` JavaScript
/*如果活动窗格为幻灯片窗格，则本示例将设置备注窗格为活动窗格。备注窗格是Panes集合中的第三个成员。*/
function test(){
    if(Application.ActiveWindow.ActivePane.ViewType == ppViewSlide){
        Application.ActiveWindow.Panes.Item(3).Activate()
    }
}
```

### DocumentWindow.Application

返回一个Application对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 DocumentWindow 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个Presentation对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行WPP的文件夹中。*/
function test(){
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1,1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

### DocumentWindow.Caption

返回在文档窗口标题栏中显示的文本。String类型，只读。

**语法**

**express.Caption**

\*express是一个代表 DocumentWindow 对象的变量。

### DocumentWindow.Height

以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Height**

\*express是一个代表 DocumentWindow 对象的变量。

**说明**

一个 HeightShape 对象的 ShapeHeight 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

**示例**

``` JavaScript
/*以下示例将第二文档窗口的高度设置为应用程序窗口高度的一半。*/
Application.Windows.Item(2).Height = Application.Windows.Item(1).Height/2
  
```

### DocumentWindow.Left

返回或设置一个 Single 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Left**

\*express是一个代表 DocumentWindow 对象的变量。

### DocumentWindow.Panes

返回一个 Panes 集合，该集合代表文档窗口中的 窗格 （窗格：文档窗口的一部分，以垂直或水平条为界限并由此与其他部分分隔开。） 只读。

**语法**

**express.Panes**

\*express是一个代表 DocumentWindow 对象的变量。

**示例**

``` JavaScript
/*本示例检测活动窗口中窗格的数目。如果其值为1，则表明是非普通视图的任意其他视图，然后本示例将激活普通视图。*/
Application.ActiveWindow.Panes.Count 
```

### DocumentWindow.Presentation

返回一个 Presentation 对象，该对象代表创建指定文档窗口或幻灯片放映窗口的演示文稿。只读。

**语法**

**express.Presentation**

\*express是一个代表 DocumentWindow 对象的变量。

**说明**

如果在第一个文档窗口中当前放映的幻灯片来自嵌入的演示文稿，则 Windows(1).View.Slide.Parent 返回该嵌入演示文稿，并且 Windows(1).Presentation 返回在其中创建第一个文档窗口的演示文稿。 如果在第一个幻灯片放映窗口中当前放映的幻灯片来自嵌入的演示文稿，则 SlideShowWindows(1).View.Slide.Parent 返回该嵌入演示文稿，并且 SlideShowWindows(1).Presentation 返回在其中创建第一个幻灯片放映窗口的演示文稿。

**示例**

``` JavaScript
/*以下示例使第一个窗口的演示文稿幻灯片编号延续到第二个窗口中。*/
function test(){
    let firstPresSlides = Application.Windows.Item(1).Presentation.Slides.Count
    Application.Windows.Item(2).Presentation.PageSetup.FirstSlideNumber  += 1
}
```

### DocumentWindow.Selection

返回一个 Selection 对象，该对象代表指定的文档窗口中的选定内容。只读。

**语法**

**express.Selection**

\*express是一个代表 DocumentWindow 对象的变量。

**示例**

``` JavaScript
/*如果在活动窗口中选择了文本，本示例则会将该文本设为斜体。*/
function test(){
    let dText = Application.ActiveWindow.Selection
    if(dText.Type == ppSelectionText){
        dText.TextRange.Font.Italic = true
    }
}
```

### DocumentWindow.Top

返回或设置一个 Single 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Top**

\*express是一个代表 DocumentWindow 对象的变量。

**示例**

``` JavaScript
/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所用水平可用空间。*/
/*要使以下示例执行，只能有两个打开的文档窗口。*/
function test(){    
    let sngHeight = Application.Windows.Item(1).Height
    let sngWidth = Application.Windows.Item(1).Width + Windows.Item(2).Width
    let dWindows1 = Application.Windows.Item(1)
    dWindows1.Width = sngWidth
    dWindows1.Height = sngHeight/2
    dWindows1.Left = 0

    let dWindows2 = Application.Windows.Item(2)
    dWindows2.Width = sngWidth
    dWindows2.Height = sngHeight/2
    dWindows2.Top = sngHeight/2
    dWindows2.Left = 0
}
```

### DocumentWindow.View

返回一个代表指定文档窗口中视图的View对象。只读。

**语法**

**express.View**

\*express是一个代表 DocumentWindow 对象的变量。

### DocumentWindow.ViewType

返回或设置指定文档窗口中所包含的视图类型。可读/写。

**语法**

**express.ViewType**

\*express是一个代表 DocumentWindow 对象的变量。

**说明**

PpViewType 可以是下列 PpViewType 常量之一。 ppViewHandoutMaster ppViewMasterThumbnails ppViewNormal ppViewNotesMaster ppViewNotesPage ppViewOutline ppViewPrintPreview ppViewSlide ppViewSlideMaster ppViewSlideSorter ppViewThumbnails ppViewTitleMaster

**示例**

``` JavaScript
/*如果当前窗口以普通视图显示，本示例将该窗口中的视图更改为幻灯片浏览视图。*/
function test(){
    let dView = Application.ActiveWindow
    if ( dView.ViewType == ppViewNormal ) {
    dView.ViewType = ppViewSlideSorter
    }
}
```

### DocumentWindow.Width

以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Width**

\*express是一个代表 DocumentWindow 对象的变量。

**示例**

``` JavaScript
/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的文档窗口。*/
function test(){
    let ah = Application.Windows.Item(1).Height                      // available height
    let aw = Application.Windows.Item(1).Width + Windows.Item(2).Width    // available width

    let dWindows1 = Application.Windows.Item(1)
    dWindows1.Width = aw
    dWindows1.Height = ah / 2
    dWindows1.Left = 0

    let dWindows2 = Application.Windows.Item(2)
    dWindows2.Width = aw
    dWindows2.Height = ah / 2
    dWindows2.Top = ah / 2
    dWindows2.Left = 0
}

/*以下示例将指定表中第一列的宽度设置为 80 磅。*/
Application.ActivePresentation.Slides.Item(2).Shapes.Item(1).Width = 80
    
```

### DocumentWindow.WindowState

返回或设置指定窗口的状态。可读/写。PpWindowState 类型。

**语法**

**express.WindowState**

\*express是一个代表 DocumentWindow 对象的变量。

**说明**

PpWindowState 可以是下列 PpWindowState 常量之一。 ppWindowMaximized ppWindowMinimized ppWindowNormal 当窗口状态为 ppWindowNormal 时，该窗口既未最大化也未最小化。

**示例**

``` JavaScript
/*以下示例将当前窗口最大化。*/
Application.ActiveWindow.WindowState = ppWindowMaximized
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/DocumentWindow](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Master 对象

## 简介

代表一个幻灯片母版、标题母版、讲义母版、备注母版或设计母版。

## 说明

若要返回一个 Master 对象，请使用 Slide 对象或 SlideRange 集合的 Master 属性，或使用 Presentation 对象的 HandoutMaster 、 NotesMaster 、 SlideMaster 或 TitleMaster 属性。请注意，这些属性中的某些也可用于 Design 对象。以下示例设置活动演示文稿的幻灯片母版的背景填充。

``` JavaScript
Application.ActivePresentation.SlideMaster.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass)
```

若要为演示文稿添加标题母版或设计并返回代表新标题母版或设计的 Master 对象，请使用 AddTitleMaster 属性。以下示例在活动演示文稿中添加一个标题母版，并将标题占位符放置在距母版顶部 10 磅的位置。

``` JavaScript
Application.ActivePresentation.AddTitleMaster.Shapes.Title.Top = 10
```

## 方法

| 序号 | 名称                             | 说明                                                                             |
|:----:|:---------------------------------|:---------------------------------------------------------------------------------|
|  1   | [ApplyTheme](#Master.ApplyTheme) | 将主题或设计模板应用于指定的幻灯片母板、标题母版、讲义母版、备注母版或设计母版。 |
|  2   | [Delete](#Master.Delete)         | 删除指定的对象。                                                                 |

## 属性

| 序号 | 名称                                               | 说明                                                                                                                                                                                                                                               |
|:----:|:---------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Master.Application)                 | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                            |
|  2   | [Background](#Master.Background)                   | 返回代表幻灯片背景的 ShapeRange 对象。                                                                                                                                                                                                             |
|  3   | [BackgroundStyle](#Master.BackgroundStyle)         | 设置或返回指定对象的背景样式。可读/写。                                                                                                                                                                                                            |
|  4   | [ColorScheme](#Master.ColorScheme)                 | 返回或设置 ColorScheme 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。                                                                                                                                                     |
|  5   | [Design](#Master.Design)                           | 返回代表设计的 Design 对象。                                                                                                                                                                                                                       |
|  6   | [HeadersFooters](#Master.HeadersFooters)           | 返回一个 HeadersFooters 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。                                                                                                                       |
|  7   | [Height](#Master.Height)                           | 以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读， Single 类型；用于所有其他对象时可读/写， Single 类型。                                                                                                                                |
|  8   | [Hyperlinks](#Master.Hyperlinks)                   | 返回一个代表指定幻灯片中的所有超链接的 Hyperlinks 集合。只读。                                                                                                                                                                                     |
|  9   | [Name](#Master.Name)                               | 返回或设置指定对象的名称。 String 类型，可读/写。                                                                                                                                                                                                  |
|  10  | [Shapes](#Master.Shapes)                           | 返回一个 Shapes 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。 |
|  11  | [SlideShowTransition](#Master.SlideShowTransition) | 返回一个 SlideShowTransition 对象，该对象代表指定幻灯片切换的特殊效果。只读。                                                                                                                                                                      |
|  12  | [TextStyles](#Master.TextStyles)                   | 返回一个 TextStyles 集合，该集合代表用于指定幻灯片母版的三种文本样式 - 标题文本、正文文本和默认文本。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。                                                                                 |
|  13  | [Theme](#Master.Theme)                             | 返回一个 Theme 对象，该对象代表指定的幻灯片母版、标题母版、讲义母版、备注母版或设计母版所使用的主题。                                                                                                                                              |
|  14  | [TimeLine](#Master.TimeLine)                       | 返回 TimeLine 对象，该对象代表幻灯片的动画时间线。                                                                                                                                                                                                 |
|  15  | [Width](#Master.Width)                             | 以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读， Single 类型；用于所有其他对象时可读/写， Single 类型。                                                                                                                                |

## 成员方法

### Master.ApplyTheme

将主题或设计模板应用于指定的幻灯片母板、标题母版、讲义母版、备注母版或设计母版。

**语法**

**express.ApplyTheme ( *themeName* )**

\*express是一个代表 Master 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                      |
|-------------|-----------|----------|---------------------------------------------------------------------------|
| *themeName* | 必选      | String   | 应用于 Master 对象的主题文件 (.thmx) 或设计模板文件 (.pot) 的路径和名称。 |

**示例**

``` JavaScript
/*以下示例将保存的主题应用于幻灯片母板：*/
Application.ActivePresentation.SlideMaster.ApplyTheme("C:\\Program Files\\ Office\\Templates\\MyTheme.thmx")

/*以下示例将保存的设计模板应用于幻灯片母板：*/
Application.ActivePresentation.SlideMaster.ApplyTheme("C:\\Program Files\\ Office\\Templates\\MyTheme.pot")
```

### Master.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 Master 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

## 成员属性

### Master.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres){
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### Master.Background

返回代表幻灯片背景的 **ShapeRange** 对象。

**语法**

**express.Background**

\*express是一个代表 Master 对象的变量。

**说明**

如果要使用 **Background** 属性设置单张幻灯片的背景而不更改幻灯片母版，则该幻灯片的 **FollowMasterBackground** 属性必须设置为 **False** 。

**示例**

``` JavaScript
/*以下示例将当前演示文稿的幻灯片母版的背景设为预设的底纹。*/
Application.ActivePresentation.SlideMaster.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset)

/*以下示例将当前演示文稿中第一张幻灯片的背景设为预设的底纹。*/
function test() {
    let sli = Application.ActivePresentation.Slides.Item(1)
    sli.FollowMasterBackground = false
    sli.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset)
}
```

### Master.BackgroundStyle

设置或返回指定对象的背景样式。可读/写。

**语法**

**express.BackgroundStyle**

\*express是一个代表 Master 对象的变量。

### Master.ColorScheme

返回或设置 **ColorScheme** 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。

**语法**

**express.ColorScheme**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
/*以下示例将当前演示文稿第一张和第三张幻灯片的标题颜色设为绿色。*/
function test() {
  let mySlides = Application.ActivePresentation.Slides.Range(Array(1, 3))
  mySlides.ColorScheme.Colors(ppTitle).RGB = (0, 255, 0)
}
```

### Master.Design

返回代表设计的 **Design** 对象。

**语法**

**express.Design**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
/*以下示例添加一个标题母版。*/
function test(){
    Application.ActivePresentation.Slides.Item(1).Design.AddTitleMaster()
}
```

### Master.HeadersFooters

返回一个 **HeadersFooters** 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。

**语法**

**express.HeadersFooters**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
以下示例在当前演示文稿的备注母版中设置页脚文本、日期和时间的格式，并将日期和时间设定为可自动修改。
function test() {
  let hea = Application.ActivePresentation.NotesMaster.HeadersFooters
  hea.Footer.Text = "Regional Sales"
  let tim = hea.DateAndTime
  tim.UseFormat = true
  tim.Format = ppDateTimeHmmss
}
```

### Master.Height

以磅为单位返回或设置指定对象的高度。用于 **Master** 对象时只读， **Single** 类型；用于所有其他对象时可读/写， **Single** 类型。

**语法**

**express.Height**

\*express是一个代表 Master 对象的变量。

**说明**

一个 ****Heigh** **t**** **Shape** 对象的 ****Shape**** **Height** 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

**示例**

``` JavaScript
/*以下示例将第二个文档窗口的高度设置为应用程序窗口高度的一半。*/
Application.Windows.Item(2).Height = Application.Height / 2
```

### Master.Hyperlinks

返回一个代表指定幻灯片中的所有超链接的 **Hyperlinks** 集合。只读。

**语法**

**express.Hyperlinks**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
/*以下示例允许用户更新当前演示文稿中所有超链接中已过时的 Internet 地址。*/
function test(){
  let oldAddr = Application.prompt("Old Internet address")
  let newAddr = Application.prompt("New Internet address")
      for(let i = 1; i <= Application.ActivePresentation.Slides.Count; i++){
          for(let ii = 1; ii <= Application.ActivePresentation.Slides.Item(i).Hyperlinks.Count; ii++){
              if(LCase(Application.ActivePresentation.Slides.Item(i).Hyperlinks.Item(ii).Address) == oldAddr){
                  Application.ActivePresentation.Slides.Item(i).Hyperlinks.Item(ii).Address = newAddr
          } 
      }
  }
}
```

### Master.Name

返回或设置指定对象的名称。 **String** 类型，可读/写。

**语法**

**express.Name**

\*express是一个代表 Master 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则 .Shapes("Rectangle 2") 将返回对该形状的引用。

### Master.Shapes

返回一个 **Shapes** 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。

**语法**

**express.Shapes**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
/*以下示例添加一个宽 100 磅、高 50 磅的矩形，它的左上角距当前演示文稿第一张幻灯片的左边 5 磅、上边 25 磅。*/
function test() {
  let firstSlide = Application.ActivePresentation.Slides.Item(1)
  firstSlide.Shapes.AddShape(msoShapeRectangle, 5, 25, 100, 50)
}

/*以下示例设置当前演示文稿第一张幻灯片的第三个形状的填充纹理。*/
function test() {
  let newRect = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
  newRect.Fill.PresetTextured(msoTextureOak)
}

/*假设当前演示文稿第一张幻灯片包含一个标题，以下示例的第二行和第三行设置该演示文稿第一张幻灯片的标题文本。*/
function test() {
  let firstSl = Application.ActivePresentation.Slides.Item(1)
  firstSl.Shapes.Title.TextFrame.TextRange.Text = "Some title text"
  firstSl.Shapes.Item(1).TextFrame.TextRange.Text = "Other title text"
}

/*假设当前演示文稿第二张幻灯片中的第二个形状包含文本框架，以下示例向该幻灯片添加一系列的段落。请注意：Chr(13) 用于在该文本中插入段落标记。*/
function test() {
  let tShape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(2)
  tShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item"
}

/*对于大多数幻灯片版式，第一个形状为文本占位符。以下示例与上例完成相同功能。*/
function test() {
  let testShape = Application.ActivePresentation.Slides.Item(2).Shapes.Placeholders.Item(2)
  testShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item"
}
```

### Master.SlideShowTransition

返回一个 **SlideShowTransition** 对象，该对象代表指定幻灯片切换的特殊效果。只读。

**语法**

**express.SlideShowTransition**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
以下示例设置在幻灯片放映过程中，当前演示文稿的第二张幻灯片在 5 秒后自动换片，并在幻灯片切换时播放犬吠声
function test() {
  let sli = Application.ActivePresentation.Slides.Item(2).SlideShowTransition
  sli.AdvanceOnTime = true
  sli.AdvanceTime = 5
  sli.SoundEffect.ImportFromFile("c:\\windows\\media\\dogbark.wav")
  Application.ActivePresentation.SlideShowSettings.AdvanceMode = ppSlideShowUseSlideTimings
}
```

### Master.TextStyles

返回一个 **TextStyles** 集合，该集合代表用于指定幻灯片母版的三种文本样式 - 标题文本、正文文本和默认文本。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**语法**

**express.TextStyles**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
本示例为活动演示文稿中幻灯片上的第一级正文文本设置字体名称和字号。
function test() {
  let sli = Application.ActivePresentation.SlideMaster.TextStyles.Item(ppBodyStyle).Levels.Item(1)
  let fon = sli.Font
  fon.Name = "arial"
  fon.Size = 36
}
```

### Master.Theme

返回一个 **Theme** 对象，该对象代表指定的幻灯片母版、标题母版、讲义母版、备注母版或设计母版所使用的主题。

**语法**

**express.Theme**

\*express是一个代表 Master 对象的变量。

### Master.TimeLine

返回 **TimeLine** 对象，该对象代表幻灯片的动画时间线。

**语法**

**express.TimeLine**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
/*以下示例将一个跳动的动画添加到第一张幻灯片的第一个形状中。*/
function test(){
    let sldFirst = Application.ActivePresentation.Slides.Item(1)
    let shpFirst = sldFirst.Shapes.Item(1)
    sldFirst.TimeLine.MainSequence.AddEffect(shpFirst, msoAnimEffectBounce)
}
```

### Master.Width

以磅为单位返回或设置指定对象的宽度。用于 **Master** 对象时只读， **Single** 类型；用于所有其他对象时可读/写， **Single** 类型。

**语法**

**express.Width**

\*express是一个代表 Master 对象的变量。

**示例**

``` JavaScript
/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的文档窗口。*/
function test() {
  let ah = Application.Windows.Item(1).Height                           // available height
  let aw = Application.Windows.Item(1).Width + Windows.Item(2).Width    // available width
  Application.Windows.Item(1).Width = aw
  Application.Windows.Item(1).Height = ah / 2
  Application.Windows.Item(1).Left = 0
  Application.Windows.Item(2).Width = aw
  Application.Windows.Item(2).Height = ah / 2
  Application.Windows.Item(2).Top = ah / 2
  Application.Windows.Item(2).Left = 0
}

/*以下示例将指定表中第一列的宽度设置为 80 磅（每英寸 72 磅）。*/
Application.ActivePresentation.Slides.Item(2).Shapes.Item(5).Table.Columns.Item(1).Width = 80
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Master](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PlaySettings 对象

## 简介

包含关于指定的媒体剪辑在幻灯片放映中的播放方式的信息

## 属性

| 序号 | 名称                                                     | 说明                                                                                                                                                                                                                                         |
|:----:|:---------------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ActionVerb](#PlaySettings.ActionVerb)                   | 返回或设置包含 OLE 动作的字符串，该动作将在幻灯片放映过程中当指定的 OLE 对象进行动画显示时运行。默认动作指定 OLE 对象在上一个动画或幻灯片切换之后运行的动作（例如，播放一波形文件，或显示数据使得用户可以进行修改）。 String 类型，可读/写。 |
|  2   | [Application](#PlaySettings.Application)                 | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                      |
|  3   | [HideWhileNotPlaying](#PlaySettings.HideWhileNotPlaying) | 决定在幻灯片放映期间指定的媒体剪辑在不播放时是否隐藏。可读/写 MsoTriState 类型。                                                                                                                                                             |
|  4   | [LoopUntilStopped](#PlaySettings.LoopUntilStopped)       | 决定指定影片或声音是否持续循环播放，直到出现下列情况：开始下一个影片或声音、用户单击幻灯片或发生幻灯片切换。 MsoTriState 类型，可读/写。                                                                                                     |
|  5   | [Parent](#PlaySettings.Parent)                           | 返回指定对象的父对象。                                                                                                                                                                                                                       |
|  6   | [PauseAnimation](#PlaySettings.PauseAnimation)           | 决定在指定的媒体剪辑播放结束前是否暂停幻灯片放映。可读/写 MsoTriState 类型。                                                                                                                                                                 |
|  7   | [PlayOnEntry](#PlaySettings.PlayOnEntry)                 | 决定指定的影片或声音后是否在激活后自动播放。可读/写 MsoTriState 类型。                                                                                                                                                                       |
|  8   | [RewindMovie](#PlaySettings.RewindMovie)                 | 决定在指定的影片播放完毕后是否自动重新显示该影片的第一帧。可读/写 MsoTriState 类型。                                                                                                                                                         |
|  9   | [StopAfterSlides](#PlaySettings.StopAfterSlides)         | 返回或设置在媒体剪辑播放完毕前要放映的幻灯片数。可读/写 Long 类型。                                                                                                                                                                          |

## 成员属性

### PlaySettings.ActionVerb

返回或设置包含 OLE 动作的字符串，该动作将在幻灯片放映过程中当指定的 OLE 对象进行动画显示时运行。默认动作指定 OLE 对象在上一个动画或幻灯片切换之后运行的动作（例如，播放一波形文件，或显示数据使得用户可以进行修改）。 **String** 类型，可读/写。

**语法**

**express.ActionVerb**

\*express是一个代表 PlaySettings 对象的变量。

**示例**

``` JavaScript
/*本示例指定当前演示文稿第一张幻灯片的第三个形状被激活时将自动打开以供编辑。但第三个形状必须是包含声音或视频对象并支持“Edit”动词的 OLE 对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　let oleobj = OLEobj.AnimationSettings.PlaySettings
  　　　　  oleobj.PlayOnEntry = true
  　　　　  oleobj.ActionVerb = "Edit"
}
```

### PlaySettings.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 PlaySettings 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### PlaySettings.HideWhileNotPlaying

决定在幻灯片放映期间指定的媒体剪辑在不播放时是否隐藏。可读/写 **MsoTriState** 类型。

**语法**

**express.HideWhileNotPlaying**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。          |
|--------------------------------------------------------|
| msoCTrue                                               |
| msoFalse                                               |
| msoTriStateMixed                                       |
| msoTriStateToggle                                      |
| msoTrue 幻灯片放映过程中指定的媒体剪辑在不播放时隐藏。 |

**示例**

``` JavaScript
/*以下示例在当前演示文稿的第一张幻灯片中插入名为“Clock.avi”的影片，设置它在幻灯片切换后自动播放，并指定该影片对象在不播放时隐藏。*/
function test(){
　　　　let mydocument = Application.ActivePresentation.Slides.Item(1).Shapes.AddOLEObject(10,10,250,250,undefined,"C:\\WINNT\\clock.avi")
　　　　let mydocument1 = mydocument.AnimationSettings.PlaySettings
 　　　　   mydocument1.PlayOnEntry = true
 　　　　   mydocument1.HideWhileNotPlaying = msoTrue
}
```

### PlaySettings.LoopUntilStopped

决定指定影片或声音是否持续循环播放，直到出现下列情况：开始下一个影片或声音、用户单击幻灯片或发生幻灯片切换。 **MsoTriState** 类型，可读/写。

**语法**

**express.LoopUntilStopped**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                                              |
|--------------------------------------------------------------------------------------------|
| msoCTrue                                                                                   |
| msoFalse                                                                                   |
| msoTriStateMixed                                                                           |
| msoTriStateToggle                                                                          |
| msoTrue 指定影片或声音持续循环，直到开始下一个影片或声音、用户单击幻灯片或发生幻灯片切换。 |

**示例**

本示例将当前演示文稿第一张幻灯片第三个形状设为以动画顺序开始播放，并持续播放直到开始下一个媒体剪辑。第三个形状必须是声音或影片对象。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).AnimationSettings.PlaySettings.LoopUntilStopped = msoTrue
```

### PlaySettings.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PlaySettings 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let mydocument = myShapes.AddShape(msoShapeOval,50,50,300,150).TextFrame
   　　　　 mydocument.TextRange.Text = "Test text"
   　　　　 mydocument.Parent.Rotation = 45
}
```

### PlaySettings.PauseAnimation

决定在指定的媒体剪辑播放结束前是否暂停幻灯片放映。可读/写 **MsoTriState** 类型。

**语法**

**express.PauseAnimation**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

要使 **PauseAnimation** 属性设置生效，指定形状的 **PlayOnEntry** 属性必须设置为 **msoTrue** 。

| MsoTriState 可以是下列 MsoTriState 常量之一。      |
|----------------------------------------------------|
| msoCTrue                                           |
| msoFalse 在背景中播放媒体剪辑时继续幻灯片放映。    |
| msoTriStateMixed                                   |
| msoTriStateToggle                                  |
| msoTrue 在指定的媒体剪辑播放结束前暂停幻灯片放映。 |

**示例**

``` JavaScript
/*本示例指定当前演示文稿第一张幻灯片的第三个形状被激活后自动播放，并且当背景中播放该影片时暂停播放幻灯片放映。第三个形状必须是声音或影片对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　let OLEobj1 = OLEobj.AnimationSettings.PlaySettings
   　　　　 OLEobj1.PlayOnEntry = msoTrue
    　　　　OLEobj1.PauseAnimation = msoTrue
}
```

### PlaySettings.PlayOnEntry

决定指定的影片或声音后是否在激活后自动播放。可读/写 **MsoTriState** 类型。

**语法**

**express.PlayOnEntry**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

将此属性设置为 **msoTrue** ，从而把 **AnimationSettings** 对象的 **Animate** 属性设为 **msoTrue** 。将 **Animate** 属性设置为 **msoFalse** 可自动将 **PlayOnEntry** 属性设为 **msoFalse** 。

使用 **ActionVerb** 属性设置媒体剪辑被激活时调用的动作。

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue 指定的影片或声音被激活后自动播放。    |

**示例**

``` JavaScript
/*本示例指定当前演示文稿第一张幻灯片的第三个形状在激活后自动播放。第三个形状必须是声音或影片对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　OLEobj.AnimationSettings.PlaySettings.PlayOnEntry = msoTrue
}
```

### PlaySettings.RewindMovie

决定在指定的影片播放完毕后是否自动重新显示该影片的第一帧。可读/写 **MsoTriState** 类型。

**语法**

**express.RewindMovie**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。              |
|------------------------------------------------------------|
| msoCTrue                                                   |
| msoFalse                                                   |
| msoTriStateMixed                                           |
| msoTriStateToggle                                          |
| msoTrue 在指定的影片放映完毕后自动重新显示该影片的第一帧。 |

**示例**

``` JavaScript
/*本示例指定影片的第一帧（由当前演示文稿第一张幻灯片的第三个形状代表）在影片播放完毕后重新显示。第三个形状必须是影片对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　OLEobj.AnimationSettings.PlaySettings.RewindMovie = msoTrue
}
```

### PlaySettings.StopAfterSlides

返回或设置在媒体剪辑播放完毕前要放映的幻灯片数。可读/写 **Long** 类型。

**语法**

**express.StopAfterSlides**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

要使 **StopAfterSlides** 属性设置生效，必须将指定幻灯片的 **PauseAnimation** 属性设为 **False** ，并且必须将 **PlayOnEntry** 属性设为 **True** 。

当指定数目的幻灯片放映结束或剪辑播放完毕时（无论哪种情况先发生），媒体剪辑将停止播放。0（零）值指定在放映完当前幻灯片后剪辑停止播放。

**示例**

``` JavaScript
/*本例指定当媒体剪辑（由活动演示文稿中第一张幻灯片上的第三个形状代表）被激活时自动播放，并且当背景中播放该媒体剪辑的同时幻灯片将继续放映，还指定当放映完三张幻灯片或该剪辑播放完毕时（无论哪种情况先发生），媒体剪辑将停止播放。第三个形状必须是声音或影片对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　let OLEobj1 = OLEobj.AnimationSettings.PlaySettings
   　　　　 OLEobj1.PlayOnEntry = true
   　　　　 OLEobj1.PauseAnimation = false
   　　　　 OLEobj1.StopAfterSlides = 3
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/PlaySettings](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SlideShowView 对象

## 简介

代表幻灯片放映窗口中的视图。

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.Exit">Exit</a></td>
<td style="text-align: left;" data-valian="middle"><p>结束指定的幻灯片放映。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.First">First</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置指定的幻灯片放映视图，显示演示文稿的第一张幻灯片。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.GetClickIndex">GetClickIndex</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回启动当前正在幻灯片上播放或者刚刚已经播放完的动画的当前鼠标单击的索引号。</p>
<p>返回值类型为Long</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.GotoSlide">GotoSlide</a></td>
<td style="text-align: left;" data-valian="middle"><p>在幻灯片放映期间切换到指定幻灯片。可指定是否要重运行动画效果。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.Last">Last</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定的幻灯片放映视图设置为显示演示文稿中的上一张幻灯片。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.Next">Next</a></td>
<td style="text-align: left;" data-valian="middle"><p>显示紧接当前显示的幻灯片之后的幻灯片。如果显示的是最后一张幻灯片，则 Next 方法会关闭演讲者模式的幻灯片放映并以展台模式返回到第一张幻灯片。使用 SlideShowWindow 对象的 View 属性可返回 SlideShowView 对象。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.Previous">Previous</a></td>
<td style="text-align: left;" data-valian="middle"><p>显示紧邻当前显示的幻灯片之前的幻灯片。如果当前以展台幻灯片放映显示第一张幻灯片，使用 Previous 方法将在幻灯片放映中切换到最后一张幻灯片；如果当前显示的是演示文稿中的第一张幻灯片，则此方法无效。使用 SlideShowWindow 对象的 View 属性可返回 SlideShowView 对象。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                                                              | 说明                                                                                                                                        |
|:----:|:------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AdvanceMode](#SlideShowView.AdvanceMode)                         | 返回一个值，该值指示如何在指定视图中切换幻灯片放映。只读 PpSlideShowAdvanceMode 类型。                                                      |
|  2   | [Application](#SlideShowView.Application)                         | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                     |
|  3   | [CurrentShowPosition](#SlideShowView.CurrentShowPosition)         | 返回当前幻灯片在指定视图下显示的幻灯片放映中的位置。只读 Long 类型。                                                                        |
|  4   | [LastSlideViewed](#SlideShowView.LastSlideViewed)                 | 返回一个 Slide 对象，该对象代表指定的幻灯片放映视图中在当前幻灯片之间观看的幻灯片。                                                         |
|  5   | [PointerColor](#SlideShowView.PointerColor)                       | 返回一个 ColorFormat 对象，该对象代表在幻灯片放映过程中指定演示文稿的指针颜色。一旦幻灯片放映结束，该颜色将转变为演示文稿的默认颜色。只读。 |
|  6   | [PointerType](#SlideShowView.PointerType)                         | 返回或设置在幻灯片放映中使用的指针类型。可读/写 PpSlideShowPointerType 类型。                                                               |
|  7   | [PresentationElapsedTime](#SlideShowView.PresentationElapsedTime) | 返回指定的幻灯片放映开始后经过的秒数。只读 Long 类型。                                                                                      |
|  8   | [Slide](#SlideShowView.Slide)                                     | 返回一个 Slide 对象，该对象代表在指定的幻灯片放映窗口视图中当前显示的幻灯片。只读。                                                         |
|  9   | [SlideElapsedTime](#SlideShowView.SlideElapsedTime)               | 返回当前幻灯片已放映的秒数。可读/写 Long 类型。                                                                                             |
|  10  | [State](#SlideShowView.State)                                     | 返回或设置幻灯片放映的状态。可读/写 PpSlideShowState 类型。                                                                                 |
|  11  | [Zoom](#SlideShowView.Zoom)                                       | 以正常大小的百分比形式返回指定幻灯片放映窗口视图的缩放设置。可以为 10% 到 400% 之间的值。只读 Integer 类型。                                |

## 成员方法

### SlideShowView.Exit

结束指定的幻灯片放映。

**语法**

**express.Exit ()**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*以下示例将结束第一个幻灯片放映窗口中的幻灯片放映。*/
Application.SlideShowWindows.Item(1).View.Exit()
```

### SlideShowView.First

设置指定的幻灯片放映视图，显示演示文稿的第一张幻灯片。

**语法**

**express.First ()**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

如果使用 **First** 方法在幻灯片放映过程中从一张幻灯片切换到另一张幻灯片，返回原幻灯片时，它的动画继续从被中断处开始放映。

**示例**

``` JavaScript
本示例设置第一个幻灯片放映窗口以显示演示文稿的第一张幻灯片。
Application.SlideShowWindows.Item(1).View.First()
```

### SlideShowView.GetClickIndex

返回启动当前正在幻灯片上播放或者刚刚已经播放完的动画的当前鼠标单击的索引号。

返回值类型为Long

**语法**

**express.GetClickIndex ()**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

使用 **GetClickCount** 方法可返回为幻灯片定义的鼠标单击次数。

如果幻灯片不包含任何动画或者用户尚未换到某个动画，则 **GetClickIndex** 方法返回 0。如果幻灯片包含自动播放的动画并且该用户移到上一页，则 **GetClickIndex** 方法返回 **msoClickStateBeforeAutomaticAnimcations** 。

**示例**

``` JavaScript
Application.SlideShowWindows.Item(1).View.GetClickIndex()
```

### SlideShowView.GotoSlide

在幻灯片放映期间切换到指定幻灯片。可指定是否要重运行动画效果。

**语法**

**express.GotoSlide ( *Index* , *ResetSlide* )**

\*express是一个代表 SlideShowView 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型    | 说明                                                                                                                                                                                                                                                                              |
|--------------|-----------|-------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Index*      | 必选      | Integer     | 要切换到的幻灯片的编号。                                                                                                                                                                                                                                                          |
| *ResetSlide* | 可选      | MsoTriState | 如果将 ResetSlide 设置为 msoFalse，那么在幻灯片放映期间，当您从一张幻灯片切换到另一张，再返回到第一张幻灯片时，动画将从刚刚中断的位置继续播放。如果将 ResetSlide 设置为 msoTrue，当从一张幻灯片切换到另一张幻灯片，再返回到第一张幻灯片时，将重新播放整个动画。默认值为 msoTrue。 |

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 默认值。                          |

**示例**

``` JavaScript
/*本示例在第一个幻灯片放映窗口中从当前幻灯片切换到第三张幻灯片。如果在幻灯片放映过程中切换回当前幻灯片，将重新播放整个动画。*/
Application.SlideShowWindows.Item(1).View.GotoSlide(3)

/*本示例在第一个幻灯片放映窗口中从当前幻灯片切换到第三张幻灯片。如果在幻灯片放映过程中切换回当前幻灯片，动画将从中断位置重新播放。*/
Application.SlideShowWindows.Item(1).View.GotoSlide(3, msoFalse)
```

### SlideShowView.Last

将指定的幻灯片放映视图设置为显示演示文稿中的上一张幻灯片。

**语法**

**express.Last ()**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

如果使用 **Last** 方法在幻灯片放映过程中从一张幻灯片切换到另一张幻灯片，则在返回原幻灯片时，其中的动画继续从被中断处开始放映。

**示例**

``` JavaScript
/*本例将第一个幻灯片放映窗口设置为显示演示文稿中的上一张幻灯片。*/
Application.SlideShowWindows.Item(1).View.Last()
```

### SlideShowView.Next

显示紧接当前显示的幻灯片之后的幻灯片。如果显示的是最后一张幻灯片，则 **Next** 方法会关闭演讲者模式的幻灯片放映并以展台模式返回到第一张幻灯片。使用 **SlideShowWindow** 对象的 **View** 属性可返回 **SlideShowView** 对象。

**语法**

**express.Next ()**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*本示例显示第一个幻灯片放映窗口中紧随在当前显示的幻灯片之后的下一张幻灯片。*/
Application.SlideShowWindows.Item(1).View.Next()
```

### SlideShowView.Previous

显示紧邻当前显示的幻灯片之前的幻灯片。如果当前以展台幻灯片放映显示第一张幻灯片，使用 **Previous** 方法将在幻灯片放映中切换到最后一张幻灯片；如果当前显示的是演示文稿中的第一张幻灯片，则此方法无效。使用 **SlideShowWindow** 对象的 **View** 属性可返回 **SlideShowView** 对象。

**语法**

**express.Previous ()**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*本示例在第一个幻灯片放映窗口中放映紧挨当前显示的幻灯片之前的幻灯片。*/
Application.SlideShowWindows.Item(1).View.Previous()
```

## 成员属性

### SlideShowView.AdvanceMode

返回一个值，该值指示如何在指定视图中切换幻灯片放映。只读 **PpSlideShowAdvanceMode** 类型。

**语法**

**express.AdvanceMode**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

|                                                                     |
|---------------------------------------------------------------------|
| PpSlideShowAdvanceMode 可以是下列 PpSlideShowAdvanceMode 常量之一。 |
| **ppSlideShowManualAdvance**                                        |
| **ppSlideShowRehearseNewTimings**                                   |
| **ppSlideShowUseSlideTimings**                                      |

### SlideShowView.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test() {
  let shpOle = Application.ActivePresentation.Slides.Item(3).Shapes
  for(let i = 1; i <= shpOle.Count; i++) {
    if(shpOle.Item(i).Type == msoLinkedOLEObject) {
        MsgBox(shpOle.Item(i).OLEFormat.Application.Name)
    }
  }
}
```

### SlideShowView.CurrentShowPosition

返回当前幻灯片在指定视图下显示的幻灯片放映中的位置。只读 **Long** 类型。

**语法**

**express.CurrentShowPosition**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

如果指定视图包含自定义放映，则 **CurrentShowPosition** 属性会返回当前幻灯片在自定义放映中的位置，而不是在整个演示文稿中的位置。

**示例**

``` JavaScript
/*本示例将一个变量设置为当前幻灯片在第一个幻灯片放映窗口中放映的幻灯片放映中的位置。*/
let lastSlideSeen = Application.SlideShowWindows.Item(1).View.CurrentShowPosition
```

### SlideShowView.LastSlideViewed

返回一个 **Slide** 对象，该对象代表指定的幻灯片放映视图中在当前幻灯片之间观看的幻灯片。

**语法**

**express.LastSlideViewed**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*本示例切换到第一个幻灯片放映窗口中当前幻灯片之前的幻灯片。*/
Application.SlideShowWindows.Item(1).View.GotoSlide(SlideShowWindows.Item(1).View.LastSlideViewed.SlideIndex)
```

### SlideShowView.PointerColor

返回一个 **ColorFormat** 对象，该对象代表在幻灯片放映过程中指定演示文稿的指针颜色。一旦幻灯片放映结束，该颜色将转变为演示文稿的默认颜色。只读。

**语法**

**express.PointerColor**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

要将指针更改为绘图笔，可将 **PointerType** 属性设置为 **ppSlideShowPointerPen** 。

### SlideShowView.PointerType

返回或设置在幻灯片放映中使用的指针类型。可读/写 **PpSlideShowPointerType** 类型。

**语法**

**express.PointerType**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

|                                                                     |
|---------------------------------------------------------------------|
| PpSlideShowPointerType 可以是下列 PpSlideShowPointerType 常量之一。 |
| **ppSlideShowPointerAlwaysHidden**                                  |
| **ppSlideShowPointerArrow**                                         |
| **ppSlideShowPointerAutoArrow**                                     |
| **ppSlideShowPointerNone**                                          |
| **ppSlideShowPointerPen**                                           |

**示例**

``` JavaScript
/*本示例执行当前演示文稿的一个幻灯片放映，将指针改变为绘图笔。*/
function test() {
  let currView = Application.ActivePresentation.SlideShowSettings.Run().View
      currView.PointerType = ppSlideShowPointerPen
}
```

### SlideShowView.PresentationElapsedTime

返回指定的幻灯片放映开始后经过的秒数。只读 **Long** 类型。

**语法**

**express.PresentationElapsedTime**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*如果幻灯片放映开始已有五分钟，本示例转向第一个幻灯片放映窗口中第七张幻灯片。*/
function test() {
  if(Application.SlideShowWindows.Item(1).View.PresentationElapsedTime > 3) {
      Application.SlideShowWindows.Item(1).View.GotoSlide(3)
  }
}
```

### SlideShowView.Slide

返回一个 **Slide** 对象，该对象代表在指定的幻灯片放映窗口视图中当前显示的幻灯片。只读。

**语法**

**express.Slide**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

如果当前显示的幻灯片来源于一个嵌入演示文稿，则可以使用 Slide 对象（从 **Slide** 属性返回）的 **Parent** 属性返回包含该幻灯片的嵌入演示文稿。（ SlideShowWindow 对象或 DocumentWindow 对象的 Presentation 属性返回创建该窗口的演示文稿，而不是该嵌入演示文稿。）

### SlideShowView.SlideElapsedTime

返回当前幻灯片已放映的秒数。可读/写 **Long** 类型。

**语法**

**express.SlideElapsedTime**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

使用 **ResetSlideTime** 方法重置当前放映的幻灯片的播放时间。

**示例**

``` JavaScript
/*本示例为正在第一个幻灯片放映窗口中放映的幻灯片的播放时间设置一个变量，并显示该变量的值。*/
function test() {
  let currTime = Application.SlideShowWindows.Item(1).View.SlideElapsedTime
  alert(currTime)
}
```

### SlideShowView.State

返回或设置幻灯片放映的状态。可读/写 **PpSlideShowState** 类型。

**语法**

**express.State**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

|                                                         |
|---------------------------------------------------------|
| PpSlideShowState 可以是下列 PpSlideShowState 常量之一。 |
| **ppSlideShowBlackScreen**                              |
| **ppSlideShowDone**                                     |
| **ppSlideShowPaused**                                   |
| **ppSlideShowRunning**                                  |
| **ppSlideShowWhiteScreen**                              |

**示例**

``` JavaScript
/*本示例将第一个幻灯片放映窗口的视图状态设为黑屏。*/
Application.SlideShowWindows.Item(1).View.State = ppSlideShowBlackScreen
```

### SlideShowView.Zoom

以正常大小的百分比形式返回指定幻灯片放映窗口视图的缩放设置。可以为 10% 到 400% 之间的值。只读 **Integer** 类型。

**语法**

**express.Zoom**

\*express是一个代表 SlideShowView 对象的变量。

> WPS 开发文档： [WPS 基础接口/演示 API 参考/SlideShowView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TextFrame2 对象

## 简介

代表 Shape 、 ShapeRange 或 ChartFormat 对象的文本框。

## 说明

该对象包含文本框中的文本，还包含控制文本框对齐方式和位置的属性和方法。使用 TextFrame2 属性可返回 TextFrame2 对象。

``` JavaScript
/*下例向 myDocument 添加一个矩形，向矩形添加文本，然后设置文本框的边距。*/
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2
    rng.TextRange.Text = "Here is some test text"
    rng.MarginBottom = 10
    rng.MarginLeft = 10
    rng.MarginRight = 10
    rng.MarginTop = 10
}
```

## 方法

| 序号 | 名称                                 | 说明                                       |
|:----:|:-------------------------------------|:-------------------------------------------|
|  1   | [DeleteText](#TextFrame2.DeleteText) | 删除文本框中的文字以及所有关联的文字属性。 |

## 属性

| 序号 | 名称                                             | 说明                                                                                                |
|:----:|:-------------------------------------------------|:----------------------------------------------------------------------------------------------------|
|  1   | [Application](#TextFrame2.Application)           | 返回 Application 对象。只读。                                                                       |
|  2   | [HorizontalAnchor](#TextFrame2.HorizontalAnchor) | 返回或设置指定文本的水平定位类型。可读/写 MsoHorizontalAnchor 类型。                                |
|  3   | [MarginBottom](#TextFrame2.MarginBottom)         | 返回或设置从文本框底端到包含文本的形状的内接矩形底端的距离（以磅为单位）。可读/写 Single 类型。     |
|  4   | [MarginLeft](#TextFrame2.MarginLeft)             | 以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形的左边界的距离。可读/写 Single 类型。   |
|  5   | [MarginRight](#TextFrame2.MarginRight)           | 返回或设置从文本框右边缘到包含文本的形状的内接矩形右边缘的距离（以磅为单位）。可读/写 Single 类型。 |
|  6   | [MarginTop](#TextFrame2.MarginTop)               | 返回或设置从文本框顶端到包含文本的形状的内接矩形顶端的距离（以磅为单位）。可读/写 Single 类型。     |
|  7   | [Orientation](#TextFrame2.Orientation)           | 返回或设置一个值，该值代表文本框方向。可读/写 MsoTextOrientation 类型。                             |
|  8   | [Ruler](#TextFrame2.Ruler)                       | 返回一个 Ruler 对象，该对象代表指定文本的标尺。只读。                                               |
|  9   | [TextRange](#TextFrame2.TextRange)               | 返回 TextRange2 对象，该对象代表对象中的文本。只读。                                                |
|  10  | [ThreeD](#TextFrame2.ThreeD)                     | 返回一个 ThreeDFormat 对象，该对象包含指定文本的三维效果格式属性。只读。                            |
|  11  | [VerticalAnchor](#TextFrame2.VerticalAnchor)     | 返回或设置指定文本的垂直定位类型。可读/写 MsoVerticalAnchor 类型。                                  |
|  12  | [WarpFormat](#TextFrame2.WarpFormat)             | 返回或设置指定文本框的弯曲类型。可读/写 MsoWarpFormat 。                                            |
|  13  | [WordWrap](#TextFrame2.WordWrap)                 | 返回或设置形状边界内外的文本换行。可读/写 MsoTriState 类型。                                        |

## 成员方法

### TextFrame2.DeleteText

删除文本框中的文字以及所有关联的文字属性。

**语法**

**express.DeleteText ()**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

关联的文字属性包括 **Font** 属性，例如加粗、下划线等等。

**示例**

``` JavaScript
/*如果文本框包含文字，则此示例将删除文本框中的文字。*/
function test() {
  let deleteText = Application.ActiveSheet.Shapes.Item(1).TextFrame2
  if ( deleteText.HasText ) {
      deleteText.DeleteText()
  }
}
```

## 成员属性

### TextFrame2.Application

返回 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

如果不与对象识别符一起使用，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可将此属性与一个 OLE 自动化对象一起使用以返回该对象的应用程序）。

### TextFrame2.HorizontalAnchor

返回或设置指定文本的水平定位类型。可读/写 **MsoHorizontalAnchor** 类型。

**语法**

**express.HorizontalAnchor**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.MarginBottom

返回或设置从文本框底端到包含文本的形状的内接矩形底端的距离（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.MarginBottom**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.MarginLeft

以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形的左边界的距离。可读/写 **Single** 类型。

**语法**

**express.MarginLeft**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.MarginRight

返回或设置从文本框右边缘到包含文本的形状的内接矩形右边缘的距离（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.MarginRight**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.MarginTop

返回或设置从文本框顶端到包含文本的形状的内接矩形顶端的距离（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.MarginTop**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.Orientation

返回或设置一个值，该值代表文本框方向。可读/写 **MsoTextOrientation** 类型。

**语法**

**express.Orientation**

\*express是一个代表 TextFrame2 对象的变量。

**说明**

此属性值可设置为从 -90 到 90 度的整数值，或者以下 **<span "="" r="" rlidawscontentredir?assetid="HV100730982052&amp;CTT=11&amp;Origin=HV103325302052&quot;" style="color: rgb(0, 0, 0); text-decoration-line: none;" target="_new">MsoTextOrientation</span>** 常量之一。

### TextFrame2.Ruler

返回一个 **Ruler** 对象，该对象代表指定文本的标尺。只读。

**语法**

**express.Ruler**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.TextRange

返回 **TextRange2** 对象，该对象代表对象中的文本。只读。

**语法**

**express.TextRange**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.ThreeD

返回一个 **ThreeDFormat** 对象，该对象包含指定文本的三维效果格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.VerticalAnchor

返回或设置指定文本的垂直定位类型。可读/写 **MsoVerticalAnchor** 类型。

**语法**

**express.VerticalAnchor**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.WarpFormat

返回或设置指定文本框的弯曲类型。可读/写 **MsoWarpFormat** 。

**语法**

**express.WarpFormat**

\*express是一个代表 TextFrame2 对象的变量。

### TextFrame2.WordWrap

返回或设置形状边界内外的文本换行。可读/写 **MsoTriState** 类型。

**语法**

**express.WordWrap**

\*express是一个代表 TextFrame2 对象的变量。

> WPS 开发文档： [WPS 基础接口/演示 API 参考/TextFrame2](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomXMLParts 对象

## 简介

代表 CustomXMLPart 对象的集合。

## 说明

存在三个始终与文档一起创建的默认部件。它们是“封面”、“文档属性”和“应用程序属性”。最后两个部件在 WPS Office 的以前版本中，但现在是在 CustomXMLParts 对象集合中以 XML 形式提供

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

``` JavaScript
//以下示例将一个节点添加到 CustomXMLPart 对象（它是 CustomXMLParts 对象集合的一部分）。
let myPart = Application.ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>")
```

## 方法

| 序号 | 名称                                                   | 说明                                                  |
|:----:|:-------------------------------------------------------|:------------------------------------------------------|
|  1   | [Add](#CustomXMLParts.Add)                             | 允许您向文件添加新的 CustomXMLPart 。                 |
|  2   | [SelectByID](#CustomXMLParts.SelectByID)               | 选择与 GUID 匹配的自定义 XML 部件。                   |
|  3   | [SelectByNamespace](#CustomXMLParts.SelectByNamespace) | 选择其命名空间与搜索条件匹配的自定义 XML 部件的集合。 |

## 属性

| 序号 | 名称                                       | 说明                                                                      |
|:----:|:-------------------------------------------|:--------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLParts.Application) | 获取一个代表 CustomXMLParts 对象的容器应用程序的 Application 对象。只读。 |
|  2   | [Count](#CustomXMLParts.Count)             | 获取一个 Long 类型的值，指示 CustomXMLParts 集合中的项数。只读。          |
|  3   | [Creator](#CustomXMLParts.Creator)         | 获取一个 32 位整数，指示创建 CustomXMLParts 对象时所使用应用程序。只读。  |
|  4   | [Item](#CustomXMLParts.Item)               | 获取 CustomXMLParts 集合中的一个 CustomXMLPart 对象。只读。               |
|  5   | [Parent](#CustomXMLParts.Parent)           | 获取 CustomXMLParts 对象的 Parent 对象。只读。                            |

## 成员方法

### CustomXMLParts.Add

允许您向文件添加新的 **CustomXMLPart** 。

**语法**

**express.Add ( *XML* , *SchemaCollection* )**

\*express是一个代表 CustomXMLParts 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型                  | 说明                                          |
|--------------------|-----------|---------------------------|-----------------------------------------------|
| *XML*              | 可选      | String                    | 包含要添加到新创建的 CustomXMLPart 中的 XML。 |
| *SchemaCollection* | 可选      | CustomXMLSchemaCollection | 代表要用于验证此流的架构集。                  |

**返回值**

CustomXMLPart

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例将添加新的 CustomXMLPart，使用搜索标准选择 CustomXMLPart，然后从该部件中选择单一节点。需要有对应命名空间的custom XML part
function test()
{
    try {
    let cxp1
    let cxn

    // Example written for Word.

    // Add a custom XML part.
    Application.ActiveDocument.CustomXMLParts.Add("<custXMLPart />")

    // Returns the first custom XML part with the given root namespace.
    cxp1 = Application.ActiveDocument.CustomXMLParts.Item("urn:invoice:namespace")        

    // Get a node using XPath.                             
    cxn = cxp1.SelectSingleNode("//*[@quantity < 4]")
    }
    
    // Exception handling. Show the message and resume.
    catch(Err) 
    {
        alert(Err.Description)
    }

}
```

### CustomXMLParts.SelectByID

选择与 GUID 匹配的自定义 XML 部件。

**语法**

**express.SelectByID ( *Id* )**

\*express是一个代表 CustomXMLParts 对象的变量。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                         |
|------|-----------|----------|------------------------------|
| *Id* | 必选      | String   | 包含自定义 XML 部件的 GUID。 |

**说明**

如果具有此 ID 的自定义 XML 部件不存在，则该方法返回 **Nothing** 。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例说明了检索与 GUID 匹配的自定义 XML 部件，然后在该部件中搜索与 XPath 表达式匹配的节点。
function test()
{
    // Returns a custom xml part by its ID.Get the node matching the XPath expression.
    let cxp1 = Application.ActivePresentation.CustomXMLParts.SelectByID("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}")
    
    // Get the node matching the XPath expression.                             
    let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")
}
```

### CustomXMLParts.SelectByNamespace

选择其命名空间与搜索条件匹配的自定义 XML 部件的集合。

**语法**

**express.SelectByNamespace ( *NamespaceURI* )**

\*express是一个代表 CustomXMLParts 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明               |
|----------------|-----------|----------|--------------------|
| *NamespaceURI* | 必选      | String   | 包含命名空间 URI。 |

**说明**

如果具有此命名空间的自定义 XML 部件不存在，则该方法将返回空的 **CustomXMLParts** 集合对象。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例说明了检索与 GUID 匹配的自定义 XML 部件，然后在该部件中搜索与 XPath 表达式匹配的节点。
function test()
{
    // Returns all of the custom xml parts with the given namespace.
    let cxp1 = Application.ActiveDocument.CustomXMLParts.SelectByNamespace("urn:invoice:namespace")          
    
    // Get the node matching the XPath expression.                             
    let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")
}
```

## 成员属性

### CustomXMLParts.Application

获取一个代表 **CustomXMLParts** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLParts 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLParts.Count

获取一个 **Long** 类型的值，指示 **CustomXMLParts** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 CustomXMLParts 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLParts.Creator

获取一个 32 位整数，指示创建 **CustomXMLParts** 对象时所使用应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLParts 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLParts.Item

获取 **CustomXMLParts** 集合中的一个 **CustomXMLPart** 对象。只读。

**语法**

**express.Item**

\*express是一个代表 CustomXMLParts 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLParts.Parent

获取 **CustomXMLParts** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLParts 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLParts](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ODSOFilter 对象

## 简介

表示应用于附加的邮件合并数据源的筛选。 ODSOFilter 对象为 ODSOFilters 对象的成员。

## 说明

每个筛选器均为查询字符串中的一行。使用 Column 、 Comparison 、 CompareTo 和 Conjunction 属性可返回或设置数据源查询条件。

## 属性

| 序号 | 名称                                   | 说明                                                                                                                              |
|:----:|:---------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ODSOFilter.Application) | 获取一个 Application 对象，代表 ODSOFilter 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Column](#ODSOFilter.Column)           | 获取或设置一个 String 类型的值，该值代表邮件合并数据源中要用于筛选的字段名。可读/写。                                             |
|  3   | [CompareTo](#ODSOFilter.CompareTo)     | 获取或设置一个 String 类型的值，该值代表查询筛选条件中要比较的文本。可读/写。                                                     |
|  4   | [Comparison](#ODSOFilter.Comparison)   | 获取或设置一个代表 Column 和 CompareTo 的比较方式的 MsoFilterComparison 常量。可读/写。                                           |
|  5   | [Conjunction](#ODSOFilter.Conjunction) | 获取或设置一个 MsoFilterConjunction 常量，该常量代表 ODSOFilters 对象中的某个筛选条件与其中的其他筛选条件的关系。可读/写。        |
|  6   | [Creator](#ODSOFilter.Creator)         | 获取一个 32 位整数，指示创建 ODSOFilter 对象时所使用的应用程序。只读。                                                            |
|  7   | [Index](#ODSOFilter.Index)             | 获取一个 Long 类型的值，代表集合中 ODSOFilter 对象的索引号。只读。                                                                |
|  8   | [Parent](#ODSOFilter.Parent)           | 获取 ODSOFilter 对象的 Parent 对象。只读。                                                                                        |

## 成员属性

### ODSOFilter.Application

获取一个 **Application** 对象，代表 **ODSOFilter** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 ODSOFilter 对象的变量。

### ODSOFilter.Column

获取或设置一个 **String** 类型的值，该值代表邮件合并数据源中要用于筛选的字段名。可读/写。

**语法**

**express.Column**

\*express是一个代表 ODSOFilter 对象的变量。

**示例**

### ODSOFilter.CompareTo

获取或设置一个 **String** 类型的值，该值代表查询筛选条件中要比较的文本。可读/写。

**语法**

**express.CompareTo**

\*express是一个代表 ODSOFilter 对象的变量。

**示例**

### ODSOFilter.Comparison

获取或设置一个代表 **Column** 和 **CompareTo** 的比较方式的 **MsoFilterComparison** 常量。可读/写。

**语法**

**express.Comparison**

\*express是一个代表 ODSOFilter 对象的变量。

**示例**

### ODSOFilter.Conjunction

获取或设置一个 **MsoFilterConjunction** 常量，该常量代表 **ODSOFilters** 对象中的某个筛选条件与其中的其他筛选条件的关系。可读/写。

**语法**

**express.Conjunction**

\*express是一个代表 ODSOFilter 对象的变量。

**示例**

### ODSOFilter.Creator

获取一个 32 位整数，指示创建 **ODSOFilter** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 ODSOFilter 对象的变量。

### ODSOFilter.Index

获取一个 **Long** 类型的值，代表集合中 **ODSOFilter** 对象的索引号。只读。

**语法**

**express.Index**

\*express是一个代表 ODSOFilter 对象的变量。

### ODSOFilter.Parent

获取 **ODSOFilter** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ODSOFilter 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/ODSOFilter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PickerDialog 对象

## 简介

提供用于选择人员或选择数据的对话框用户界面功能。

## 说明

通过 Application 对象的 PickerDialog 属性获取 PickerDialog 对象。

下面的代码设置选取器对话框的属性，然后显示选取器对话框。

``` JavaScript
// Configure the Picker Dialog properties.
function test()
{
    let objPickerDialog = Application.PickerDialog
    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
    objPickerDialog.Title = "Sample Picker Dialog"
    let objPickerProperties = objPickerDialog.Properties
    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)
    let objPickerExistingResults = objPickerDialog.CreatePickerResults()
    let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")
 
    // Show the Picker Dialog and get the results.
    let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)
}
```

## 方法

| 序号 | 名称                                                     | 说明                                                   |
|:----:|:---------------------------------------------------------|:-------------------------------------------------------|
|  1   | [CreatePickerResults](#PickerDialog.CreatePickerResults) | 创建一个空的 PickerResults 对象。                      |
|  2   | [Resolve](#PickerDialog.Resolve)                         | 使用选取器对话框解析标记，并检索结果。                 |
|  3   | [Show](#PickerDialog.Show)                               | 使用已指定的数据处理程序和给定的选项显示选取器对话框。 |

## 属性

| 序号 | 名称                                         | 说明                                                                         |
|:----:|:---------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#PickerDialog.Application)     | 获取一个 Application 对象，该对象代表 PickerDialog 对象的容器应用程序。只读  |
|  2   | [Creator](#PickerDialog.Creator)             | 获取一个 32 位整数，该整数指示在其中创建了 PickerDialog 对象的应用程序。只读 |
|  3   | [DataHandlerId](#PickerDialog.DataHandlerId) | 设置或获取选取器对话框数据处理程序组件的 GUID。可读/写                       |
|  4   | [Properties](#PickerDialog.Properties)       | 返回 PickerProperties 对象，以指定数据处理程序组件的自定义属性。只读         |
|  5   | [Title](#PickerDialog.Title)                 | 设置或返回在选取器对话框中显示的选取器对话框的标题。可读/写                  |

## 成员方法

### PickerDialog.CreatePickerResults

创建一个空的 **PickerResults** 对象。

**语法**

**express.CreatePickerResults ()**

\*express是一个代表 PickerDialog 对象的变量。

**说明**

可以将 PickerResult 添加到返回的对象中，并如同 **PickerDialog** 对象已存在的结果一样，将其指定为 **Show** 方法的第二个参数。

**示例**

``` JavaScript
//下面的代码设置选取器对话框的各种属性，并将已存在的 PickerResults 添加到结果中。
function test()
{
   let objPickerDialog = Application.PickerDialog
   objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
   objPickerDialog.Title = "Sample Picker Dialog"
   let objPickerExistingResults = objPickerDialog.CreatePickerResults()
   let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")
   let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)
}
```

### PickerDialog.Resolve

使用选取器对话框解析标记，并检索结果。

**语法**

**express.Resolve ( *TokenText* , *duplicateDlgMode* )**

\*express是一个代表 PickerDialog 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                 |
|--------------------|-----------|----------|----------------------|
| *TokenText*        | 必选      | String   | 要解析的文本字符串。 |
| *duplicateDlgMode* | 必选      | Integer  |                      |

**返回值**

PickerResults

**示例**

``` JavaScript
//使用选取器对话框对象解析实体。
function test()
{
   // Configure the Picker Dialog properties.
   let objPickerDialog = Application.PickerDialog
   objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
   objPickerDialog.Title = "Sample Picker Dialog"
   let objPickerProperties = objPickerDialog.Properties
   let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)
   // Resolve the token by using Picker Dialog and get the results.
   let objPickerResults = objPickerDialog.Resolve("johndoe", false)
}
```

### PickerDialog.Show

使用已指定的数据处理程序和给定的选项显示选取器对话框。

**语法**

**express.Show ( *sMultiSelect* , *ExistingResults* )**

\*express是一个代表 PickerDialog 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型      | 说明                                                                           |
|-------------------|-----------|---------------|--------------------------------------------------------------------------------|
| *sMultiSelect*    | 可选      | Boolean       | 指定选取器对话框用户界面是否提供多项选择功能。                                 |
| *ExistingResults* | 必选      | PickerResults | 包含选取器对话框用户界面内的现有 PickerResults。这些结果在选定的项控件中显示。 |

**返回值**

PickerResults

**示例**

``` JavaScript
//下面的代码设置选取器对话框的属性，然后显示选取器对话框。
function test()
{
   // Configure the Picker Dialog properties.
   let objPickerDialog = Application.PickerDialog
   objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
   objPickerDialog.Title = "Sample Picker Dialog"
   let objPickerProperties = objPickerDialog.Properties
   let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)
   let objPickerExistingResults = objPickerDialog.CreatePickerResults()
   let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")
 
   // Show the Picker Dialog and get the results.
   let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)
}
```

## 成员属性

### PickerDialog.Application

获取一个 **Application** 对象，该对象代表 **PickerDialog** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PickerDialog 对象的变量。

### PickerDialog.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PickerDialog** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PickerDialog 对象的变量。

### PickerDialog.DataHandlerId

设置或获取选取器对话框数据处理程序组件的 GUID。可读/写

**语法**

**express.DataHandlerId**

\*express是一个代表 PickerDialog 对象的变量。

**说明**

必须先指定 **DataHandlerID** ，之后才能调用选取器对话框。

**示例**

``` JavaScript
//下面的代码设置选取器对话框的属性，然后显示选取器对话框。
function test()
{
   // Configure the Picker Dialog properties.
   let objPickerDialog = Application.PickerDialog
   objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
   objPickerDialog.Title = "Sample Picker Dialog"
   let objPickerProperties = objPickerDialog.Properties
   let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)
   let objPickerExistingResults = objPickerDialog.CreatePickerResults()
   let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")
 
   // Show the Picker Dialog and get the results.
   let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)
}
```

### PickerDialog.Properties

返回 **PickerProperties** 对象，以指定数据处理程序组件的自定义属性。只读

**语法**

**express.Properties**

\*express是一个代表 PickerDialog 对象的变量。

**说明**

**PickerProperties** 对象的属性将传递给数据处理程序。

**示例**

``` JavaScript
//下面的代码设置选取器对话框的各种属性并检索结果。
function test()
{
   let objPickerDialog = Application.PickerDialog
   objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
   objPickerDialog.Title = "Sample Picker Dialog"
   let objPickerProperties = objPickerDialog.Properties
   let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)
   
   // Show the Picker Dialog with no existing result.
   let objPickerResults = objPickerDialog.Show(true)
}
```

### PickerDialog.Title

设置或返回在选取器对话框中显示的选取器对话框的标题。可读/写

**语法**

**express.Title**

\*express是一个代表 PickerDialog 对象的变量。

**示例**

``` JavaScript
//下面的代码设置选取器对话框的属性，然后显示选取器对话框。
function test()
{
   // Configure the Picker Dialog properties.
   let objPickerDialog = Application.PickerDialog
   objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
   objPickerDialog.Title = "Sample Picker Dialog"
   let objPickerProperties = objPickerDialog.Properties
   let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)
   let objPickerExistingResults = objPickerDialog.CreatePickerResults()
   let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")
 
   // Show the Picker Dialog and get the results.
   let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)
}
```

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PickerDialog](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PictureEffect 对象

## 简介

代表图片效果。

## 说明

图片效果被处理为由各个项构成的链，以便创下面的代码将设置 Wpp 幻灯片中某个形状的几个图片效果填充属性。

``` JavaScript
function test() {
    // Setup a slide with one picture shape.
    let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects
    // Insert a 150% Saturation effect.
    pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5
    // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.
    let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)
    brightnessContrast.EffectParameters.Item(1).Value = -0.5
    brightnessContrast.EffectParameters.Item(2).Value = 0.25
}
```

建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。

## 方法

| 序号 | 名称                            | 说明           |
|:----:|:--------------------------------|:---------------|
|  1   | [Delete](#PictureEffect.Delete) | 删除图片效果。 |

## 属性

| 序号 | 名称                                                | 说明                                                                          |
|:----:|:----------------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#PictureEffect.Application)           | 获取一个 Application 对象，该对象代表 PictureEffect 对象的容器应用程序。只读  |
|  2   | [Creator](#PictureEffect.Creator)                   | 获取一个 32 位整数，该整数指示在其中创建了 PictureEffect 对象的应用程序。只读 |
|  3   | [EffectParameters](#PictureEffect.EffectParameters) | 返回一个 EffectParameter 对象。只读                                           |
|  4   | [Position](#PictureEffect.Position)                 | 指定图片效果在复合效果链中的位置。可读/写                                     |
|  5   | [Type](#PictureEffect.Type)                         | 指定图片效果的类型。只读                                                      |
|  6   | [Visible](#PictureEffect.Visible)                   | 获取或设置一个代表图片效果的可视状态的 Boolean 值。可读/写                    |

## 成员方法

### PictureEffect.Delete

删除图片效果。

**语法**

**express.Delete ()**

\*express是一个代表 PictureEffect 对象的变量。

## 成员属性

### PictureEffect.Application

获取一个 **Application** 对象，该对象代表 **PictureEffect** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PictureEffect 对象的变量。

**说明**

### PictureEffect.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PictureEffect** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PictureEffect 对象的变量。

### PictureEffect.EffectParameters

返回一个 **EffectParameter** 对象。只读

**语法**

**express.EffectParameters**

\*express是一个代表 PictureEffect 对象的变量。

**说明**

图片效果被处理为由各个项构成的链，以便创建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。效果参数指定这些效果的属性。

**示例**

``` JavaScript
//下面的代码将设置 Microsoft PowerPoint 幻灯片中某个形状的几个图片效果填充属性。
function test() {
    // Setup a slide with one picture shape.
    let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects
    // Insert a 150% Saturation effect.
    pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5
    // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.
    let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)
    brightnessContrast.EffectParameters.Item(1).Value = -0.5
    brightnessContrast.EffectParameters.Item(2).Value = 0.25
}
```

### PictureEffect.Position

指定图片效果在复合效果链中的位置。可读/写

**语法**

**express.Position**

\*express是一个代表 PictureEffect 对象的变量。

**说明**

图片效果被处理为由各个项构成的链，以便创建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。

**示例**

``` JavaScript
//下面的代码将设置 Microsoft PowerPoint 幻灯片中某个形状的几个图片效果填充属性。
function test() {
    // Setup a slide with one picture shape.
    let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects
    // Insert a 150% Saturation effect.
    pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5
    // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.
    let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)
    brightnessContrast.EffectParameters.Item(1).Value = -0.5    brightnessContrast.EffectParameters.Item(2).Value = 0.25
}
```

### PictureEffect.Type

指定图片效果的类型。只读

**语法**

**express.Type**

\*express是一个代表 PictureEffect 对象的变量。

**说明**

此属性使用 **MsoPictureEffectType** 枚举。

### PictureEffect.Visible

获取或设置一个代表图片效果的可视状态的 Boolean 值。可读/写

**语法**

**express.Visible**

\*express是一个代表 PictureEffect 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PictureEffect](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SignatureInfo 对象

## 简介

代表用于创建数字签名或文档内签名的信息。

## 方法

| 序号 | 名称                                                                                      | 说明                                                                     |
|:----:|:------------------------------------------------------------------------------------------|:-------------------------------------------------------------------------|
|  1   | [GetCertificateDetail](#SignatureInfo.GetCertificateDetail)                               | 显示与数字证书相关的指定详细信息。                                       |
|  2   | [GetSignatureDetail](#SignatureInfo.GetSignatureDetail)                                   | 显示与签名相关的指定详细信息。                                           |
|  3   | [SelectCertificateDetailByThumbprint](#SignatureInfo.SelectCertificateDetailByThumbprint) | 在通过个性特征对用户进行验证后，显示一个包含有关数字证书的信息的对话框。 |
|  4   | [SelectSignatureCertificate](#SignatureInfo.SelectSignatureCertificate)                   | 显示一个对话框，允许用户选择要用于签署文档的签名证书。                   |
|  5   | [ShowSignatureCertificate](#SignatureInfo.ShowSignatureCertificate)                       |                                                                          |

## 属性

| 序号 | 名称                                                                            | 说明                                                                                        |
|:----:|:--------------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------|
|  1   | [Application](#SignatureInfo.Application)                                       | 获取一个代表 SignatureInfo 对象的容器应用程序的 Application 对象。只读。                    |
|  2   | [CertificateVerificationResults](#SignatureInfo.CertificateVerificationResults) | 获取验证数字证书的结果。只读。                                                              |
|  3   | [ContentVerificationResults](#SignatureInfo.ContentVerificationResults)         | 获取一个代表已签名文档的散列内容的验证结果的值。                                            |
|  4   | [Creator](#SignatureInfo.Creator)                                               | 获取一个 32 位整数，指示创建 SignatureInfo 对象时所使用的应用程序。只读。                   |
|  5   | [IsCertificateExpired](#SignatureInfo.IsCertificateExpired)                     | 获取一个 Boolean 类型的值，指示数字证书是否已到期。只读。                                   |
|  6   | [IsCertificateRevoked](#SignatureInfo.IsCertificateRevoked)                     | 获取一个 Boolean 类型的值，指示数字证书是否已吊销。只读。                                   |
|  7   | [IsCertificateUntrusted](#SignatureInfo.IsCertificateUntrusted)                 | 获取一个 Boolean 类型的值，指示用于以数字方式签署文档的数字证书是否来自受信任的来源。只读。 |
|  8   | [IsValid](#SignatureInfo.IsValid)                                               | 获取一个 Boolean 类型的值，指示在执行了签名验证过程后签名是否已成功获得验证。只读。         |
|  9   | [ReadOnly](#SignatureInfo.ReadOnly)                                             | 获取一个 Boolean 类型的值，指示 SignatureInfo 对象是否为只读。只读。                        |
|  10  | [SignatureComment](#SignatureInfo.SignatureComment)                             | 获取或设置包含纳入签名数据包中的注释的值。可读/写。                                         |
|  11  | [SignatureImage](#SignatureInfo.SignatureImage)                                 | 获取或设置用于签署文档的图像的值。可读/写。                                                 |
|  12  | [SignatureProvider](#SignatureInfo.SignatureProvider)                           | 获取一个标识所安装的签名提供程序加载项的值。只读。                                          |
|  13  | [SignatureText](#SignatureInfo.SignatureText)                                   | 获取或设置用于签署此文档的签名文字的值。可读/写。                                           |

## 成员方法

### SignatureInfo.GetCertificateDetail

显示与数字证书相关的指定详细信息。

**语法**

**express.GetCertificateDetail ( *certdet* )**

\*express是一个代表 SignatureInfo 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型          | 说明                                   |
|-----------|-----------|-------------------|----------------------------------------|
| *certdet* | 必选      | CertificateDetail | 一个指定要显示的证书详细信息的枚举值。 |

**返回值**

Variant

**示例**

### SignatureInfo.GetSignatureDetail

显示与签名相关的指定详细信息。

**语法**

**express.GetSignatureDetail ( *sigdet* )**

\*express是一个代表 SignatureInfo 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型        | 说明                                   |
|----------|-----------|-----------------|----------------------------------------|
| *sigdet* | 必选      | SignatureDetail | 一个指定要显示的签名详细信息的枚举值。 |

**返回值**

Variant

**示例**

### SignatureInfo.SelectCertificateDetailByThumbprint

在通过个性特征对用户进行验证后，显示一个包含有关数字证书的信息的对话框。

**语法**

**express.SelectCertificateDetailByThumbprint ( *bstrThumbprint* )**

\*express是一个代表 SignatureInfo 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                   |
|------------------|-----------|----------|----------------------------------------|
| *bstrThumbprint* | 必选      | String   | 包含有关由个性特征标识的签署人的信息。 |

**示例**

### SignatureInfo.SelectSignatureCertificate

显示一个对话框，允许用户选择要用于签署文档的签名证书。

**语法**

**express.SelectSignatureCertificate ( *ParentWindow* )**

\*express是一个代表 SignatureInfo 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型   | 说明                               |
|----------------|-----------|------------|------------------------------------|
| *ParentWindow* | 必选      | IOleWindow | 包含证书选择对话框所在窗口的句柄。 |

**示例**

### SignatureInfo.ShowSignatureCertificate

**语法**

**express.ShowSignatureCertificate ( *ParentWindow* )**

\*express是一个代表 SignatureInfo 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型   | 说明                             |
|----------------|-----------|------------|----------------------------------|
| *ParentWindow* | 必选      | IOleWindow | 包含“证书”对话框所在窗口的句柄。 |

**示例**

## 成员属性

### SignatureInfo.Application

获取一个代表 **SignatureInfo** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.CertificateVerificationResults

获取验证数字证书的结果。只读。

**语法**

**express.CertificateVerificationResults**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.ContentVerificationResults

获取一个代表已签名文档的散列内容的验证结果的值。

**语法**

**express.ContentVerificationResults**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.Creator

获取一个 32 位整数，指示创建 **SignatureInfo** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.IsCertificateExpired

获取一个 **Boolean** 类型的值，指示数字证书是否已到期。只读。

**语法**

**express.IsCertificateExpired**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.IsCertificateRevoked

获取一个 **Boolean** 类型的值，指示数字证书是否已吊销。只读。

**语法**

**express.IsCertificateRevoked**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.IsCertificateUntrusted

获取一个 **Boolean** 类型的值，指示用于以数字方式签署文档的数字证书是否来自受信任的来源。只读。

**语法**

**express.IsCertificateUntrusted**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.IsValid

获取一个 **Boolean** 类型的值，指示在执行了签名验证过程后签名是否已成功获得验证。只读。

**语法**

**express.IsValid**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.ReadOnly

获取一个 **Boolean** 类型的值，指示 **SignatureInfo** 对象是否为只读。只读。

**语法**

**express.ReadOnly**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.SignatureComment

获取或设置包含纳入签名数据包中的注释的值。可读/写。

**语法**

**express.SignatureComment**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.SignatureImage

获取或设置用于签署文档的图像的值。可读/写。

**语法**

**express.SignatureImage**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.SignatureProvider

获取一个标识所安装的签名提供程序加载项的值。只读。

**语法**

**express.SignatureProvider**

\*express是一个代表 SignatureInfo 对象的变量。

### SignatureInfo.SignatureText

获取或设置用于签署此文档的签名文字的值。可读/写。

**语法**

**express.SignatureText**

\*express是一个代表 SignatureInfo 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SignatureInfo](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# WorkflowTasks 对象

## 简介

代表 WorkflowTask 对象的集合。

## 说明

``` JavaScript
/*以下示例显示当前文档中每个工作流任务的名称，然后显示特定任务的工作流任务编辑用户界面。应该注意到，调用 GetWorkflowTasks 方法涉及到服务器的往返行程。*/
function test(){
    const objWorkflowTasks = Application.ActiveWorkbook.GetWorkflowTasks();
    for(let i = 1; i <= objWorkflowTasks.Count; i++)
        Debug.Print(objWorkflowTasks.Item(i).Name);
        
    var objWorkflowTask = objWorkflowTasks.Item(1);
    objWorkflowTask.Show();
}
```

## 方法

| 序号 | 名称                        | 说明                                                |
|:----:|:----------------------------|:----------------------------------------------------|
|  1   | [Item](#WorkflowTasks.Item) | 获取 WorkflowTasks 集合中的一个 WorkflowTask 对象。 |

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorkflowTasks.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取一个代表 WorkflowTasks 对象的容器应用程序的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorkflowTasks.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; } &lt;STYLE TYPE="text/CSS"&gt; .ACICollapsed { display: inline; } .ACECollapsed { display: block; } &lt;/STYLE&gt;</p>
<p>获取一个 Long 类型的值，指示 WorkflowTasks 集合中的项数。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorkflowTasks.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取一个 32 位整数，指示创建 WorkflowTasks 对象时所使用的应用程序。只读。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### WorkflowTasks.Item

获取 **WorkflowTasks** 集合中的一个 **WorkflowTask** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 WorkflowTasks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                 |
|---------|-----------|----------|--------------------------------------|
| *Index* | 必选      | Long     | 要返回的 WorkflowTask 对象的索引号。 |

**返回值**

WorkflowTask

## 成员属性

### WorkflowTasks.Application

获取一个代表 **WorkflowTasks** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 WorkflowTasks 对象的变量。

### WorkflowTasks.Count

table { word-break:break-all; } \<STYLE TYPE="text/CSS"\> .ACICollapsed { display: inline; } .ACECollapsed { display: block; } \</STYLE\>

获取一个 **Long** 类型的值，指示 **WorkflowTasks** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 WorkflowTasks 对象的变量。

### WorkflowTasks.Creator

获取一个 32 位整数，指示创建 **WorkflowTasks** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 WorkflowTasks 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/WorkflowTasks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# WebShape 对象

## 简介

代表容器应用程序的一个图形，该图形是一个浏览器对象，可以直接渲染Html网页。

## 说明

使用AddWebShapeEx(str)可以为当前文档添加一个 WebShape 对象，其中str表示需要渲染的网址。下面的示例表示像当前的图形中添加一个 WebShape 对象，渲染 "https://www.wps.cn"。

``` JavaScript
//需要用户至少有一个sheet
Application.ActiveWorkbook.Sheets.Item(1).Shapes.AddWebShapeEx("https://www.wps.cn")
```

## 方法

| 序号 | 名称                                               | 说明                                                 |
|:----:|:---------------------------------------------------|:-----------------------------------------------------|
|  1   | [Copy](#WebShape.Copy)                             | 复制当前WebShape对象至剪切板。                       |
|  2   | [Cut](#WebShape.Cut)                               | 剪切当前WebShape对象至剪切板。                       |
|  3   | [Delete](#WebShape.Delete)                         | 删除当前WebShape对象。                               |
|  4   | [DeleteProperty](#WebShape.DeleteProperty)         | 删除WebShape对象中对象属性的值。                     |
|  5   | [DeleteWatchingData](#WebShape.DeleteWatchingData) | 删除对应名字的监听区域。（已废除）                   |
|  6   | [GetAllProperty](#WebShape.GetAllProperty)         | 获取当前WebShape对象的所有属性，返回JSON格式的字符串 |
|  7   | [GetAllWatchingData](#WebShape.GetAllWatchingData) | 获取所有的监听区域。（已废除）                       |
|  8   | [GetProperty](#WebShape.GetProperty)               | 获取WebShape对象对应键的属性。                       |
|  9   | [GetWatchingData](#WebShape.GetWatchingData)       | 获取对应的监听区域。（已废除）                       |
|  10  | [SaveAsPicture](#WebShape.SaveAsPicture)           | 将当前WebShape对象保存为图片。                       |
|  11  | [Select](#WebShape.Select)                         | 选中当前WebShape对象。                               |
|  12  | [SetImageData](#WebShape.SetImageData)             | 读取图片，并将其展示在WebShape对象上。               |
|  13  | [SetProperty](#WebShape.SetProperty)               | 设置WebShape属性的键值对。                           |
|  14  | [SetWatchingData](#WebShape.SetWatchingData)       | 设置监听区域。（已废除）                             |

## 属性

| 序号 | 名称                               | 说明                                         |
|:----:|:-----------------------------------|:---------------------------------------------|
|  1   | [DataSource](#WebShape.DataSource) | 返回DataSource对象。只读。                   |
|  2   | [Height](#WebShape.Height)         | 返回或设置WebShape对象的高。可读写。         |
|  3   | [ID](#WebShape.ID)                 | 返回WebShape对象的ID值。只读。               |
|  4   | [Type](#WebShape.Type)             | 返回WebShape对象的类型。                     |
|  5   | [Url](#WebShape.Url)               | 返回WebShape对象所渲染的网页网址。只读。     |
|  6   | [Version](#WebShape.Version)       | 返回当前版本。只读。                         |
|  7   | [Visible](#WebShape.Visible)       | 返回或设置当前WebShape对象的可见性。可读写。 |
|  8   | [Width](#WebShape.Width)           | 返回或设置WebShape对象的宽度。               |

## 成员方法

### WebShape.Copy

复制当前WebShape对象至剪切板。

**语法**

**express.Copy ()**

\*express是一个代表 WebShape 对象的变量。

**示例**

``` JavaScript
//以下示例复制Sheet中第一个WebShape并粘贴
function test(){
    let webShape = Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape;
    webShape.Copy();
    Application.ActiveWorkbook.ActiveSheet.Paste();
}
```

### WebShape.Cut

剪切当前WebShape对象至剪切板。

**语法**

**express.Cut ()**

\*express是一个代表 WebShape 对象的变量。

**示例**

``` JavaScript
//以下示例剪切Sheet中第一个WebShape并粘贴
function test(){
    let webShape = Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape;
    webShape.Cut();
    Application.ActiveWorkbook.ActiveSheet.Paste();
}
```

### WebShape.Delete

删除当前WebShape对象。

**语法**

**express.Delete ()**

\*express是一个代表 WebShape 对象的变量。

**示例**

``` JavaScript
//以下示例删除Sheet中第一个WebShape
function test(){
    let webShape = Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape;
    webShape.Delete();
}
```

### WebShape.DeleteProperty

删除WebShape对象中对象属性的值。

**语法**

**express.DeleteProperty ( *Key* )**

\*express是一个代表 WebShape 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明       |
|-------|-----------|----------|------------|
| *Key* | 必选      | String   | 属性的键。 |

**示例**

``` JavaScript
//以下示例删除WebShape中键为wps的属性
function test(){
    let webShape = Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape;
    webShape.DeleteProperty("wps");
}
```

### WebShape.DeleteWatchingData

删除对应名字的监听区域。（已废除）

**语法**

**express.DeleteWatchingData ( *Name* )**

\*express是一个代表 WebShape 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Name* | 必选      | String   | 监听区域的名字 |

### WebShape.GetAllProperty

获取当前WebShape对象的所有属性，返回JSON格式的字符串

**语法**

**express.GetAllProperty ()**

\*express是一个代表 WebShape 对象的变量。

**返回值**

String

**示例**

``` JavaScript
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.GetAllProperty()
```

### WebShape.GetAllWatchingData

获取所有的监听区域。（已废除）

**语法**

**express.GetAllWatchingData ()**

\*express是一个代表 WebShape 对象的变量。

### WebShape.GetProperty

获取WebShape对象对应键的属性。

**语法**

**express.GetProperty ( *Key* )**

\*express是一个代表 WebShape 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明       |
|-------|-----------|----------|------------|
| *Key* | 必选      | String   | 属性的键。 |

**示例**

``` JavaScript
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.GetProperty("wps")
```

### WebShape.GetWatchingData

获取对应的监听区域。（已废除）

**语法**

**express.GetWatchingData ()**

\*express是一个代表 WebShape 对象的变量。

### WebShape.SaveAsPicture

将当前WebShape对象保存为图片。

**语法**

**express.SaveAsPicture ()**

\*express是一个代表 WebShape 对象的变量。

**示例**

``` JavaScript
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.SaveAsPicture()
```

### WebShape.Select

选中当前WebShape对象。

**语法**

**express.Select ()**

\*express是一个代表 WebShape 对象的变量。

**示例**

``` JavaScript
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.Select()
```

### WebShape.SetImageData

读取图片，并将其展示在WebShape对象上。

**语法**

**express.SetImageData ( *Path* , *X* , *Y* , *Width* , *Height* )**

\*express是一个代表 WebShape 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明             |
|----------|-----------|----------|------------------|
| *Path*   | 必选      | String   | 图片路径         |
| *X*      | 必选      | Number   | 图片展示的横坐标 |
| *Y*      | 必选      | Number   | 图片展示的纵坐标 |
| *Width*  | 必选      | Number   | 图片展示的宽     |
| *Height* | 必选      | Number   | 图片展示的高     |

**示例**

``` JavaScript
//以下示例将c://tmp.png的图片读取为边长为100的正方形，并显示在WebShape的坐标（100，100）处
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.SetImageData("c://tmp.png", 100, 100, 100, 100)
```

### WebShape.SetProperty

设置WebShape属性的键值对。

**语法**

**express.SetProperty ( *Key* , *Value* )**

\*express是一个代表 WebShape 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明     |
|---------|-----------|----------|----------|
| *Key*   | 必选      | String   | 属性的键 |
| *Value* | 可选      | String   | 属性的值 |

**示例**

``` JavaScript
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.SetProperty("wps", "wps")
```

### WebShape.SetWatchingData

设置监听区域。（已废除）

**语法**

**express.SetWatchingData ()**

\*express是一个代表 WebShape 对象的变量。

## 成员属性

### WebShape.DataSource

返回DataSource对象。只读。

**语法**

**express.DataSource**

\*express是一个代表 WebShape 对象的变量。

**类型**

Object

### WebShape.Height

返回或设置WebShape对象的高。可读写。

**语法**

**express.Height**

\*express是一个代表 WebShape 对象的变量。

**类型**

Number

**示例**

``` JavaScript
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.Height = 300
```

### WebShape.ID

返回WebShape对象的ID值。只读。

**语法**

**express.ID**

\*express是一个代表 WebShape 对象的变量。

### WebShape.Type

返回WebShape对象的类型。

**语法**

**express.Type**

\*express是一个代表 WebShape 对象的变量。

**类型**

Number

**说明**

| 枚举           | 值  |
|----------------|-----|
| WET_Normal     | 0   |
| WET_DataSource | 1   |

**示例**

``` JavaScript
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.Type
```

### WebShape.Url

返回WebShape对象所渲染的网页网址。只读。

**语法**

**express.Url**

\*express是一个代表 WebShape 对象的变量。

**类型**

String

**示例**

``` JavaScript
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.Url
```

### WebShape.Version

返回当前版本。只读。

**语法**

**express.Version**

\*express是一个代表 WebShape 对象的变量。

**类型**

String

### WebShape.Visible

返回或设置当前WebShape对象的可见性。可读写。

**语法**

**express.Visible**

\*express是一个代表 WebShape 对象的变量。

**类型**

Boolean

**说明**

应注意在jsapi返回时，-1代表真，0代表假。

**示例**

``` JavaScript
//以下示例隐藏WebShape对象
function test(){
    Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.Visible = 0;
}
```

### WebShape.Width

返回或设置WebShape对象的宽度。

**语法**

**express.Width**

\*express是一个代表 WebShape 对象的变量。

**类型**

Number

**示例**

``` JavaScript
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.Width = 300
```

> WPS 开发文档： [WPS 基础接口/WPS 扩展 API/公共/WebShape](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
