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
