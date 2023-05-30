# DataSource 对象

## 简介

代表容器应用程序的一个数据源对象。

## 方法

| 序号 | 名称                                                   | 说明                                      |
|:----:|:-------------------------------------------------------|:------------------------------------------|
|  1   | [Activate](#DataSource.Activate)                       | 激活数据源，跨进程表现为激活ET文档。      |
|  2   | [ApplyDemoRange](#DataSource.ApplyDemoRange)           | 往活动单元格填充示例数据。返回DataRange。 |
|  3   | [CreateDataRange](#DataSource.CreateDataRange)         | 为选定的单元格创建DataRange对象。         |
|  4   | [GetAllRangeInfo](#DataSource.GetAllRangeInfo)         | 获取所有已注册的DataRange。               |
|  5   | [GetDataRange](#DataSource.GetDataRange)               | 返回已注册的对应DataRange。               |
|  6   | [InitDemoRange](#DataSource.InitDemoRange)             | 初始化示例数据的行列数                    |
|  7   | [RegisterDataRange](#DataSource.RegisterDataRange)     | 向当前DataSource对象注册DataRange对象。   |
|  8   | [SetDemoRangeItem](#DataSource.SetDemoRangeItem)       | 填充示例数据。                            |
|  9   | [UnRegisterDataRange](#DataSource.UnRegisterDataRange) | 反注册DataRange。                         |

## 属性

| 序号 | 名称                     | 说明                         |
|:----:|:-------------------------|:-----------------------------|
|  1   | [Type](#DataSource.Type) | 返回DataSource的类型。只读。 |

## 成员方法

### DataSource.Activate

激活数据源，跨进程表现为激活ET文档。

**语法**

**express.Activate ()**

\*express是一个代表 DataSource 对象的变量。

**返回值**

Boolean

**示例**

``` JavaScript
//以下示例为WPP激活数据源对象，调起ET文档
Application.ActivePresentation.Slides(1).Shapes.Item(3).WebShape.DataSource.Activate()
```

### DataSource.ApplyDemoRange

往活动单元格填充示例数据。返回DataRange。

**语法**

**express.ApplyDemoRange ()**

\*express是一个代表 DataSource 对象的变量。

**返回值**

Object

### DataSource.CreateDataRange

为选定的单元格创建DataRange对象。

**语法**

**express.CreateDataRange ( *Range* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                              |
|---------|-----------|----------|-----------------------------------|
| *Range* | 可选      | Variant  | 代表单元格的字符串或者Range对象。 |

**返回值**

Object

**说明**

参数Range可以为代表单元格的字符串或者Range对象，当为空或字符串时，会弹出对话框输入单元格地址。形如“sheet1!\$a\$1”的字符串表示sheet1中A1单元格。返回值为DataRange。

**示例**

``` JavaScript
//下列示例为Sheet1中的A1单元格创建DataRange对象
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.DataSource.CreateDataRange("sheet1!$a$1")
```

### DataSource.GetAllRangeInfo

获取所有已注册的DataRange。

**语法**

**express.GetAllRangeInfo ()**

\*express是一个代表 DataSource 对象的变量。

**返回值**

String

**说明**

返回的字符串以JSON模式记录所有DataRange的Key。

### DataSource.GetDataRange

返回已注册的对应DataRange。

**语法**

**express.GetDataRange ( *Index* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明       |
|---------|-----------|----------|------------|
| *Index* | 必选      | Variant  | 对象的索引 |

**返回值**

Object

**说明**

目前Index仅支持字符串，代表DataRange对象的键。当输入的Index为数字时，不会有类型检查报错，但是也不会有对象返回。

### DataSource.InitDemoRange

初始化示例数据的行列数

**语法**

**express.InitDemoRange ( *rows* , *cols* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明 |
|--------|-----------|----------|------|
| *rows* | 必选      | Number   | 行数 |
| *cols* | 必选      | Number   | 列数 |

### DataSource.RegisterDataRange

向当前DataSource对象注册DataRange对象。

**语法**

**express.RegisterDataRange ( *Key* , *Value* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明              |
|---------|-----------|----------|-------------------|
| *Key*   | 必选      | String   | DataRange对象的键 |
| *Value* | 必选      | Object   | DataRange对象     |

**返回值**

Boolean

### DataSource.SetDemoRangeItem

填充示例数据。

**语法**

**express.SetDemoRangeItem ( *rowsIdx* , *colsIdx* , *value* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明       |
|-----------|-----------|----------|------------|
| *rowsIdx* | 必选      | Number   | 行数索引   |
| *colsIdx* | 必选      | Number   | 列数索引   |
| *value*   | 必选      | Variant  | 要填充的值 |

### DataSource.UnRegisterDataRange

反注册DataRange。

**语法**

**express.UnRegisterDataRange ( *Key* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明           |
|-------|-----------|----------|----------------|
| *Key* | 必选      | String   | DataRange的Key |

**返回值**

Boolean

## 成员属性

### DataSource.Type

返回DataSource的类型。只读。

**语法**

**express.Type**

\*express是一个代表 DataSource 对象的变量。

**类型**

Number

> WPS 开发文档： [WPS 基础接口/WPS 扩展 API/公共/DataSource](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
