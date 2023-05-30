# DataRange 对象

## 简介

代表容器应用程序的一个数据区域对象。

## 说明

DataRange为类似表格Range概念，对应于表格中的一个区域，用于存储缓存数据，如果注册进DataSource，表格中的对应区域数据修改，将会触发事件。

## 属性

| 序号 | 名称                                  | 说明                                                                     |
|:----:|:--------------------------------------|:-------------------------------------------------------------------------|
|  1   | [Address](#DataRange.Address)         | 代表DataRange对象的地址。形如“sheet1!\$a\$1”表示sheet1中A1单元格。只读。 |
|  2   | [TextInJSON](#DataRange.TextInJSON)   | json格式表示区域的格式化后的值。只读。                                   |
|  3   | [ValueInJSON](#DataRange.ValueInJSON) | json格式表示区域的值。只读。                                             |

## 成员属性

### DataRange.Address

代表DataRange对象的地址。形如“sheet1!\$a\$1”表示sheet1中A1单元格。只读。

**语法**

**express.Address**

\*express是一个代表 DataRange 对象的变量。

**类型**

String

### DataRange.TextInJSON

json格式表示区域的格式化后的值。只读。

**语法**

**express.TextInJSON**

\*express是一个代表 DataRange 对象的变量。

**类型**

String

### DataRange.ValueInJSON

json格式表示区域的值。只读。

**语法**

**express.ValueInJSON**

\*express是一个代表 DataRange 对象的变量。

> WPS 开发文档： [WPS 基础接口/WPS 扩展 API/公共/DataRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
