# AllowEditRanges 对象

## 简介

所有 AllowEditRange 对象的集合，这些对象代表受保护工作表上的可编辑单元格。

## 方法

| 序号 | 名称                        | 说明                                                                 |
|:----:|:----------------------------|:---------------------------------------------------------------------|
|  1   | [Add](#AllowEditRanges.Add) | 在受保护的工作表中添加可编辑的单元格区域。返回 AllowEditRange 对象。 |

## 属性

| 序号 | 名称                            | 说明                                       |
|:----:|:--------------------------------|:-------------------------------------------|
|  1   | [Count](#AllowEditRanges.Count) | 返回一个 Long 值，它代表集合中对象的数量。 |
|  2   | [Item](#AllowEditRanges.Item)   | 从集合中返回一个对象。                     |

## 成员方法

### AllowEditRanges.Add

在受保护的工作表中添加可编辑的单元格区域。返回 AllowEditRange 对象。

**语法**

**express.Add ( *Title* , *Range* , *Password* )**

\*express是一个代表 AllowEditRanges 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                               |
|------------|-----------|----------|------------------------------------|
| *Title*    | 必选      | String   | 单元格区域的标题。                 |
| *Range*    | 必选      | Object   | Range 对象。允许编辑的单元格区域。 |
| *Password* | 可选      | Variant  | 单元格区域的密码。                 |

## 成员属性

### AllowEditRanges.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 AllowEditRanges 对象的变量。

### AllowEditRanges.Item

从集合中返回一个对象。

**语法**

**express.Item**

\*express是一个代表 AllowEditRanges 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AllowEditRanges](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
