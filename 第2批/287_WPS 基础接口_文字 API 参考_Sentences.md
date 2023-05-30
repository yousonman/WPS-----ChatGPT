# Sentences 对象

## 简介

Range 对象的集合，这些对象代表选定内容、区域或文档中的所有句子。不存在 Sentence 对象。

## 方法

| 序号 | 名称                    | 说明                                     |
|:----:|:------------------------|:-----------------------------------------|
|  1   | [Item](#Sentences.Item) | 返回集合中的单个 Range 对象。返回Range值 |

## 属性

| 序号 | 名称                                  | 说明                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------|
|  1   | [Application](#Sentences.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                  |
|  2   | [Count](#Sentences.Count)             | 返回一个 Long 类型的值，该值代表集合中句子的数量。只读。                        |
|  3   | [Creator](#Sentences.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。    |
|  4   | [First](#Sentences.First)             | 返回一个 Range 对象，该对象代表文档、区域或选定内容中的句子集合中的第一个句子。 |
|  5   | [Last](#Sentences.Last)               | 返回一个 Range 对象，该对象代表文档、所选内容或范围中的最后一句。               |
|  6   | [Parent](#Sentences.Parent)           | 返回一个 Object 类型值，该值代表指定 Sentences 对象的父对象。                   |

## 成员方法

### Sentences.Item

返回集合中的单个 **Range** 对象。返回Range值

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Sentences 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

## 成员属性

### Sentences.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Sentences 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Sentences.Count

返回一个 **Long** 类型的值，该值代表集合中句子的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Sentences 对象的变量。

### Sentences.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Sentences 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Sentences.First

返回一个 Range 对象，该对象代表文档、区域或选定内容中的句子集合中的第一个句子。

**语法**

**express.First**

\*express是一个代表 Sentences 对象的变量。

**示例**

``` JavaScript
返回一个 Range 对象，该对象代表文档、区域或选定内容中的句子集合中的第一个句子。
```

### Sentences.Last

返回一个 **Range** 对象，该对象代表文档、所选内容或范围中的最后一句。

**语法**

**express.Last**

\*express是一个代表 Sentences 对象的变量。

### Sentences.Parent

返回一个 **Object** 类型值，该值代表指定 **Sentences** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Sentences 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Sentences](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
