# Characters 对象

## 简介

代表包含文本的对象中的字符。

## 说明

使用 Characters 对象可修改包含在全文本字符串中的任意字符序列。

使用 Characters ( *start* , *length* )（其中 *start* 为起始字符号，而 *length* 为要返回的字符个数）返回 Characters 对象。

``` JavaScript
/*下例向单元格 B1 中添加文本，并将第二个单词设置为加粗。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("B1").Value = "New Title"
    Application.Worksheets.Item("Sheet1").Range("B1").Characters(5, 5).Font.Bold = true
}
```

仅当需要更改对象中文本的一部分而不影响其余部分时，才有必要使用 Characters 方法（如果对象不支持格式文本，则不能使用 Characters 方法对文本中的一部分单独设置格式）。要同时更改所有文本，通常可以对该对象直接应用某一适当的方法或属性。

``` JavaScript
/*下例将单元格 A5 的内容设置为倾斜。*/
Application.Worksheets.Item("Sheet1").Range("A5").Font.Italic = true
```

## 方法

| 序号 | 名称                         | 说明                             |
|:----:|:-----------------------------|:---------------------------------|
|  1   | [Delete](#Characters.Delete) | 删除对象。                       |
|  2   | [Insert](#Characters.Insert) | 在所选的字符前面插入一个字符串。 |

## 属性

| 序号 | 名称                                                 | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Characters.Application)               | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Caption](#Characters.Caption)                       | 返回一个 String 值，它代表此字符区域的文本。                                                                                                                                                                                    |
|  3   | [Count](#Characters.Count)                           | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  4   | [Creator](#Characters.Creator)                       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Font](#Characters.Font)                             | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                      |
|  6   | [Parent](#Characters.Parent)                         | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  7   | [PhoneticCharacters](#Characters.PhoneticCharacters) | 返回或设置指定 Characters 对象中的拼音文本。String 类型，可读写。                                                                                                                                                               |
|  8   | [Text](#Characters.Text)                             | 返回或设置指定对象中的文本。 String 型，可读写。                                                                                                                                                                                |

## 成员方法

### Characters.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Characters 对象的变量。

**说明**

返回值 Variant

### Characters.Insert

在所选的字符前面插入一个字符串。

**语法**

**express.Insert ( *String* )**

\*express是一个代表 Characters 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明             |
|----------|-----------|----------|------------------|
| *String* | 必选      | String   | 要插入的字符串。 |

**说明**

返回值 Variant

## 成员属性

### Characters.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Characters 对象的变量。

### Characters.Caption

返回一个 **String** 值，它代表此字符区域的文本。

**语法**

**express.Caption**

\*express是一个代表 Characters 对象的变量。

### Characters.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Characters 对象的变量。

**示例**

``` JavaScript
/*此示例将 A1 单元格的最后一个字符设为上标字符。*/
function test(){
    let n = Application.Worksheets.Item("Sheet1").Range("A1").Characters.Count
    Application.Worksheets.Item("Sheet1").Range("A1").Characters(n, 1).Font.Superscript = true
}
```

### Characters.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Characters 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Characters.Font

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 Characters 对象的变量。

### Characters.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Characters 对象的变量。

### Characters.PhoneticCharacters

返回或设置指定 **Characters** 对象中的拼音文本。String 类型，可读写。

**语法**

**express.PhoneticCharacters**

\*express是一个代表 Characters 对象的变量。

**说明**

要在单元格中添加拼音信息，请不要使用该属性，应使用 **Phonetics** 集合的 **Add** 方法；另外，应使用 **Phonetic** 对象的 **Text** 属性返回或设置单元格中的拼音文本字符串。

**示例**

``` JavaScript
/*本示例将活动单元格中从文本起始位置开始的第四个字符替换。*/
ActiveCell.Characters(1,3).PhoneticCharacters = "替换字符"
```

### Characters.Text

返回或设置指定对象中的文本。 **String** 型，可读写。

**语法**

**express.Text**

\*express是一个代表 Characters 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Characters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
