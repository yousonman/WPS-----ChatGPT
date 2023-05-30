# Phonetic 对象

## 简介

包含单元格中特定拼音文本串的有关信息。

## 说明

在 ET 97 中，此对象包含指定区域中任意拼音文本的格式属性。

## 属性

| 序号 | 名称                                     | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Alignment](#Phonetic.Alignment)         | 返回或设置一个 Long 值，它代表指定的拼音文本或刻度线标签的对齐方式。                                                                                                                                                            |
|  2   | [Application](#Phonetic.Application)     | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  3   | [CharacterType](#Phonetic.CharacterType) | 返回或设置指定单元格中拼音文本的类型。 XlPhoneticCharacterType 类型，可读写。                                                                                                                                                   |
|  4   | [Creator](#Phonetic.Creator)             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Font](#Phonetic.Font)                   | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                      |
|  6   | [Parent](#Phonetic.Parent)               | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  7   | [Text](#Phonetic.Text)                   | 返回或设置指定对象中的文本。 String 型，可读写。                                                                                                                                                                                |
|  8   | [Visible](#Phonetic.Visible)             | 返回或设置一个 Boolean 值，它确定对象是否可见。可读写。                                                                                                                                                                         |

## 成员属性

### Phonetic.Alignment

返回或设置一个 **Long** 值，它代表指定的拼音文本或刻度线标签的对齐方式。

**语法**

**express.Alignment**

\*express是一个代表 Phonetic 对象的变量。

### Phonetic.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Phonetic 对象的变量。

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

### Phonetic.CharacterType

返回或设置指定单元格中拼音文本的类型。 **XlPhoneticCharacterType** 类型，可读写。

**语法**

**express.CharacterType**

\*express是一个代表 Phonetic 对象的变量。

**示例**

本示例将活动单元格中的第一个拼音文本字符串从 Furigana 更改为 Hiragana。

``` JavaScript
ActiveCell.Phonetics.Item(1).CharacterType = xlHiragana
```

### Phonetic.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Phonetic 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Phonetic.Font

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 Phonetic 对象的变量。

### Phonetic.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Phonetic 对象的变量。

### Phonetic.Text

返回或设置指定对象中的文本。 **String** 型，可读写。

**语法**

**express.Text**

\*express是一个代表 Phonetic 对象的变量。

**说明**

对于 **Phonetic** 对象，此属性返回或设置其拼音文本。不能将此属性设为 **Null** 。

### Phonetic.Visible

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 Phonetic 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Phonetic](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
