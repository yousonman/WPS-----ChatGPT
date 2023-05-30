# Style 对象

## 简介

代表区域的样式说明。

## 说明

Style 对象包含样式的所有属性（字体、数字格式、对齐方式，等等）。有几种内置样式，包括“常规”、“货币”和“百分比”。同时对多个单元格修改单元格格式属性时，使用 Style 对象是快捷高效的方法。

对于 Workbook 对象， Style 对象是 Styles 集合的成员。 Styles 集合包含该工作簿的所有已定义样式。

通过更改应用于单元格的样式的属性可更改单元格的外观。但要记住，更改样式的属性将影响所有以该样式格式化了的单元格。

样式按照名称的字母顺序排序。样式编号表明指定样式在样式名排序列表中的位置。 ` Styles(1) ` 是排序列表中的第一个样式，而 ` Styles(Styles.Count) ` 是最后一个。

有关创建和修改样式的详细信息，请参阅 Styles 对象。

使用 Style 属性可返回一个用于 Range 对象的 Style 对象。下例将“百分比”样式应用于 Sheet1 中的单元格区域 A1:A10。

``` JavaScript
Application.Worksheets.Item("Sheet1").Range("A1:A10").Style = "Percent"
```

使用 Styles ( *index* )（其中 *index* 是样式索引号或名称）可从工作簿的 Style 集合中返回一个 Styles 对象。下例通过设置活动工作簿中“常规”样式的 Bold 属性来更改该样式。

``` JavaScript
Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true
```

## 方法

| 序号 | 名称                    | 说明                    |
|:----:|:------------------------|:------------------------|
|  1   | [Delete](#Style.Delete) | 删除对象。返回Variant值 |

## 属性

| 序号 | 名称                                              | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddIndent](#Style.AddIndent)                     | 返回或设置一个 Boolean 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。                                                                                                                           |
|  2   | [Application](#Style.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  3   | [Borders](#Style.Borders)                         | 返回一个 Bor ders 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。                                                                                                                                        |
|  4   | [BuiltIn](#Style.BuiltIn)                         | 如果样式为内置样式，则为 True 。只读 Boolean 类型。                                                                                                                                                                             |
|  5   | [Creator](#Style.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  6   | [Font](#Style.Font)                               | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                      |
|  7   | [FormulaHidden](#Style.FormulaHidden)             | 返回或设置一个 Boolean 值，它指明在工作表处于保护状态时是否隐藏公式。                                                                                                                                                           |
|  8   | [HorizontalAlignment](#Style.HorizontalAlignment) | 返回或设置一个 XlHAlign 值，它代表指定对象的水平对齐方式。                                                                                                                                                                      |
|  9   | [IncludeAlignment](#Style.IncludeAlignment)       | 如果样式包含 AddIndent 、 HorizontalAlignment 、 VerticalAlignment 、 WrapText 、 IndentLevel 和 Orientation 属性，则该属性值为 True 。 Boolean 类型，可读写。                                                                  |
|  10  | [IncludeBorder](#Style.IncludeBorder)             | 如果指定样式中包含 Color 、 ColorIndex 、 LineStyle 和 Weight 边框属性，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                             |
|  11  | [IncludeFont](#Style.IncludeFont)                 | 如果样式包含 Background 、 Bold 、 Color 、 ColorIndex 、 FontStyle 、 Italic 、 Name 、 Size 、 Strikethrough 、 Subscript 、 Superscript 和 Underline 字体属性，则此属性为 True 。 Boolean 类型，可读写。                     |
|  12  | [IncludeNumber](#Style.IncludeNumber)             | 如果样式中包含 NumberFormat 属性，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                   |
|  13  | [IncludePatterns](#Style.IncludePatterns)         | 如果指定样式中包含 Color 、 ColorIndex 、 InvertIfNegative 、 Pattern 、 PatternColor 和 PatternColorIndex 对象的内部属性，则该属性值为 True 。 Boolean 类型，可读写。                                                          |
|  14  | [IncludeProtection](#Style.IncludeProtection)     | 如果指定样式中包含 FormulaHidden 和 Locked 保护属性，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                |
|  15  | [IndentLevel](#Style.IndentLevel)                 | 返回或设置一个 Long 值，它代表样式的缩进量。                                                                                                                                                                                    |
|  16  | [Interior](#Style.Interior)                       | 返回一个 Inte rior 对象，它代表指定对象的内部。                                                                                                                                                                                 |
|  17  | [Locked](#Style.Locked)                           | 返回或设置一个 Boolean 值，它指明对象是否已被锁定。                                                                                                                                                                             |
|  18  | [MergeCells](#Style.MergeCells)                   | 如果样式包含合并的单元格，则为 True 。 Variant 型，可读写。                                                                                                                                                                     |
|  19  | [Name](#Style.Name)                               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  20  | [NameLocal](#Style.NameLocal)                     | 以用户语言返回或设置对象的名称。 String 型，只读。                                                                                                                                                                              |
|  21  | [NumberFormat](#Style.NumberFormat)               | 返回或设置一个 String 值，它代表对象的格式代码。                                                                                                                                                                                |
|  22  | [NumberFormatLocal](#Style.NumberFormatLocal)     | 以采用用户语言字符串的形式返回或设置一个 String 值，它代表对象的格式代码。                                                                                                                                                      |
|  23  | [Orientation](#Style.Orientation)                 | 返回或设置一个 XlOrientation 值，它代表文本方向。                                                                                                                                                                               |
|  24  | [Parent](#Style.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  25  | [ReadingOrder](#Style.ReadingOrder)               | 返回或设置指定对象的阅读次序。可为以下常量之一： xlRTL （从右到左）、 xlLTR （从左到右）或 xlContext 。 Long 类型，可读写。                                                                                                     |
|  26  | [ShrinkToFit](#Style.ShrinkToFit)                 | 返回或设置一个 Boolean 值，它指明文本是否可以自动收缩为适当尺寸以适应可用列宽。                                                                                                                                                 |
|  27  | [Value](#Style.Value)                             | 返回一个 String 值，它代表指定样式的名称。                                                                                                                                                                                      |
|  28  | [VerticalAlignment](#Style.VerticalAlignment)     | 返回或设置一个 XlVAlign 值，它代表指定对象的垂直对齐方式。                                                                                                                                                                      |
|  29  | [WrapText](#Style.WrapText)                       | 返回或设置一个 Boolean 值，它指明 ET 是否为对象中的文本自动换行。                                                                                                                                                               |

## 成员方法

### Style.Delete

删除对象。返回Variant值

**语法**

**express.Delete ()**

\*express是一个代表 Style 对象的变量。

## 成员属性

### Style.AddIndent

返回或设置一个 **Boolean** 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。

**语法**

**express.AddIndent**

\*express是一个代表 Style 对象的变量。

**说明**

如果将此属性的值设为 **True** ，那么在单元格中文本的对齐方式设为水平或垂直等距分布时，将自动缩进文本。

要将文本对齐方式设为等距分布，可在 Orientation 属性的值为 xlVertical 时，将 VerticalAlignment 属性设为 **xlVAlignDistributed** ；在 Orientation 属性的值为 xlHorizontal 时，将 HorizontalAlignment 属性设为 **xlHAlignDistributed** 。

### Style.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本示例显示一条有关创建 myObject 的应用程序的消息。*/
  let myObject = Application.ActiveWorkbook
  if(myObject.Application.Value == "ET"){
      alert("This is an ET Application object.")
  }
  else{
      alert("This is not an ET Application object.")
  }
}
```

### Style.Borders

返回一个 **Bor ders** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。

**语法**

**express.Borders**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*此示例将 Sheet1 中单元格 B2 的底部边框颜色设置为红色细边框。*/
function test(){

    let rng = Application.Worksheets.Item("Sheet1").Range("B2").Borders.Item(xlEdgeBottom)
    rng.LineStyle = xlContinuous
    rng.Weight = xlThin
    rng.ColorIndex = 3
                                
}
```

### Style.BuiltIn

如果样式为内置样式，则为 **True** 。只读 **Boolean** 类型。

**语法**

**express.BuiltIn**

\*express是一个代表 Style 对象的变量。

### Style.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Style 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Style.Font

返回一个 Font 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 Style 对象的变量。

### Style.FormulaHidden

返回或设置一个 **Boolean** 值，它指明在工作表处于保护状态时是否隐藏公式。

**语法**

**express.FormulaHidden**

\*express是一个代表 Style 对象的变量。

**说明**

请不要将此属性与 **Hidd en** 属性混淆。如果工作簿受保护，而工作表不受保护，将不会隐藏公式。只有在工作表受保护时，才会隐藏公式。

### Style.HorizontalAlignment

返回或设置一个 XlHAlign 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 Style 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### Style.IncludeAlignment

如果样式包含 **AddIndent** 、 **HorizontalAlignment** 、 **VerticalAlignment** 、 **WrapText** 、 **IndentLevel** 和 **Orientation** 属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeAlignment**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入对齐格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeAlignment = true
```

### Style.IncludeBorder

如果指定样式中包含 **Color** 、 **ColorIndex** 、 **LineStyle** 和 **Weight** 边框属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeBorder**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入边框格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeBorder = true
```

### Style.IncludeFont

如果样式包含 **Background** 、 **Bold** 、 **Color** 、 **ColorIndex** 、 **FontStyle** 、 **Italic** 、 **Name** 、 **Size** 、 **Strikethrough** 、 **Subscript** 、 **Superscript** 和 **Underline** 字体属性，则此属性为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeFont**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入字体格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeFont = true
```

### Style.IncludeNumber

如果样式中包含 **NumberFormat** 属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeNumber**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入数字格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeNumber = true
```

### Style.IncludePatterns

如果指定样式中包含 **Color** 、 **ColorIndex** 、 **InvertIfNegative** 、 **Pattern** 、 **PatternColor** 和 **PatternColorIndex** 对象的内部属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludePatterns**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入图案格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludePatterns = true
```

### Style.IncludeProtection

如果指定样式中包含 **FormulaHidden** 和 **Locked** 保护属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeProtection**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入保护格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeProtection = true
```

### Style.IndentLevel

返回或设置一个 **Long** 值，它代表样式的缩进量。

**语法**

**express.IndentLevel**

\*express是一个代表 Style 对象的变量。

**说明**

使用此属性将缩进量设为小于 0 或者大于 15 的数字，将导致发生错误。

### Style.Interior

返回一个 **Inte rior** 对象，它代表指定对象的内部。

**语法**

**express.Interior**

\*express是一个代表 Style 对象的变量。

### Style.Locked

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

\*express是一个代表 Style 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True** ；如果在工作表处于受保护状态时仍能修改对象，则返回 **False** 。

### Style.MergeCells

如果样式包含合并的单元格，则为 **True** 。 **Variant** 型，可读写。

**语法**

**express.MergeCells**

\*express是一个代表 Style 对象的变量。

### Style.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*首先用宏语言，然后用用户语言显示活动工作簿中的第一种样式的名称。*/
function test(){
  let sty = Application.ActiveWorkbook.Styles.Item(1)
  alert("The name of the style: " + sty.Name)
  alert("The localized name of the style: " + sty.NameLocal)
}  

/*显示活动工作簿的 Sheet1 中的默认 ListObject 对象的名称。*/
function Test(){  
    let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let oListObj = wrksht.ListObjects.Item(1)                                       
    alert(oListObj.Name)                                 
}
```

### Style.NameLocal

以用户语言返回或设置对象的名称。 **String** 型，只读。

**语法**

**express.NameLocal**

\*express是一个代表 Style 对象的变量。

**说明**

如果样式为内置样式，则此属性以当前系统所使用的地区语言返回样式的名称。

**示例**

``` JavaScript
function test(){
  /*此示例显示活动工作簿上样式一的原名称和本地化以后的名称。*/
  let sty = Application.ActiveWorkbook.Styles.Item(1)
  alert("The name of the style is " + sty.Name)
  alert("The localized name of the style is " + sty.NameLocal)
}
```

### Style.NumberFormat

返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

\*express是一个代表 Style 对象的变量。

**说明**

格式代码与 **“设置单元格格式”** 对话框中的 **“格式代码”** 选项是同一个字符串。 **Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### Style.NumberFormatLocal

以采用用户语言字符串的形式返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormatLocal**

\*express是一个代表 Style 对象的变量。

**说明**

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### Style.Orientation

返回或设置一个 **XlOrientation** 值，它代表文本方向。

**语法**

**express.Orientation**

\*express是一个代表 Style 对象的变量。

### Style.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Style 对象的变量。

### Style.ReadingOrder

返回或设置指定对象的阅读次序。可为以下常量之一： **xlRTL** （从右到左）、 **xlLTR** （从左到右）或 **xlContext** 。 **Long** 类型，可读写。

**语法**

**express.ReadingOrder**

\*express是一个代表 Style 对象的变量。

### Style.ShrinkToFit

返回或设置一个 **Boolean** 值，它指明文本是否可以自动收缩为适当尺寸以适应可用列宽。

**语法**

**express.ShrinkToFit**

\*express是一个代表 Style 对象的变量。

### Style.Value

返回一个 **String** 值，它代表指定样式的名称。

**语法**

**express.Value**

\*express是一个代表 Style 对象的变量。

### Style.VerticalAlignment

返回或设置一个 **XlVAlign** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 Style 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### Style.WrapText

返回或设置一个 **Boolean** 值，它指明 ET 是否为对象中的文本自动换行。

**语法**

**express.WrapText**

\*express是一个代表 Style 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Style](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
