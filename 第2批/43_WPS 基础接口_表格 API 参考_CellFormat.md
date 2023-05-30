# CellFormat 对象

## 简介

代表单元格格式的搜索条件。

## 说明

使用 Application 对象的 FindFormat 或 ReplaceFormat 属性可返回 CellFormat 对象。

使用 CellFormat 对象的 Borders 、 Fon t 属性或 CellFormat 对象的 Interi or 属性可定义单元格格式的搜索条件。

下例设置单元格格式内部的搜索条件。

``` JavaScript
function test(){
    //Set the interior of cell A1 to yellow.
    Application.Range("A1").Select()
    Application.Selection.Interior.ColorIndex = 36
    alert("The cell format for cell A1 is a yellow interior.")

    //Set the CellFormat object to replace yellow with green.
    Application.FindFormat.Interior.ColorIndex = 36
    Application.ReplaceFormat.Interior.ColorIndex = 35       //Find and replace cell A1's yellow interior with green.
    Application.ActiveCell.Replace("", "", xlPart, xlByRows, false, null, true, true)
    alert("The cell format for cell A1 is replaced with a green interior.")
} 
```

## 方法

| 序号 | 名称                       | 说明                                               |
|:----:|:---------------------------|:---------------------------------------------------|
|  1   | [Clear](#CellFormat.Clear) | 清除 FindFormat 和 ReplaceFor mat 属性中的条件集。 |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddIndent](#CellFormat.AddIndent)                     | 返回或设置一个 Variant 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。                                                                                                                           |
|  2   | [Application](#CellFormat.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  3   | [Borders](#CellFormat.Borders)                         | 返回或设置一个 Borders 集合，它代表基于单元格边框格式的搜索条件。                                                                                                                                                               |
|  4   | [Creator](#CellFormat.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Font](#CellFormat.Font)                               | 返回一个 Font 对象，该对象允许用户根据单元格的字体格式设置或返回搜索条件。                                                                                                                                                      |
|  6   | [FormulaHidden](#CellFormat.FormulaHidden)             | 返回或设置一个 Variant 值，它指明在工作表处于保护状态时是否隐藏公式。                                                                                                                                                           |
|  7   | [HorizontalAlignment](#CellFormat.HorizontalAlignment) | 返回或设置一个 Variant 值，它代表指定对象的水平对齐方式。                                                                                                                                                                       |
|  8   | [IndentLevel](#CellFormat.IndentLevel)                 | 返回或设置一个 Variant 值，它代表单元格或单元格区域的缩进量。可为 0 到 15 之间的整数。                                                                                                                                          |
|  9   | [Interior](#CellFormat.Interior)                       | 返回一个 Interior 对象，该对象允许用户根据单元格内部格式设置或返回搜索条件。                                                                                                                                                    |
|  10  | [Locked](#CellFormat.Locked)                           | 返回或设置一个 Variant 值，它指明对象是否已被锁定。                                                                                                                                                                             |
|  11  | [MergeCells](#CellFormat.MergeCells)                   | 如果区域或样式包含合并单元格，则为 True 。 Variant 。可读写。                                                                                                                                                                   |
|  12  | [NumberFormat](#CellFormat.NumberFormat)               | 返回或设置一个 Variant 值，它代表对象的格式代码。                                                                                                                                                                               |
|  13  | [NumberFormatLocal](#CellFormat.NumberFormatLocal)     | 以采用用户语言字符串的形式返回或设置一个 Variant 值，它代表对象的格式代码。                                                                                                                                                     |
|  14  | [Orientation](#CellFormat.Orientation)                 | 返回或设置一个 Variant 值，它代表文本方向。                                                                                                                                                                                     |
|  15  | [Parent](#CellFormat.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  16  | [ShrinkToFit](#CellFormat.ShrinkToFit)                 | 返回或设置一个 Variant 值。 该值表示文本是否自动缩小以适应可用列宽                                                                                                                                                              |
|  17  | [VerticalAlignment](#CellFormat.VerticalAlignment)     | 返回或设置一个 Variant 值，它代表指定对象的垂直对齐方式。                                                                                                                                                                       |
|  18  | [WrapText](#CellFormat.WrapText)                       | 返回或设置一个 Variant 值，它指明 ET 是否为对象中的文本自动换行。                                                                                                                                                               |

## 成员方法

### CellFormat.Clear

清除 **FindFormat** 和 **ReplaceFor mat** 属性中的条件集。

**语法**

**express.Clear ()**

\*express是一个代表 CellFormat 对象的变量。

## 成员属性

### CellFormat.AddIndent

返回或设置一个 **Variant** 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。

**语法**

**express.AddIndent**

\*express是一个代表 CellFormat 对象的变量。

**说明**

如果将此属性的值设为 **True** ，那么在单元格中文本的对齐方式设为水平或垂直等距分布时，将自动缩进文本。

要将文本对齐方式设为等距分布，可在 Orientation 属性的值为 **xlVertical** 时，将 **VerticalAlignment** 属性设为 **xlVAlignDistributed** ；在 **Orientation** 属性的值为 **xlHorizontal** 时，将 **HorizontalAlignment** 属性设为 **xlHAlignDistributed** 。

### CellFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CellFormat 对象的变量。

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

### CellFormat.Borders

返回或设置一个 **Borders** 集合，它代表基于单元格边框格式的搜索条件。

**语法**

**express.Borders**

\*express是一个代表 CellFormat 对象的变量。

**示例**

``` JavaScript
function test(){
    // Set the search criteria for the border of the cell format.
    let borders = Application.FindFormat.Borders.Item(xlEdgeBottom)
    borders.LineStyle = xlContinuous
    borders.Weight = xlThick

    // Create a continuous thick bottom-edge border for cell A5.
    Application.Range("A5").Select()
    Application.Selection.Borders.Item(xlEdgeBottom).LineStyle = xlContinuous
    Application.Selection.Borders.Item(xlEdgeBottom).Weight = xlThick
    Application.Range("A1").Select()
    alert("Cell A5 has a continuous thick bottom-edge border")

    // Find the cells based on the search criteria.
    Cells.Find("", ActiveCell, xlFormulas, xlPart, xlByRows, xlNext, false , null, true).Activate()
    alert("ET has found this cell matching the search criteria.")

}
```

### CellFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CellFormat 对象的变量。

### CellFormat.Font

返回一个 **Font** 对象，该对象允许用户根据单元格的字体格式设置或返回搜索条件。

**语法**

**express.Font**

\*express是一个代表 CellFormat 对象的变量。

**示例**

``` JavaScript
/* 此示例设置搜索条件以识别包含红色字体的单元格，应用该条件创建一个单元格，找到该单元格，并通知用户*/
function test(){
    // Set the search criteria for the font of the cell format.
    Application.FindFormat.Font.ColorIndex = 3

    // Set the color index of the font for cell A5 to red.
    Application.Range("A5").Font.ColorIndex = 3
    Application.Range("A5").Formula = "Red font"
    Application.Range("A1").Select()
    alert("Cell A5 has red font")

    // Find the cells based on the search criteria.
    Application.Cells.Find("", ActiveCell, xlFormulas, xlPart, xlByRows, xlNext, false , null, true).Activate()
    alert("ET has found this cell matching the search criteria.")
}
```

### CellFormat.FormulaHidden

返回或设置一个 **Variant** 值，它指明在工作表处于保护状态时是否隐藏公式。

**语法**

**express.FormulaHidden**

\*express是一个代表 CellFormat 对象的变量。

**说明**

如果在工作表处于保护状态时要隐藏公式，此属性将返回 **True** ；如果指定区域中有些单元格的 **FormulaHidden** 为 **True** ，而有些单元格的 **FormulaHidden** 为 **False** ，则返回 **Null** 。

请不要将此属性与 **Hidden** 属性混淆。如果工作簿受保护，而工作表不受保护，将不会隐藏公式。只有在工作表受保护时，才会隐藏公式。

### CellFormat.HorizontalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 CellFormat 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlLeft**        |
| **xlRight**       |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### CellFormat.IndentLevel

返回或设置一个 **Variant** 值，它代表单元格或单元格区域的缩进量。可为 0 到 15 之间的整数。

**语法**

**express.IndentLevel**

\*express是一个代表 CellFormat 对象的变量。

**说明**

使用此属性将缩进量设为小于 0 或者大于 15 的数字，将导致发生错误。

### CellFormat.Interior

返回一个 **Interior** 对象，该对象允许用户根据单元格内部格式设置或返回搜索条件。

**语法**

**express.Interior**

\*express是一个代表 CellFormat 对象的变量。

**示例**

``` JavaScript
/* 此示例设置搜索条件以识别内部为纯黄色的单元格，创建满足该条件的单元格，找到该单元格，并通知给用户*/
function test(){
    // Set the search criteria for the interior of the cell format.
    let interior = Application.FindFormat.Interior
    interior.ColorIndex = 6
    interior.Pattern = xlSolid
    interior.PatternColorIndex = xlAutomatic

    // Create a yellow interior for cell A5.
    Application.Range("A5").Select()
    Application.Selection.Interior.ColorIndex = 6
    Application.Selection.Interior.Pattern = xlSolid
    Application.Selection.Interior.PatternColorIndex = xlAutomatic
    Range("A1").Select()
    alert("Cell A5 has a yellow interior.")

    // Find the cells based on the search criteria.
    Application.Cells.Find("", ActiveCell, xlFormulas, xlPart, xlByRows, xlNext, false , null, true).Activate()
    alert("ET has found this cell matching the search criteria.")
}
```

### CellFormat.Locked

返回或设置一个 **Variant** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

\*express是一个代表 CellFormat 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True** ；如果在工作表处于受保护状态时仍能修改对象，则返回 **False** ；如果指定区域既包含锁定单元格又包含不锁定单元格，则返回 **Null** 。

**示例**

``` JavaScript
/* 此示例解除对 Sheet1 中 A1:G37 区域单元格的锁定，以便当该工作表受保护时也可对这些单元格进行修改*/
function test(){
    Application.Worksheets.Item("Sheet1").Range("A1:G37").Locked = false
    Application.Worksheets.Item("Sheet1").Protect()
}
```

### CellFormat.MergeCells

如果区域或样式包含合并单元格，则为 **True** 。 **Variant** 。可读写。

**语法**

**express.MergeCells**

\*express是一个代表 CellFormat 对象的变量。

### CellFormat.NumberFormat

返回或设置一个 **Variant** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

\*express是一个代表 CellFormat 对象的变量。

**说明**

如果指定区域中的所有单元格包含不同的数字格式，则此属性返回 **Null** 。

格式代码与 **“设置单元格格式”** 对话框中的 **“格式代码”** 选项是同一个字符串。 **Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### CellFormat.NumberFormatLocal

以采用用户语言字符串的形式返回或设置一个 **Variant** 值，它代表对象的格式代码。

**语法**

**express.NumberFormatLocal**

\*express是一个代表 CellFormat 对象的变量。

**说明**

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### CellFormat.Orientation

返回或设置一个 **Variant** 值，它代表文本方向。

**语法**

**express.Orientation**

\*express是一个代表 CellFormat 对象的变量。

### CellFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CellFormat 对象的变量。

### CellFormat.ShrinkToFit

返回或设置一个 **Variant** 值。 该值表示文本是否自动缩小以适应可用列宽

**语法**

**express.ShrinkToFit**

\*express是一个代表 CellFormat 对象的变量。

**说明**

如果文本自动收缩以适应可用列宽，此属性将返回 **True** ；如果没有将指定区域中所有单元格的这一属性设为相同的值，则返回 **Null** 。

### CellFormat.VerticalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 CellFormat 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlBottom**      |
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlTop**         |

**示例**

``` JavaScript
/* 此示例将 Sheet1 上第二行的行高设置为标准行高的两倍，然后使该行的内容垂直居中*/
function test(){
    Application.Worksheets.Item("Sheet1").Rows.Item(2).RowHeight = 2 * Worksheets.Item("Sheet1").StandardHeight
    Application.Worksheets.Item("Sheet1").Rows.Item(2).VerticalAlignment = xlVAlignCenter
}
```

### CellFormat.WrapText

返回或设置一个 **Variant** 值，它指明 ET 是否为对象中的文本自动换行。

**语法**

**express.WrapText**

\*express是一个代表 CellFormat 对象的变量。

**说明**

如果指定区域内所有单元格中的文本都自动换行，此属性将返回 **True** ；如果指定区域内所有单元格中的文本都不自动换行，则返回 **False** ；如果指定区域内有些单元格中的文本自动换行，而另一些单元格中的文本不自动换行，则返回 **Null** 。

ET 会在必要的时候改变区域的行高以容纳其中的文字。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/CellFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
