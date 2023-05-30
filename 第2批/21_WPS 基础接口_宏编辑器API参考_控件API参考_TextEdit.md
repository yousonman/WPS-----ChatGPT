# TextEdit 对象

## 简介

用于输入或改变文本内容。

## 说明

用于输入或改变文本内容。

## 方法

| 序号 | 名称                   | 说明                       |
|:----:|:-----------------------|:---------------------------|
|  1   | [Move](#TextEdit.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                     | 说明                                                       |
|:----:|:-----------------------------------------|:-----------------------------------------------------------|
|  1   | [AutoSize](#TextEdit.AutoSize)           | 指定该对象是否根据内容自动设置它的大小                     |
|  2   | [BackColor](#TextEdit.BackColor)         | 获取或设置指定对象的内部颜色                               |
|  3   | [BackStyle](#TextEdit.BackStyle)         | 您可以使用该属性指定控件是否透明                           |
|  4   | [BorderColor](#TextEdit.BorderColor)     | 您可以通过该属性指定控件边框的颜色                         |
|  5   | [BorderStyle](#TextEdit.BorderStyle)     | 指定控件边框的显示方式                                     |
|  6   | [ControlSource](#TextEdit.ControlSource) | 您可以使用该属性指定控件中显示的数据                       |
|  7   | [Enabled](#TextEdit.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                   |
|  8   | [Font](#TextEdit.Font)                   | 指定对象的字体，只读                                       |
|  9   | [ForeColor](#TextEdit.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                       |
|  10  | [Height](#TextEdit.Height)               | 获取或设置指定对象的高度                                   |
|  11  | [HeightPolicy](#TextEdit.HeightPolicy)   | 获取或设置指定对像的高度策略                               |
|  12  | [Left](#TextEdit.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置               |
|  13  | [MaxLength](#TextEdit.MaxLength)         | 指定用户可以在文本框或组合框中输入的最大字符数             |
|  14  | [Name](#TextEdit.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读       |
|  15  | [PasswordChar](#TextEdit.PasswordChar)   | 指定 TextBox中 显示的占位符而不是实际在TextBox中输入的字符 |
|  16  | [TabIndex](#TextEdit.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置  |
|  17  | [TabStop](#TextEdit.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上      |
|  18  | [Text](#TextEdit.Text)                   | 您可以通过该属性设置或返回文本框中包含的文本值             |
|  19  | [TextAlign](#TextEdit.TextAlign)         | 该属性用于指定对象中文本的对齐方式                         |
|  20  | [Top](#TextEdit.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置               |
|  21  | [Value](#TextEdit.Value)                 | 确定或指定文本框中的文本                                   |
|  22  | [Visible](#TextEdit.Visible)             | 返回或设置该对象是否可见                                   |
|  23  | [Width](#TextEdit.Width)                 | 获取或设置指定对象的宽度                                   |
|  24  | [WidthPolicy](#TextEdit.WidthPolicy)     | 获取或设置指定对象的宽度策略                               |
|  25  | [WordWrap](#TextEdit.WordWrap)           | 指定控件的内容是否在行尾自动换行                           |

## 成员方法

### TextEdit.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 TextEdit 对象的变量。

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
/*将TextEdit1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.TextEdit1.Move(100,100,100,100)
}
```

## 成员属性

### TextEdit.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 TextEdit 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置TextEdit1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.TextEdit1.AutoSize = true
}
```

### TextEdit.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 TextEdit 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将TextEdit1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.TextEdit1.BackColor = 0x665898
}
```

### TextEdit.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

**示例**

``` JavaScript
/*将TextEdit1设置为透明*/
function func1()
{
    UserForm1.TextEdit1.BackStyle = 0
}
```

### TextEdit.BorderColor

您可以通过该属性指定控件边框的颜色

**语法**

**express.BorderColor**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以通过该属性指定控件边框的颜色

**示例**

``` JavaScript
/*将TextEdit1的边框设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.TextEdit1.BorderColor = 0x665898
}
```

### TextEdit.BorderStyle

指定控件边框的显示方式

**语法**

**express.BorderStyle**

\*express是一个代表 TextEdit 对象的变量。

**说明**

指定控件边框的显示方式，可以设置为以下值：

| 设置   | 值  | 描述     |
|--------|-----|----------|
| None   | 0   | 无边框   |
| Single | 1   | 实线边框 |

**示例**

``` JavaScript
/*将TextEdit1的边框显示方式设置为实线*/
function func1()
{
    UserForm1.TextEdit1.BorderStyle = 1
}
```

### TextEdit.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### TextEdit.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.TextEdit1.Enabled = true
}
```

### TextEdit.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 TextEdit 对象的变量。

**说明**

指定对象的字体，只读

### TextEdit.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将TextEdit1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.TextEdit1.ForeColor = 0x658978
}
```

### TextEdit.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 TextEdit 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将TextEdit1控件的高度设置为100*/
function func1()
{
    UserForm1.TextEdit1.Height = 100
}
```

### TextEdit.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 TextEdit 对象的变量。

**说明**

获取或设置指定对像的高度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将TextEdit1的高度设置为固定*/
function func1()
{
    UserForm1.TextEdit1.HeightPolicy = 2
}
```

### TextEdit.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将TextEdit1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.TextEdit1.Left = 100
}
```

### TextEdit.MaxLength

指定用户可以在文本框或组合框中输入的最大字符数

**语法**

**express.MaxLength**

\*express是一个代表 TextEdit 对象的变量。

**说明**

指定用户可以在文本框或组合框中输入的最大字符数

**示例**

``` JavaScript
/*指定TextEdit1控件中输入的最大字符数为32767*/
function func1()
{
    UserForm1.TextEdit1.MaxLength = 32767
}
```

### TextEdit.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### TextEdit.PasswordChar

指定 TextBox中 显示的占位符而不是实际在TextBox中输入的字符

**语法**

**express.PasswordChar**

\*express是一个代表 TextEdit 对象的变量。

**说明**

指定 TextBox中 显示的占位符而不是实际在TextBox中输入的字符

**示例**

``` JavaScript
/*将用户输入的字符用*替换*/
function func1()
{
    UserForm1.TextEdit1.PasswordChar = "*"
}
```

### TextEdit.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将TextEdit1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.TextEdit1.TabIndex = 2
}
```

### TextEdit.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将TextEdit1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.TextEdit1.TabStop = true
}
```

### TextEdit.Text

您可以通过该属性设置或返回文本框中包含的文本值

**语法**

**express.Text**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以通过该属性设置或返回文本框中包含的文本值

**示例**

``` JavaScript
/*将TextEdit1控件中的文本设置为test test*/
function func1()
{
    UserForm1.TextEdit1.Text = "test test"
}
```

### TextEdit.TextAlign

该属性用于指定对象中文本的对齐方式

**语法**

**express.TextAlign**

\*express是一个代表 TextEdit 对象的变量。

**说明**

该属性用于指定对象中文本的对齐方式，可设置为以下值：

| 设置   | 值  | 描述                     |
|--------|-----|--------------------------|
| Left   | 1   | 文本数字和日期都左对齐   |
| Center | 2   | 文本数字和日期都居中对齐 |
| Right  | 3   | 文本数字和日期都右对齐   |

**示例**

``` JavaScript
/*将TextEdit1中的文本设置为居中对齐*/
function func1()
{
    UserForm1.TextEdit1.TextAlign = 2
}
```

### TextEdit.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 TextEdit 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将TextEdit1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.TextEdit1.Top = 100
}
```

### TextEdit.Value

确定或指定文本框中的文本

**语法**

**express.Value**

\*express是一个代表 TextEdit 对象的变量。

**说明**

确定或指定文本框中的文本

**示例**

``` JavaScript
/*将TextEdit1控件中的文本指定为test test*/
function func1()
{
    UserForm1.TextEdit1.Value = "test test"
}
```

### TextEdit.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 TextEdit 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将TextEdit1设置为可见*/
function func1()
{
    UserForm1.TextEdit1.Visible = true
}
```

### TextEdit.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 TextEdit 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将TextEdit1的宽度设置为100*/
function func1()
{
    UserForm1.TextEdit1.Width = 100
}
```

### TextEdit.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 TextEdit 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将TextEdit1的宽度设置为固定*/
function func1()
{
    UserForm1.TextEdit1.WidthPolicy = 2
}
```

### TextEdit.WordWrap

指定控件的内容是否在行尾自动换行

**语法**

**express.WordWrap**

\*express是一个代表 TextEdit 对象的变量。

**说明**

指定控件的内容是否在行尾自动换行

**示例**

``` JavaScript
/*将TextEdit1设置为控件的内容在行尾自动换行*/
function func1()
{
    UserForm1.TextEdit1.WordWrap = true
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/TextEdit](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
