# Debug 对象

## 简介

宏编辑器的调试对象，提供调试辅助接口。

## 说明

宏编辑器的调试对象，提供调试辅助接口。

## 方法

| 序号 | 名称                  | 说明                             |
|:----:|:----------------------|:---------------------------------|
|  1   | [GC](#Debug.GC)       | 触发JavaScript引擎进行垃圾回收。 |
|  2   | [Print](#Debug.Print) | 在立即窗口中输出文本。           |

## 成员方法

### Debug.GC

触发JavaScript引擎进行垃圾回收。

**语法**

**express.GC ()**

\*express是一个代表 Debug 对象的变量。

**说明**

触发JavaScript引擎进行垃圾回收。

### Debug.Print

在立即窗口中输出文本。

**语法**

**express.Print ( *outputlist* )**

\*express是一个代表 Debug 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                |
|--------------|-----------|----------|-----------------------------------------------------|
| *outputlist* | 可选      | Variant  | 要打印的表达式或表达式列表。 如果省略，将打印空行。 |

**说明**

在立即窗口中输出文本。

**示例**

``` JavaScript
/*打印当前表格中B10单元格的值到立即窗口*/
function abcd()
{
    Debug.Print(Range("B10").Value2)
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/Debug](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# MultiPage 对象

## 简介

包含一个选项卡栏和一个页面区的容器类组件，通过切换选项卡来切换不同页面，从而达到在同一个窗口中自由切换不同的内容。

## 说明

包含一个选项卡栏和一个页面区的容器类组件，通过切换选项卡来切换不同页面，从而达到在同一个窗口中自由切换不同的内容。

## 方法

| 序号 | 名称                    | 说明                       |
|:----:|:------------------------|:---------------------------|
|  1   | [Move](#MultiPage.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                    | 说明                                                      |
|:----:|:----------------------------------------|:----------------------------------------------------------|
|  1   | [BackColor](#MultiPage.BackColor)       | 获取或设置指定对象的内部颜色                              |
|  2   | [Enabled](#MultiPage.Enabled)           | 您可以通过该属性设置或返回是否启用该控件                  |
|  3   | [Font](#MultiPage.Font)                 | 指定对象的字体，只读                                      |
|  4   | [ForeColor](#MultiPage.ForeColor)       | 您可以通过该属性指定控件中文本的颜色                      |
|  5   | [Height](#MultiPage.Height)             | 获取或设置指定对象的高度                                  |
|  6   | [HeightPolicy](#MultiPage.HeightPolicy) | 获取或设置指定对像的高度策略                              |
|  7   | [Left](#MultiPage.Left)                 | 您可以通过该属性指定控件在对象或报表中的位置              |
|  8   | [Name](#MultiPage.Name)                 | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  9   | [TabIndex](#MultiPage.TabIndex)         | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  10  | [TabStop](#MultiPage.TabStop)           | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  11  | [Top](#MultiPage.Top)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  12  | [Value](#MultiPage.Value)               | 您可以通过该属性选择激活第几个页面                        |
|  13  | [Visible](#MultiPage.Visible)           | 返回或设置该对象是否可见                                  |
|  14  | [Width](#MultiPage.Width)               | 获取或设置指定对象的宽度                                  |
|  15  | [WidthPolicy](#MultiPage.WidthPolicy)   | 获取或设置指定对象的宽度策略                              |

## 成员方法

### MultiPage.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 MultiPage 对象的变量。

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
/*将MultiPage1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.MultiPage1.Move(100,100,100,100)
}
```

## 成员属性

### MultiPage.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 MultiPage 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将MultiPage1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.MultiPage1.BackColor = 0x665898
}
```

### MultiPage.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 MultiPage 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.MultiPage1.Enabled = true
}
```

### MultiPage.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 MultiPage 对象的变量。

**说明**

指定对象的字体，只读

### MultiPage.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 MultiPage 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将MultiPage1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.MultiPage1.ForeColor = 0x658978
}
```

### MultiPage.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 MultiPage 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将MultiPage1控件的高度设置为100*/
function func1()
{
    UserForm1.MultiPage1.Height = 100
}
```

### MultiPage.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 MultiPage 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将MultiPage1的高度设置为固定*/
function func1()
{
    UserForm1.MultiPage1.HeightPolicy = 2
}
```

### MultiPage.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 MultiPage 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将MultiPage1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.MultiPage1.Left = 100
}
```

### MultiPage.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 MultiPage 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### MultiPage.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 MultiPage 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将MultiPage1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.MultiPage1.TabIndex = 2
}
```

### MultiPage.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 MultiPage 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将MultiPage1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.MultiPage1.TabStop = true
}
```

### MultiPage.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 MultiPage 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将MultiPage1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.MultiPage1.Top = 100
}
```

### MultiPage.Value

您可以通过该属性选择激活第几个页面

**语法**

**express.Value**

\*express是一个代表 MultiPage 对象的变量。

**说明**

您可以通过该属性选择激活第几个页面

**示例**

``` JavaScript
/*激活MultiPage1的第一个页面*/
function func1()
{
    UserForm1.MultiPage1.Value = 0
}
```

### MultiPage.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 MultiPage 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将MultiPage1设置为可见*/
function func1()
{
    UserForm1.MultiPage1.Visible = true
}
```

### MultiPage.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 MultiPage 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将MultiPage1的宽度设置为100*/
function func1()
{
    UserForm1.MultiPage1.Width = 100
}
```

### MultiPage.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 MultiPage 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将MultiPage1的宽度设置为固定*/
function func1()
{
    UserForm1.MultiPage1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/MultiPage](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# ChartView 对象

## 简介

代表图表的视图。

## 说明

ChartView 对象是可由 SheetViews 集合（类似于 Sheets 集合）返回的对象之一。 ChartView 对象只适用于图表工作表。

``` JavaScript
/*以下示例返回 ChartView 对象。*/

ActiveWindow.SheetViews.Item(1)
```

``` JavaScript
/*以下示例返回 Chart 对象。*/
ActiveWindow.SheetViews.Item(1).Sheet
```

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                               |
|:----:|:--------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartView.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#ChartView.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Parent](#ChartView.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  4   | [Sheet](#ChartView.Sheet)             | 返回指定的 ChartView 对象的工作表名称。只读。                                                                                                                      |

## 成员属性

### ChartView.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ChartView 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### ChartView.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartView 对象的变量。

### ChartView.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartView 对象的变量。

### ChartView.Sheet

返回指定的 **ChartView** 对象的工作表名称。只读。

**语法**

**express.Sheet**

\*express是一个代表 ChartView 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ColorStops 对象

## 简介

指定的数据系列中所有 ColorStop 对象的集合。

## 说明

``` JavaScript
/*以下示例显示具有 LinearGradients 的 ColorStops。*/
function test(){
let myInterior = Selection.Interior
    myInterior.Pattern = xlPatternLinearGradient
    myInterior.Gradient.Degree = 90
    myInterior.Gradient.ColorStops.Clear()

    //adds stops after any have been cleared
let myColorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)
    myColorStops1.ThemeColor = xlThemeColorDark1
    myColorStops1.TintAndShade = 0

let myColorStops2 = Selection.Interior.Gradient.ColorStops.Add(1)
    myColorStops2.ThemeColor = xlThemeColorAccent1
    myColorStops2.TintAndShade = 0
}
```

每个 ColorStop 对象代表一个区域或选定内容中渐变填充的一个颜色光圈。

## 方法

| 序号 | 名称                           | 说明 |
|:----:|:-------------------------------|:-----|
|  1   | [Reserve](#ColorStops.Reserve) |      |

## 属性

| 序号 | 名称                           | 说明 |
|:----:|:-------------------------------|:-----|
|  1   | [Reserve](#ColorStops.Reserve) |      |

## 成员方法

### ColorStops.Reserve

**语法**

**express.Reserve ()**

\*express是一个代表 ColorStops 对象的变量。

## 成员属性

### ColorStops.Reserve

**语法**

**express.Reserve**

\*express是一个代表 ColorStops 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ColorStops](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DropLines 对象

## 简介

代表图表组中的垂直线。

## 说明

垂直线将图表中的数据点与 x 轴连接起来。只有折线图和面积图组可以有垂直线。此对象不是集合。没有代表单个垂直线的对象；或者打开图表组中所有数据点的垂直线，或者将其全部关闭。

如果 HasDropLines 属性为 False ， DropLines 对象的绝大部分属性将被禁用。

使用 DropLines 属性可返回 DropLines 对象。

``` JavaScript
/*下例打开嵌入的第一个图表的第一个图表组的垂直线，并将垂直线的颜色设置为红色。*/
function test(){
Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
ActiveChart.ChartGroups(1).HasDropLines = true
ActiveChart.ChartGroups(1).DropLines.Border.ColorIndex = 3
}
```

## 方法

| 序号 | 名称                        | 说明       |
|:----:|:----------------------------|:-----------|
|  1   | [Delete](#DropLines.Delete) | 删除对象。 |
|  2   | [Select](#DropLines.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#DropLines.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Border](#DropLines.Border)           | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  3   | [Creator](#DropLines.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Format](#DropLines.Format)           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  5   | [Name](#DropLines.Name)               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  6   | [Parent](#DropLines.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### DropLines.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 DropLines 对象的变量。

**返回值**

Variant

### DropLines.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 DropLines 对象的变量。

**返回值**

Variant

## 成员属性

### DropLines.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 DropLines 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### DropLines.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 DropLines 对象的变量。

### DropLines.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DropLines 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### DropLines.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 DropLines 对象的变量。

### DropLines.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 DropLines 对象的变量。

### DropLines.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DropLines 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/DropLines](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Font 对象

## 简介

包含对象的字体属性（字体名称、字号、颜色等等）。

## 说明

如果不想将单元格中的文本或图形设为相同的格式，则使用 Characters 属性返回文本的子集。

使用 Font 属性可返回 Font 对象。下例将单元格区域 A1:C5 的格式设置为加粗。

``` JavaScript
Application.Worksheets.Item("Sheet1").Range("A1:C5").Font.Bold = true
```

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Font.Application)     | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Background](#Font.Background)       | 返回或设置图表中使用的文本的背景类型。 Variant 类型，可读写。可将其设置为 XlBackground 常量之一。                                                                                                                               |
|  3   | [Bold](#Font.Bold)                   | 如果字体格式为加粗，则该属性值为 True 。 Variant 类型，可读写。                                                                                                                                                                 |
|  4   | [Color](#Font.Color)                 | 返回或设置对象的主要颜色，如注释部分中的表格所示。使用 RGB 函数可创建颜色值。 Variant 型，可读写。                                                                                                                              |
|  5   | [ColorIndex](#Font.ColorIndex)       | 返回或设置一个 Variant 值，它代表字体的颜色。                                                                                                                                                                                   |
|  6   | [Creator](#Font.Creator)             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  7   | [FontStyle](#Font.FontStyle)         | 返回或设置字体样式。 String 类型，可读写。                                                                                                                                                                                      |
|  8   | [Italic](#Font.Italic)               | 如果字体样式为倾斜，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                 |
|  9   | [Name](#Font.Name)                   | 返回或设置一个 Variant 值，它代表对象的名称。                                                                                                                                                                                   |
|  10  | [Parent](#Font.Parent)               | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  11  | [Size](#Font.Size)                   | 返回或设置字号。可读/写 Variant 类型。                                                                                                                                                                                          |
|  12  | [Strikethrough](#Font.Strikethrough) | 如果文字中间有一条水平删除线，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                       |
|  13  | [Subscript](#Font.Subscript)         | 如果字体格式设置为下标，则该属性值为 True 。默认值为 False 。 Variant 类型，可读写。                                                                                                                                            |
|  14  | [Superscript](#Font.Superscript)     | 如果字体格式设置为上标字符，则该属性值为 True ；默认值为 False 。 Variant 类型，可读写。                                                                                                                                        |
|  15  | [ThemeColor](#Font.ThemeColor)       | 返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 Variant 类型。                                                                                                                                      |
|  16  | [ThemeFont](#Font.ThemeFont)         | 返回或设置与指定对象关联的应用字体方案中的主题字体。可读/写 XlThemeFont 类型。                                                                                                                                                  |
|  17  | [TintAndShade](#Font.TintAndShade)   | 返回或设置一个 Single ，使颜色变深或变浅。                                                                                                                                                                                      |
|  18  | [Underline](#Font.Underline)         | 返回或设置应用于字体的下划线类型。可为以下 XlUnderlineStyle 常量之一。 Variant 类型，可读写。                                                                                                                                   |

## 成员属性

### Font.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
```

### Font.Background

返回或设置图表中使用的文本的背景类型。 **Variant** 类型，可读写。可将其设置为 **XlBackground** 常量之一。

**语法**

**express.Background**

\*express是一个代表 Font 对象的变量。

**说明**

| XlBackground 可为以下常量之一。 | 文本背景描述                                                                                       |
|---------------------------------|----------------------------------------------------------------------------------------------------|
| **xlBackgroundAutomatic**       | 字体背景将自动更改文本周围的背景区域颜色，以使在文本下方元素所应用的颜色之上显示的文本达到最佳效果 |
| **xlBackgroundOpaque**          | 如果文本颜色与文本下的填充颜色非常接近或者为同一颜色，则字体背景会设置为空，以便不显示文本         |
| **xlBackgroundTransparent**     | 字体背景设置为透明，以便当文本颜色接近文本下方的颜色时，文本背景不会更改                           |

**示例**

``` JavaScript
/* 本示例为第一张工作表上的第一张嵌入图表添加图表标题，然后设置标题的字体大小和背景类型。本示例假定图表在第一张工作表上*/
function test() {

let objects1 = Application.Worksheets.Item(1).ChartObjects(1).Chart
objects1.HasTitle = true
objects1.ChartTitle.Text = "Rainfall Totals by Month"
    let objects2 = objects1.ChartTitle.Font
    objects2.Size = 10
    objects2.Background = xlBackgroundTransparent

}
```

### Font.Bold

如果字体格式为加粗，则该属性值为 **True** 。 **Variant** 类型，可读写。

**语法**

**express.Bold**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/* 本示例将工作表 Sheet1 上 A1:A5 区域中的字体设为加粗*/
Application.Worksheets.Item("Sheet1").Range("A1:A5").Font.Bold = true
```

### Font.Color

返回或设置对象的主要颜色，如注释部分中的表格所示。使用 **RGB** 函数可创建颜色值。 **Variant** 型，可读写。

**语法**

**express.Color**

\*express是一个代表 Font 对象的变量。

**说明**

| 对象         | 对应颜色                                                                            |
|--------------|-------------------------------------------------------------------------------------|
| **Border**   | 边框的颜色。                                                                        |
| **Borders**  | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font**     | 字体的颜色。                                                                        |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。                                              |
| **Tab**      | 选项卡的颜色。                                                                      |

**示例**

``` JavaScript
/* 此示例对 Chart1 中数值坐标轴的刻度线标志颜色进行设置*/
Application.Charts.Item("Chart1").Axes(xlValue).TickLabels.Font.Color = (0, 255, 0)
```

### Font.ColorIndex

返回或设置一个 **Variant** 值，它代表字体的颜色。

**语法**

**express.ColorIndex**

\*express是一个代表 Font 对象的变量。

**说明**

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 **XlColorIndex** 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

**示例**

``` JavaScript
/* 本示例将 Sheet1 上 A1 单元格的字体颜色更改为红色*/
Application.Worksheets.Item("Sheet1").Range("A1").Font.ColorIndex = 3
```

### Font.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Font 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Font.FontStyle

返回或设置字体样式。 **String** 类型，可读写。

**语法**

**express.FontStyle**

\*express是一个代表 Font 对象的变量。

**说明**

更改该属性可能会影响其他 **F ont** 属性（例如 **Bold** 和 **Ital ic** ）。

**示例**

``` JavaScript
/* 本示例将 Sheet1 中 A1 单元格的字体样式设置为加粗和倾斜*/
Application.Worksheets.Item("Sheet1").Range("A1").Font.FontStyle = "Bold Italic"
```

### Font.Italic

如果字体样式为倾斜，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Italic**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/* 本示例将 Sheet1 中 A1:A5 区域的字形设置为倾斜*/
Application.Worksheets.Item("Sheet1").Range("A1:A5").Font.Italic = true
```

### Font.Name

返回或设置一个 **Variant** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Font 对象的变量。

**说明**

**Fon t** 对象的名称为 **String** 。

### Font.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Font 对象的变量。

### Font.Size

返回或设置字号。可读/写 **Variant** 类型。

**语法**

**express.Size**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/* 本示例将 Sheet1 的 A1:D10 单元格的字体大小设为 12 磅*/
function test() {
    let size = Application.Worksheets.Item("Sheet1").Range("A1:D10")
    size.Value2 = "Test"
    size.Font.Size = 12
}
```

### Font.Strikethrough

如果文字中间有一条水平删除线，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Strikethrough**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/* 本示例将 Sheet1 中活动单元格的字体设置为加删除线*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
Application.ActiveCell.Font.Strikethrough = true
}
```

### Font.Subscript

如果字体格式设置为下标，则该属性值为 **True** 。默认值为 **False** 。 **Variant** 类型，可读写。

**语法**

**express.Subscript**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/* 本示例将 A1 单元格的第二个字符设置为下标字符*/
Application.Worksheets.Item("Sheet1").Range("A1").Characters(2, 1).Font.Subscript = true
```

### Font.Superscript

如果字体格式设置为上标字符，则该属性值为 **True** ；默认值为 **False** 。 **Variant** 类型，可读写。

**语法**

**express.Superscript**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/* 本示例将 A1 单元格的最后一个字符设置为上标字符*/
function test() {
    let n = Application.Worksheets.Item("Sheet1").Range("A1").Characters().Count
    Application.Worksheets.Item("Sheet1").Range("A1").Characters(n, 1).Font.Superscript = true
}
```

### Font.ThemeColor

返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。

**语法**

**express.ThemeColor**

\*express是一个代表 Font 对象的变量。

**说明**

如果对象当前未应用主题颜色，试图访问其主题颜色则会引起无效请求运行时错误。

### Font.ThemeFont

返回或设置与指定对象关联的应用字体方案中的主题字体。可读/写 **XlThemeFont** 类型。

**语法**

**express.ThemeFont**

\*express是一个代表 Font 对象的变量。

### Font.TintAndShade

返回或设置一个 **Single** ，使颜色变深或变浅。

**语法**

**express.TintAndShade**

\*express是一个代表 Font 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

### Font.Underline

返回或设置应用于字体的下划线类型。可为以下 **XlUnderlineStyle** 常量之一。 **Variant** 类型，可读写。

**语法**

**express.Underline**

\*express是一个代表 Font 对象的变量。

**说明**

|                                                               |
|---------------------------------------------------------------|
| **XlUnderlineStyle** 可为以下 **XlUnderlineStyle** 常量之一。 |
| **xlUnderlineStyleNone**                                      |
| **xlUnderlineStyleSingle**                                    |
| **xlUnderlineStyleDouble**                                    |
| **xlUnderlineStyleSingleAccounting**                          |
| **xlUnderlineStyleDoubleAccounting**                          |

**示例**

``` JavaScript
/* 本示例将 Sheet1 中活动单元格的字体设置为单下划线*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Font.Underline = xlUnderlineStyleSingle
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Font](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListObjects 对象

## 简介

工作表上所有 ListObject 对象的集合。每个 ListObject 对象都代表工作表中的一个表格。

## 说明

使用 Worksheet 对象的 ListObjects 属性可返回 ListObjects 集合。

``` JavaScript
/*下例创建一个新 ListObjects 集合，该集合代表工作表中所有的表格.*/
let myWorksheetLists = Application.Worksheets.Item(1).ListObjects
```

## 方法

| 序号 | 名称                    | 说明               |
|:----:|:------------------------|:-------------------|
|  1   | [Add](#ListObjects.Add) | 创建新的列表对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ListObjects.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#ListObjects.Count)             | 返回一个 Integer 值，它代表集合中对象的数量。                                                                                                                                                                                   |
|  3   | [Creator](#ListObjects.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Item](#ListObjects.Item)               | 从集合中返回一个对象。                                                                                                                                                                                                          |
|  5   | [Parent](#ListObjects.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### ListObjects.Add

创建新的列表对象。

**语法**

**express.Add ( *SourceType* , *Source* , *LinkSource* , *TableStyleName* , *Destination* )**

\*express是一个代表 ListObjects 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型               | 说明                                                                                                                                                                                                                                                                                                                                                                                                     |
|------------------|-----------|------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *SourceType*     | 可选      | XlListObjectSourceType | 表示查询的来源类型。                                                                                                                                                                                                                                                                                                                                                                                     |
| *Source*         | 可选      | Variant                | 当 SourceType 为 xlSrcRange 时，它是代表数据源的 Range 对象。如果省略该参数，Source 的默认值将是列表区域检测代码返回的区域。当 SourceType 为 xlSrcExternal 时，它是一个 String 值数组，用于指定与数据源的连接，包含以下元素： 0 ― SharePoint 网站的 URL 1 ― ListName 2 ― ViewGUID                                                                                                                        |
| *LinkSource*     | 可选      | Variant                | Boolean 型。表示外部数据源是否链接到 ListObject 对象。如果 SourceType 为 xlSrcExternal，则默认值为 True。如果 SourceType 为 xlSrcRange，则此参数无效，如果不省略它，将会返回一个错误。                                                                                                                                                                                                                   |
| *TableStyleName* | 可选      | Variant                | 一个 XlYesNoGuess 常量，它指示正在导入的数据是否有列标签。如果 Source 没有标题，ET 将自动生成标题。                                                                                                                                                                                                                                                                                                      |
| *Destination*    | 可选      | Variant                | 一个 Range 对象，作用是将一个单元格引用指定为新列表对象左上角的目标区域。如果 Range 对象引用多个单元格，则会产生错误。当 SourceType 设置为 xlSrcExternal 时，必须指定 Destination 参数。如果 SourceType 设置为 xlSrcRange，则忽略 Destination 参数。目标区域所在的工作表必须是包含 expression 所指定的 ListObjects 集合的工作表。新列的插入位置将是 Destination 以适应新列表。因此，现有数据不会被覆盖。 |

**返回值**

一个 ListObject 对象，它代表新的列表对象。

**说明**

当列表有标题时，第一行单元格将转换为 **Text** （如果还未被设为文本）。转换将基于单元格的可见文本。这意味着，如果有一个日期值，该日期值的 **Date** 格式随区域设置的更改而更改，则对列表的转换可能产生不同的结果，具体取决于当前的系统区域设置。而且，如果标题行中有两个单元格包含相同的可见文本，则会追加递增的 **Integer** 以使每个列标题唯一。

**示例**

下例在默认 **ListObject** 集合中添加新的 **ListObjects** 对象（该对象基于 Microsoft SharePoint Foundation 网站的数据），并将列表放在工作簿中第一个工作表的 A1 单元格中。

| 注释                                                                                                                                                            |
|-----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 以下代码示例假设您会将 ` strServerName ` 和 ` strListGUID ` 变量替换为有效的服务器名称和列表 GUID。此外，服务器名称后面必须是“/\_vti_bin”，否则本示例将不运行。 |

``` JavaScript
let objListObject = Application.ActiveWorkbook.Worksheets.Item(1).ListObjects.Add(xlSrcExternal, Array(strServerName, strListName, strListGUID), true, xlGuess, Range("A10"))
```

## 成员属性

### ListObjects.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ListObjects 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### ListObjects.Count

返回一个 **Integer** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 ListObjects 对象的变量。

### ListObjects.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ListObjects 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ListObjects.Item

从集合中返回一个对象。

**语法**

**express.Item**

\*express是一个代表 ListObjects 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动工作簿的 Sheet1 上的默认列表对象的名称。*/
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
let oListObj = wrksht.ListObjects.Item(1).Name
}
```

### ListObjects.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ListObjects 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ListObjects](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListRow 对象

## 简介

代表表格中的一行。 ListRow 对象是 ListRows 集合的成员。

## 说明

ListRows 集合包含列表对象中的所有行。 ListRows

使用 ListObject 对象的 ListRows 属性可返回一个 ListRows 集合。

``` JavaScript
/*下例给活动工作簿的第一张工作表的默认 ListObject 对象添加新 ListRow 对象。由于未指定位置，因此在表格结束处添加一个新行。*/
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
let oListRow = wrksht.ListObjects.Item(1).ListRows.Add()
}
```

## 方法

| 序号 | 名称                      | 说明                                                                                                                                                                                |
|:----:|:--------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Delete](#ListRow.Delete) | 删除列表行的单元格并向上移动被删除行下面的任何剩余单元格。即使列表被链接到 SharePoint 网站，您也可以删除该列表中的行。但是，在同步更改之前，SharePoint 网站上的列表将不会进行更新。 |

## 属性

| 序号 | 名称                                | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ListRow.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#ListRow.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Index](#ListRow.Index)             | 返回一个 Long 值，它代表 ListRow 对象在 ListRows 集合中的索引号。                                                                                                                                                               |
|  4   | [Parent](#ListRow.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [Range](#ListRow.Range)             | 返回一个 Range 对象，它代表在以上列表中的指定列表对象所应用的区域。                                                                                                                                                             |

## 成员方法

### ListRow.Delete

删除列表行的单元格并向上移动被删除行下面的任何剩余单元格。即使列表被链接到 SharePoint 网站，您也可以删除该列表中的行。但是，在同步更改之前，SharePoint 网站上的列表将不会进行更新。

**语法**

**express.Delete ()**

\*express是一个代表 ListRow 对象的变量。

## 成员属性

### ListRow.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ListRow 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### ListRow.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ListRow 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ListRow.Index

返回一个 **Long** 值，它代表 **ListRow** 对象在 **ListRows** 集合中的索引号。

**语法**

**express.Index**

\*express是一个代表 ListRow 对象的变量。

### ListRow.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ListRow 对象的变量。

### ListRow.Range

返回一个 **Range** 对象，它代表在以上列表中的指定列表对象所应用的区域。

**语法**

**express.Range**

\*express是一个代表 ListRow 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ListRow](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PageSetup 对象

## 简介

代表页面设置说明。

## 说明

PageSetup 对象包含所有页面设置的属性（左边距、底部边距、纸张大小等）。

使用 PageSetup 属性可返回一个 PageSetup 对象。下例将打印方向设置为横向，然后打印工作表。

``` JavaScript
funtion test(){
    let wss = Application.Worksheets.Item("Sheet1")
    wss.PageSetup.Orientation = xlLandscape
    wss.PrintOut()
}
```

With 语句使同时设置若干属性变得简单而迅速。下例设置第一张工作表的所有页边距。

``` JavaScript
function test(){
    let psp = Application.Worksheets.Item(1).PageSetup
    psp.LeftMargin = Application.InchesToPoints(0.5)
    psp.RightMargin = Application.InchesToPoints(0.75)
    psp.TopMargin = Application.InchesToPoints(1.5)
    psp.BottomMargin = Application.InchesToPoints(1)
    psp.HeaderMargin = Application.InchesToPoints(0.5)
    psp.FooterMargin = Application.InchesToPoints(0.5)
}
```

## 方法

| 序号 | 名称                                    | 说明                                        |
|:----:|:----------------------------------------|:--------------------------------------------|
|  1   | [PrintQuality](#PageSetup.PrintQuality) | 返回或设置打印质量。 Variant 类型，可读写。 |

## 属性

| 序号 | 名称                                                                        | 说明                                                                                                                                                                                                                                                                                                   |
|:----:|:----------------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AlignMarginsHeaderFooter](#PageSetup.AlignMarginsHeaderFooter)             | 如果 ET 以页面设置选项中设置的边距对齐页眉和页脚，则返回 True 。可读/写 Boolean 类型。                                                                                                                                                                                                                 |
|  2   | [Application](#PageSetup.Application)                                       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。                                                                        |
|  3   | [BlackAndWhite](#PageSetup.BlackAndWhite)                                   | 如果指定文档中的元素以黑白方式打印，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                        |
|  4   | [BottomMargin](#PageSetup.BottomMargin)                                     | 以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置底端边距的大小。 Double 类型，可读写。       |
|  5   | [CenterFooter](#PageSetup.CenterFooter)                                     | 居中对齐 PageSetup 对象中的页脚信息。可读/写 String 类型。                                                                                                                                                                                                                                             |
|  6   | [CenterFooterPicture](#PageSetup.CenterFooterPicture)                       | 返回一个 Graphic 对象，该对象表示页脚中间部分的图片。用于设置图片的属性。                                                                                                                                                                                                                              |
|  7   | [CenterHeader](#PageSetup.CenterHeader)                                     | 居中对齐 PageSetup 对象中的页眉信息。可读/写 String 类型。                                                                                                                                                                                                                                             |
|  8   | [CenterHeaderPicture](#PageSetup.CenterHeaderPicture)                       | 返回一个 Graphic 对象，该对象表示页眉中间部分的图片。用于设置图片的属性。                                                                                                                                                                                                                              |
|  9   | [CenterHorizontally](#PageSetup.CenterHorizontally)                         | 如果在页面的水平居中位置打印指定工作表，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                    |
|  10  | [CenterVertically](#PageSetup.CenterVertically)                             | 如果在页面的垂直居中位置打印指定工作表，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                    |
|  11  | [Creator](#PageSetup.Creator)                                               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                                                                                             |
|  12  | [DifferentFirstPageHeaderFooter](#PageSetup.DifferentFirstPageHeaderFooter) | 如果在第一页使用不同的页眉或页脚，则为 True 。可读/写 Boolean 类型。                                                                                                                                                                                                                                   |
|  13  | [Draft](#PageSetup.Draft)                                                   | 如果打印工作表时不打印其中的图形，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                          |
|  14  | [EvenPage](#PageSetup.EvenPage)                                             | 返回或设置工作簿或节的偶数页上的文本对齐方式。                                                                                                                                                                                                                                                         |
|  15  | [FirstPage](#PageSetup.FirstPage)                                           | 返回或设置工作簿或节的第一页上的文本对齐方式。                                                                                                                                                                                                                                                         |
|  16  | [FirstPageNumber](#PageSetup.FirstPageNumber)                               | 返回或设置打印指定工作表时第一页的页号。如果设为 xlAutomatic ，则 ET 采用第一页的页号。默认值为 xlAutomatic 。 Long 类型，可读写。                                                                                                                                                                     |
|  17  | [FitToPagesTall](#PageSetup.FitToPagesTall)                                 | 返回或设置打印工作表时，对工作表进行缩放使用的页高。仅应用于工作表。 Variant 类型，可读写。                                                                                                                                                                                                            |
|  18  | [FitToPagesWide](#PageSetup.FitToPagesWide)                                 | 返回或设置打印工作表时，对工作表进行缩放使用的页宽。仅应用于工作表。 Variant 类型，可读写。                                                                                                                                                                                                            |
|  19  | [FooterMargin](#PageSetup.FooterMargin)                                     | 以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置页脚到页面底端的距离。 Double 类型，可读写。 |
|  20  | [HeaderMargin](#PageSetup.HeaderMargin)                                     | 以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置页面顶端到页眉的距离。 Double 类型，可读写。 |
|  21  | [LeftFooter](#PageSetup.LeftFooter)                                         | 返回或设置工作簿或节的左页脚上的文本对齐方式。                                                                                                                                                                                                                                                         |
|  22  | [LeftFooterPicture](#PageSetup.LeftFooterPicture)                           | 返回一个 Graphic 对象，该对象表示页脚左边的图片。用于设置图片的属性。                                                                                                                                                                                                                                  |
|  23  | [LeftHeader](#PageSetup.LeftHeader)                                         | 返回或设置工作簿或节的左页眉上的文本对齐方式。                                                                                                                                                                                                                                                         |
|  24  | [LeftHeaderPicture](#PageSetup.LeftHeaderPicture)                           | 返回一个 Graphic 对象，该对象表示页眉左边的图片。用于设置图片的属性。                                                                                                                                                                                                                                  |
|  25  | [LeftMargin](#PageSetup.LeftMargin)                                         | 以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置左边距的大小。 Double 类型，可读写。         |
|  26  | [OddAndEvenPagesHeaderFooter](#PageSetup.OddAndEvenPagesHeaderFooter)       | 如果指定的 PageSetup 对象的奇数页和偶数页具有不同的页眉和页脚，则为 True 。可读/写 Boolean 类型。                                                                                                                                                                                                      |
|  27  | [Order](#PageSetup.Order)                                                   | 返回或设置一个 XlOrder 值，它代表 ET 打印一张大工作表时所使用的页编号的次序。                                                                                                                                                                                                                          |
|  28  | [Orientation](#PageSetup.Orientation)                                       | 返回或设置一个 XlPageOrientation 值，它代表纵向或横向打印模式。                                                                                                                                                                                                                                        |
|  29  | [Pages](#PageSetup.Pages)                                                   | 返回或设置 Pages 集合中的页数。                                                                                                                                                                                                                                                                        |
|  30  | [PaperSize](#PageSetup.PaperSize)                                           | 返回或设置纸张的大小。 XlPaperSize 类型，可读写。                                                                                                                                                                                                                                                      |
|  31  | [Parent](#PageSetup.Parent)                                                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                                                                                           |
|  32  | [PrintArea](#PageSetup.PrintArea)                                           | 以字符串返回或设置要打印的区域，该字符串使用宏语言的 A1 样式的引用。 String 类型，可读写。                                                                                                                                                                                                             |
|  33  | [PrintComments](#PageSetup.PrintComments)                                   | 返回或设置批注随工作表打印的方式。 XlPrintLocation 类型，可读写。                                                                                                                                                                                                                                      |
|  34  | [PrintErrors](#PageSetup.PrintErrors)                                       | 设置或返回一个 XlPrintErrors 常量，该常量指定显示的打印错误类型。该功能允许用户在打印工作表时取消错误显示。可读写。                                                                                                                                                                                    |
|  35  | [PrintGridlines](#PageSetup.PrintGridlines)                                 | 如果在页面上打印单元格网格线，则该值为 True 。仅应用于工作表。 Boolean 类型，可读写。                                                                                                                                                                                                                  |
|  36  | [PrintHeadings](#PageSetup.PrintHeadings)                                   | 如果打印本页时同时打印行标题和列标题，则该值为 True 。仅应用于工作表。 Boolean 类型，可读写。                                                                                                                                                                                                          |
|  37  | [PrintNotes](#PageSetup.PrintNotes)                                         | 如果打印工作表时单元格批注作为尾注一起打印，则该值为 True 。仅应用于工作表。 Boolean 类型，可读写。                                                                                                                                                                                                    |
|  38  | [PrintTitleColumns](#PageSetup.PrintTitleColumns)                           | 返回或设置包含在每一页的左边重复出现的单元格的列，用宏语言 A1-样式中的字符串表示。 String 类型，可读写。                                                                                                                                                                                               |
|  39  | [PrintTitleRows](#PageSetup.PrintTitleRows)                                 | 返回或设置那些包含在每一页顶部重复出现的单元格的行，用宏语言字符串以 A1 样式表示法表示。 String 类型，可读写。                                                                                                                                                                                         |
|  40  | [RightFooter](#PageSetup.RightFooter)                                       | 返回或设置页面右边缘与页脚右边界之间的距离（以磅为单位）。可读/写 String 类型。                                                                                                                                                                                                                        |
|  41  | [RightFooterPicture](#PageSetup.RightFooterPicture)                         | 返回一个 Graphic 对象，该对象代表页脚右边的图片，用于设置图片的属性。                                                                                                                                                                                                                                  |
|  42  | [RightHeader](#PageSetup.RightHeader)                                       | 返回或设置页眉的右边部分内容。可读/写 String 类型。                                                                                                                                                                                                                                                    |
|  43  | [RightHeaderPicture](#PageSetup.RightHeaderPicture)                         | 指定右页眉中应显示的图形图像。只读。                                                                                                                                                                                                                                                                   |
|  44  | [RightMargin](#PageSetup.RightMargin)                                       | 以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置右边距的大小。 Double 类型，可读写。         |
|  45  | [ScaleWithDocHeaderFooter](#PageSetup.ScaleWithDocHeaderFooter)             | 返回或设置页眉和页脚是否在文档大小更改时随文档缩放。可读/写 Boolean 类型。                                                                                                                                                                                                                             |
|  46  | [TopMargin](#PageSetup.TopMargin)                                           | 以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置上边距的大小。 Double 类型，可读写。         |
|  47  | [Zoom](#PageSetup.Zoom)                                                     | 返回或设置一个 Variant 值，它代表一个数值在 10% 到 400% 之间的百分比，该百分比为 ET 打印工作表时的缩放比例。                                                                                                                                                                                           |

## 成员方法

### PageSetup.PrintQuality

返回或设置打印质量。 **Variant** 类型，可读写。

**语法**

**express.PrintQuality ( *Index* )**

\*express是一个代表 PageSetup 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                      |
|---------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Index* | 可选      | Variant  | 水平方向打印质量 (1) 或垂直方向打印质量 (2)。有些打印机可能不支持垂直方向打印质量设置。如果不指定该参数，则 PrintQuality 属性将返回（或可被设置为）一个二元数组，该数组包含水平方向和垂直方向的打印质量。 |

**示例**

``` JavaScript
function test(){
  /*用非正方形像素设置打印机上的打印质量。数组指定水平方向和垂直方向的打印质量。本示例是否出现错误取决于使用的打印机驱动程序。*/
  Application.Worksheets.Item("Sheet1").PageSetup.PrintQuality = [240, 140]
  
  /*显示水平方向打印质量的当前设置。*/
  alert("Horizontal Print Quality is " + Application.Worksheets.Item("Sheet1").PageSetup.PrintQuality(1))
}
```

## 成员属性

### PageSetup.AlignMarginsHeaderFooter

如果 ET 以页面设置选项中设置的边距对齐页眉和页脚，则返回 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.AlignMarginsHeaderFooter**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本示例显示一条有关创建 myObject 的应用程序的消息。*/
  let myObject = Application.ActiveWorkbook
  if (myObject.Application.Value == "ET"){
      alert("This is an ET Application object.")
  }
  else{
      alert("This is not an ET Application object.")
  }
}
```

### PageSetup.BlackAndWhite

如果指定文档中的元素以黑白方式打印，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.BlackAndWhite**

\*express是一个代表 PageSetup 对象的变量。

**说明**

该属性仅适用于工作表页面。

**示例**

``` JavaScript
/*本示例设置工作表 Sheet1 以黑白方式打印。*/
Application.Worksheets.Item("Sheet1").PageSetup.BlackAndWhite = true
```

### PageSetup.BottomMargin

以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置底端边距的大小。 **Double** 类型，可读写。

**语法**

**express.BottomMargin**

\*express是一个代表 PageSetup 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **I nchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

``` JavaScript
function test(){  
  /*将工作表 Sheet1 的底端边距设为 0.5 英寸（36 磅）。*/
  Application.Worksheets.Item("Sheet1").PageSetup.BottomMargin = Application.InchesToPoints(0.5)
  Application.Worksheets.Item("Sheet1").PageSetup.BottomMargin = 36
  
  /*显示 Sheet1 底端边距的当前设定值。*/
  let marginInches = Application.Worksheets.Item("Sheet1").PageSetup.BottomMargin / 
  Application.InchesToPoints(1)
  alert("The current bottom margin is " + marginInches + " inches")
}
```

### PageSetup.CenterFooter

居中对齐 **PageSetup** 对象中的页脚信息。可读/写 **String** 类型。

**语法**

**express.CenterFooter**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.CenterFooterPicture

返回一个 **Graphic** 对象，该对象表示页脚中间部分的图片。用于设置图片的属性。

**语法**

**express.CenterFooterPicture**

\*express是一个代表 PageSetup 对象的变量。

**说明**

**CenterFooterPicture** 属性为只读，但并非它的所有属性都为只读。

要求“&G”为 **CenterFooter** 属性字符串的组成部分，以便在页脚的中间显示图像。

**示例**

``` JavaScript
function InsertPicture() {
  /*本示例将 C:\ 驱动器中名为“Sample.jpg”的图片添加到页脚的中间部分(“Sample.jpg”文件位于 C:\ 驱动器上。)*/
  let picture = Application.ActiveSheet.PageSetup.CentertFooterPicture
  picture.FileName = "C:\\Sample.jpg"
  picture.Height = 275.25
  picture.Width = 463.5
  picture.Brightness = 0.36
  picture.ColorType = msoPictureGrayscale
  picture.Contrast = 0.39
  picture.CropBottom = -14.4
  picture.CropLeft = -28.8
  picture.CropRight = -14.4
  picture.CropTop = 21.6
                                    
  // Enable the image to show up in the center footer.
  ActiveSheet.PageSetup.CenterFooter = "&G"
  

}
```

### PageSetup.CenterHeader

居中对齐 **PageSetup** 对象中的页眉信息。可读/写 **String** 类型。

**语法**

**express.CenterHeader**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.CenterHeaderPicture

返回一个 **Graphic** 对象，该对象表示页眉中间部分的图片。用于设置图片的属性。

**语法**

**express.CenterHeaderPicture**

\*express是一个代表 PageSetup 对象的变量。

**说明**

**CenterHeaderPicture** 属性为只读，但并非它的所有属性都为只读。

要求“&G”为 **CenterHeader** 属性字符串的组成部分，以便在中间的页眉中显示图像。

**示例**

``` JavaScript
function InsertPicture() {
  /*本示例将 C:\ 驱动器中名为“Sample.jpg”的图片添加到页眉的中间部分(“Sample.jpg”文件位于 C:\ 驱动器上。)*/
  let picture = Application.ActiveSheet.PageSetup.CentertHeaderPicture
  picture.FileName = "C:\\Sample.jpg"
  picture.Height = 275.25
  picture.Width = 463.5
  picture.Brightness = 0.36
  picture.ColorType = msoPictureGrayscale
  picture.Contrast = 0.39
  picture.CropBottom = -14.4
  picture.CropLeft = -28.8
  picture.CropRight = -14.4
  picture.CropTop = 21.6
                                        
  // Enable the image to show up in the center footer.
  Application.ActiveSheet.PageSetup.CenterFooter = "&G"
  
                                        
}
```

### PageSetup.CenterHorizontally

如果在页面的水平居中位置打印指定工作表，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CenterHorizontally**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将工作表 Sheet1 设置成水平居中打印。*/
Application.Worksheets.Item("Sheet1").PageSetup.CenterHorizontally = true
```

### PageSetup.CenterVertically

如果在页面的垂直居中位置打印指定工作表，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CenterVertically**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将 Sheet1 设为垂直居中打印。*/
Application.Worksheets.Item("Sheet1").PageSetup.CenterVertically = true
```

### PageSetup.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PageSetup.DifferentFirstPageHeaderFooter

如果在第一页使用不同的页眉或页脚，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.DifferentFirstPageHeaderFooter**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.Draft

如果打印工作表时不打印其中的图形，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Draft**

\*express是一个代表 PageSetup 对象的变量。

**说明**

将该属性设置为 **True** 可加快打印速度（但是不打印其中的图形）。

**示例**

``` JavaScript
/*本示例关闭对 Sheet1 中图形的打印。*/
Application.Worksheets.Item("Sheet1").PageSetup.Draft = true
```

### PageSetup.EvenPage

返回或设置工作簿或节的偶数页上的文本对齐方式。

**语法**

**express.EvenPage**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.FirstPage

返回或设置工作簿或节的第一页上的文本对齐方式。

**语法**

**express.FirstPage**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.FirstPageNumber

返回或设置打印指定工作表时第一页的页号。如果设为 **xlAutomatic** ，则 ET 采用第一页的页号。默认值为 **xlAutomatic** 。 **Long** 类型，可读写。

**语法**

**express.FirstPageNumber**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将 Sheet1 打印时的第一页的页号设置为 100。*/
Application.Worksheets.Item("Sheet1").PageSetup.FirstPageNumber = 100
```

### PageSetup.FitToPagesTall

返回或设置打印工作表时，对工作表进行缩放使用的页高。仅应用于工作表。 **Variant** 类型，可读写。

**语法**

**express.FitToPagesTall**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果该属性值为 **False** ，ET 将根据 **FitToPagesWide** 属性缩放工作表。

如果 **Zoom** 属性值为 **True** ，将忽略 **FitToPagesTall** 属性。

**示例**

``` JavaScript
function test(){  
  /*本示例设置 ET 准确按照一页的宽度和高度打印 Sheet1。*/
  let pagesetup = Application.Worksheets.Item("Sheet1").PageSetup
  pagesetup.Zoom = false
  pagesetup.FitToPagesTall = 1
  pagesetup.FitToPagesWide = 1
}
```

### PageSetup.FitToPagesWide

返回或设置打印工作表时，对工作表进行缩放使用的页宽。仅应用于工作表。 **Variant** 类型，可读写。

**语法**

**express.FitToPagesWide**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果该属性值为 **False** ，ET 将根据 **FitToPagesTall** 属性缩放工作表。

如果 **Zoom** 属性值为 **True** ，则忽略 **FitToPagesWide** 属性。

**示例**

``` JavaScript
function test(){
  /*本示例设置 ET 准确按照一页的宽度和高度打印 Sheet1。*/
  let pagesetup = Application.Worksheets.Item("Sheet1").PageSetup
  pagesetup.Zoom = false
  pagesetup.FitToPagesTall = 1
  pagesetup.FitToPagesWide = 1
}
```

### PageSetup.FooterMargin

以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置页脚到页面底端的距离。 **Double** 类型，可读写。

**语法**

**express.FooterMargin**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将 Sheet1 的页脚边矩设置为 0.5 英寸。*/
Application.Worksheets.Item("Sheet1").PageSetup.FooterMargin = Application.InchesToPoints(0.5)
```

### PageSetup.HeaderMargin

以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置页面顶端到页眉的距离。 **Double** 类型，可读写。

**语法**

**express.HeaderMargin**

\*express是一个代表 PageSetup 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **InchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

``` JavaScript
/*本示例将 Sheet1 的页眉边距设置为 0.5 英寸。*/
Application.Worksheets.Item("Sheet1").PageSetup.HeaderMargin = Application.InchesToPoints(0.5)
```

### PageSetup.LeftFooter

返回或设置工作簿或节的左页脚上的文本对齐方式。

**语法**

**express.LeftFooter**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.LeftFooterPicture

返回一个 **Graphic** 对象，该对象表示页脚左边的图片。用于设置图片的属性。

**语法**

**express.LeftFooterPicture**

\*express是一个代表 PageSetup 对象的变量。

**说明**

**LeftFooterPicture** 属性为只读，但并非其所有属性都为只读。

| 注释                                                                           |
|--------------------------------------------------------------------------------|
| 要求 "&G" 为 **LeftFooter** 属性字符串的组成部分，以便在左边的页脚中显示图像。 |

**示例**

``` JavaScript
function InsertPicture() {
  /*本示例将 C:\ 驱动器中名为“Sample.jpg”的图片添加到页脚的左边(假定“Sample.jpg”文件位于 C:\ 驱动器上。)*/
  let picture = Application.ActiveSheet.PageSetup.LeftFooterPicture
  picture.FileName = "C:\\Sample.jpg"
  picture.Height = 275.25
  picture.Width = 463.5
  picture.Brightness = 0.36
  picture.ColorType = msoPictureGrayscale
  picture.Contrast = 0.39
  picture.CropBottom = -14.4
  picture.CropLeft = -28.8
  picture.CropRight = -14.4
  picture.CropTop = 21.6
                                        
  // Enable the image to show up in the center footer.
  Application.ActiveSheet.PageSetup.CenterFooter = "&G"
  
                                        
}
```

### PageSetup.LeftHeader

返回或设置工作簿或节的左页眉上的文本对齐方式。

**语法**

**express.LeftHeader**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.LeftHeaderPicture

返回一个 **Graphic** 对象，该对象表示页眉左边的图片。用于设置图片的属性。

**语法**

**express.LeftHeaderPicture**

\*express是一个代表 PageSetup 对象的变量。

**说明**

**LeftHeaderPicture** 属性为只读，但并非其所有属性都为只读。

| 注释                                                                           |
|--------------------------------------------------------------------------------|
| 要求 "&G" 为 **LeftHeader** 属性字符串的组成部分，以便在左边的页眉中显示图像。 |

**示例**

``` JavaScript
function test() {
  /*本示例将 C: 驱动器中名为“Sample.jpg”的图片添加到页眉的左边。(假定“Sample.jpg”文件位于 C: 驱动器上。)*/
  let picture = Application.ActiveSheet.PageSetup.LeftHeaderPicture
  picture.FileName = "C:\\Sample.jpg"
  picture.Height = 275.25
  picture.Width = 463.5
  picture.Brightness = 0.36
  picture.ColorType = msoPictureGrayscale
  picture.Contrast = 0.39
  picture.CropBottom = -14.4
  picture.CropLeft = -28.8
  picture.CropRight = -14.4
  picture.CropTop = 21.6
                                        
  // Enable the image to show up in the center footer.
  Application.ActiveSheet.PageSetup.CenterFooter = "&G"
  
                                        
}
```

### PageSetup.LeftMargin

以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置左边距的大小。 **Double** 类型，可读写。

**语法**

**express.LeftMargin**

\*express是一个代表 PageSetup 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **InchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

``` JavaScript
function test(){ 
  /*将 Sheet1 的左边距设置为 1.5 英寸。*/
  Application.Worksheets.Item("Sheet1").PageSetup.LeftMargin = Application.InchesToPoints(1.5)
  
  /*将 Sheet1 的左页边距设置为 2 厘米。*/
  Application.Worksheets.Item("Sheet1").PageSetup.LeftMargin = Application.CentimetersToPoints(2)
  
  /*显示 Sheet1 的左边距的当前设定值。*/
  let marginInches = Application.Worksheets.Item("Sheet1").PageSetup.LeftMargin /
  Application.InchesToPoints(1)
  alert("The current left margin is " + marginInches + " inches")
}
```

### PageSetup.OddAndEvenPagesHeaderFooter

如果指定的 **PageSetup** 对象的奇数页和偶数页具有不同的页眉和页脚，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.OddAndEvenPagesHeaderFooter**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.Order

返回或设置一个 **XlOrder** 值，它代表 ET 打印一张大工作表时所使用的页编号的次序。

**语法**

**express.Order**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例设置打印 Sheet1 时将其分为多页打印。按照从左到右，从上到下的顺序编号和打印。*/
Application.Worksheets.Item("Sheet1").PageSetup.Order = xlOverThenDown
```

### PageSetup.Orientation

返回或设置一个 **XlPageOrientation** 值，它代表纵向或横向打印模式。

**语法**

**express.Orientation**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例将 Sheet1 设置为横向打印。*/
Application.Worksheets.Item("Sheet1").PageSetup.Orientation = xlLandscape
```

### PageSetup.Pages

返回或设置 **Pages** 集合中的页数。

**语法**

**express.Pages**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.PaperSize

返回或设置纸张的大小。 XlPaperSize 类型，可读写。

**语法**

**express.PaperSize**

\*express是一个代表 PageSetup 对象的变量。

**说明**

|                                                                     |
|---------------------------------------------------------------------|
| **XlPaperSize** 可为以下 **XlPaperSize** 常量之一。                 |
| **xlPaper11x17** ：11 in. x 17 in.                                  |
| **xlPaperA4** ：A4 (210 mm x 297 mm)                                |
| **xlPaperA5** ：A5 (148 mm x 210 mm)                                |
| **xlPaperB5** ：A5 (148 mm x 210 mm)                                |
| **xlPaperDsheet** ：D 型纸                                          |
| **xlPaperEnvelope11** ：信封 \#11 (4-1/2 in. x 10-3/8 in.)          |
| **xlPaperEnvelope14** ：信封 \#14 (5 in. x 11-1/2 in.)              |
| **xlPaperEnvelopeB4** ：信封 B4 (250 mm x 353 mm)                   |
| **xlPaperEnvelopeB6** ：信封 (176 mm x 125 mm)                      |
| **xlPaperEnvelopeC4** ：信封 C4 (229 mm x 324 mm)                   |
| **xlPaperEnvelopeC6** ：信封 C6 (114 mm x 162 mm)                   |
| **xlPaperEnvelopeDL** ：信封 DL (110 mm x 220 mm)                   |
| **xlPaperEnvelopeMonarch** ：君主式信封 (3-7/8 in. x 7-1/2 in.)     |
| **xlPaperEsheet** ：E 型纸                                          |
| **xlPaperFanfoldLegalGerman** ：德国正规复写簿 (8-1/2 in. x 13 in.) |
| **xlPaperFanfoldUS** ：美国标准复写簿 (14-7/8 in. x 11 in.)         |
| **xlPaperLedger** ：帐单 (17 in. x 11 in.)                          |
| **xlPaperLetter** ：信函 (8-1/2 in. x 11 in.)                       |
| **xlPaperNote** ：便笺 (8-1/2 in. x 11 in.)                         |
| **xlPaperStatement** ：报告单 (5-1/2 in. x 8-1/2 in.)               |
| **xlPaperUser** ：用户自定义                                        |
| **xlPaper10x14** ：10 in. x 14 in.                                  |
| **xlPaperA3** ：A3 (297 mm x 420 mm)                                |
| **xlPaperA4Small** ：A4（小）(210 mm x 297 mm)                      |
| **xlPaperB4** ：B4 (250 mm x 354 mm)                                |
| **xlPaperCsheet** ：C 型纸                                          |
| **xlPaperEnvelope10** ：信封 \#10 (4-1/8 in. x 9-1/2 in.)           |
| **xlPaperEnvelope12** ：信封 \#12 (4-1/2 in. x 11 in.)              |
| **xlPaperEnvelope9** ：信封 \#9 (3-7/8 in. x 8-7/8 in.)             |
| **xlPaperEnvelopeB5** ：信封 B5 (176 mm x 250 mm)                   |
| **xlPaperEnvelopeC3** ：信封 C3 (324 mm x 458 mm)                   |
| **xlPaperEnvelopeC5** ：信封 C5 (162 mm x 229 mm)                   |
| **xlPaperEnvelopeC65** ：信封 C65 (114 mm x 229 mm)                 |
| **xlPaperEnvelopeItaly** ：信封 (110 mm x 230 mm)                   |
| **xlPaperEnvelopePersonal** ：信封 (3-5/8 in. x 6-1/2 in.)          |
| **xlPaperExecutive** ：行政公文纸 (7-1/2 in. x 10-1/2 in.)          |
| **xlPaperFanfoldStdGerman** ：德国正规复写簿 (8-1/2 in. x 13 in.)   |
| **xlPaperFolio** ：对开 (8-1/2 in. x 13 in.)                        |
| **xlPaperLegal** ：正式信封 (8-1/2 in. x 14 in.)                    |
| **xlPaperLetterSmall** ：简式信纸 (8-1/2 in. x 11 in.)              |
| **xlPaperQuarto** ：四开 (215 mm x 275 mm)                          |
| **xlPaperTabloid** ：文摘 (11 in. x 17 in.)                         |

| 注释                                     |
|------------------------------------------|
| 上面这些纸张大小可能不为所有打印机支持。 |

**示例**

``` JavaScript
/*本示例将 Sheet1 的纸张大小设置为正式信封。*/
Application.Worksheets.Item("Sheet1").PageSetup.PaperSize = xlPaperLegal
```

### PageSetup.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.PrintArea

以字符串返回或设置要打印的区域，该字符串使用宏语言的 A1 样式的引用。 **String** 类型，可读写。

**语法**

**express.PrintArea**

\*express是一个代表 PageSetup 对象的变量。

**说明**

将该属性设置为 **False** 或空字符串 ("")，可打印整个工作表。

该属性仅适用于工作表页面。

**示例**

``` JavaScript
function test(){
  /*将打印区域设置为 Sheet1 上的单元格 A1:C5。*/
  Application.Worksheets.Item("Sheet1").PageSetup.PrintArea = "$A$1:$C$5"
  
  /*将 sheet1 的打印区域设置为当前区。注意使用 Address 属性返回 A1-样式的地址。*/
  Application.Worksheets.Item("Sheet1").Activate()
  Application.ActiveSheet.PageSetup.PrintArea = ActiveCell.CurrentRegion.Address
}
```

### PageSetup.PrintComments

返回或设置批注随工作表打印的方式。 **XlPrintLocation** 类型，可读写。

**语法**

**express.PrintComments**

\*express是一个代表 PageSetup 对象的变量。

**说明**

|                                                             |
|-------------------------------------------------------------|
| **XlPrintLocation** 可为以下 **XlPrintLocation** 常量之一。 |
| **xlPrintInPlace**                                          |
| **xlPrintNoComments**                                       |
| **xlPrintSheetEnd**                                         |

**示例**

``` JavaScript
/*打印第一张工作表时，本示例使批注以尾注的方式打印。*/
Application.Worksheets.Item(1).PageSetup.PrintComments = xlPrintSheetEnd
```

### PageSetup.PrintErrors

设置或返回一个 **XlPrintErrors** 常量，该常量指定显示的打印错误类型。该功能允许用户在打印工作表时取消错误显示。可读写。

**语法**

**express.PrintErrors**

\*express是一个代表 PageSetup 对象的变量。

**说明**

|                                                         |
|---------------------------------------------------------|
| **XlPrintErrors** 可为以下 **XlPrintErrors** 常量之一。 |
| **xlPrintErrorsBlank**                                  |
| **xlPrintErrorsDash**                                   |
| **xlPrintErrorsDisplayed**                              |
| **xlPrintErrorsNA**                                     |

**示例**

``` JavaScript
function UsePrintErrors() {
    /*在本示例中，ET 在活动工作表中使用返回错误的公式。设置 PrintErrors 属性以显示短划线。打印预览窗口显示打印错误的短划线。(假定已安装了打印驱动器)*/
    let wksOne = Application.ActiveSheet
                                    
    // Create a formula that returns an error value.
    Range("A1").Value = 1
    Range("A2").Value = 0
    Range("A3").Formula = "=A1/A2"
                                    
    // Change print errors to display dashes.
    wksOne.PageSetup.PrintErrors = xlPrintErrorsDash
                                    
    // Use the Print Preview window to see the dashes used for print errors.
    ActiveWindow.SelectedSheets.PrintPreview()
}
```

### PageSetup.PrintGridlines

如果在页面上打印单元格网格线，则该值为 **True** 。仅应用于工作表。 **Boolean** 类型，可读写。

**语法**

**express.PrintGridlines**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例在打印 Sheet1 时同时打印单元格网格线。*/
Application.Worksheets.Item("Sheet1").PageSetup.PrintGridlines = true
```

### PageSetup.PrintHeadings

如果打印本页时同时打印行标题和列标题，则该值为 **True** 。仅应用于工作表。 **Boolean** 类型，可读写。

**语法**

**express.PrintHeadings**

\*express是一个代表 PageSetup 对象的变量。

**说明**

**DisplayHeadings** 属性控制标题在屏幕上的显示。

**示例**

``` JavaScript
/*本示例关闭打印 Sheet1 中的标题。*/
Application.Worksheets.Item("Sheet1").PageSetup.PrintHeadings = false
```

### PageSetup.PrintNotes

如果打印工作表时单元格批注作为尾注一起打印，则该值为 **True** 。仅应用于工作表。 **Boolean** 类型，可读写。

**语法**

**express.PrintNotes**

\*express是一个代表 PageSetup 对象的变量。

**说明**

使用 **PrintComments** 属性可以文本框或尾注形式打印批注。

**示例**

``` JavaScript
/*本示例关闭打印批注。*/
Application.Worksheets.Item("Sheet1").PageSetup.PrintNotes = false
```

### PageSetup.PrintTitleColumns

返回或设置包含在每一页的左边重复出现的单元格的列，用宏语言 A1-样式中的字符串表示。 **String** 类型，可读写。

**语法**

**express.PrintTitleColumns**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果仅指定列的一部分，ET 将自动把该区域扩展为整个列。

将该属性设置为 **False** 或空字符串 ("")，将会关闭标题列。

该属性仅适用于工作表页面。

**示例**

``` JavaScript
function test(){
  /*本示例将第三行定义为标题行，将第一列到第三列定义为标题列。*/
  Application.Worksheets.Item("Sheet1").Activate()
  Application.ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows.Item(3).Address
  Application.ActiveSheet.PageSetup.PrintTitleColumns = ActiveSheet.Columns.Item("A:C").Address
}
```

### PageSetup.PrintTitleRows

返回或设置那些包含在每一页顶部重复出现的单元格的行，用宏语言字符串以 A1 样式表示法表示。 **String** 类型，可读写。

**语法**

**express.PrintTitleRows**

\*express是一个代表 PageSetup 对象的变量。

**说明**

如果仅指定行的一部分，ET 将把该区域扩展为整个行。

将该属性设置为 **False** 或空字符串 ("")，将会关闭标题行。

该属性仅适用于工作表页面。

**示例**

``` JavaScript
function test(){  
  /*本示例将第三行定义为标题行，将第一列到第三列定义为标题列。*/
  Application.Worksheets.Item("Sheet1").Activate()
  Application.ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows.Item(3).Address
  Application.ActiveSheet.PageSetup.PrintTitleColumns = ActiveSheet.Columns.Item("A:C").Address
}
```

### PageSetup.RightFooter

返回或设置页面右边缘与页脚右边界之间的距离（以磅为单位）。可读/写 **String** 类型。

**语法**

**express.RightFooter**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.RightFooterPicture

返回一个 **Graphic** 对象，该对象代表页脚右边的图片，用于设置图片的属性。

**语法**

**express.RightFooterPicture**

\*express是一个代表 PageSetup 对象的变量。

**说明**

**RightFooterPicture** 属性为只读，但并非其所有属性都为只读。

**示例**

``` JavaScript
function InsertPicture() {
    /*本示例将 C: 驱动器中名为“Sample.jpg”的图片添加到页脚的右边。（假定“Sample.jpg”文件位于 C: 驱动器上。）*/
    et picture = Application.ActiveSheet.PageSetup.RightFooterPicture
    picture.FileName = "C:\\Sample.jpg"
    picture.Height = 275.25
    picture.Width = 463.5
    picture.Brightness = 0.36
    picture.ColorType = msoPictureGrayscale
    picture.Contrast = 0.39
    picture.CropBottom = -14.4
    picture.CropLeft = -28.8
    picture.CropRight = -14.4
    picture.CropTop = 21.6
                                        
    // Enable the image to show up in the center footer.
    Application.ActiveSheet.PageSetup.CenterFooter = "&G"
                                        
}
```

### PageSetup.RightHeader

返回或设置页眉的右边部分内容。可读/写 **String** 类型。

**语法**

**express.RightHeader**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*本示例在每页的右上角打印文件名。*/
Application.Worksheets.Item("Sheet1").PageSetup.RightHeader = "&F"
```

### PageSetup.RightHeaderPicture

指定右页眉中应显示的图形图像。只读。

**语法**

**express.RightHeaderPicture**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.RightMargin

以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置右边距的大小。 **Double** 类型，可读写。

**语法**

**express.RightMargin**

\*express是一个代表 PageSetup 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **InchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

``` JavaScript
function test(){
  /*将 Sheet1 的右边距设置为 1.5 英寸。*/
  Application.Worksheets.Item("Sheet1").PageSetup.RightMargin = Application.InchesToPoints(1.5)
  
  /*将 Sheet1 的右边距设置为 2 厘米。*/
  Application.Worksheets.Item("Sheet1").PageSetup.RightMargin = Application.CentimetersToPoints(2)
  
  /*显示 Sheet1 中右边距的当前设置。*/
  let marginInches = Application.Worksheets.Item("Sheet1").PageSetup.RightMargin /
  Application.InchesToPoints(1)
  alert("The current right margin is " + marginInches + " inches")
}
```

### PageSetup.ScaleWithDocHeaderFooter

返回或设置页眉和页脚是否在文档大小更改时随文档缩放。可读/写 **Boolean** 类型。

**语法**

**express.ScaleWithDocHeaderFooter**

\*express是一个代表 PageSetup 对象的变量。

### PageSetup.TopMargin

以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置上边距的大小。 **Double** 类型，可读写。

**语法**

**express.TopMargin**

\*express是一个代表 PageSetup 对象的变量。

**说明**

边距的设置和返回均以磅为单位。可使用 **InchesToPoints** 方法进行英寸到磅值的转换，也可使用 **CentimetersToPoints** 方法进行厘米到磅值的转换。

**示例**

``` JavaScript
function test(){
  /*将 Sheet1 的上边距设置为 0.5 英寸（36 磅）*/
  Application.Worksheets.Item("Sheet1").PageSetup.TopMargin = Application.InchesToPoints(0.5)
  Application.Worksheets.Item("Sheet1").PageSetup.TopMargin = 36
  
  /*本示例显示上边距的当前设置值。*/
  let marginInches = Application.ActiveSheet.PageSetup.TopMargin /
  Application.InchesToPoints(1)
  alert("The current top margin is " + marginInches + " inches")
}
```

### PageSetup.Zoom

返回或设置一个 **Variant** 值，它代表一个数值在 10% 到 400% 之间的百分比，该百分比为 ET 打印工作表时的缩放比例。

**语法**

**express.Zoom**

\*express是一个代表 PageSetup 对象的变量。

**说明**

此属性仅适用于工作表。

如果此属性设置为 **False** ，则 **FitToPagesWide** 和 **FitToPagesTall** 属性控制工作表缩放的方式。

所有缩放均保持原文档的长宽比例。

**示例**

``` JavaScript
/*此示例设置打印 Sheet1 时的缩放比例为 150%。*/
Application.Worksheets.Item("Sheet1").PageSetup.Zoom = 150
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PageSetup](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Parameter 对象

## 简介

代表一个在参数查询中使用的参数。

## 说明

Parameter 对象是 Parameters 集合的成员。

## 方法

| 序号 | 名称                            | 说明                   |
|:----:|:--------------------------------|:-----------------------|
|  1   | [SetParam](#Parameter.SetParam) | 定义指定查询表的参数。 |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Parameter.Application)         | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#Parameter.Creator)                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [DataType](#Parameter.DataType)               | 返回或设置一个 XlParameterDataType 值，它代表指定查询参数的数据类型。                                                                                                                                                           |
|  4   | [Name](#Parameter.Name)                       | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                    |
|  5   | [Parent](#Parameter.Parent)                   | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [PromptString](#Parameter.PromptString)       | 返回提示用户在参数查询输入参数值的短语。 String 类型，只读。                                                                                                                                                                    |
|  7   | [RefreshOnChange](#Parameter.RefreshOnChange) | 如果每次更改参数查询的参数值时都要刷新指定的查询表，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                 |
|  8   | [SourceRange](#Parameter.SourceRange)         | 返回一个 Range 对象，该对象表示包含指定查询参数值的单元格。只读。                                                                                                                                                               |
|  9   | [Type](#Parameter.Type)                       | 返回一个代表参数类型的 XlParameterType 值。                                                                                                                                                                                     |
|  10  | [Value](#Parameter.Value)                     | 返回一个代表参数值的 Variant 值。                                                                                                                                                                                               |

## 成员方法

### Parameter.SetParam

定义指定查询表的参数。

**语法**

**express.SetParam ( *Type* , *Value* )**

\*express是一个代表 Parameter 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型        | 说明                                      |
|---------|-----------|-----------------|-------------------------------------------|
| *Type*  | 必选      | XlParameterType | 指定参数类型的 XlParameterType 常量之一。 |
| *Value* | 必选      | Variant         | 指定参数的值，如 Type 参数的描述所示。    |

**说明**

| XlParameterType 可为以下 XlParameterType 常量之一。                          |
|------------------------------------------------------------------------------|
| xlConstant。使用 Value 参数指定的值。                                        |
| xlPrompt。显示提示用户输入值的对话框。Value 参数指定的是对话框中显示的文字。 |
| xlRange。使用区域左上角单元格的值。Value 参数指定的是一个 Range 对象。       |

**示例**

``` JavaScript
/*本示例更改第一张查询表的 SQL 语句。子句“(city=?)”表示此查询为参数查询，本示例将城市值设置为常量“Oakland”。*/function test(){
　　　　Set qt = Sheets.Item("sheet1").QueryTables.Item(1)
　　　　qt.Sql = "SELECT * FROM authors  WHERE (city=?)"
　　　　let param1 = qt.Parameters.Add("City Parameter",xlParamTypeVarChar)
　　　　param1.SetParam(xlConstant, "Oakland")
　　　　qt.Refresh()}
```

## 成员属性

### Parameter.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Parameter 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
　　　　if(ActiveWorkbook.Application.Value == "ET") {
   　　　　 MsgBox("This is an ET Application object.")
　　　　}
　　　　else {
　　　　    MsgBox("This is not an ET Application object.")
　　　　}
}
```

### Parameter.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Parameter 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Parameter.DataType

返回或设置一个 **XlParameterDataType** 值，它代表指定查询参数的数据类型。

**语法**

**express.DataType**

\*express是一个代表 Parameter 对象的变量。

### Parameter.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Parameter 对象的变量。

### Parameter.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Parameter 对象的变量。

### Parameter.PromptString

返回提示用户在参数查询输入参数值的短语。 **String** 类型，只读。

**语法**

**express.PromptString**

\*express是一个代表 Parameter 对象的变量。

**示例**

``` JavaScript
/*本示例修改查询表一的参数提示字符串。*/
function test(){
　　　　let pa = Worksheets.Item(1).QueryTables.Item(1).Parameters.Item(1)
　　　　pa.SetParam(xlPrompt, "Please " + pa.PromptString)
}
```

### Parameter.RefreshOnChange

如果每次更改参数查询的参数值时都要刷新指定的查询表，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.RefreshOnChange**

\*express是一个代表 Parameter 对象的变量。

**说明**

只有当使用 **True** xlRange 类型的参数且引用的参数值位于单个单元格中时，才能将此属性设置为 **xlRange** True。每次更改该单元格的值时，都会刷新。

**示例**

``` JavaScript
/*本示例更改 Sheet1 上第一张查询表的 SQL 语句。语句“(ContactTitle=?)”表明该查询是参数查询，标题的值设置为单元格 D4 的值。每次该单元格中的值发生变化时，查询表将会自动刷新。*/
function test(){
　　　　let objQT = Worksheets.Item("Sheet1").QueryTables.Item(1)
　　　　objQT.CommandText = "Select * From Customers Where (ContactTitle=?)"
　　　　let objParam1 = objQT.Parameters.Add("Contact Title", xlParamTypeVarChar)
　　　　objParam1.RefreshOnChange = true
　　　　objParam1.SetParam(xlRange, Range("D4"))
}
```

### Parameter.SourceRange

返回一个 **Range** 对象，该对象表示包含指定查询参数值的单元格。只读。

**语法**

**express.SourceRange**

\*express是一个代表 Parameter 对象的变量。

**示例**

``` JavaScript
/*本示例对作为查询源区域的单元格的值进行更改。*/
function test(){
　　　　let qt = Sheets.Item("sheet1").QueryTables.Item(1)
　　　　let param1 = qt.Parameters.Item(1)
　　　　let r = param1.SourceRange
　　　　r.Value = "New York"
　　　　qt.Refresh()
}
```

### Parameter.Type

返回一个代表参数类型的 **XlParameterType** 值。

**语法**

**express.Type**

\*express是一个代表 Parameter 对象的变量。

### Parameter.Value

返回一个代表参数值的 **Variant** 值。

**语法**

**express.Value**

\*express是一个代表 Parameter 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Parameter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotAxis 对象

## 简介

PivotAxis 对象用于在数据透视表中进行不对称深化。

## 说明

PivotAxis 对象包含诸如 PivotRowAxis 和 PivotRowAxis 之类的属性，用于处理数据透视表中的行和列。

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                               |
|:----:|:--------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#PivotAxis.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#PivotAxis.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Parent](#PivotAxis.Parent)           | 返回指定 PivotAxis 对象的父对象。只读。                                                                                                                            |
|  4   | [PivotLines](#PivotAxis.PivotLines)   | 返回附加到指定 PivotAxis 对象的 PivotLines。只读。                                                                                                                 |

## 成员属性

### PivotAxis.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotAxis 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### PivotAxis.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotAxis 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotAxis.Parent

返回指定 **PivotAxis** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotAxis 对象的变量。

### PivotAxis.PivotLines

返回附加到指定 **PivotAxis** 对象的 PivotLines。只读。

**语法**

**express.PivotLines**

\*express是一个代表 PivotAxis 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotAxis](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# QueryTables 对象

## 简介

QueryTable 对象的集合。

## 说明

每一个 QueryTable 对象都代表一个利用从外部数据源返回的数据生成的工作表表格。

## 方法

| 序号 | 名称                      | 说明                   |
|:----:|:--------------------------|:-----------------------|
|  1   | [Add](#QueryTables.Add)   | 新建一个查询表。       |
|  2   | [Item](#QueryTables.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#QueryTables.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#QueryTables.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#QueryTables.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#QueryTables.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### QueryTables.Add

新建一个查询表。

**语法**

**express.Add ( *Connection* , *Destination* , *Sql* )**

\*express是一个代表 QueryTables 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
|---------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Connection*  | 必选      | Variant  | 查询表的数据源。可为以下情形之一： 一个包含 OLE DB 或 ODBC 连接字符串的字符串。ODBC 连接字符串的格式为“ODBC;\<连接字符串\>”。 一个 QueryTable 对象，该对象表示查询信息的初始复制源，包括连接字符串和 SQL 文本，但不包括 Destination 区域。指定 QueryTable 对象会导致 Sql 参数被忽略。 一个 ADO 或 DAO Recordset 对象。从 ADO 或 DAO 记录集中读取数据。ET 会保留该记录集，直到该查询表被删除或 SQL 连接发生更改。不能对产生的查询表进行编辑。 一个 Web 查询，“URL; ”格式的字符串，其中“URL;”是必需的，但不进行本地化，字符串的其余部分作为 Web 查询的 URL。 数据查找程序。“FINDER;\<数据查找程序文件路径\>”格式的字符串，其中“FINDER;”是必需的，但不能本地化。字符串的其余部分为数据查找程序文件（\*.dqy 或 \*.iqy）的路径和名称。使用 Add 方法时将读取该文件；之后，对查询表的 Connection 属性的调用将相应地返回以“ODBC”或“URL”开头的字符串。 一个文本文件。“TEXT;\<文本文件路径和名称\>”格式的字符串，其中 TEXT 是必需的，但不能本地化。 |
| *Destination* | 必选      | Range    | 查询表目标区域（生成的查询表的放置区域）左上角的单元格。目标区域必须位于 expression 指定的 QueryTables 对象所在的工作表中。                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| *Sql*         | 可选      | Variant  | 在 ODBC 数据源上运行的 SQL 查询字符串。当使用的数据源为 ODBC 数据源时，该参数可选（如果不在此处指定该参数，则应该在查询表刷新之前使用查询表的 Sql 属性进行设置）。当将 QueryTable 对象、文本文件、ADO 或 DAO Recordset 对象指定为数据源时，不能使用该参数。                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |

**返回值**

一个代表新查询表的 QueryTable 对象。

**说明**

由本方法创建的查询将在调用 **Refresh** 方法后运行。

**示例**

``` JavaScript
/*此示例基于 ADO 记录集创建一个查询表。为了向后兼容，此示例保留了现有的列排序和筛选设置以及布局信息。*/
function test（）{
　　　　Dim cnnConnect As ADODB.Connection
　　　　Dim rstRecordset As ADODB.Recordset

　　　　Set cnnConnect = New ADODB.Connection
　　　　cnnConnect.Open "Provider=SQLOLEDB;" & _
　　　　    "Data Source=srvdata;" & _
　　　　    "User ID=testac;Password=4me2no;"

　　　　Set rstRecordset = New ADODB.Recordset
　　　　rstRecordset.Open _
　　　　    Source:="Select Name, Quantity, Price From Products", _
　　　　    ActiveConnection:=cnnConnect, _
 　　　　   CursorType:=adOpenDynamic, _
  　　　　  LockType:=adLockReadOnly, _
 　　　　   Options:=adCmdText

　　　　With ActiveSheet.QueryTables.Add( _
      　　　　  Connection:=rstRecordset, _
      　　　　  Destination:=Range("A1"))
　　　　    .Name = "Contact List"
   　　　　 .FieldNames = True
   　　　　 .RowNumbers = False
   　　　　 .FillAdjacentFormulas = False
    　　　　.PreserveFormatting = True
   　　　　 .RefreshOnFileOpen = False
   　　　　 .BackgroundQuery = True
   　　　　 .RefreshStyle = xlInsertDeleteCells
   　　　　 .SavePassword = True
   　　　　 .SaveData = True
  　　　　  .AdjustColumnWidth = True
   　　　　 .RefreshPeriod = 0
  　　　　  .PreserveColumnInfo = True
   　　　　 .Refresh BackgroundQuery:=False
　　　　End With
}
```

``` JavaScript
/*此示例向新的查询表中导入固定宽度的文本文件。该文本文件的第一列为 5 个字符宽，作为文本导入。第二列为 4 个字符宽，被跳过。其余部分则导入第三列中，并对其应用常规格式。*/
function test(){
　　　　let shFirstQtr = Workbooks.Item(1).Worksheets.Item(1)
　　　　let qtQtrResults = shFirstQtr.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",shFirstQtr.Cells.Item(1,1))
   　　　　 qtQtrResults.TextFileParsingType = xlFixedWidth
   　　　　 qtQtrResults.TextFileFixedColumnWidths = [5,4]
   　　　　 qtQtrResults.TextFileColumnDataTypes = [xlTextFormat, xlSkipColumn, xlGeneralFormat]
  　　　　  qtQtrResults.Refresh()
}
```

``` JavaScript
/*此示例在活动工作表上新建查询表。*/
function test(){
　　　　let sqlstring = "select 96Sales.totals from 96Sales where profit < 5"
　　　　let connstring = "ODBC;DSN=96SalesData;UID=Rep21;PWD=NUyHwYQI;Database=96Sales"
　　　　let qtadd = ActiveSheet.QueryTables.Add(connstring,Range("B1"),sqlstring)
 　　　　   qtadd.Refresh()
}
```

### QueryTables.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 QueryTables 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 QueryTable 对象。

**示例**

本示例对查询表进行设置，以便查询表右侧的公式在每次刷新时都可以进行自动更新。

``` JavaScript
Sheets.Item("sheet1").QueryTables.Item(1).FillAdjacentFormulas = true
```

## 成员属性

### QueryTables.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 QueryTables 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
　　　　let myObject = ActiveWorkbook
    if(myObject.Application.Value == "ET") {
　　　　    MsgBox("This is an ET Application object.")
　　　　}else {
　　　　    MsgBox("This is not an ET Application object.")
　　　　}
}
```

### QueryTables.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 QueryTables 对象的变量。

### QueryTables.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 QueryTables 对象的变量。

**说明**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

### QueryTables.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 QueryTables 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/QueryTables](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Scenario 对象

## 简介

代表工作表上的一个方案。

## 说明

方案是一组经过命名和保存的输入值（被称为 *changing cells* ）。 Scenario 对象是 Scenarios 集合的成员。 Scenarios 集合包含某个工作表的所有定义方案。

## 方法

| 序号 | 名称                                       | 说明                                                                     |
|:----:|:-------------------------------------------|:-------------------------------------------------------------------------|
|  1   | [ChangeScenario](#Scenario.ChangeScenario) | 更改方案以得到一组新的可变单元格及方案值（可选）。                       |
|  2   | [Delete](#Scenario.Delete)                 | 删除对象。                                                               |
|  3   | [Show](#Scenario.Show)                     | 通过在工作表中插入方案的值来显示方案。受影响的单元格为方案的更改单元格。 |

## 属性

| 序号 | 名称                                     | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Scenario.Application)     | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [ChangingCells](#Scenario.ChangingCells) | 返回一个 Range 对象，该对象表示方案的可变单元格。只读。                                                                                                                                                                         |
|  3   | [Comment](#Scenario.Comment)             | 返回或设置一个 String 值，它代表与方案相关联的批注。                                                                                                                                                                            |
|  4   | [Creator](#Scenario.Creator)             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Hidden](#Scenario.Hidden)               | 返回或设置一个 Boolean 值，它指明方案是否已被隐藏。                                                                                                                                                                             |
|  6   | [Index](#Scenario.Index)                 | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                    |
|  7   | [Locked](#Scenario.Locked)               | 返回或设置一个 Boolean 值，它指明对象是否已被锁定。                                                                                                                                                                             |
|  8   | [Name](#Scenario.Name)                   | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                    |
|  9   | [Parent](#Scenario.Parent)               | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  10  | [Values](#Scenario.Values)               | 返回一个 Variant 数组，它包含方案的可变单元格的当前值。                                                                                                                                                                         |

## 成员方法

### Scenario.ChangeScenario

更改方案以得到一组新的可变单元格及方案值（可选）。

**语法**

**express.ChangeScenario ( *ChangingCells* , *Values* )**

\*express是一个代表 Scenario 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                   |
|-----------------|-----------|----------|----------------------------------------------------------------------------------------|
| *ChangingCells* | 必选      | Variant  | 指定方案中新的可变单元格集合的 Range 对象。这些可变单元格必须与方案位于同一工作表中。  |
| *Values*        | 可选      | Variant  | 包含可变单元格的新方案值的数组。如果省略此参数，将假定方案值为这些可变单元格的当前值。 |

**返回值**

Variant

**说明**

如果指定 *Values* 参数，则数组中每一元素必须对应于 *ChangingCells* 区域中的相应单元格；否则，ET 就会产生错误。

**示例**

本示例将方案一的可变单元格设为工作表 Sheet1 上的区域 A1:A10。

``` JavaScript
Worksheets.Item("Sheet1").Scenarios(1).ChangeScenario(Worksheets.Item("Sheet1").Range("A1:A10"))
```

### Scenario.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Scenario 对象的变量。

**返回值**

Variant

### Scenario.Show

通过在工作表中插入方案的值来显示方案。受影响的单元格为方案的更改单元格。

**语法**

**express.Show ()**

\*express是一个代表 Scenario 对象的变量。

**返回值**

Variant

## 成员属性

### Scenario.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Scenario 对象的变量。

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

### Scenario.ChangingCells

返回一个 **Range** 对象，该对象表示方案的可变单元格。只读。

**语法**

**express.ChangingCells**

\*express是一个代表 Scenario 对象的变量。

**示例**

``` JavaScript
/*本示例选定 Sheet1 中方案一的可变单元格。*/
function test(){
　　　　Worksheets.Item("Sheet1").Activate()
　　　　ActiveSheet.Scenarios(1).ChangingCells.Select()
}
```

### Scenario.Comment

返回或设置一个 **String** 值，它代表与方案相关联的批注。

**语法**

**express.Comment**

\*express是一个代表 Scenario 对象的变量。

**说明**

批注文本不能超过 255 个字符。

**示例**

此示例设置 Sheet1 中第一个方案的批注。

``` JavaScript
Worksheets.Item("Sheet1").Scenarios(1).Comment ="Worst case July 1993 sales"
```

### Scenario.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Scenario 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Scenario.Hidden

返回或设置一个 **Boolean** 值，它指明方案是否已被隐藏。

**语法**

**express.Hidden**

\*express是一个代表 Scenario 对象的变量。

**说明**

此属性的默认设置为 **False** 。

请不要将此属性与 **FormulaHidden** 属性混淆。

### Scenario.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 Scenario 对象的变量。

### Scenario.Locked

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

\*express是一个代表 Scenario 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True** ；如果在工作表处于受保护状态时仍能修改对象，则返回 **False** 。

### Scenario.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Scenario 对象的变量。

### Scenario.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Scenario 对象的变量。

### Scenario.Values

返回一个 **Variant** 数组，它包含方案的可变单元格的当前值。

**语法**

**express.Values**

\*express是一个代表 Scenario 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Scenario](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Styles 对象

## 简介

指定的数据透视表中所有 Style

## 说明

每一个 Style 对象都代表对某区域的样式描述。 Style 对象包含样式的所有属性（字体、数字格式、对齐方式，等等）。有几种内置的样式，包括“常规”、“货币”和“百分比”。

使用 Styles 属性可返回 Styles 集合。下例创建活动工作簿中工作表一上的样式名的列表。

``` JavaScript
function test(){
    for(let i = 1;i <= Application.ActiveWorkbook.Styles.Count; i++){
        Application.Worksheets.Item(1).Cells.Item(i, 1).Value2 = ActiveWorkbook.Styles.Item(i).Name
    }
}
```

使用 Add 方法可创建一个新的样式并将它添加到集合。下例基于“常规”样式创建一个新的样式，修改边框和字体，然后将该新样式应用到单元格 A25:A30。

``` JavaScript
function test(){
    let rng = Application.ActiveWorkbook.Styles.Add("Bookman Top Border")
    rng.Borders.Item(xlTop).LineStyle = xlDouble
    rng.Font.Bold = true
    rng.Font.Name = "Bookman"
    Application.Worksheets.Item(1).Range("A25:A30").Style = "Bookman Top Border"
}
```

使用 Styles ( *index* )（其中 *index* 是样式索引号或名称）可从工作簿的 Style 集合中返回一个 Styles 对象。下例通过设置活动工作簿中“常规”样式的 Bold 属性来更改该样式。

``` JavaScript
Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true
```

## 方法

| 序号 | 名称                   | 说明                                                                               |
|:----:|:-----------------------|:-----------------------------------------------------------------------------------|
|  1   | [Add](#Styles.Add)     | 新建样式并将其添加到当前工作簿的可用样式列表中。返回 一个代表新样式的 Style 对象。 |
|  2   | [Item](#Styles.Item)   | 从集合中返回一个对象。                                                             |
|  3   | [Merge](#Styles.Merge) | 将另一张工作簿中的样式合并到 Styles 集合中。返回 Variant值                         |

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Styles.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Styles.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Styles.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Styles.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Styles.Add

新建样式并将其添加到当前工作簿的可用样式列表中。返回 一个代表新样式的 **Style** 对象。

**语法**

**express.Add ( *Name* )**

\*express是一个代表 Styles 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Name* | 必选      | String   | 新样式的名称。 |

**示例**

``` JavaScript
/*本示例基于工作表 Sheet1 中的单元格 A1 定义新样式。*/
function test(){
  let Sty = Application.ActiveWorkbook.Styles.Add("theNewStyle")
  Sty.IncludeNumber = false
  Sty.IncludeFont = true
  Sty.IncludeAlignment = false
  Sty.IncludeBorder = false
  Sty.IncludePatterns = false
  Sty.IncludeProtection = false
  Sty.Font.Name = "Arial"
  Sty.Font.Size = 18
}
```

### Styles.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Styles 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
/*此示例通过设置“常规”样式的 Bold 属性来更改活动工作簿中的该样式。*/
Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true
```

### Styles.Merge

将另一张工作簿中的样式合并到 **Styles** 集合中。返回 Variant值

**语法**

**express.Merge ( *Workbook* )**

\*express是一个代表 Styles 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                               |
|------------|-----------|----------|----------------------------------------------------|
| *Workbook* | 必选      | Variant  | 一个 Workbook 对象，它代表包含待合并样式的工作簿。 |

**示例**

``` JavaScript
/*此示例将工作簿 Template.xls 中的样式合并到活动工作簿中。*/
Application.ActiveWorkbook.Styles.Merge(Workbooks.Item("TEMPLATE.XLS"))
```

## 成员属性

### Styles.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Styles 对象的变量。

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

### Styles.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Styles 对象的变量。

### Styles.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Styles 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Styles.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Styles 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Styles](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TextEffectFormat 对象

## 简介

包含应用于艺术字对象的属性和方法。

## 说明

使用 TextEffect 属性可返回一个 TextEffectFormat 对象。

下例为 *myDocument* 上的形状一设置字体名称及格式。要运行本示例，形状一必须是艺术字对象。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let TextEffect = myDocument.Shapes(1).TextEffect
    TextEffect.FontName = "Courier New"
    TextEffect.FontBold = true
    TextEffect.FontItalic = true
}
```

## 方法

| 序号 | 名称                                                       | 说明                                                     |
|:----:|:-----------------------------------------------------------|:---------------------------------------------------------|
|  1   | [ToggleVerticalText](#TextEffectFormat.ToggleVerticalText) | 将指定艺术字中的文字排列方向从水平切换为垂直，反之亦然。 |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Alignment](#TextEffectFormat.Alignment)               | 返回或设置一个 MsoTextEffectAlignment 值，它代表艺术字的对齐方式。                                                                                                                                                              |
|  2   | [Application](#TextEffectFormat.Application)           | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  3   | [Creator](#TextEffectFormat.Creator)                   | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [FontBold](#TextEffectFormat.FontBold)                 | 如果指定艺术字中的字体格式为加粗，则该属性值为 True 。 MsoTriState 类型，可读写。                                                                                                                                               |
|  5   | [FontItalic](#TextEffectFormat.FontItalic)             | 如果指定艺术字中的字体是倾斜的，则返回 msoTrue 。 MsoTriState 类型，可读写。                                                                                                                                                    |
|  6   | [FontName](#TextEffectFormat.FontName)                 | 返回或设置指定艺术字的字体名称。 String 类型，可读写。                                                                                                                                                                          |
|  7   | [FontSize](#TextEffectFormat.FontSize)                 | 以磅为单位返回或设置指定艺术字的字体大小。 Single 类型，可读写。                                                                                                                                                                |
|  8   | [KernedPairs](#TextEffectFormat.KernedPairs)           | 如果指定艺术字中的字符对是自动缩紧的，则该属性值为 True 。 MsoTriState 类型，可读写。                                                                                                                                           |
|  9   | [NormalizedHeight](#TextEffectFormat.NormalizedHeight) | 如果指定艺术字中所有字符（无论大小写）均为同一高度，则该属性值为 True 。 MsoTriState 类型，可读写。                                                                                                                             |
|  10  | [Parent](#TextEffectFormat.Parent)                     | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  11  | [PresetShape](#TextEffectFormat.PresetShape)           | 返回或设置指定艺术字的形状样式。 MsoPresetTextEffectShape 类型，可读写。                                                                                                                                                        |
|  12  | [PresetTextEffect](#TextEffectFormat.PresetTextEffect) | 返回或设置指定艺术字的样式。 MsoPresetTextEffect 类型，可读写。                                                                                                                                                                 |
|  13  | [RotatedChars](#TextEffectFormat.RotatedChars)         | 如果指定艺术字对象中的字符相对于该对象旋转了 90 度，则该值为 True 。如果指定艺术字对象中的字符相对于该对象保持原有方向，则该值为 False 。 MsoTriState 类型，可读/写。                                                           |
|  14  | [Text](#TextEffectFormat.Text)                         | 返回或设置指定对象中的文本。 String 型，可读写。                                                                                                                                                                                |
|  15  | [Tracking](#TextEffectFormat.Tracking)                 | 返回或设置指定艺术字中每个字符的水平间隔相对于字符宽度的比例。该比例可为从 0 到 5 之间的值。（为该属性指定的值越大，字符间的空间越大；小于 1 的值会产生字符重叠）。可读写。 Single 类型。                                       |

## 成员方法

### TextEffectFormat.ToggleVerticalText

将指定艺术字中的文字排列方向从水平切换为垂直，反之亦然。

**语法**

**express.ToggleVerticalText ()**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

使用 **ToggleVerticalText** 方法可以交换代表艺术字的 **Shape** 对象的 **Width** 和 **Height** 属性的值，但不改变 **Left** 和 **Top** 属性。

**Shape** 对象的 **Flip** 方法和 **Rotation** 属性以及 **TextEffectFormat** 对象的 **RotatedChars** 属性和 **ToggleVerticalText** 方法都会影响代表艺术字的 **Shape** 对象中字符方向和文本排列方式。可能必须经过试验才能找到组合使用这些属性和方法的最佳途径以达到预期效果。

**示例**

本示例向 ` myDocument ` 中添加了包含文本“Test”的艺术字，并将其水平文字排列方向（对于指定的艺术字样式 **msoTextEffect1** ，这是默认的文字排列方向）更改为垂直文字排列方向。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let newWordArt = myDocument.Shapes.AddTextEffect(Application.Enum.msoTextEffect1, "Test", "Arial Black", 36, false, false, 100, 100)
    newWordArt.TextEffect.ToggleVerticalText()
}
```

## 成员属性

### TextEffectFormat.Alignment

返回或设置一个 **MsoTextEffectAlignment** 值，它代表艺术字的对齐方式。

**语法**

**express.Alignment**

\*express是一个代表 TextEffectFormat 对象的变量。

**示例**

此示例向第一张工作表中添加艺术字对象，然后将该对象右对齐

``` JavaScript
function test() {
    let mySh = Application.Worksheets.Item(1).Shapes
    let myTE = mySh.AddTextEffect(Application.Enum.msoTextEffect1, "Test Text", "Palatino", 54, true, false, 100, 50)
    myTE.TextEffect.Alignment = Application.Enum.msoTextEffectAlignmentRight
}
```

### TextEffectFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TextEffectFormat 对象的变量。

**示例**

本示例显示一条有关创建 myObject 的应用程序的消息。

``` JavaScript
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### TextEffectFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### TextEffectFormat.FontBold

如果指定艺术字中的字体格式为加粗，则该属性值为 **True** 。 **MsoTriState** 类型，可读写。

**语法**

**express.FontBold**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

|                                                       |
|-------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue**                                          |
| **msoFalse**                                          |
| **msoTriStateMixed**                                  |
| **msoTriStateToggle**                                 |
| **msoTrue** 指定的艺术字为粗体。                      |

**示例**

本示例将 myDocument 中第三个形状（如果该形状为艺术字）的字体设置为粗体。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let rng = myDocument.Shapes.Item(3)
    if (rng.Type == Application.Enum.msoTextEffect) {
        rng.TextEffect.FontBold = Application.Enum.msoTrue
    }
}
```

### TextEffectFormat.FontItalic

如果指定艺术字中的字体是倾斜的，则返回 **msoTrue** 。 **MsoTriState** 类型，可读写。

**语法**

**express.FontItalic**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

|                                                       |
|-------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue** 不应用于该属性。                         |
| **msoFalse** 指定的艺术字对象中的文字不是倾斜的。     |
| **msoTriStateMixed** 不应用于该属性。                 |
| **msoTriStateToggle** 不应用于该属性。                |
| **msoTrue** 指定的艺术字对象中的文字是倾斜的。        |

**示例**

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.Item("WordArt 4").TextEffect.FontItalic = Application.Enum.msoTrue
}
```

本示例将 ` myDocument ` 中形状“WordArt 4”的字体设置为倾斜

### TextEffectFormat.FontName

返回或设置指定艺术字的字体名称。 **String** 类型，可读写。

**语法**

**express.FontName**

\*express是一个代表 TextEffectFormat 对象的变量。

**示例**

本示例将 myDocument 中第三个形状（如果该形状是艺术字）字体名称设置为“Courier New”。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let rng = myDocument.Shapes.Item(3)
    if (rng.Type == Application.Enum.msoTextEffect) {        
        rng.TextEffect.FontName = "Courier New"
    }
}
```

### TextEffectFormat.FontSize

以磅为单位返回或设置指定艺术字的字体大小。 **Single** 类型，可读写。

**语法**

**express.FontSize**

\*express是一个代表 TextEffectFormat 对象的变量。

**示例**

本示例将 ` myDocument ` 中名为“WordArt 4”的形状中的字号设置为 16 磅。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.Item("WordArt 4").TextEffect.FontSize = 16
}
```

### TextEffectFormat.KernedPairs

如果指定艺术字中的字符对是自动缩紧的，则该属性值为 **True** 。 **MsoTriState** 类型，可读写。

**语法**

**express.KernedPairs**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

|                                                       |
|-------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue**                                          |
| **msoFalse**                                          |
| **msoTriStateMixed**                                  |
| **msoTriStateToggle**                                 |
| **msoTrue** 指定艺术字中的字符对自动紧缩。            |

**示例**

本示例打开 ` myDocument ` 中第三个形状（如果该形状为艺术字）的字符对自动紧缩。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let rng = myDocument.Shapes.Item(3)
    if (rng.Type == Application.Enum.msoTextEffect) {
        rng.TextEffect.KernedPairs = Application.Enum.msoTrue
    }
}
```

### TextEffectFormat.NormalizedHeight

如果指定艺术字中所有字符（无论大小写）均为同一高度，则该属性值为 **True** 。 **MsoTriState** 类型，可读写。

**语法**

**express.NormalizedHeight**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

|                                                                |
|----------------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。          |
| **msoCTrue**                                                   |
| **msoFalse**                                                   |
| **msoTriStateMixed**                                           |
| **msoTriStateToggle**                                          |
| **msoTrue** 指定艺术字中的所有字符（无论大小写）均为同一高度。 |

**示例**

本示例将包含文字“Test Effect”的艺术字添加至 myDocument ，并将新艺术字命名为“texteff1”。然后，此代码将名为“texteff1”的形状中的所有字符设置为同一高度。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.AddTextEffect(Application.Enum.msoTextEffect1, "Test Effect", "Courier New", 44, true, false, 10, 10).Name = "texteff1"
    myDocument.Shapes.Item("texteff1").TextEffect.NormalizedHeight = Application.Enum.msoTrue
}
```

### TextEffectFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 TextEffectFormat 对象的变量。

### TextEffectFormat.PresetShape

返回或设置指定艺术字的形状样式。 **MsoPresetTextEffectShape** 类型，可读写。

**语法**

**express.PresetShape**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

|                                                                                 |
|---------------------------------------------------------------------------------|
| **MsoPresetTextEffectShape** 可以是下列 **MsoPresetTextEffectShape** 常量之一。 |
| **msoTextEffectShapeArchDownCurve**                                             |
| **msoTextEffectShapeArchDownPour**                                              |
| **msoTextEffectShapeArchUpCurve**                                               |
| **msoTextEffectShapeArchUpPour**                                                |
| **msoTextEffectShapeButtonCurve**                                               |
| **msoTextEffectShapeButtonPour**                                                |
| **msoTextEffectShapeCanDown**                                                   |
| **msoTextEffectShapeCanUp**                                                     |
| **msoTextEffectShapeCascadeDown**                                               |
| **msoTextEffectShapeCascadeUp**                                                 |
| **msoTextEffectShapeChevronDown**                                               |
| **msoTextEffectShapeChevronUp**                                                 |
| **msoTextEffectShapeCircleCurve**                                               |
| **msoTextEffectShapeCirclePour**                                                |
| **msoTextEffectShapeCurveDown**                                                 |
| **msoTextEffectShapeCurveUp**                                                   |
| **msoTextEffectShapeDeflate**                                                   |
| **msoTextEffectShapeDeflateBottom**                                             |
| **msoTextEffectShapeDeflateInflateDeflate**                                     |
| **msoTextEffectShapeDoubleWave1**                                               |
| **msoTextEffectShapeFadeDown**                                                  |
| **msoTextEffectShapeFadeRight**                                                 |
| **msoTextEffectShapeInflate**                                                   |
| **msoTextEffectShapeInflateTop**                                                |
| **msoTextEffectShapePlainText**                                                 |
| **msoTextEffectShapeRingOutside**                                               |
| **msoTextEffectShapeSlantUp**                                                   |
| **msoTextEffectShapeTriangleDown**                                              |
| **msoTextEffectShapeWave1**                                                     |
| **msoTextEffectShapeDeflateInflate**                                            |
| **msoTextEffectShapeDeflateTop**                                                |
| **msoTextEffectShapeDoubleWave2**                                               |
| **msoTextEffectShapeFadeLeft**                                                  |
| **msoTextEffectShapeFadeUp**                                                    |
| **msoTextEffectShapeInflateBottom**                                             |
| **msoTextEffectShapeMixed**                                                     |
| **msoTextEffectShapeRingInside**                                                |
| **msoTextEffectShapeSlantDown**                                                 |
| **msoTextEffectShapeStop**                                                      |
| **msoTextEffectShapeTriangleUp**                                                |
| **msoTextEffectShapeWave2**                                                     |

设置 **PresetTextEffect** 属性会自动设置 **PresetShape** 属性。

**示例**

本示例将 ` myDocument ` 中的所有艺术字形状设置为中心指向下方的 V 字形。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    for (let s = 1; s <= myDocument.Shapes.Count; s++) {
        if (myDocument.Shapes.Item(s).Type == Application.Enum.msoTextEffect) {
            myDocument.Shapes.Item(s).TextEffect.PresetShape = Application.Enum.msoTextEffectShapeChevronDown
        }
    }
}
```

### TextEffectFormat.PresetTextEffect

返回或设置指定艺术字的样式。 **MsoPresetTextEffect** 类型，可读写。

**语法**

**express.PresetTextEffect**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

该属性的值与 **“艺术字库”** 对话框内的格式（按从左至右和从上至下的顺序编号）相对应。

|                                                                       |
|-----------------------------------------------------------------------|
| **MsoPresetTextEffect** 可以是下列 **MsoPresetTextEffect** 常量之一。 |
| **msoTextEffect1**                                                    |
| **msoTextEffect10**                                                   |
| **msoTextEffect11**                                                   |
| **msoTextEffect12**                                                   |
| **msoTextEffect13**                                                   |
| **msoTextEffect14**                                                   |
| **msoTextEffect15**                                                   |
| **msoTextEffect16**                                                   |
| **msoTextEffect17**                                                   |
| **msoTextEffect18**                                                   |
| **msoTextEffect19**                                                   |
| **msoTextEffect2**                                                    |
| **msoTextEffect20**                                                   |
| **msoTextEffect21**                                                   |
| **msoTextEffect22**                                                   |
| **msoTextEffect23**                                                   |
| **msoTextEffect24**                                                   |
| **msoTextEffect25**                                                   |
| **msoTextEffect26**                                                   |
| **msoTextEffect27**                                                   |
| **msoTextEffect28**                                                   |
| **msoTextEffect29**                                                   |
| **msoTextEffect3**                                                    |
| **msoTextEffect30**                                                   |
| **msoTextEffect4**                                                    |
| **msoTextEffect5**                                                    |
| **msoTextEffect6**                                                    |
| **msoTextEffect7**                                                    |
| **msoTextEffect8**                                                    |
| **msoTextEffect9**                                                    |
| **msoTextEffectMixed**                                                |

设置 **PresetTextEffect** 属性会自动设置指定形状的许多其他格式属性。

**示例**

本示例将 ` myDocument ` 中所有的艺术字样式设置为“艺术字库” 对话框中列出的第一种样式。

``` JavaScript
function test() {
    let myDocument =  Application.Worksheets.Item(1)
    for (let s = 1; s <= myDocument.Shapes.Count; s++) {
        if (myDocument.Shapes.Item(s).Type == Application.Enum.msoTextEffect) {
            myDocument.Shapes.Item(s).TextEffect.PresetTextEffect = Application.Enum.msoTextEffect1
        }
    }
}
```

### TextEffectFormat.RotatedChars

如果指定艺术字对象中的字符相对于该对象旋转了 90 度，则该值为 **True** 。如果指定艺术字对象中的字符相对于该对象保持原有方向，则该值为 **False** 。 **MsoTriState** 类型，可读/写。

**语法**

**express.RotatedChars**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

|                                                                  |
|------------------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。            |
| **msoCTrue**                                                     |
| **msoFalse** 指定艺术字中的字符保留原有的相对于边框形状的方向。  |
| **msoTriStateMixed**                                             |
| **msoTriStateToggle**                                            |
| **msoTrue** 将指定艺术字中的字符相对于艺术字边框形状旋转 90 度。 |

如果艺术字中含有水平文字，将 **RotatedChars** 属性设置为 **msoTrue** 可将该字符逆时针旋转 90 度。如果艺术字中含有垂直文字，将 **RotatedChars** 属性设置为 **msoFalse** 可将该字符顺时针旋转 90 度。使用 **ToggleVerticalText** 方法在水平文字流和垂直文字流之间切换。

**Shape** 对象的 **Flip** 方法和 **Rotation** 属性以及 **TextEffectFormat** 对象的 **RotatedChars** 属性和 **ToggleVerticalText** 方法会对表示艺术字的 **Shape** 对象中的字符的方向和文字排列的方向产生影响。可能必须经过试验才能找到组合使用这些属性和方法的最佳途径以达到预期效果。

**示例**

本例向 ` myDocument ` 中添加带有“Test”文本的艺术字，并将该字符逆时针旋转 90 度。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let newWordArt = myDocument.Shapes.AddTextEffect(Application.Enum.msoTextEffect1, "Test", "Arial Black", 36, false, false, 10, 10)
    newWordArt.TextEffect.RotatedChars = Application.Enum.msoTrue
}
```

### TextEffectFormat.Text

返回或设置指定对象中的文本。 **String** 型，可读写。

**语法**

**express.Text**

\*express是一个代表 TextEffectFormat 对象的变量。

### TextEffectFormat.Tracking

返回或设置指定艺术字中每个字符的水平间隔相对于字符宽度的比例。该比例可为从 0 到 5 之间的值。（为该属性指定的值越大，字符间的空间越大；小于 1 的值会产生字符重叠）。可读写。 **Single** 类型。

**语法**

**express.Tracking**

\*express是一个代表 TextEffectFormat 对象的变量。

**说明**

以下表格相对于用户界面上的可用设置给出 **Tracking** 属性的值。

| 用户界面设置 | 相应的 Tracking 属性值 |
|--------------|------------------------|
| 很紧         | 0.8                    |
| 紧密         | 0.9                    |
| 正常         | 1.0                    |
| 稀疏         | 1.2                    |
| 很松         | 1.5                    |

**示例**

本例向 ` myDocument ` 中添加包含“Test”文本的艺术字，并且指定这些字符紧密排列。

``` JavaScript
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let newWordArt = myDocument.Shapes.AddTextEffect(Application.Enum.msoTextEffect1, "Test", "Arial Black", 36, false, false, 100, 100)
    newWordArt.TextEffect.Tracking = 0.8
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/TextEffectFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# UserAccess 对象

## 简介

代表对受保护区域的用户访问。

## 说明

使用 UserAccessList 集合的 Add 方法或 Item 属性可返回一个 UserAccess 对象。

一旦返回了 UserAccess 对象，您就可以使用 AllowEdit 属性来确定是否允许访问工作表中某个特定区域。下例添加一个在受保护的工作表上可编辑的区域，并通知用户该区域的标题。

``` JavaScript
function test(){
    let wksSheet = Application.ActiveSheet
    
    //Add a range that can be edited on the protected worksheet.
    wksSheet.Protection.AllowEditRanges.Add("Test", Range("A1"))
    
    //Notify the user the title of the range that can be edited.
    MsgBox(wksSheet.Protection.AllowEditRanges.Item(1).Title)
}
```

## 方法

| 序号 | 名称                         | 说明       |
|:----:|:-----------------------------|:-----------|
|  1   | [Delete](#UserAccess.Delete) | 删除对象。 |

## 属性

| 序号 | 名称                               | 说明                                                                      |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------|
|  1   | [AllowEdit](#UserAccess.AllowEdit) | 返回或设置一个 Boolean 值，它指明是否允许用户访问受保护工作表的指定区域。 |
|  2   | [Name](#UserAccess.Name)           | 返回或设置一个 String 值，它代表对象的名称。                              |

## 成员方法

### UserAccess.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 UserAccess 对象的变量。

## 成员属性

### UserAccess.AllowEdit

返回或设置一个 **Boolean** 值，它指明是否允许用户访问受保护工作表的指定区域。

**语法**

**express.AllowEdit**

\*express是一个代表 UserAccess 对象的变量。

### UserAccess.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 UserAccess 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/UserAccess](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Bookmark 对象

## 简介

代表文档、选定内容或区域中的单个书签。 Bookmark 对象是 Bookmarks 集合的成员。 Bookmarks 集合包括 “书签” 对话框中（ “插入” 菜单）列出的所有书签。

## 说明

使用 Bookmarks ( *索引* ) （其中 *index* 是书签名称或索引号）可返回一个 Bookmark 对象。 您必须完全匹配拼写 （但不必大小写匹配） 的书签名称。 下面的示例选择名为活动文档中的"temp"的书签。

``` JavaScript
Application.ActiveDocument.Bookmarks.Item("temp").Select()
```

索引编号代表书签在选区Selection或者区域Range 对象中的位置。对 Document 对象， 索引编号代表书签在 “书签对话框”中按名称字母排序的位置。 下面的示例显示 书签 集合中的第二个书签的名称。

``` JavaScript
Application.ActiveDocument.Bookmarks.Item(1).Name
```

使用 Add 方法将书签添加到文档范围。 以下示例通过添加名为“Abc”的书签来标记选定内容。

``` JavaScript
Application.ActiveDocument.Bookmarks.Add("Abc", Application.Selection.Range)
```

可将 BookmarkID 属性与区域或选定内容对象配合使用，以返回 Bookmark 对象在 Bookmarks 集合中的索引编号。以下示例显示活动文档中名为“ pe ak ”的书签的索引编号。

``` JavaScript
Application.ActiveDocument.Bookmarks.Item("peak").Range.BookmarkID
```

若要确定所选内容、 范围或文档中是否已存在一个书签的方法。 下面的示例可确保在选定该书签之前名为"temp"的书签存在于活动文档。

``` JavaScript
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    if (bks.Exists("temp1") === true) {
        bks.Item("temp").Select();
    }
}
```

## 方法

| 序号 | 名称                       | 说明                                                               |
|:----:|:---------------------------|:-------------------------------------------------------------------|
|  1   | [Copy](#Bookmark.Copy)     | 将书签复制到 Name 参数中指定的新书签处，并返回一个 Bookmark 对象。 |
|  2   | [Delete](#Bookmark.Delete) | 删除指定的书签。                                                   |
|  3   | [Select](#Bookmark.Select) | 选择指定书签。                                                     |

## 属性

| 序号 | 名称                                 | 说明                                                                         |
|:----:|:-------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Bookmark.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                         |
|  2   | [Column](#Bookmark.Column)           | 如果指定的书签是表格的一列，则该属性值为 True 。 Boolean 类型，只读。        |
|  3   | [Creator](#Bookmark.Creator)         | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。 |
|  4   | [Empty](#Bookmark.Empty)             | 如果指定书签为空，则该属性值为 True 。只读 Boolean 类型。                    |
|  5   | [End](#Bookmark.End)                 | 返回或设置选定内容、区域或书签中结束字符的位置。 Number 类型，可读写。       |
|  6   | [Name](#Bookmark.Name)               | 返回指定对象的名称。 String 类型，只读。                                     |
|  7   | [Parent](#Bookmark.Parent)           | 返回一个 Object 类型的值，该值代表指定 Bookmark 对象的父对象。               |
|  8   | [Range](#Bookmark.Range)             | 返回一个 Range 对象，该对象代表指定对象所含的部分文档。                      |
|  9   | [Start](#Bookmark.Start)             | 返回或设置书签起始字符的位置。 Number 类型，可读写。                         |
|  10  | [StoryType](#Bookmark.StoryType)     | 返回指定范围、选定内容或书签的 文章 类型。 WdStoryType 类型，只读。          |

## 成员方法

### Bookmark.Copy

将书签复制到 **Name** 参数中指定的新书签处，并返回一个 **Bookmark** 对象。

**语法**

**express.Copy ( *Name* )**

\*express是一个代表 Bookmark 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Name* | 必选      | String   | 新书签的名称。 |

**示例**

``` JavaScript
/* 在书签temp处复制生成一个新的书签 */
Application.ActiveDocument.Bookmarks.Item("temp").Copy("test")
```

### Bookmark.Delete

删除指定的书签。

**语法**

**express.Delete ()**

\*express是一个代表 Bookmark 对象的变量。

**示例**

``` JavaScript
/* 如果书签"temp"存在，则删除掉 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    if (bks.Exists("temp") === true) {
        bks.Item("temp").Delete();
    }
}
```

### Bookmark.Select

选择指定书签。

**语法**

**express.Select ()**

\*express是一个代表 Bookmark 对象的变量。

**说明**

使用该方法后，可用 **Selection** 对象来处理选定项目。有关详细信息，请参阅 **Selection** 对象。

## 成员属性

### Bookmark.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Bookmark 对象的变量。

### Bookmark.Column

如果指定的书签是表格的一列，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.Column**

\*express是一个代表 Bookmark 对象的变量。

**示例**

``` JavaScript
/* 获取"peak"书签是否为表格的一列 */
Application.ActiveDocument.Bookmarks.Item("peak").Column
```

### Bookmark.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Bookmark 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

### Bookmark.Empty

如果指定书签为空，则该属性值为 **True** 。只读 **Boolean** 类型。

**语法**

**express.Empty**

\*express是一个代表 Bookmark 对象的变量。

**说明**

空书签可标记位置（折叠选定内容），而不标记任何文本。如果指定书签不存在，则会出错。

**示例**

``` JavaScript
/* 本示例查看名为“temp”的书签是否存在和是否为空。 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    if(bks.Exists("temp") === true) {
        if(bks.Item("temp").Empty === true) {
            alert("The Temp bookmark is empty");
        }
    }
}
```

### Bookmark.End

返回或设置选定内容、区域或书签中结束字符的位置。 **Number** 类型，可读写。

**语法**

**express.End**

\*express是一个代表 Bookmark 对象的变量。

**说明**

如果本属性设置的值小于 [**Start**](#Bookmark.Start) 属性的值，则 **Start** 属性将被设成同一值（即 **Start** 与 **End** 属性值相等）。

该属性返回结束字符相对于文档开头部分的位置。文档主要文字部分 (wdMainTextStory) 的起始字符位置为 0。通过设置该属性可以更改书签的大小。

**示例**

``` JavaScript
/* 返回书签temp的长度 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    let start = bks.Item("temp").Start;
    let end = bks.Item("temp").End;
    alert(end - start);
}
```

### Bookmark.Name

返回指定对象的名称。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 Bookmark 对象的变量。

**示例**

``` JavaScript
/* 返回索引为1的书签名字 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    let bk = bks.Item(1);
    alert(bk.Name);
}
```

### Bookmark.Parent

返回一个 **Object** 类型的值，该值代表指定 **Bookmark** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Bookmark 对象的变量。

### Bookmark.Range

返回一个 **Range** 对象，该对象代表指定对象所含的部分文档。

**语法**

**express.Range**

\*express是一个代表 Bookmark 对象的变量。

**说明**

有关从文档返回一个区域或从图形集合返回图形区域的信息，请参阅 **Range** 方法。

### Bookmark.Start

返回或设置书签起始字符的位置。 **Number** 类型，可读写。

**语法**

**express.Start**

\*express是一个代表 Bookmark 对象的变量。

**说明**

如果将该属性的值设置为大于 **[End](#Bookmark.End)** 属性值，则 **End** 属性值会设置为与 **Start** 属性值相同。

书签对象包括起始字符和结束字符位置。起始字符位置距文档开头部分最近。

该属性返回起始字符相对于文档开头部分的位置。文字部分 ( **wdMainTextStory** ) 的起始字符位置为 0。通过设置本属性可以更改书签的大小。

**示例**

``` JavaScript
/* 返回书签temp的长度 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    let start = bks.Item("temp").Start;
    let end = bks.Item("temp").End;
    alert(end - start);
}
```

### Bookmark.StoryType

返回指定范围、选定内容或书签的 文章 类型。 **WdStoryType** 类型，只读。

**语法**

**express.StoryType**

\*express是一个代表 Bookmark 对象的变量。

**示例**

``` JavaScript
/*如果活动文档的主体部分包括名为“temp”的书签，本示例将选中该书签。*/
function test()
{
if(ActiveDocument.Bookmarks.Exists("temp") == true) {
    let myBookmark = ActiveDocument.Bookmarks.Item("temp")
    if(myBookmark.StoryType == wdMainTextStory) {
        myBookmark.Select()
    }
}
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Bookmark](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Comment 对象

## 简介

代表单个备注。 Comment 对象是 Comments 集合的成员。 Comments 集合中包括选定内容、区域或文档中的备注。

## 说明

使用 [Add](#Comments.Add) 方法在指定区域添加一个备注。以下示例在紧挨着选定内容后添加一个备注。

``` JavaScript
Application.Selection.Collapse(wdCollapseEnd)
Application.ActiveDocument.Comments.Add(Application.Selection.Range, "review this")
```

使用 [Reference](#Comment.Reference) 属性返回与指定备注关联的引用标记。使用 [Range](#Comment.Range) 属性返回与指定备注关联的文本。以下示例显示与活动文档的第一个备注关联的文本。

``` JavaScript
alert(Application.ActiveDocument.Comments.Item(1).Range.Text)
```

## 方法

| 序号 | 名称                      | 说明                                                      |
|:----:|:--------------------------|:----------------------------------------------------------|
|  1   | [Delete](#Comment.Delete) | 删除指定的批注。                                          |
|  2   | [Edit](#Comment.Edit)     | 在创建指定 OLE 对象的应用程序中打开该对象，以便进行编辑。 |

## 属性

| 序号 | 名称                                | 说明                                                                                  |
|:----:|:------------------------------------|:--------------------------------------------------------------------------------------|
|  1   | [Application](#Comment.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                        |
|  2   | [Author](#Comment.Author)           | 返回或设置一个 String 类型的值，该值代表某一批注的作者姓名。可读写。                  |
|  3   | [Date](#Comment.Date)               | 返回代表批注的插入日期和时间的 Date 类型。只读。                                      |
|  4   | [Index](#Comment.Index)             | 返回一个 Long 类型的值，该值代表项目在集合中的位置。只读。                            |
|  5   | [Initial](#Comment.Initial)         | 返回或设置与指定批注相关联的用户的姓名缩写。 String 类型，可读写。                    |
|  6   | [IsInk](#Comment.IsInk)             | 返回一个 Boolean 类型的值，该值代表批注是否为手写批注。                               |
|  7   | [Parent](#Comment.Parent)           | 返回一个 Object 类型值，该值代表指定 Comment 对象的父对象。                           |
|  8   | [Range](#Comment.Range)             | 返回代表备注内容的 Range 对象。                                                       |
|  9   | [Reference](#Comment.Reference)     | 返回代表备注的引用标记的 Range 对象。                                                 |
|  10  | [Scope](#Comment.Scope)             | 返回一个 Range 对象，该对象代表指定批注所标记的文本范围。                             |
|  11  | [ShowTip](#Comment.ShowTip)         | 如果在“屏幕提示”中显示与批注相关联的文本，则该属性值为 True 。 Boolean 类型，可读写。 |

## 成员方法

### Comment.Delete

删除指定的批注。

**语法**

**express.Delete ()**

\*express是一个代表 Comment 对象的变量。

### Comment.Edit

在创建指定 OLE 对象的应用程序中打开该对象，以便进行编辑。

**语法**

**express.Edit ()**

\*express是一个代表 Comment 对象的变量。

**示例**

``` JavaScript
/*本示例打开（为了进行编辑）活动文档中第一个嵌入式 OLE 对象（定义为图形）。*/
function test() {
    let shapesAll = Application.ActiveDocument.Shapes
    if (shapesAll.Count >= 1) {
        if (shapesAll.Item(1).Type == msoEmbeddedOLEObject) {
            shapesAll.Item(1).OLEFormat.Edit()
        }
    }
}
```

``` JavaScript
/*本示例打开（为了进行编辑）活动文档中第一个链接的 OLE 对象（定义为嵌入式图形）。*/
function test() {
    let colIS = Application.ActiveDocument.InlineShapes
    if (colIS.Count >= 1) {
        if (colIS.Item(1).Type == wdInlineShapeLinkedOLEObject) {
            colIS.Item(1).OLEFormat.Edit()
        }
    }
}
```

## 成员属性

### Comment.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Comment 对象的变量。

### Comment.Author

返回或设置一个 **String** 类型的值，该值代表某一批注的作者姓名。可读写。

**语法**

**express.Author**

\*express是一个代表 Comment 对象的变量。

**示例**

``` JavaScript
/* 本示例设置作者姓名，并且为活动文档的第一个批注设置作者姓名缩写。 */
function test() {
    if (Application.ActiveDocument.Comments.Count >= 1) {
        let rng = Application.ActiveDocument.Comments.Item(1)
        rng.Author = ("Joe Smith")
        rng.Initial = ("JAS")
    }
}
```

### Comment.Date

返回代表批注的插入日期和时间的 **Date** 类型。只读。

**语法**

**express.Date**

\*express是一个代表 Comment 对象的变量。

### Comment.Index

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Comment 对象的变量。

### Comment.Initial

返回或设置与指定批注相关联的用户的姓名缩写。 **String** 类型，可读写。

**语法**

**express.Initial**

\*express是一个代表 Comment 对象的变量。

**示例**

``` JavaScript
/*本示例显示选定内容的第一条批注的用户姓名缩写。*/
function test() {
    if (Application.Selection.Comments.Count >= 1) {
        alert("Comment made by " + Selection.Comments.Item(1).Initial)
    }
}
```

``` JavaScript
/*本示例检查文档第一节中每条备注的作者姓名缩写。如果作者姓名缩写为“WPSOffice”，本示例则将其改为“KAE”。*/
function test() {
    let rngTemp = Application.ActiveDocument.Sections.Item(1).Range
    for (let i = 1; i <= rngTemp.Comments.Count; i++) {
        if (rngTemp.Comments.Item(i).Initial == "MSOffice") {
            rngTemp.Comments.Item(i).Initial = "KAE"
        }
    }
}
```

### Comment.IsInk

返回一个 **Boolean** 类型的值，该值代表批注是否为手写批注。

**语法**

**express.IsInk**

\*express是一个代表 Comment 对象的变量。

**示例**

``` JavaScript
/* 删除文档中的手写签批 */
function test() {
    for(let i = 1; i <=  Application.ActiveDocument.Comments.count; i++){
    if(Application.ActiveDocument.Comments.Item(i).IsInk == true){
        objComment.Delete()
    }
}
```

### Comment.Parent

返回一个 **Object** 类型值，该值代表指定 **Comment** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Comment 对象的变量。

### Comment.Range

返回代表备注内容的 **Range** 对象。

**语法**

**express.Range**

\*express是一个代表 Comment 对象的变量。

**示例**

``` JavaScript
/* 修改第一个备注的文本，在前面附加上test文本 */
Application.ActiveDocument.Comments.Item(1).Range.InsertBefore("test")
```

### Comment.Reference

返回代表备注的引用标记的 **Range** 对象。

**语法**

**express.Reference**

\*express是一个代表 Comment 对象的变量。

### Comment.Scope

返回一个 **Range** 对象，该对象代表指定批注所标记的文本范围。

**语法**

**express.Scope**

\*express是一个代表 Comment 对象的变量。

**示例**

``` JavaScript
/*本示例显示与选定内容中第一条批注相关联的文本。*/
function test() {
    if (Application.Selection.Comments.Count >= 1) {
        let myRange = Application.Selection.Comments.Item(1).Scope
        alert(myRange.Text)
    }
}
```

``` JavaScript
/*本示例复制与活动文档中最后一条批注相关联的文本。*/
function test() {
    total = Application.ActiveDocument.Comments.Count
    if (total >= 1) {
        Application.ActiveDocument.Comments.Item(total).Scope.Copy()
    }
}
```

### Comment.ShowTip

如果在“屏幕提示”中显示与批注相关联的文本，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowTip**

\*express是一个代表 Comment 对象的变量。

**说明**

当单击鼠标或按下任意键时，“屏幕提示”才会消失。

**示例**

``` JavaScript
/*本示例显示活动文档中的第一条批注的“屏幕提示”。*/
function test() {
    if (Application.ActiveDocument.Comments.Count >= 1) {
        Application.ActiveDocument.Comments.Item(1).ShowTip = true
    }
}
```

``` JavaScript
/*本示例显示活动文档中下一条批注的“屏幕提示”。*/
function test() {
    if (Application.ActiveDocument.Comments.Count >= 1) {
        Application.Selection.GoTo(wdGoToComment, wdGoToNext)
        Application.Selection.MoveEnd(wdWord, 1)
        Application.Selection.Comments.Item(1).ShowTip == true
    }
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Comment](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Document 对象

## 简介

代表一个文档。 Document 对象是 Documents 集合的成员。 Documents 集合中包含当前在 WPS 中打开的所有 Document 对象。

## 说明

使用 Documents ( *Index* ) 返回单个 Document 对象（其中 *Index* 是文档名称或索引编号）。以下示例关闭名为“Report.doc”的文档但不保存更改。

``` JavaScript
Application.Documents.Item("Report.doc").Close(wdDoNotSaveChanges)
```

索引编号代表文档在 Documents 集合中的位置。以下示例激活 Documents 集合中的第一篇文档。

``` JavaScript
Application.Documents.Item(1).Activate()
```

可以使用 ActiveDocument 属性引用具有焦点的文档。以下示例使用 Activate 方法激活名为“Document 1”的文档，然后将页面方向设置为横向，并打印该文档。

``` JavaScript
function test(){
  Application.Documents.Item("Document1").Activate()
  Application.ActiveDocument.PageSetup.Orientation = wdOrientLandscape
  Application.ActiveDocument.PrintOut()
}
```

## 方法

| 序号 | 名称                                                                   | 说明                                                                                                                                                                  |
|:----:|:-----------------------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AcceptAllRevisions](#Document.AcceptAllRevisions)                     | 接受对指定文档的所有修订。                                                                                                                                            |
|  2   | [AcceptAllRevisionsShown](#Document.AcceptAllRevisionsShown)           | 接受显示在屏幕上的指定文档中的所有修订。                                                                                                                              |
|  3   | [Activate](#Document.Activate)                                         | 激活指定的文档，使其成为活动文档。                                                                                                                                    |
|  4   | [AddToFavorites](#Document.AddToFavorites)                             | 创建文档或超链接的快捷方式，并将其添加到“收藏夹”。                                                                                                                    |
|  5   | [ApplyQuickStyleSet2](#Document.ApplyQuickStyleSet2)                   | 将指定的快速样式集应用于文档。                                                                                                                                        |
|  6   | [ApplyTheme](#Document.ApplyTheme)                                     | 将 主题 应用于打开的文档。                                                                                                                                            |
|  7   | [AutoFormat](#Document.AutoFormat)                                     | 自动给文档套用格式。                                                                                                                                                  |
|  8   | [CanCheckin](#Document.CanCheckin)                                     | 如果该方法返回值为 True ，则 WPS 可将指定文档签入服务器。可读/写 Boolean 类型。                                                                                       |
|  9   | [CheckConsistency](#Document.CheckConsistency)                         | 搜索日语文档中的所有文字，并显示对相同单词的字符用法不一致的实例。                                                                                                    |
|  10  | [CheckGrammar](#Document.CheckGrammar)                                 | 开始对指定的文档或区域执行拼写和语法检查。                                                                                                                            |
|  11  | [CheckIn](#Document.CheckIn)                                           | 该方法将文档从本地计算机返回到服务器，并将本地文档设为只读，使之无法在本地进行编辑。                                                                                  |
|  12  | [CheckInWithVersion](#Document.CheckInWithVersion)                     | 将文档从本地计算机保存到服务器，并将本地文档设为只读，使之无法在本地进行编辑。                                                                                        |
|  13  | [CheckSpelling](#Document.CheckSpelling)                               | 开始对指定的文档或区域进行拼写检查。                                                                                                                                  |
|  14  | [Close](#Document.Close)                                               | 关闭指定的文档。                                                                                                                                                      |
|  15  | [ClosePrintPreview](#Document.ClosePrintPreview)                       | 将指定文档从打印预览视图切换到以前的视图。                                                                                                                            |
|  16  | [Compare](#Document.Compare)                                           | 显示修订标记，以表明指定的文档与另一个文档的区别。                                                                                                                    |
|  17  | [Compatibility](#Document.Compatibility)                               | 如果为 True ，则启用由 *Type* 参数指定的兼容性选项。兼容性选项将影响 WPS 中文档的显示方式。 Boolean 类型，可读写。                                                    |
|  18  | [ComputeStatistics](#Document.ComputeStatistics)                       | 返回基于指定文档内容的统计值。 Long 类型。                                                                                                                            |
|  19  | [Convert](#Document.Convert)                                           | 将文件转换成最新的文件格式，并启用所有新功能。                                                                                                                        |
|  20  | [ConvertAutoHyphens](#Document.ConvertAutoHyphens)                     | 将由自动断字功能创建的连字符转换为手动连字符。                                                                                                                        |
|  21  | [ConvertNumbersToText](#Document.ConvertNumbersToText)                 | 将指定文档中的列表编号和 LISTNUM 域更改为文本。                                                                                                                       |
|  22  | [ConvertVietDoc](#Document.ConvertVietDoc)                             | 使用非默认的代码页将一篇越南语文档重新转换为 Unicode。                                                                                                                |
|  23  | [CopyStylesFromTemplate](#Document.CopyStylesFromTemplate)             | 将指定模板中的样式复制到某篇文档。                                                                                                                                    |
|  24  | [CountNumberedItems](#Document.CountNumberedItems)                     | 返回指定 Document 对象中项目符号项或编号项以及 LISTNUM 域的数目。                                                                                                     |
|  25  | [CreateLetterContent](#Document.CreateLetterContent)                   | 生成并返回一个 LetterContent 对象，该对象基于指定信函内容。 LetterContent 对象。                                                                                      |
|  26  | [DataForm](#Document.DataForm)                                         | 显示 “数据表单” 对话框，在此对话框中可添加、删除或修改记录。                                                                                                          |
|  27  | [DeleteAllComments](#Document.DeleteAllComments)                       | 删除文档中 Comments 集合内的所有备注。                                                                                                                                |
|  28  | [DeleteAllCommentsShown](#Document.DeleteAllCommentsShown)             | 删除显示在屏幕上的指定文档的所有修订。                                                                                                                                |
|  29  | [DeleteAllEditableRanges](#Document.DeleteAllEditableRanges)           | 删除（指定用户或用户组有权修改其权限的）所有区域中的权限。                                                                                                            |
|  30  | [DeleteAllInkAnnotations](#Document.DeleteAllInkAnnotations)           | 删除文档中所有手写墨迹注释。                                                                                                                                          |
|  31  | [DetectLanguage](#Document.DetectLanguage)                             | 分析指定文本，以确定书写文本的语言类型。                                                                                                                              |
|  32  | [DowngradeDocument](#Document.DowngradeDocument)                       | 将文档降级为 WPS 97-2003 文档格式，以便可以在以前版本的 WPS 中进行编辑。                                                                                              |
|  33  | [EndReview](#Document.EndReview)                                       | 终止审阅使用 SendForReview 方法发送以进行审阅的文件，或通过电子邮件将文档发送给另一用户而自动进行审阅的文件。                                                         |
|  34  | [ExportAsFixedFormat](#Document.ExportAsFixedFormat)                   | 将文档保存为 PDF 或 XPS 格式。                                                                                                                                        |
|  35  | [FitToPages](#Document.FitToPages)                                     | 适当缩小文本的字体大小，使文档的页数减少。                                                                                                                            |
|  36  | [FollowHyperlink](#Document.FollowHyperlink)                           | 显示一个缓存文档（如果该文档已下载）。否则，此方法将解析该超链接，下载目标文档，并在相应的应用程序中显示此文档。                                                      |
|  37  | [ForwardMailer](#Document.ForwardMailer)                               | 您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 Document 对象的 ForwardMailer 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。 |
|  38  | [FreezeLayout](#Document.FreezeLayout)                                 | 在 Web 视图中，按照文档当前所显示的样子固定文档的版式，以使换行保持不变，且在调整窗口大小时墨迹不会移动。                                                             |
|  39  | [GetCrossReferenceItems](#Document.GetCrossReferenceItems)             | 返回一个项目数组，根据指定的交叉引用类型可对该数组中的项目进行交叉引用。                                                                                              |
|  40  | [GetLetterContent](#Document.GetLetterContent)                         | 检索指定文档中的信函内容并返回一个 LetterContent 对象。                                                                                                               |
|  41  | [GetWorkflowTasks](#Document.GetWorkflowTasks)                         | 返回一个 WorkflowTasks 集合，该集合表示分配给文档的工作流任务。                                                                                                       |
|  42  | [GetWorkflowTemplates](#Document.GetWorkflowTemplates)                 | 返回一个 WorkflowTemplates 集合，该集合表示附加到文档的工作流模板。                                                                                                   |
|  43  | [GoTo](#Document.GoTo)                                                 | 返回一个 Range 对象，该对象代表指定项（如页、书签或域）的起始位置。                                                                                                   |
|  44  | [LockServerFile](#Document.LockServerFile)                             | 在服务器上锁定文件，以避免任何其他人进行编辑。                                                                                                                        |
|  45  | [MakeCompatibilityDefault](#Document.MakeCompatibilityDefault)         | 兼容性选项位于 “选项” 对话框的 “兼容性” 选项卡上，它们是新文档的默认设置。                                                                                            |
|  46  | [ManualHyphenation](#Document.ManualHyphenation)                       | 启动文档的手动断字设置，每次一行。                                                                                                                                    |
|  47  | [Merge](#Document.Merge)                                               | 将文档中用修订标记标识的修改合并到另一篇文档。                                                                                                                        |
|  48  | [Post](#Document.Post)                                                 | 将指定的文档传送到 Microsoft Exchange 中的公用文件夹。                                                                                                                |
|  49  | [PresentIt](#Document.PresentIt)                                       | 打开 WPP，并加载指定的 WPS 文档。                                                                                                                                     |
|  50  | [PrintOut](#Document.PrintOut)                                         | 打印指定文档的全部或部分内容。                                                                                                                                        |
|  51  | [PrintPreview](#Document.PrintPreview)                                 | 将视图切换到打印预览。                                                                                                                                                |
|  52  | [Protect](#Document.Protect)                                           | 帮助保护指定的文档不被更改。保护文档后，用户只能进行有限的更改，例如添加注释，进行修订或填写表格。                                                                    |
|  53  | [Range](#Document.Range)                                               | 通过使用指定的开始和结束字符位置返回一个 Range 对象。                                                                                                                 |
|  54  | [Redo](#Document.Redo)                                                 | 重复最后一次撤消的操作（ Undo 方法的逆操作）。如果重复操作成功，则返回 True 。                                                                                        |
|  55  | [RejectAllRevisions](#Document.RejectAllRevisions)                     | 拒绝指定文档的所有修订。                                                                                                                                              |
|  56  | [RejectAllRevisionsShown](#Document.RejectAllRevisionsShown)           | 拒绝显示在屏幕上的文档的所有修订。                                                                                                                                    |
|  57  | [Reload](#Document.Reload)                                             | 通过解析指向缓冲文档的超链接再下载该文档来重新加载该文档。                                                                                                            |
|  58  | [ReloadAs](#Document.ReloadAs)                                         | 使用指定的编码重新加载一篇基于 HTML 文档的文档。                                                                                                                      |
|  59  | [RemoveDocumentInformation](#Document.RemoveDocumentInformation)       | 从文档中删除敏感性信息、属性、批注和其他元数据。                                                                                                                      |
|  60  | [RemoveLockedStyles](#Document.RemoveLockedStyles)                     | 当文档中应用了格式设置限制时，清除锁定样式的文档。                                                                                                                    |
|  61  | [RemoveNumbers](#Document.RemoveNumbers)                               | 删除指定文档中的编号或项目符号。                                                                                                                                      |
|  62  | [RemoveTheme](#Document.RemoveTheme)                                   | 删除当前文档中的活动 主题 。                                                                                                                                          |
|  63  | [Repaginate](#Document.Repaginate)                                     | 对整篇文档重新分页。                                                                                                                                                  |
|  64  | [Reply](#Document.Reply)                                               | 打开一封新的电子邮件（ “收件人” 行已含发件人地址）以答复活动邮件。                                                                                                    |
|  65  | [ReplyAll](#Document.ReplyAll)                                         | 打开一封新的电子邮件（ “收件人” 和 “抄送” 行已含发件人和所有其他收件人的地址）以答复活动邮件。                                                                        |
|  66  | [ReplyWithChanges](#Document.ReplyWithChanges)                         | 为文档（该文档以被发送出去进行审阅）的作者发送电子邮件，通知他们审阅者已经审阅完文档。                                                                                |
|  67  | [ResetFormFields](#Document.ResetFormFields)                           | 清除文档中的所有窗体域，准备再次填充窗体。                                                                                                                            |
|  68  | [RunAutoMacro](#Document.RunAutoMacro)                                 | 运行存储于指定文档的自动宏。如果没有自动宏，则不做任何操作。                                                                                                          |
|  69  | [RunLetterWizard](#Document.RunLetterWizard)                           | 对指定的文档运行英文信函向导。                                                                                                                                        |
|  70  | [Save](#Document.Save)                                                 | 保存当前 文档。                                                                                                                                                       |
|  71  | [SaveAs2](#Document.SaveAs2)                                           | 使用新的名称或格式保存指定的文档。此方法的一些参数与 “另存为” 对话框（ “文件” 选项卡）中的选项相对应。                                                                |
|  72  | [SaveAsQuickStyleSet](#Document.SaveAsQuickStyleSet)                   | 保存当前所用的快速样式组。                                                                                                                                            |
|  73  | [Select](#Document.Select)                                             | 选择指定文档的内容。                                                                                                                                                  |
|  74  | [SelectAllEditableRanges](#Document.SelectAllEditableRanges)           | 选择指定用户或用户组有权修改的所有区域。                                                                                                                              |
|  75  | [SelectContentControlsByTag](#Document.SelectContentControlsByTag)     | 返回一个 ContentControls 集合，该集合代表文档中具有 *Tag* 参数中指定标记值的所有内容控件。只读。                                                                      |
|  76  | [SelectContentControlsByTitle](#Document.SelectContentControlsByTitle) | 返回一个 ContentControls 集合，该集合代表文档中具有 *Title* 参数中指定标题的所有内容控件。只读。                                                                      |
|  77  | [SelectLinkedControls](#Document.SelectLinkedControls)                 | 返回一个 ContentControls 集合，该集合代表文档中根据 Node 参数的指定，链接到文档 XML 数据存储区中特定自定义 XML 节点的所有内容控件。只读。                             |
|  78  | [SelectNodes](#Document.SelectNodes)                                   | 返回一个 XMLNodes 集合，代表所有与 *XPath* 参数相匹配的节点并按其在文档或区域中显示的顺序排列。                                                                       |
|  79  | [SelectSingleNode](#Document.SelectSingleNode)                         | 返回一个 XMLNode 对象，该对象代表画布与指定文档中 *XPath* 参数相匹配的第一个节点。                                                                                    |
|  80  | [SelectUnlinkedControls](#Document.SelectUnlinkedControls)             | 返回一个 ContentControls 集合，该集合代表文档中没有链接到文档 XML 数据存储区中 XML 节点的所有内容控件。只读。                                                         |
|  81  | [SendFax](#Document.SendFax)                                           | 将指定的文档作为传真发送，且无需用户交互。                                                                                                                            |
|  82  | [SendFaxOverInternet](#Document.SendFaxOverInternet)                   | 将文档发送给传真服务提供商，该提供商将该文档传真给一个或多个指定收件人。                                                                                              |
|  83  | [SendForReview](#Document.SendForReview)                               | 将文档通过电子邮件发送给指定收件人进行审阅。                                                                                                                          |
|  84  | [SendMail](#Document.SendMail)                                         | 打开一个邮件窗口以通过 Microsoft Exchange 发送指定文档。                                                                                                              |
|  85  | [SendMailer](#Document.SendMailer)                                     | 您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 Document 对象的 SendMailer 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。    |
|  86  | [SetCompatibilityMode](#Document.SetCompatibilityMode)                 | 设置文档的兼容模式。                                                                                                                                                  |
|  87  | [SetDefaultTableStyle](#Document.SetDefaultTableStyle)                 | 指定用于文档中新建表的表样式。                                                                                                                                        |
|  88  | [SetLetterContent](#Document.SetLetterContent)                         | 将指定 LetterContent 对象的内容插入文档。                                                                                                                             |
|  89  | [SetPasswordEncryptionOptions](#Document.SetPasswordEncryptionOptions) | 设置 WPS 用于通过密码加密文档的选项                                                                                                                                   |
|  90  | [ToggleFormsDesign](#Document.ToggleFormsDesign)                       | 在打开和关闭窗体设计模式之间切换。                                                                                                                                    |
|  91  | [TransformDocument](#Document.TransformDocument)                       | 将 Extensible Stylesheet Language Transformation (XSLT) 文件应用于指定文档，并用结果替换该文档。                                                                      |
|  92  | [Undo](#Document.Undo)                                                 | 消最后一次操作或最后一系列操作，这些操作显示在 “撤消” 列表中。返回 Boolean 。 如果撤消操作成功，则返回 True 。                                                        |
|  93  | [UndoClear](#Document.UndoClear)                                       | 清除可对指定文档撤消的操作列表。                                                                                                                                      |
|  94  | [Unprotect](#Document.Unprotect)                                       | 清除对指定文档的保护。                                                                                                                                                |
|  95  | [UpdateStyles](#Document.UpdateStyles)                                 | 将所有样式从附加模板复制到文档中，同时覆盖文档中所有现有同名的样式。                                                                                                  |
|  96  | [UpdateSummaryProperties](#Document.UpdateSummaryProperties)           | 更新 “文件” 菜单 “属性” 对话框中的关键词和备注文字，以反映为指定文档编写的自动摘要内容。                                                                              |
|  97  | [ViewCode](#Document.ViewCode)                                         | 显示指定文档中所选 Microsoft ActiveX 控件的代码窗口。                                                                                                                 |
|  98  | [ViewPropertyBrowser](#Document.ViewPropertyBrowser)                   | 显示指定文档中所选 Microsoft ActiveX 控件的属性窗口。                                                                                                                 |
|  99  | [WebPagePreview](#Document.WebPagePreview)                             | 显示将当前文档保存为网页时的预览。                                                                                                                                    |

## 属性

| 序号 | 名称                                                                           | 说明                                                                                                                                                                                 |
|:----:|:-------------------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ActiveTheme](#Document.ActiveTheme)                                           | 返回指定文档的活动 主题 名称及该主题的格式选项。 String 类型，只读。                                                                                                                 |
|  2   | [ActiveThemeDisplayName](#Document.ActiveThemeDisplayName)                     | 返回指定文档的活动 主题 的显示名称。 String 类型，只读。                                                                                                                             |
|  3   | [ActiveWindow](#Document.ActiveWindow)                                         | 返回一个 Window 对象，该对象代表活动窗口（焦点所在的窗口）。只读。                                                                                                                   |
|  4   | [ActiveWritingStyle](#Document.ActiveWritingStyle)                             | 返回或设置指定文档中指定语言的写作风格。 String 类型，可读写。                                                                                                                       |
|  5   | [Application](#Document.Application)                                           | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                                       |
|  6   | [AttachedTemplate](#Document.AttachedTemplate)                                 | 返回一个 Template 对象，该对象代表与指定文档相关联的模板。可读/写 Variant 类型。                                                                                                     |
|  7   | [AutoFormatOverride](#Document.AutoFormatOverride)                             | 返回或设置一个 Boolean 类型的值，该数值代表是否在已应用格式设置限制的文档中使用自动设置格式替代格式设置限制。                                                                        |
|  8   | [AutoHyphenation](#Document.AutoHyphenation)                                   | 如果代表指定文档已启动自动断字功能，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                      |
|  9   | [Background](#Document.Background)                                             | 返回一个 Shape 对象，该对象代表指定文档的背景图像。只读。                                                                                                                            |
|  10  | [Bibliography](#Document.Bibliography)                                         | 返回一个 Bibliography 对象，代表文档中包含的书目参考。只读。                                                                                                                         |
|  11  | [Bookmarks](#Document.Bookmarks)                                               | 返回一个 Bookmarks 集合，该集合代表文档中的所有书签。只读。                                                                                                                          |
|  12  | [BuiltInDocumentProperties](#Document.BuiltInDocumentProperties)               | 返回一个 DocumentProperties 集合，该集合代表指定文档的所有内置的文档属性。                                                                                                           |
|  13  | [Characters](#Document.Characters)                                             | 返回一个 Characters 集合，该集合代表文档中的字符。只读。                                                                                                                             |
|  14  | [ChildNodeSuggestions](#Document.ChildNodeSuggestions)                         | 返回一个 XMLChildNodeSuggestions 集合，该集合代表允许在文档中使用的元素列表。                                                                                                        |
|  15  | [ClickAndTypeParagraphStyle](#Document.ClickAndTypeParagraphStyle)             | 返回或设置“即点即输”功能在指定文档中应用于文字的默认段落样式。 Variant 类型，可读写。                                                                                                |
|  16  | [CoAuthoring](#Document.CoAuthoring)                                           | 返回 CoAuthoring 对象，该对象提供文档中与共同创作相关的对象模型的入口点。只读。                                                                                                      |
|  17  | [CodeName](#Document.CodeName)                                                 | 返回指定文档的代码名称。 String 类型，只读。                                                                                                                                         |
|  18  | [CommandBars](#Document.CommandBars)                                           | 返回一个 CommandBars 集合，该集合代表 WPS 中的菜单栏以及所有工具栏。                                                                                                                 |
|  19  | [Comments](#Document.Comments)                                                 | 返回一个 Comments 集合，该集合代表指定文档的所有批注。只读。                                                                                                                         |
|  20  | [CompatibilityMode](#Document.CompatibilityMode)                               | 返回 Number 类型值，该值指定打开文档时 WPS 2015 使用的兼容模式。只读。                                                                                                               |
|  21  | [ConsecutiveHyphensLimit](#Document.ConsecutiveHyphensLimit)                   | 返回或设置以连字符结束的连续行的最大数目。 Long 类型，可读写。                                                                                                                       |
|  22  | [Container](#Document.Container)                                               | 返回一个对象，该对象代表指定文档的容器应用程序。 Object 类型，只读。                                                                                                                 |
|  23  | [Content](#Document.Content)                                                   | 返回一个 Range 对象，该对象代表主文档 文章 。只读。                                                                                                                                  |
|  24  | [ContentControls](#Document.ContentControls)                                   | 返回一个 ContentControls 集合，该集合代表文档中的所有内容控件。只读。                                                                                                                |
|  25  | [ContentTypeProperties](#Document.ContentTypeProperties)                       | 返回一个 MetaProperties 集合，该集合代表文档中所存储的元数据，如作者姓名、主题和公司。只读。                                                                                         |
|  26  | [Creator](#Document.Creator)                                                   | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                                         |
|  27  | [CurrentRsid](#Document.CurrentRsid)                                           | 返回一个 Long 类型的值，该值代表 WPS 为文档中的更改所分配的随机数。只读。                                                                                                            |
|  28  | [CustomDocumentProperties](#Document.CustomDocumentProperties)                 | 返回一个 DocumentProperties 集合，该集合代表指定文档的所有自定义文档属性。                                                                                                           |
|  29  | [CustomXMLParts](#Document.CustomXMLParts)                                     | 返回一个 CustomXMLParts 集合，该集合代表 XML 数据存储区中的自定义 XML。只读。                                                                                                        |
|  30  | [DefaultTableStyle](#Document.DefaultTableStyle)                               | 返回一个 Variant 类型的值，该值代表可以应用于文档中新创建的所有表格的表格样式。只读。                                                                                                |
|  31  | [DefaultTabStop](#Document.DefaultTabStop)                                     | 以磅为单位返回或设置指定文档中的默认制表位间隔。 Single 类型，可读写。                                                                                                               |
|  32  | [DefaultTargetFrame](#Document.DefaultTargetFrame)                             | 返回或设置一个 String 类型的值，该值代表通过超链接可到达显示网页的浏览器框架。可读写。                                                                                               |
|  33  | [DisableFeatures](#Document.DisableFeatures)                                   | 如果为 True ，则禁用 DisableFeaturesIntroducedAfter 属性中指定的版本之后的所有功能。默认值为 False 。 Boolean 类型，可读写。                                                         |
|  34  | [DisableFeaturesIntroducedAfter](#Document.DisableFeaturesIntroducedAfter)     | 仅对该文档禁用指定的 WPS 版本后引入的所有功能。 WdDisableFeaturesIntroducedAfter 类型，可读写。                                                                                      |
|  35  | [DocumentInspectors](#Document.DocumentInspectors)                             | 返回一个 DocumentInspectors 集合，该集合使您能够查找隐藏的个人信息，如作者姓名、公司名称和修订日期。只读。                                                                           |
|  36  | [DocumentLibraryVersions](#Document.DocumentLibraryVersions)                   | 返回一个 DocumentLibraryVersions 集合，该集合表示共享文档的版本集合，该共享文档启用了版本控制并存储在服务器的文档库中。                                                              |
|  37  | [DocumentTheme](#Document.DocumentTheme)                                       | 返回一个 OfficeTheme 对象，代表应用于文档的 WPS Office主题。只读。                                                                                                                   |
|  38  | [DoNotEmbedSystemFonts](#Document.DoNotEmbedSystemFonts)                       | 如果为 True ，则 WPS 不嵌入常规系统字体。 Boolean 类型，可读写。                                                                                                                     |
|  39  | [Email](#Document.Email)                                                       | 返回一个 Email 对象，该对象包含当前文档的所有与电子邮件相关的属性。只读。                                                                                                            |
|  40  | [EmbedLinguisticData](#Document.EmbedLinguisticData)                           | 如果为 True ，则 WPS 会嵌入语音和手写功能，以便数据可以转换回语音或手写。 Boolean 类型，可读写。                                                                                     |
|  41  | [EmbedTrueTypeFonts](#Document.EmbedTrueTypeFonts)                             | 如果在保存文档时，WPS 在文档中嵌入 TrueType 字体，则该属性值为 True 。 Boolean 类型，可读写。                                                                                        |
|  42  | [EncryptionProvider](#Document.EncryptionProvider)                             | 返回一个 String 类型的值，该值指定 WPS 在对文档进行加密时使用的算法加密提供程序的名称。可读写。                                                                                      |
|  43  | [Endnotes](#Document.Endnotes)                                                 | 返回一个 Endnotes 集合，该集合代表文档中的所有尾注。只读。                                                                                                                           |
|  44  | [EnforceStyle](#Document.EnforceStyle)                                         | 返回或设置一个 Boolean 类型的值，该值代表是否在受保护的文档中实施格式设置限制。                                                                                                      |
|  45  | [Envelope](#Document.Envelope)                                                 | 返回一个 Envelope 对象，该对象代表文档中的信封和信封功能。只读。                                                                                                                     |
|  46  | [FarEastLineBreakLanguage](#Document.FarEastLineBreakLanguage)                 | 返回或设置一个 WdFarEastLineBreakLanguageID 类型的值，该值代表在指定文档或模板中对文本进行换行时使用的东亚语言。可读写。                                                             |
|  47  | [FarEastLineBreakLevel](#Document.FarEastLineBreakLevel)                       | 返回或设置一个 WdFarEastLineBreakLevel 类型的值，该值代表指定文档的换行控制级别。可读写。                                                                                            |
|  48  | [Fields](#Document.Fields)                                                     | 返回一个 Fields 集合，该集合代表文档中的所有域。只读。                                                                                                                               |
|  49  | [Final](#Document.Final)                                                       | 返回或设置 Boolean 值，该值指示文档是否是最终的。可读/写。                                                                                                                           |
|  50  | [Footnotes](#Document.Footnotes)                                               | 返回一个 Footnotes 集合，该集合代表文档中的所有脚注。只读。                                                                                                                          |
|  51  | [FormattingShowClear](#Document.FormattingShowClear)                           | 如果为 True ，则 WPS 在 “样式和格式” 任务窗格中显示“清除格式”。 Boolean 类型，可读写。                                                                                               |
|  52  | [FormattingShowFilter](#Document.FormattingShowFilter)                         | 设置或返回一个 WdShowFilter 常量，代表 “样式和格式” 任务窗格中显示的样式和格式。 Boolean 类型，可读写                                                                                |
|  53  | [FormattingShowFont](#Document.FormattingShowFont)                             | 如果为 True ，则 WPS 显示 “样式和格式” 任务窗格中的字体格式。 Boolean 类型，可读写。                                                                                                 |
|  54  | [FormattingShowNextLevel](#Document.FormattingShowNextLevel)                   | 返回或设置 Boolean 值，该值表示 WPS 在使用上一个标题级别时是否显示下一个标题级别。可读/写。                                                                                          |
|  55  | [FormattingShowNumbering](#Document.FormattingShowNumbering)                   | 如果为 True ，则 WPS 在 “样式和格式” 任务窗格中显示编号格式。 Boolean 类型，可读写。                                                                                                 |
|  56  | [FormattingShowParagraph](#Document.FormattingShowParagraph)                   | 如果为 True ，则 WPS 在 “样式和格式” 任务窗格中显示段落格式。 Boolean 类型，可读写。                                                                                                 |
|  57  | [FormattingShowUserStyleName](#Document.FormattingShowUserStyleName)           | 返回或设置 Boolean 值，该值表示是否显示用户定义的样式。可读/写。                                                                                                                     |
|  58  | [FormFields](#Document.FormFields)                                             | 返回一个 FormFields 集合，该集合代表文档中的所有窗体域。只读。                                                                                                                       |
|  59  | [FormsDesign](#Document.FormsDesign)                                           | 如果指定的文档处于窗体设计模式，则该属性值为 True 。 Boolean 类型，只读。                                                                                                            |
|  60  | [FormsDesignFormsDesign](#Document.FormsDesignFormsDesign)                     | 如果指定的文档处于窗体设计模式，则该属性值为 True 。 Boolean 类型，只读。                                                                                                            |
|  61  | [Frames](#Document.Frames)                                                     | 返回一个 Frames 集合，该集合代表文档中的所有图文框。只读。                                                                                                                           |
|  62  | [Frameset](#Document.Frameset)                                                 | 返回一个 Frameset 对象，该对象代表一个整体框架页或框架页的单个框架。只读。                                                                                                           |
|  63  | [FullName](#Document.FullName)                                                 | 返回一个 String 类型的值，该值代表包括路径的文档的名称。只读。                                                                                                                       |
|  64  | [GrammarChecked](#Document.GrammarChecked)                                     | 如果已经检查了指定范围或文档的语法，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                      |
|  65  | [GrammaticalErrors](#Document.GrammaticalErrors)                               | 返回一个 ProofreadingErrors 集合，该集合代表指定的文档中有语法检查错误的句子。只读。                                                                                                 |
|  66  | [GridDistanceHorizontal](#Document.GridDistanceHorizontal)                     | 在指定文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置代表 WPS 所用的不可见的网格线之间的水平距离的 Single 类型，可读写。                                                  |
|  67  | [GridDistanceVertical](#Document.GridDistanceVertical)                         | 在指定的文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置代表 WPS 所用的不可见网格线之间的垂直距离的 Single 类型。可读写。                                                  |
|  68  | [GridOriginFromMargin](#Document.GridOriginFromMargin)                         | 如果 WPS 将在页面左上角开始使用字符网格，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                 |
|  69  | [GridOriginHorizontal](#Document.GridOriginHorizontal)                         | 返回或设置一个 Single 类型的值，该值代表不可见网格相对于页面左边的起点位置，当在指定的文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。可读写。                           |
|  70  | [GridOriginVertical](#Document.GridOriginVertical)                             | 返回或设置一个 Single 类型的值，该值代表不可见网格相对于页面顶边的起点位置，当在指定的文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。可读写。                           |
|  71  | [GridSpaceBetweenHorizontalLines](#Document.GridSpaceBetweenHorizontalLines)   | 返回或设置 WPS 在页面视图中显示的水平字符网格的间隔。 Long 类型，可读写。                                                                                                            |
|  72  | [GridSpaceBetweenVerticalLines](#Document.GridSpaceBetweenVerticalLines)       | 返回或设置 WPS 在页面视图中显示的垂直字符网格线间隔。 Long 类型，可读写。                                                                                                            |
|  73  | [HasMailer](#Document.HasMailer)                                               | 您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 Document 对象的 HasMailer 属性的信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。                      |
|  74  | [HasPassword](#Document.HasPassword)                                           | 如果需要密码才能打开指定文档，则该属性值为 True 。 Boolean 类型，只读。                                                                                                              |
|  75  | [HasVBProject](#Document.HasVBProject)                                         | 返回一个 Boolean 类型的值，该值代表文档是否具有附加的 示例代码 项目。只读。                                                                                                          |
|  76  | [HTMLDivisions](#Document.HTMLDivisions)                                       | 返回一个 HTMLDivisions 集合，该集合代表 Web 文档中的 HTML DIV 元素。                                                                                                                 |
|  77  | [Hyperlinks](#Document.Hyperlinks)                                             | 返回一个 Hyperlinks 集合，该集合代表指定文档中的所有超链接。只读。                                                                                                                   |
|  78  | [HyphenateCaps](#Document.HyphenateCaps)                                       | 如果全是大写字母时可对其进行断字，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                        |
|  79  | [HyphenationZone](#Document.HyphenationZone)                                   | 返回或设置断字区的宽度（以磅为单位）。 Long 类型，可读写。                                                                                                                           |
|  80  | [Indexes](#Document.Indexes)                                                   | 返回一个 Indexes 集合，该集合代表指定文档的所有索引。只读。                                                                                                                          |
|  81  | [InlineShapes](#Document.InlineShapes)                                         | 返回一个 InlineShapes 集合，该集合代表文档中的所有 InlineShape 对象。只读。                                                                                                          |
|  82  | [IsMasterDocument](#Document.IsMasterDocument)                                 | 如果指定的文档为一篇主文档，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                |
|  83  | [IsSubdocument](#Document.IsSubdocument)                                       | 如果指定文档是主文档的子文档，则该属性值为 True 。 Boolean 类型，只读。                                                                                                              |
|  84  | [JustificationMode](#Document.JustificationMode)                               | 返回或设置指定文档字符间距的调整量。可读/写 WdJustificationMode 类型。                                                                                                               |
|  85  | [KerningByAlgorithm](#Document.KerningByAlgorithm)                             | 如果 WPS 在指定文档中调整半角拉丁字符和标点符号的间距，则为 True 。可读/写 Boolean 类型。                                                                                            |
|  86  | [Kind](#Document.Kind)                                                         | 返回或设置 WPS 在自动设置指定文档的格式时使用的格式类型。 WdDocumentKind 类型，可读写。                                                                                              |
|  87  | [LanguageDetected](#Document.LanguageDetected)                                 | 返回或设置一个指定 WPS 是否对指定文本的语言进行检测的值。 Boolean 类型，可读写。                                                                                                     |
|  88  | [ListParagraphs](#Document.ListParagraphs)                                     | 返回一个 ListParagraphs 对象，该对象代表文档中的所有编号段落。只读。                                                                                                                 |
|  89  | [Lists](#Document.Lists)                                                       | 返回一个 Lists 集合，该集合含有指定文档中的所有项目符号和编号列表。只读。                                                                                                            |
|  90  | [ListTemplates](#Document.ListTemplates)                                       | 返回一个 ListTemplates 集合，该集合代表指定文档中的所有列表格式。只读。                                                                                                              |
|  91  | [LockQuickStyleSet](#Document.LockQuickStyleSet)                               | 返回或设置一个 Boolean ，它代表用户是否可以更改使用的快速样式集。可读写。                                                                                                            |
|  92  | [LockTheme](#Document.LockTheme)                                               | 返回或设置一个 Boolean 类型的值，该值代表用户是否可以更改文档主题。可读/写。                                                                                                         |
|  93  | [MailEnvelope](#Document.MailEnvelope)                                         | 返回一个 MsoEnvelope 对象，该对象代表文档的电子邮件标题。                                                                                                                            |
|  94  | [Mailer](#Document.Mailer)                                                     | 您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 Document 对象的 Mailer 属性的信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。                         |
|  95  | [MailMerge](#Document.MailMerge)                                               | 返回一个 MailMerge 对象，该对象代表指定文档的邮件合并功能。只读。                                                                                                                    |
|  96  | [Name](#Document.Name)                                                         | 返回指定文档的名称。 String 类型，只读。                                                                                                                                             |
|  97  | [NoLineBreakAfter](#Document.NoLineBreakAfter)                                 | 返回或设置 WPS 不在其后换行的首尾字符。可读/写 String 类型。                                                                                                                         |
|  98  | [NoLineBreakBefore](#Document.NoLineBreakBefore)                               | 返回或设置 WPS 不在其前换行的首尾字符。可读/写 String 类型。                                                                                                                         |
|  99  | [OMathBreakBin](#Document.OMathBreakBin)                                       | 返回或设置一个 WdOMathBreakBin 常量，该常量代表当公式跨越两行或多行时 WPS 用来放置二元运算符的位置。可读写。                                                                         |
| 100  | [OMathBreakSub](#Document.OMathBreakSub)                                       | 返回或设置一个 WdOMathBreakSub 常量，该常量代表 WPS 如何处理位于换行符之前的减法运算符。可读写。                                                                                     |
| 101  | [OMathFontName](#Document.OMathFontName)                                       | 返回或设置一个 String 类型的值，该值代表文档中用于显示公式的字体的名称。可读写。                                                                                                     |
| 102  | [OMathIntSubSupLim](#Document.OMathIntSubSupLim)                               | 返回或设置一个 Boolean 类型的值，该值代表积分的默认极限位置。可读写。                                                                                                                |
| 103  | [OMathJc](#Document.OMathJc)                                                   | 返回或设置一个 WdOMathJc 常量，该常量代表一组公式的默认对齐方式：左对齐、右对齐、居中或作为一个组居中。可读写。                                                                      |
| 104  | [OMathLeftMargin](#Document.OMathLeftMargin)                                   | 返回或设置一个代表公式的左边距的 Single 。可读写。                                                                                                                                   |
| 105  | [OMathNarySupSubLim](#Document.OMathNarySupSubLim)                             | 返回或设置一个 Boolean 类型的值，该值代表 *n* 元对象（而不是积分）的默认极限位置。可读写。                                                                                           |
| 106  | [OMathRightMargin](#Document.OMathRightMargin)                                 | 返回或设置一个 Single 类型的值，该值代表公式的右边距。可读写。                                                                                                                       |
| 107  | [OMaths](#Document.OMaths)                                                     | 返回一个 OMaths 集合，该集合代表指定区域内的 OMath 对象。只读。                                                                                                                      |
| 108  | [OMathSmallFrac](#Document.OMathSmallFrac)                                     | 返回或设置一个 Boolean 类型的值，该值代表是否在文档所包含的公式中使用小型分数。可读写。                                                                                              |
| 109  | [OMathWrap](#Document.OMathWrap)                                               | 返回或设置一个 Single 类型的值，该值代表换到新的一行的公式的第二行位置。可读写。                                                                                                     |
| 110  | [OpenEncoding](#Document.OpenEncoding)                                         | 返回用于打开指定文档的编码。 MsoEncoding 类型，只读。                                                                                                                                |
| 111  | [OptimizeForWord97](#Document.OptimizeForWord97)                               | 如果 WPS 通过禁用所有不兼容的格式来优化当前文档，以便在 WPS 97 中查看，则该属性值为 True 。 Boolean 类型，可读写。                                                                   |
| 112  | [OriginalDocumentTitle](#Document.OriginalDocumentTitle)                       | 返回一个 String 类型的值，该值代表在运行“精确比较”文档比较功能之后原文档的文档标题。只读。                                                                                           |
| 113  | [PageSetup](#Document.PageSetup)                                               | 返回一个与指定文档相关联的 PageSetup 对象。                                                                                                                                          |
| 114  | [Paragraphs](#Document.Paragraphs)                                             | 返回一个 Paragraphs 集合，该集合代表指定文档中的所有段落。只读。                                                                                                                     |
| 115  | [Parent](#Document.Parent)                                                     |                                                                                                                                                                                      |
| 116  | [Password](#Document.Password)                                                 | 设置在打开文档必须使用的密码。 String 类型，只写。                                                                                                                                   |
| 117  | [PasswordEncryptionAlgorithm](#Document.PasswordEncryptionAlgorithm)           | 返回一个 String 类型的值，该值代表 WPS 用密码加密文档时所用的算法。只读。                                                                                                            |
| 118  | [PasswordEncryptionFileProperties](#Document.PasswordEncryptionFileProperties) | 如果 WPS 为密码保护的文档的文件属性加密，则该属性值为 True 。 Boolean 类型，只读。                                                                                                   |
| 119  | [PasswordEncryptionKeyLength](#Document.PasswordEncryptionKeyLength)           | 返回一个 Long 类型的值，该值代表 WPS 用密码加密文档时所用的密码键长。只读。                                                                                                          |
| 120  | [PasswordEncryptionProvider](#Document.PasswordEncryptionProvider)             | 返回一个 String 类型的值，该值指定 WPS 用密码加密文档时所用的加密算法提供程序的名称。只读。                                                                                          |
| 121  | [Path](#Document.Path)                                                         | 返回文档的磁盘或 Web 路径。只读 String 类型。                                                                                                                                        |
| 122  | [Permission](#Document.Permission)                                             | 返回一个 Permission 对象，该对象代表指定文档中的权限设置。                                                                                                                           |
| 123  | [PrintFormsData](#Document.PrintFormsData)                                     | 如果 WPS 仅将相应联机窗体中的键入的数据打印在预打印窗体上，则该属性值为 True 。 Boolean 类型，可读写。                                                                               |
| 124  | [PrintFractionalWidths](#Document.PrintFractionalWidths)                       | 如果设置指定文档的格式，使用分式间距来显示和打印字符，则该属性值为 True 。 Boolean 类型，可读写。                                                                                    |
| 125  | [PrintPostScriptOverText](#Document.PrintPostScriptOverText)                   | 如果为 True ，则使用 PostScript 打印机时，文档中的 PRINT 域指令（例如，PostScript 命令）将打印在文字和图形的上方。 Boolean 类型，可读写。                                            |
| 126  | [PrintRevisions](#Document.PrintRevisions)                                     | 如果为 True ，则在打印文档的同时打印修订标记。如果返回 False ，则不打印修订标记（即打印接受修订后的状态）。 Boolean 类型，可读写。                                                   |
| 127  | [Protect](#Document.Protect)                                                   | 返回或设置与指定传送名单相关联的文档的保护类型。 WdProtectionType 类型，可读写。                                                                                                     |
| 128  | [ProtectionType](#Document.ProtectionType)                                     | 该属性返回指定文档的保护类型。可以是下列 WdProtectionType 常量之一： wdAllowOnlyComments 、 wdAllowOnlyFormFields 、 wdAllowOnlyReading 、 wdAllowOnlyRevisions 或 wdNoProtection 。 |
| 129  | [ReadabilityStatistics](#Document.ReadabilityStatistics)                       | 返回一个 ReadabilityStatistics 集合，该集合代表指定文档或区域的可读性统计信息。只读。                                                                                                |
| 130  | [ReadingLayoutSizeX](#Document.ReadingLayoutSizeX)                             | 设置或返回一个 Long 类型的值，该值表示当在阅读版式视图中显示文档并对其冻结了输入手写标记时文档中页面的宽度。                                                                         |
| 131  | [ReadingLayoutSizeY](#Document.ReadingLayoutSizeY)                             | 设置或返回一个 Long 类型的值，该值表示当在阅读版式视图中显示文档并对其冻结了输入手写标记时文档中页面的高度。                                                                         |
| 132  | [ReadingModeLayoutFrozen](#Document.ReadingModeLayoutFrozen)                   | 设置或返回一个 Boolean 类型的值，该值表示是否将阅读版式视图中显示的页面冻结为指定大小以向文档插入手写标记。                                                                          |
| 133  | [ReadOnly](#Document.ReadOnly)                                                 | 如果对文档所作修改不能保存到原始文档中，则该属性为 True 。只读 Boolean 类型。                                                                                                        |
| 134  | [ReadOnlyRecommended](#Document.ReadOnlyRecommended)                           | 如果无论用户何时打开文档，WPS 都显示一个消息框，提示以只读方式打开，则该属性值为 True 。 Boolean 类型，可读写。                                                                      |
| 135  | [RemoveDateAndTime](#Document.RemoveDateAndTime)                               | 设置或返回一个 Boolean 类型的值，该值指示文档是否存储修订的日期和时间元数据。                                                                                                        |
| 136  | [RemovePersonalInformation](#Document.RemovePersonalInformation)               | 如果 WPS 在保存文档时从批注、修订和“属性”对话框中删除所有用户信息，则该属性值为 True 。 Boolean 类型，可读写。                                                                       |
| 137  | [Research](#Document.Research)                                                 | 返回一个 Research 对象，该对象代表文档的检索服务。只读。                                                                                                                             |
| 138  | [RevisedDocumentTitle](#Document.RevisedDocumentTitle)                         | 返回一个 String 类型的值，该值代表在运行“精确比较”文档比较功能之后修订的文档的文档标题。只读。                                                                                       |
| 139  | [Revisions](#Document.Revisions)                                               | 返回一个 Revisions 集合，该集合代表文档或区域中的修订。只读。                                                                                                                        |
| 140  | [Saved](#Document.Saved)                                                       | 如果指定的文档或模板从上次保存后一直没有更改，则该属性值为 True 。如果关闭文档时，WPS 提示保存对文档所做的更改，则该属性值为 False 。 Boolean 类型，可读写。                         |
| 141  | [SaveEncoding](#Document.SaveEncoding)                                         | 返回或设置保存文档时使用的编码。 MsoEncoding 类型，可读写。                                                                                                                          |
| 142  | [SaveFormat](#Document.SaveFormat)                                             | 返回指定文档或文件转换器的文件格式。只读 Long 类型。                                                                                                                                 |
| 143  | [SaveFormsData](#Document.SaveFormsData)                                       | 如果 WPS 将输入到一个窗体中的数据作为以制表位分隔的数据库记录保存，则该属性值为 True 。 Boolean 类型，可读写。                                                                       |
| 144  | [SaveSubsetFonts](#Document.SaveSubsetFonts)                                   | 如果 WPS 将嵌入的 TrueType 字体子集和文档一起保存，则该属性值为 True 。 Boolean 类型，可读写。                                                                                       |
| 145  | [Scripts](#Document.Scripts)                                                   | 返回一个 Scripts 集合，该集合代表指定对象中的 HTML 脚本的集合。                                                                                                                      |
| 146  | [Sections](#Document.Sections)                                                 | 返回一个 Section 集合，该集合代表指定文档中的节。只读。                                                                                                                              |
| 147  | [Sentences](#Document.Sentences)                                               | 返回一个 Sentences 集合，该集合代表文档中的所有句子。只读。                                                                                                                          |
| 148  | [ServerPolicy](#Document.ServerPolicy)                                         | 返回一个 ServerPolicy 对象，该对象代表为运行 SharePoint Server 2007 的服务器上存储的文档指定的策略。只读。                                                                           |
| 149  | [Shapes](#Document.Shapes)                                                     | 返回一个 Shapes 集合，该集合代表指定文档中的所有 Shape 对象。只读。                                                                                                                  |
| 150  | [ShowGrammaticalErrors](#Document.ShowGrammaticalErrors)                       | 如果指定文档中的语法错误用绿色波浪线标记，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                |
| 151  | [ShowRevisions](#Document.ShowRevisions)                                       | 如果在屏幕上显示对指定文档的修订，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                        |
| 152  | [ShowSpellingErrors](#Document.ShowSpellingErrors)                             | 如果 WPS 为文档中的拼写错误添加下划线，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                   |
| 153  | [Signatures](#Document.Signatures)                                             | 返回一个 SignatureSet 集合，该集合代表文档的数字签名。                                                                                                                               |
| 154  | [SmartDocument](#Document.SmartDocument)                                       | 返回一个 SmartDocument 对象，该对象代表智能文档解决方案的设置。                                                                                                                      |
| 155  | [SmartTagsAsXMLProps](#Document.SmartTagsAsXMLProps)                           | 如果该属性的值为 True ，当将包含智能标记的文档保存为 HTML 格式时，WPS 会创建一个包含智能标记信息的 XML 页眉。 Boolean 类型，可读写。                                                 |
| 156  | [SnapToGrid](#Document.SnapToGrid)                                             | 如果在指定文档中绘制、移动自选图形/东亚字符或者调整它们的大小时，自选图形/东亚字符自动与不可见的网格线对齐，则该属性为 True 。可读/写 Boolean 类型。                                 |
| 157  | [SnapToShapes](#Document.SnapToShapes)                                         | 如果 WPS 在指定文档中使自选图形或东亚字符自动与不可见网格线对齐，这些网格线穿过其他自选图形或东亚字符的水平和垂直边缘，则该属性为 True 。可读/写 Boolean 类型。                      |
| 158  | [SpellingChecked](#Document.SpellingChecked)                                   | 如果已对指定的区域或文档完成拼写检查，则该属性值为 True 。如果所有或部分区域或文档尚未进行拼写检查，则该属性值为 False 。 Boolean 类型，可读写。                                     |
| 159  | [SpellingErrors](#Document.SpellingErrors)                                     | 返回一个 ProofreadingErrors 集合，该集合代表在指定文档或区域中标识为拼写错误的单词。只读。                                                                                           |
| 160  | [StoryRanges](#Document.StoryRanges)                                           | 返回一个 StoryRanges 集合，该集合代表指定文档中的所有的文字部分。只读。                                                                                                              |
| 161  | [Styles](#Document.Styles)                                                     | 本示例返回一个指定文档的 Styles 集合。只读。                                                                                                                                         |
| 162  | [StyleSortMethod](#Document.StyleSortMethod)                                   | 返回或设置一个 WdStyleSort 常量，该常量代表要在对 “样式” 任务窗格中的样式进行排序时使用的排序方法。可读写。                                                                          |
| 163  | [Subdocuments](#Document.Subdocuments)                                         | 返回一个 Subdocuments 集合，该集合代表指定文档中的所有子文档。只读。                                                                                                                 |
| 164  | [Sync](#Document.Sync)                                                         | 返回一个 Sync 对象，该对象提供对构成“文档工作区”的文档的方法和属性的访问。                                                                                                           |
| 165  | [Tables](#Document.Tables)                                                     | 返回一个 Table 集合，该集合代表指定文档中的所有表格。只读。                                                                                                                          |
| 166  | [TablesOfAuthorities](#Document.TablesOfAuthorities)                           | 返回一个 TableOfAuthorities 集合，该集合代表指定文档中的引文目录。只读。                                                                                                             |
| 167  | [TablesOfAuthoritiesCategories](#Document.TablesOfAuthoritiesCategories)       | 返回一个 TablesOfAuthoritiesCategories 集合，该集合代表指定文档中可用的引文目录类别。只读。                                                                                          |
| 168  | [TablesOfContents](#Document.TablesOfContents)                                 | 返回一个 TablesOfContents 集合，该集合代表指定文档中的目录。只读。                                                                                                                   |
| 169  | [TablesOfFigures](#Document.TablesOfFigures)                                   | 返回一个 TablesOfFigures 集合，该集合代表指定文档中的图表目录。只读。                                                                                                                |
| 170  | [TextEncoding](#Document.TextEncoding)                                         | 返回或设置代码页或字符集，WPS 应用该代码页或字符集将文档保存为编码文本文件。 MsoEncoding 类型，可读写。                                                                              |
| 171  | [TextLineEnding](#Document.TextLineEnding)                                     | 返回或设置一个 WdLineEndingType 常量，该常量代表 WPS 在保存为文本文件的文档中标记直线和段落分隔符的方式。可读写。                                                                    |
| 172  | [TrackFormatting](#Document.TrackFormatting)                                   | 返回或设置 Boolean 值，该值表示在打开更改跟踪功能时是否跟踪格式更改。可读/写。                                                                                                       |
| 173  | [TrackMoves](#Document.TrackMoves)                                             | 返回或设置 Boolean 值，该值表示在打开跟踪更改功能时是否对移动文本进行标记。可读/写。                                                                                                 |
| 174  | [TrackRevisions](#Document.TrackRevisions)                                     | 如果标记对指定文档的修改，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                |
| 175  | [Type](#Document.Type)                                                         | 返回文档的类型（模板或文档）。只读 WdDocumentType 类型。                                                                                                                             |
| 176  | [UpdateStylesOnOpen](#Document.UpdateStylesOnOpen)                             | 如果在每次打开指定文档时会依照其附属的模板更新文档样式，则该属性值为 True 。 Boolean 类型，可读写。                                                                                  |
| 177  | [UseMathDefaults](#Document.UseMathDefaults)                                   | 返回或设置一个代表在创建新的公式时是否使用默认的数学设置的 Boolean 。可读写。                                                                                                        |
| 178  | [UserControl](#Document.UserControl)                                           | 如果文档是由用户创建或打开的，则该属性为 True 。可读/写 Boolean 类型。                                                                                                               |
| 179  | [Variables](#Document.Variables)                                               | 返回一个 Variables 集合，该集合代表指定文档所存的变量。只读。                                                                                                                        |
| 180  | [VBASigned](#Document.VBASigned)                                               | 如果指定文档的 示例代码 (VBA) 项目具有数字签名，则该属性值为 True 。 Boolean 类型，只读。                                                                                            |
| 181  | [VBProject](#Document.VBProject)                                               | 返回指定模板或文档的 VBProject 对象。                                                                                                                                                |
| 182  | [WebOptions](#Document.WebOptions)                                             | 将文档保存为网页或打开网页时，该属性将返回 WebOptions 对象，该对象包含 WPS 所用的文档属性。只读。                                                                                    |
| 183  | [Windows](#Document.Windows)                                                   | 返回一个 Windows 集合，该集合代表指定文档的所有窗口。只读。                                                                                                                          |
| 184  | [WordOpenXML](#Document.WordOpenXML)                                           | 返回一个 String 类型的值，该值代表文档的 WPS Open XML 内容的 Flat XML 格式 。只读。                                                                                                  |
| 185  | [Words](#Document.Words)                                                       | 返回一个 Words 集合，该集合代表文档中的所有字词。只读。                                                                                                                              |
| 186  | [WritePassword](#Document.WritePassword)                                       | 该属性设置一个保存对指定文档所做的修改时所需的密码。 String 类型，只写。                                                                                                             |
| 187  | [WriteReserved](#Document.WriteReserved)                                       | 如果指定文档设有写权限密码，则该属性值为 True 。 Boolean 类型，只读。                                                                                                                |
| 188  | [XMLHideNamespaces](#Document.XMLHideNamespaces)                               | 返回一个 Boolean 类型的值，该值表示是否隐藏 XML 命名空间，该命名空间位于“XML 结构”任务窗格的元素列表中。                                                                             |
| 189  | [XMLNodes](#Document.XMLNodes)                                                 | 返回一个 XMLNodes 集合，该集合代表文档中的所有 XML 元素。                                                                                                                            |
| 190  | [XMLSaveDataOnly](#Document.XMLSaveDataOnly)                                   | 设置或返回一个 Boolean 类型的值，保存文档时包含 XML 标记还是仅保存文本。                                                                                                             |
| 191  | [XMLSaveThroughXSLT](#Document.XMLSaveThroughXSLT)                             | 设置或返回一个 String 类型的值，指定用户保存文档时，Extensible Stylesheet Language Transformation (XSLT) 要应用的路径和文件名。                                                      |
| 192  | [XMLSchemaReferences](#Document.XMLSchemaReferences)                           | 返回一个 XMLSchemaReferences 集合，代表附加到文档的架构。                                                                                                                            |
| 193  | [XMLSchemaViolations](#Document.XMLSchemaViolations)                           | 返回一个 XMLNodes 集合，代表文档中所有存在验证错误的节点。                                                                                                                           |
| 194  | [XMLShowAdvancedErrors](#Document.XMLShowAdvancedErrors)                       | 返回或设置一个 Boolean 类型的值，该值表示是由内置的 WPS 错误消息还是由包括在 Office 中的 Microsoft XML Core Services (MSXML) 5.0 组件来生成错误消息文本。                            |
| 195  | [XMLUseXSLTWhenSaving](#Document.XMLUseXSLTWhenSaving)                         | 返回一个 Boolean 类型的值，代表是否通过 Extensible Stylesheet Language Transformation (XSLT) 保存文档。如果通过 XSLT 保存文档，则该值为 True 。                                      |

## 成员方法

### Document.AcceptAllRevisions

接受对指定文档的所有修订。

**语法**

**express.AcceptAllRevisions ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 本示例检查文档是否有修订，有的话接受所有修订 */
function test() {
    if(Application.ActiveDocument.Revisions.Count >= 1){
        Application.ActiveDocument.AcceptAllRevisions()
    }   
}
```

### Document.AcceptAllRevisionsShown

接受显示在屏幕上的指定文档中的所有修订。

**语法**

**express.AcceptAllRevisionsShown ()**

\*express是一个代表 Document 对象的变量。

**说明**

可用 **RejectAllRevisionsShown** 方法来拒绝显示在屏幕上的指定文档中的所有修订。

### Document.Activate

激活指定的文档，使其成为活动文档。

**语法**

**express.Activate ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 激活第一个文档 */
Application.Documents.Item(1).Activate()
```

### Document.AddToFavorites

创建文档或超链接的快捷方式，并将其添加到“收藏夹”。

**语法**

**express.AddToFavorites ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档的所有超链接创建快捷方式，并将其添加到“收藏夹”文件夹。*/
function test() {
    for (let i = 1; i <= Application.ActiveDocument.Hyperlinks.Count; i++) {
        Application.ActiveDocument.Hyperlinks.Item(i).AddToFavorites()
    }
}
```

``` JavaScript
/*本示例为 Sales.doc 创建快捷方式，并将其添加至“收藏夹”文件夹。如果 Sales.doc 还未打开，本示例将从 C:\Documents 文件夹打开该文档。*/
function test() {
    let doc
    for (let i = 1; i <= Application.Documents.Count; i++) {
        doc = Application.Documents.Item(i)
        if (doc.Name.toLowerCase() == "sales.doc") {
            isOpen = true
        }
    }
    if (isOpen != true) {
        Application.Documents.Open("C:\\Documents\\Sales.doc")
    }
    Application.Documents.Item("Sales.doc").AddToFavorites()
}
```

### Document.ApplyQuickStyleSet2

将指定的快速样式集应用于文档。

**语法**

**express.ApplyQuickStyleSet2 ( *Style* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                 |
|---------|-----------|----------|----------------------------------------------------------------------------------------------------------------------|
| *Style* | 必选      | Variant  | 可以是指定要使用的集合名称（与“样式集”列表中列出的名称对应）的 String，也可以是 WdApplyQuickStyleSets 枚举中的常量。 |

### Document.ApplyTheme

将 主题 应用于打开的文档。

**语法**

**express.ApplyTheme ( *Name* )**

\*express是一个代表 Document 对象的变量。

**参数**

<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th>名称</th>
<th>必选/可选</th>
<th>数据类型</th>
<th>说明</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td class="mainsection"><em>Name</em></td>
<td class="mainsection">必选</td>
<td class="mainsection">String</td>
<td class="mainsection" style="word-break: break-all">主题名称及要应用的任何主题格式选项。该字符串格式为“themennn”，其中 theme 和 nnn 的定义如下：
主题: 包含请求主题的数据的文件夹名称（主题数据文件夹的默认位置是 C:\Program Files\Common Files\WPS Shared\Themes）。主题必须使用该文件夹名，不能使用“主题”对话框中显示的名称。
nnn: 三个数字的字符串，表示要激活的主题格式选项（1 为激活，0 为停用）。每个数字分别对应“主题”对话框中的“鲜艳颜色”、“动态图形”和“背景图像”复选框。如果省略该字符串，则 nnn 的默认值为“011”（“动态图形”和“背景图像”均被激活）。</td>
</tr>
</tbody>
</table>

**示例**

``` JavaScript
/*本示例实现的功能是：向活动文档添加“Artsy”主题，并激活“鲜艳颜色”选项。*/
Application.ActiveDocument.ApplyTheme("artsy 100")
```

### Document.AutoFormat

自动给文档套用格式。

**语法**

**express.AutoFormat ()**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **Kind** 属性指定文档类型。

### Document.CanCheckin

如果该方法返回值为 **True** ，则 WPS 可将指定文档签入服务器。可读/写 **Boolean** 类型。

**语法**

**express.CanCheckin ()**

\*express是一个代表 Document 对象的变量。

**返回值**

Boolean

**说明**

要利用 WPS 中内置的协作功能，文档必须存储在 Microsoft SharePoint Portal Server 上。

**示例**

``` JavaScript
/*本示例检查服务器并确定指定的文档是否可以签入；如果可以，则关闭文档并将其签回服务器。*/
function CheckInOut(docCheckIn) {
    if (Application.Documents.Item(docCheckIn).CanCheckin() == true) {
        Application.Documents.Item(docCheckIn).CheckIn()
        alert(docCheckIn + " has been checked in.")
    }
    else {
        alert("This file cannot be checked in " + "at this time.  Please try again later.")
    }
}
```

``` JavaScript
/*若要调用上述 CheckInOut 子程序，请使用下列子程序并将文件名“http://servername/workspace/report.doc”替换为以上“说明”一节中提到的服务器上实际的文件。*/
function CheckDocInOut(){
    CheckInOut ("http://servername/workspace/report.doc")
}
```

### Document.CheckConsistency

搜索日语文档中的所有文字，并显示对相同单词的字符用法不一致的实例。

**语法**

**express.CheckConsistency ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例检查活动文档中日语字符的一致性。*/
ActiveDocument.CheckConsistency()
```

### Document.CheckGrammar

开始对指定的文档或区域执行拼写和语法检查。

**语法**

**express.CheckGrammar ()**

\*express是一个代表 Document 对象的变量。

**说明**

如果文档或区域中包含错误，则该方法显示 **“拼写和语法”** 对话框，其中 **“检查语法”** 复选框处于选定状态。该方法应用于文档时，将检查文档中所有可用文字的拼写（如页眉、页脚和文本框）。

**示例**

``` JavaScript
/*本示例开始对活动文档中的所有文字执行拼写和语法检查。*/
ActiveDocument.CheckGrammar()
```

### Document.CheckIn

该方法将文档从本地计算机返回到服务器，并将本地文档设为只读，使之无法在本地进行编辑。

**语法**

**express.CheckIn ( *SaveChanges* , *MakePublic* , *Comments* , *VersionType* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                  |
|---------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *SaveChanges* | 可选      | Boolean  | 如果该属性值为 True，则将文档保存到服务器位置。默认值为 True。                                                                                                                        |
| *MakePublic*  | 可选      | Boolean  | 如果该属性值为 True，则允许用户在文档签入后进行文档发布。这将提交文档以进行批准，并最终生成发布给用户的一个文档版本，用户对该文档只具有可读权限（仅当 SaveChanges 为 True 时才应用)。 |
| *Comments*    | 可选      | Variant  | 正在签入的文档的修订的批注（仅当 SaveChanges 为 True 时才应用）。                                                                                                                     |
| *VersionType* | 可选      | Variant  | 指定文档的版本信息。                                                                                                                                                                  |

**说明**

要利用 WPS 中内置的协作功能，文档必须存储在 Microsoft SharePoint Portal Server 上。

**示例**

``` JavaScript
/*本示例检查服务器并确定指定的文档是否可以签入。如果可以，则保存并关闭文档，然后将文档签回服务器。*/
function CheckInOut(docCheckIn){
  if(Documents.Item(docCheckIn).CanCheckin() == true){
    Documents.Item(docCheckIn).CheckIn()
    MsgBox(docCheckIn + " has been checked in.")
  }
  else{
    MsgBox("This file cannot be checked in at this time.  Please try again later.")
  }
}
```

``` JavaScript
/*若要调用 CheckInOut 子程序，请使用下列子程序并将“http://servername/workspace/report.doc”替换为以上“注解”一节中提到的服务器上实际文件的文件名。*/
function CheckDocInOut(){
    CheckInOut("http://servername/workspace/report.doc")
}
```

### Document.CheckInWithVersion

将文档从本地计算机保存到服务器，并将本地文档设为只读，使之无法在本地进行编辑。

**语法**

**express.CheckInWithVersion ( *SaveChanges* , *Comments* , *MakePublic* , *VersionType* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                  |
|---------------|-----------|----------|-----------------------------------------------------------------------|
| *SaveChanges* | 可选      | Boolean  | 如果值为 True，则将文档保存到服务器位置。默认值为 True。              |
| *Comments*    | 可选      | Variant  | 正在签入的文档的修订批注（仅适用于 SaveChanges 设置为 True 的情况）。 |
| *MakePublic*  | 可选      | Boolean  | 如果值为 True，则允许用户在文档签入后发布文档。                       |
| *VersionType* | 可选      | Variant  | 指定文档的版本信息。                                                  |

**说明**

如果将 *MakePublic* 参数设置为 **True** ，则会提交文档以进行审批，审批过程最终产生的文档版本将发布给对该文档具有只读权限的用户（仅适用于 *SaveChanges* 设置为 **True** 的情况）。

要利用 WPS 中内置的协作功能，文档必须存储在 Microsoft SharePoint Server 上。

**示例**

``` JavaScript
/*以下示例使用 CanCheckin 方法确定文档是否已存储在 Microsoft SharePoint Server 上。如果文档已存储在服务器上，则此示例调用 CheckInWithVersion 方法签入文档以及指定的批注和版本号、将更改保存到服务器位置并提交文档以进行审批。 

此示例适用于文档级自定义。 
*/
function DocumentCheckIn() {
    if (Application.ActiveDocument.CanCheckin()) {
        Application.ActiveDocument.CheckInWithVersion(true, "My updates.", true, wdCheckInMinorVersion)
    }
    else {
        alert("This document cannot be checked in")
    }
}
```

### Document.CheckSpelling

开始对指定的文档或区域进行拼写检查。

**语法**

**express.CheckSpelling ( *IgnoreUppercase* , *AlwaysSuggest* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型 | 说明                                                                                                             |
|-------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------|
| *IgnoreUppercase* | 可选      | Variant  | 如果忽略大小写，则该属性值为 True。如果省略该参数，则使用 IgnoreUppercase 属性的当前值。                         |
| *AlwaysSuggest*   | 可选      | Variant  | 如果该属性值为 True，则 WPS 总是给出建议的拼写。如果省略该参数，则使用 SuggestSpellingCorrections 属性的当前值。 |

**说明**

如果该文档或区域中包含错误，则该方法显示 **“拼写和语法”** 对话框（ **“工具”** 菜单），其中 **“检查语法”** 复选框处于清除状态。该方法应用于文档时，将检查文档中所有可用文字的拼写（如页眉、页脚和文本框）。

**示例**

``` JavaScript
/*以下示例检查活动文档中的拼写。*/
Application.ActiveDocument.CheckSpelling()
```

### Document.Close

关闭指定的文档。

**语法**

**express.Close ( *SaveChanges* , *OriginalFormat* , *RouteDocument* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型         | 说明                                                                                                                 |
|------------------|-----------|------------------|----------------------------------------------------------------------------------------------------------------------|
| *SaveChanges*    | 可选      | WdSaveOptions    | 指定保存文档的操作。可以是下列 WdSaveOptions 常量之一：wdDoNotSaveChanges、wdPromptToSaveChanges 或 wdSaveChanges。  |
| *OriginalFormat* | 可选      | WdOriginalFormat | 指定保存文档的格式。可以是下列 WdOriginalFormat 常量之一：wdOriginalDocumentFormat、wdPromptUser 或 wdWordDocument。 |
| *RouteDocument*  | 可选      | Boolean          | 如果为 True，则将文档传送给下一个收件人。如果文档没有附加的传送名单，则忽略该参数。                                  |

**示例**

``` JavaScript
/* 本示例在关闭活动文档前提示用户保存该文档 */
Application.ActiveDocument.Close(wdPromptToSaveChanges,wdPromptUser)
```

### Document.ClosePrintPreview

将指定文档从打印预览视图切换到以前的视图。

**语法**

**express.ClosePrintPreview ()**

\*express是一个代表 Document 对象的变量。

**说明**

如果指定文档未处于打印预览视图，将出现错误。

**示例**

``` JavaScript
/*本示例将活动窗口从打印预览视图切换至普通视图。*/
function test() {
    if (Application.ActiveDocument.PrintPreview == true) {
        Application.ActiveDocument.ClosePrintPreview()
    }
    Application.ActiveDocument.ActiveWindow.View.Type = wdNormalView
}
```

### Document.Compare

显示修订标记，以表明指定的文档与另一个文档的区别。

**语法**

**express.Compare ( *Name* , *AuthorName* , *CompareTarget* , *DetectFormatChanges* , *IgnoreAllComparisonWarnings* , *AddToRecentFiles* , *RemovePersonalInformation* , *RemoveDateAndTime* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                          | 必选/可选 | 数据类型 | 说明                                                                                                                           |
|-------------------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------|
| *Name*                        | 可选      | String   | 与指定文档作比较的文档的名称。                                                                                                 |
| *AuthorName*                  | 可选      | Variant  | 与比较生成的区别相关联的审阅者姓名。如果没有指定，则默认值为修订的文档的作者姓名或字符串“Comparison”（如果没有提供作者信息）。 |
| *CompareTarget*               | 可选      | Variant  | 用于比较的目标文档。可以是任意 WdCompareTarget 常量。                                                                          |
| *DetectFormatChanges*         | 可选      | Boolean  | 如果该属性值为 True（默认值），则比较结果包括格式变化检测。                                                                    |
| *IgnoreAllComparisonWarnings* | 可选      | Variant  | 如果该属性值为 True，则比较文档，而不通知用户问题。默认值为 False。                                                            |
| *AddToRecentFiles*            | 可选      | Variant  | 如果该属性值为 True，则将文档添加到“文件”菜单上最近使用的文件列表中。                                                          |
| *RemovePersonalInformation*   | 可选      | Boolean  | 如果该属性值为 True，则从返回的 Document 对象的备注、修订和属性对话框中删除所有用户信息。默认值为 False。                      |
| *RemoveDateAndTime*           | 可选      | Boolean  | 如果该属性值为 True，则从返回的 Document 对象的跟踪更改中删除日期和时间戳记。默认值为 False。                                  |

**示例**

``` JavaScript
/*本示例将活动文档与位于 Draft 文件夹中名为“FirstRev.doc”的文档进行比较，并将比较结果区别置于一个新文档中。*/
function CompareDocument(){
  Application.ActiveDocument.Compare("C:\\Draft\\FirstRev.doc", null, wdCompareTargetNew)
}
```

### Document.Compatibility

如果为 **True** ，则启用由 *Type* 参数指定的兼容性选项。兼容性选项将影响 WPS 中文档的显示方式。 **Boolean** 类型，可读写。

**语法**

**express.Compatibility ( *Type* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型        | 说明         |
|--------|-----------|-----------------|--------------|
| *Type* | 必选      | WdCompatibility | 兼容性选项。 |

**说明**

由于您选择或安装的语言支持不同（例如美国英语），上述部分常量可能无法使用。

**示例**

``` JavaScript
/*本示例为活动文档启用“工具”菜单中“选项”对话框“兼容性”选项卡上的“取消硬分页符或分栏符前后的空格”选项。*/
Applicatiion.ActiveDocument.Compatibility(wdSuppressSpBfAfterPgBrk) = true
```

``` JavaScript
/*本示例在启用或关闭“不为悬挂式缩进添加自动制表位”选项之间进行切换。*/
Application.ActiveDocument.Compatibility(wdNoTabHangIndent) !Application.ActiveDocument.Compatibility(wdNoTabHangIndent)
```

### Document.ComputeStatistics

返回基于指定文档内容的统计值。 **Long** 类型。

**语法**

**express.ComputeStatistics ( *Statistic* , *IncludeFootnotesAndEndnotes* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                          | 必选/可选 | 数据类型    | 说明                                                                      |
|-------------------------------|-----------|-------------|---------------------------------------------------------------------------|
| *Statistic*                   | 必选      | WdStatistic | 要计算的统计值。                                                          |
| *IncludeFootnotesAndEndnotes* | 可选      | Variant     | 如果为 True，则统计时将包括脚注和尾注；如果省略该参数，则默认值为 False。 |

**说明**

由于您选择或安装的语言支持不同（例如美国英语），上述部分常量可能无法使用。

**示例**

``` JavaScript
/*本示例显示活动文档的字数（包括脚注）。*/
alert(Application.ActiveDocument.ComputeStatistics(wdStatisticWords, true) + " words")
```

### Document.Convert

将文件转换成最新的文件格式，并启用所有新功能。

**语法**

**express.Convert ()**

\*express是一个代表 Document 对象的变量。

### Document.ConvertAutoHyphens

将由自动断字功能创建的连字符转换为手动连字符。

**语法**

**express.ConvertAutoHyphens ()**

\*express是一个代表 Document 对象的变量。

### Document.ConvertNumbersToText

将指定文档中的列表编号和 LISTNUM 域更改为文本。

**语法**

**express.ConvertNumbersToText ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中的列表编号和 LISTNUM 域转换为文本。*/
Application.ActiveDocument.ConvertNumbersToText()
```

### Document.ConvertVietDoc

使用非默认的代码页将一篇越南语文档重新转换为 Unicode。

**语法**

**express.ConvertVietDoc ( *CodePageOrigin* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                             |
|------------------|-----------|----------|----------------------------------|
| *CodePageOrigin* | 必选      | Long     | 用于对文档进行编码的原始代码页。 |

**说明**

如果要使一篇文档在另一台计算机或平台上具有可浏览性的话，可以使用 **ConvertVietDoc** 方法。

**示例**

``` JavaScript
//本示例将活动文档从越南语 ABC 代码页转换为 Unicode。本示例假定活动文档是使用越南语 ABC 代码页进行编码。function ConvertToVietCodePage(){
  Application.ActiveDocument.ConvertVietDoc(5)
}
```

### Document.CopyStylesFromTemplate

将指定模板中的样式复制到某篇文档。

**语法**

**express.CopyStylesFromTemplate ( *Template* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明         |
|------------|-----------|----------|--------------|
| *Template* | 必选      | String   | 模板文件名。 |

**说明**

在将某个模板的样式复制到文档时，文档中同名的样式将重新定义为与模板中描述的样式一致。模板的特有样式将复制到文档，文档特有的样式则保持不变。

**示例**

``` JavaScript
/*本示例将活动文档的模板的样式复制到文档。*/
Application.ActiveDocument.CopyStylesFromTemplate(ActiveDocument.AttachedTemplate.FullName)
```

``` JavaScript
/*本示例将模板 Sales96.dot 的样式复制到 Sales.doc。*/
Application.Documents.Item("Sales.doc").CopyStylesFromTemplate ("C:\\WPSOffice\\Templates\\Sales96.dot")
```

### Document.CountNumberedItems

返回指定 **Document** 对象中项目符号项或编号项以及 LISTNUM 域的数目。

**语法**

**express.CountNumberedItems ( *NumberType* , *Level* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                          |
|--------------|-----------|----------|-------------------------------------------------------------------------------|
| *NumberType* | 可选      | Variant  | 要计数的编号类型。可以是 WdNumberType 常量之一。默认值为 wdNumberAllNumbers。 |
| *Level*      | 可选      | Variant  | 与要计数的编号级别相对应的数。如果省略该参数，将计算所有级别的数目。          |

**说明**

将 *NumberType* 指定为 **wdNumberParagraph** 或 **wdNumberAllNumbers** （默认值）时，将计算项目符号项的数目。

有两种编号类型：一种是预设编号 ( **wdNumberParagraph** )，这种编号可通过在 **“项目符号和编号”** 对话框中选择模板来添加到段落中；另一种是 LISTNUM 域 ( **wdNumberListNum** )，这种编号允许为每段添加多个编号。

### Document.CreateLetterContent

生成并返回一个 **LetterContent** 对象，该对象基于指定信函内容。 **LetterContent** 对象。

**语法**

**express.CreateLetterContent ( *DateFormat* , *IncludeHeaderFooter* , *PageDesign* , *LetterStyle* , *Letterhead* , *LetterheadLocation* , *LetterheadSize* , *RecipientName* , *RecipientAddress* , *Salutation* , *SalutationType* , *RecipientReference* , *MailingInstructions* , *AttentionLine* , *Subject* , *CCList* , *ReturnAddress* , *SenderName* , *Closing* , *SenderCompany* , *SenderJobTitle* , *SenderInitials* , *EnclosureNumber* , *InfoBlock* , *RecipientCode* , *RecipientGender* , *ReturnAddressShortForm* , *SenderCity* , *SenderCode* , *SenderGender* , *SenderReference* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型             | 说明                                                           |
|--------------------------|-----------|----------------------|----------------------------------------------------------------|
| *DateFormat*             | 必选      | String               | 信函的日期。                                                   |
| *IncludeHeaderFooter*    | 必选      | Boolean              | 如果该属性值为 True，则页面设计模板中将包括页眉和页脚。        |
| *PageDesign*             | 必选      | String               | 附加到文档的模板的名称。                                       |
| *LetterStyle*            | 必选      | WdLetterStyle        | 文档版式。                                                     |
| *Letterhead*             | 必选      | String               | 如果该属性值为 True，则为预先打印的信头保留空间。              |
| *LetterheadLocation*     | 必选      | WdLetterheadLocation | 预先打印信头的位置。                                           |
| *LetterheadSize*         | 必选      | Single               | 为预先打印的信头保留的空间（以磅为单位）。                     |
| *RecipientName*          | 必选      | String               | 收信人的姓名。                                                 |
| *RecipientAddress*       | 必选      | String               | 收信人的地址。                                                 |
| *Salutation*             | 必选      | String               | 信函的称呼文本。                                               |
| *SalutationType*         | 必选      | WdSalutationType     | 信函称呼的类型。                                               |
| *RecipientReference*     | 必选      | String               | 信函的参考词组（例如，“In reply to”；）。                      |
| *MailingInstructions*    | 必选      | String               | 信函邮寄格式（例如，“Certified Mail”）。                       |
| *AttentionLine*          | 必选      | String               | 信函注意用语（例如，“Attention”；）。                          |
| *Subject*                | 必选      | String               | 指定信函的主题文本。                                           |
| *CCList*                 | 必选      | String               | 信函副本收件人的姓名（转寄人）。                               |
| *ReturnAddress*          | 必选      | String               | 回信的通讯地址。                                               |
| *SenderName*             | 必选      | String               | 发信人姓名。                                                   |
| *Closing*                | 必选      | String               | 信函的结束语。                                                 |
| *SenderCompany*          | 必选      | String               | 写信人所在公司的名称。                                         |
| *SenderJobTitle*         | 必选      | String               | 写信人的职务。                                                 |
| *SenderInitials*         | 必选      | String               | 写信人的姓名缩写。                                             |
| *EnclosureNumber*        | 必选      | Long                 | 信函的附件数。                                                 |
| *InfoBlock*              | 可选      | Variant              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *RecipientCode*          | 可选      | Variant              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *RecipientGender*        | 可选      | Variant              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *ReturnAddressShortForm* | 可选      | Variant              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *SenderCity*             | 可选      | Variant              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *SenderCode*             | 可选      | Variant              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *SenderGender*           | 可选      | Variant              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *SenderReference*        | 可选      | Variant              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |

**示例**

``` JavaScript
/*本示例实现的功能是：利用 CreateLetterContent 方法在活动文档中新建 LetterContent 对象，然后用 RunLetterWizard 方法使用该对象。*/
function test() {
    let myLetter = Application.ActiveDocument.CreateLetterContent("July 31,1996", false, "", wdFullBlock, true, wdLetterTop, InchesToPoints(1.5), "Dave Edson", "436 SE Main St." + "\r" + "Bellevue, WA 98004", "Dear Dave,", +
        wdSalutationInformal, "", "", "", "End of year report", "", "", "", "Sincerely yours,", "", "", "", 0)
    Application.ActiveDocument.RunLetterWizard(myLetter)
}
```

### Document.DataForm

显示 **“数据表单”** 对话框，在此对话框中可添加、删除或修改记录。

**语法**

**express.DataForm ()**

\*express是一个代表 Document 对象的变量。

**说明**

可在邮件合并主文档、邮件合并数据源或其他含有以表格单元格或分隔符分开的数据的文档中应用此方法。

**示例**

``` JavaScript
/*当活动文档为邮件合并文档时，本示例将显示“数据表单”对话框。*/
function test() {
    if (Application.ActiveDocument.MailMerge.State != wdNormalDocument) {
        Application.ActiveDocument.DataForm()
    }
}
```

``` JavaScript
/*本示例在一个新文档中创建一个表格并显示“数据表单”对话框。*/
function test() {
    let aDoc = Application.Documents.Add()
    aDoc.Tables.Add(aDoc.Content, 2, 2)
    aDoc.Tables.Item(1).Cell(1, 1).Range.Text = "Name"
    aDoc.Tables.Item(1).Cell(1, 2).Range.Text = "Age"
    aDoc.DataForm()
}
```

### Document.DeleteAllComments

删除文档中 **Comments** 集合内的所有备注。

**语法**

**express.DeleteAllComments ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 删除当前文档所有批注 */
Application.ActiveDocument.DeleteAllComments()
```

### Document.DeleteAllCommentsShown

删除显示在屏幕上的指定文档的所有修订。

**语法**

**express.DeleteAllCommentsShown ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 删除当前文档所有显示的批注 */
Application.ActiveDocument.DeleteAllCommentsShown()
```

### Document.DeleteAllEditableRanges

删除（指定用户或用户组有权修改其权限的）所有区域中的权限。

**语法**

**express.DeleteAllEditableRanges ( *EditorID* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                                           |
|------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *EditorID* | 可选      | Variant  | 可以是代表用户电子邮件别名（如果在相同的域中）和电子邮件地址的 String 类型值，或者是代表用户组的 WdEditorType 常量。如果省略该参数，则不删除文档中的任何权限。 |

**说明**

还可以使用 **DeleteAll** 方法来删除（指定用户或用户组有权修改其权限的）所有区域中的权限。

**示例**

``` JavaScript
/*以下示例删除当前用户在所有区域中的所有权限。*/
Application.ActiveDocument.DeleteAllEditableRanges(wdEditorCurrent)
```

### Document.DeleteAllInkAnnotations

删除文档中所有手写墨迹注释。

**语法**

**express.DeleteAllInkAnnotations ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*要处理墨迹注释，必须在平板电脑上运行 WPS 。*/
Application.ActiveDocument.DeleteAllInkAnnotations()
```

### Document.DetectLanguage

分析指定文本，以确定书写文本的语言类型。

**语法**

**express.DetectLanguage ()**

\*express是一个代表 Document 对象的变量。

**说明**

应用于 **Document** 对象时， **DetectLanguage** 方法将检查文档中所有可用文本（页眉、页脚、文本框等）。如果指定文本包含了某个句子的一部分，则选定内容或区域将扩展到该句的句末。

如果指定文本已应用了 **DetectLanguage** 方法，则 **LanguageDetected** 属性将设置为 **True** 。要重新检测指定文本的语言，必须先将 **LanguageDetected** 属性设置为 **False** 。

**示例**

``` JavaScript
/*本示例检查活动文档，以确定其所用的语言类型并显示检查结果。*/
function test() {
    if (Application.ActiveDocument.LanguageDetected == true) {
        alert("This document has already ")
        Application.ActiveDocument.LanguageDetected = false
        Application.ActiveDocument.DetectLanguage()

    }
    else {
        Application.ActiveDocument.DetectLanguage()
    }
    if (Application.ActiveDocument.Range.LanguageID == wdEnglishUS) {
        alert("This is a U.S. English document.")
    }
    else {
        alert("This is not a U.S. English document.")
    }
}
```

### Document.DowngradeDocument

将文档降级为 WPS 97-2003 文档格式，以便可以在以前版本的 WPS 中进行编辑。

**语法**

**express.DowngradeDocument ()**

\*express是一个代表 Document 对象的变量。

### Document.EndReview

终止审阅使用 **SendForReview** 方法发送以进行审阅的文件，或通过电子邮件将文档发送给另一用户而自动进行审阅的文件。

**语法**

**express.EndReview ()**

\*express是一个代表 Document 对象的变量。

**说明**

当执行 **EndReview** 方法时将显示一笑消息，询问用户是否要终止审阅。

**示例**

``` JavaScript
/*本示例终止对活动文档的审阅，本示例假定活动文档处于审阅循环中。*/
function EndDocRev(){
  Application.ActiveDocument.EndReview()
}
```

### Document.ExportAsFixedFormat

将文档保存为 PDF 或 XPS 格式。

**语法**

**express.ExportAsFixedFormat ( *OutputFileName* , *ExportFormat* , *OpenAfterExport* , *OptimizeFor* , *Range* , *From* , *To* , *Item* , *IncludeDocProps* , *KeepIRM* , *CreateBookmarks* , *DocStructureTags* , *BitmapMissingFonts* , *UseISO19005_1* , *FixedFormatExtClassPtr* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型                | 说明                                                                                                                                                                                                                     |
|--------------------------|-----------|-------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *OutputFileName*         | 必选      | String                  | 新的 PDF 或 XPS 文件的路径和文件名。                                                                                                                                                                                     |
| *ExportFormat*           | 可选      | WdExportFormat          | 指定采用 PDF 格式或 XPS 格式。                                                                                                                                                                                           |
| *OpenAfterExport*        | 可选      | Boolean                 | 导出内容后打开新文件。                                                                                                                                                                                                   |
| *OptimizeFor*            | 可选      | WdExportOptimizeFor     | 指定是针对屏幕显示还是打印进行优化。                                                                                                                                                                                     |
| *Range*                  | 可选      | WdExportRange           | 指定导出区域是整个文档、当前页面、文本区域还是当前所选内容。默认情况下导出整个文档。                                                                                                                                     |
| *From*                   | 可选      | Long                    | 如果 Range 参数设置为 wdExportFromTo，则指定起始页码。                                                                                                                                                                   |
| *To*                     | 可选      | Long                    | 如果 Range 参数设置为 wdExportFromTo，则指定结束页码。                                                                                                                                                                   |
| *Item*                   | 可选      | WdExportItem            | 指定导出过程是只包括文本还是包括文本和标记。                                                                                                                                                                             |
| *IncludeDocProps*        | 可选      | Boolean                 | 指定在最新导出的文件中是否包括文档属性。                                                                                                                                                                                 |
| *KeepIRM*                | 可选      | Boolean                 | 指定在源文档受 IRM 保护的情况下，是否将 IRM 权限复制到 XPS 文档。默认值为 True。                                                                                                                                         |
| *CreateBookmarks*        | 可选      | WdExportCreateBookmarks | 指定是否导出书签和要导出的书签的类型。                                                                                                                                                                                   |
| *DocStructureTags*       | 可选      | Boolean                 | 指定是否包含额外数据（如有关内容的流和逻辑组织的信息）以便为屏幕读取器提供帮助。默认值为 True。                                                                                                                          |
| *BitmapMissingFonts*     | 可选      | Boolean                 | 指定是否包含文本的位图。当字体许可不允许在 PDF 文件中嵌入某一字体时，将此参数设置为 True。如果为 False，则引用该字体，且查看者的计算机会替换适当的字体（如果此创作的字体不可用）。默认值为 True。                        |
| *UseISO19005_1*          | 可选      | Boolean                 | 指定是否将 PDF 的使用范围限制在按照 ISO 19005-1 标准化的 PDF 子集。如果为 True，生成的文件会更加可靠地自我包含，但由于受到格式的限制，可能会导致该文件较大或显示更多的可视项目。默认值为 False。                         |
| *FixedFormatExtClassPtr* | 可选      | Variant                 | 指定一个指针以指向一个允许对代码的备用实现进行调用的加载项。代码的备用实现将对应用程序生成的 EMF 和 EMF+ 页面描述进行解释，以生成其自身的 PDF 或 XPS。有关详细信息，请参阅 MSDN 上的“扩展 WPS (2015) 固定格式导出功能”。 |

### Document.FitToPages

适当缩小文本的字体大小，使文档的页数减少。

**语法**

**express.FitToPages ()**

\*express是一个代表 Document 对象的变量。

**说明**

如果 WPS 无法将文档的页数减少一页，将出现错误。

**示例**

``` JavaScript
/*本示例试图将活动文档的页数减少一页。*/
function test() {
    try {
        Application.ActiveDocument.FitToPages()
    }
    catch (exception) {
        alert("Fit to pages failed")
    }
}
```

``` JavaScript
/*本示例试图将每一打开文档的页数减少一页。*/
function test() {
    for (let i = 1; i <= Application.Documents.Count; i++) {
        Application.Documents.Item(i).FitToPages()
    }
}
```

### Document.FollowHyperlink

显示一个缓存文档（如果该文档已下载）。否则，此方法将解析该超链接，下载目标文档，并在相应的应用程序中显示此文档。

**语法**

**express.FollowHyperlink ( *Address* , *SubAddress* , *NewWindow* , *AddHistory* , *ExtraInfo* , *Method* , *HeaderInfo* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                  |
|--------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Address*    | 必选      | String   | String 目标文档的地址。                                                                                                                                                                                                               |
| *SubAddress* | 可选      | Variant  | Variant 目标文档中的位置。默认值为一个空字符串。                                                                                                                                                                                      |
| *NewWindow*  | 可选      | Variant  | 如果该属性值为 True，则在新窗口中显示目标位置。默认值为 False。                                                                                                                                                                       |
| *AddHistory* | 可选      | Variant  | 如果该属性值为 True，则向当前的历史记录文件夹中加入链接。                                                                                                                                                                             |
| *ExtraInfo*  | 可选      | Variant  | 指定用于解析超链接的 HTTP 附加信息的字符串或字节数组。例如，您可以使用 ExtraInfo 来指定图像映射的坐标、窗体的内容或 FAT 文件的名称。根据 Method 的值来决定是发布还是附加该字符串。可使用 ExtraInfoRequired 属性确定是否需要额外信息。 |
| *Method*     | 可选      | Variant  | 指定 HTTP 附加信息的处理方式。可以是任何 MsoExtraInfoMethod 常量。                                                                                                                                                                    |
| *HeaderInfo* | 可选      | Variant  | 指定 HTTP 请求的标题信息的字符串。默认值为一个空字符串。可以使用以下语法将几个标题行合并成一个字符串："string1" & vbCr & "string2"。指定的字符串自动转换为 ANSI 字符。请注意，HeaderInfo 参数可能会覆盖默认的 HTTP 标题字段。         |

**示例**

``` JavaScript
/*本示例根据指定的 URL 地址在新窗口显示 WPS 主页。*/
Application.ActiveDocument.FollowHyperlink("http://www.wps.cn", null, true, true)
```

``` JavaScript
/*本示例显示名为“Default.htm”的 HTML 文档。/
Application.ActiveDocument.FollowHyperlink("file:C:\\Pages\\Default.htm")
```

### Document.ForwardMailer

您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **ForwardMailer** 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。

**语法**

**express.ForwardMailer ()**

\*express是一个代表 Document 对象的变量。

### Document.FreezeLayout

在 Web 视图中，按照文档当前所显示的样子固定文档的版式，以使换行保持不变，且在调整窗口大小时墨迹不会移动。

**语法**

**express.FreezeLayout ()**

\*express是一个代表 Document 对象的变量。

### Document.GetCrossReferenceItems

返回一个项目数组，根据指定的交叉引用类型可对该数组中的项目进行交叉引用。

**语法**

**express.GetCrossReferenceItems ( *ReferenceType* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                              |
|-----------------|-----------|----------|-------------------------------------------------------------------|
| *ReferenceType* | 必选      | Variant  | 要在其中插入交叉引用的项目类型。可以是任何 WdReferenceType 常量。 |

**说明**

该方法返回的数组与 **“交叉引用”** 对话框的 **“引用哪一个”** 框中所列的项目相对应。该方法返回的值可以用作 **Range** 或 **Selection** 对象的 **InsertCrossReference** 方法的 *ReferenceWhich* 参数的值。

**示例**

``` JavaScript
/*本示例显示活动文档中第一个可以进行交叉引用的书签名称。*/
function test() {
    if (Application.ActiveDocument.Bookmarks.Count >= 1) {
        myBookmarks = Application.ActiveDocument.GetCrossReferenceItems(wdRefTypeBookmark)
        alert(myBookmarks(1))
    }
}
```

``` JavaScript
/*本示例首先用 GetCrossReferenceItems 方法检索可以进行交叉引用的标题的列表，然后在包含标题“Introduction”页中插入一条交叉引用。*/
function test() {
    let myHeadings = Application.ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
    for (let i = 0; i <= myHeadings.length - 1; i++) {
        if (myHeadings[i].toLowerCase().indexOf("introduction") != -1) {
            Application.Selection.InsertCrossReference(wdRefTypeHeading, wdPageNumber, i + 1)
            Application.Selection.InsertParagraphAfter()
        }
    }
}
```

### Document.GetLetterContent

检索指定文档中的信函内容并返回一个 **LetterContent** 对象。

**语法**

**express.GetLetterContent ()**

\*express是一个代表 Document 对象的变量。

**返回值**

LetterContent

**示例**

``` JavaScript
/*本示例显示活动文档信件中的问候语以及收信人的姓名。*/
alert(Application.ActiveDocument.GetLetterContent().Salutation + Application.ActiveDocument.GetLetterContent().RecipientName)
```

``` JavaScript
/*本示例从活动文档获取信函元素，并通过设置“CClist”属性来改变抄送人列表，然后用 SetLetterContent 方法来更新活动文档，以便反映这些更改。*/
function test() {
    let myLetterContent = Application.ActiveDocument.GetLetterContent()
    myLetterContent.CCList = "J. Burns, L. Scarpaczyk, K. Wong"
    myLetterContent.RecipientName = "Amy Anderson"
    myLetterContent.RecipientAddress = "123 Main" + "\r" + "Bellevue, WA  98004"
    myLetterContent.LetterStyle = wdFullBlock
    Application.ActiveDocument.SetLetterContent(myLetterContent)
}
```

### Document.GetWorkflowTasks

返回一个 **WorkflowTasks** 集合，该集合表示分配给文档的工作流任务。

**语法**

**express.GetWorkflowTasks ()**

\*express是一个代表 Document 对象的变量。

**返回值**

WorkflowTasks

### Document.GetWorkflowTemplates

返回一个 **WorkflowTemplates** 集合，该集合表示附加到文档的工作流模板。

**语法**

**express.GetWorkflowTemplates ()**

\*express是一个代表 Document 对象的变量。

**返回值**

WorkflowTemplates

### Document.GoTo

返回一个 **Range** 对象，该对象代表指定项（如页、书签或域）的起始位置。

**语法**

**express.GoTo ( *What* , *Which* , *Count* , *Name* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型        | 说明                                                                                                                                                                                                       |
|---------|-----------|-----------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *What*  | 可选      | WdGoToItem      | 指定区域或选定内容要移动到的项的类别。可以是下列 WdGoToItem 常量之一。                                                                                                                                     |
| *Which* | 可选      | WdGoToDirection | 指定区域或选定内容要移动到的项。可以是下列 WdGoToDirection 常量之一。                                                                                                                                      |
| *Count* | 可选      | Number          | 文档中的项数。默认值为 1。只有正值是有效的。若要指定一个在该区域或选定内容之前的项，可将 Which 参数指定为 wdGoToPrevious，并指定一个 Count 值。                                                            |
| *Name*  | 可选      | Variant         | 如果 What 参数是 wdGoToBookmark、wdGoToComment、wdGoToField 或 wdGoToObject，则此参数指定一个名称。若要指定一个在该区域或选定内容之前的项，可将 Which 参数指定为 wdGoToPrevious，并指定一个 Count 参数值。 |

**说明**

将 **GoTo** 方法与 **wdGoToGrammaticalError** 、 **wdGoToProofreadingError** 或 **wdGoToSpellingError** 常量一起使用时，返回的 **Range** 对象中包括所有含语法或拼写错误的文本。

**示例**

``` JavaScript
/* 切换到第一个作者名为"authorname"的备注开始位置 */
Application.ActiveDocument.GoTo(wdGoToComment,wdGoToFirst,1,"authorname").Select()

/* 选中当前文档的第一个脚注引用处 */
function test() {
    if (Application.ActiveDocument.Footnotes.Count >= 1) {
        let r1 = Application.ActiveDocument.GoTo(wdGoToFootnote, wdGoToFirst);
        r1.Select();
    }
}
```

### Document.LockServerFile

在服务器上锁定文件，以避免任何其他人进行编辑。

**语法**

**express.LockServerFile ()**

\*express是一个代表 Document 对象的变量。

**说明**

此方法仅适用于在运行 SharePoint Server 2007 的服务器上存储的文档。

### Document.MakeCompatibilityDefault

兼容性选项位于 **“选项”** 对话框的 **“兼容性”** 选项卡上，它们是新文档的默认设置。

**语法**

**express.MakeCompatibilityDefault ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例首先为活动文档设置兼容性选项，然后将当前选项设置为默认设置。*/
function test() {
    Application.ActiveDocument.Compatibility(wdSuppressSpBfAfterPgBrk)
    Application.ActiveDocument.Compatibility(wdExpandShiftReturn) 
    Application.ActiveDocument.Compatibility(wdUsePrinterMetrics)
    Application.ActiveDocument.Compatibility(wdNoLeading)
    Application.ActiveDocument.MakeCompatibilityDefault()
}
```

### Document.ManualHyphenation

启动文档的手动断字设置，每次一行。

**语法**

**express.ManualHyphenation ()**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **ManualHyphenation** 方法时， WPS 提示用户接受还是拒绝建议的断字。

**示例**

``` JavaScript
/*本示例启动活动文档的手动断字操作。*/
ActiveDocument.ManualHyphenation()
```

``` JavaScript
/*本示例先设置断字选项，然后对 MyDoc.doc 进行手动断字。*/
function test() {
    let rng = Application.Documents.Item("MyDoc.doc")
    rng.HyphenationZone = InchesToPoints(0.25)
    rng.HyphenateCaps = false
    rng.ManualHyphenation()
}
```

### Document.Merge

将文档中用修订标记标识的修改合并到另一篇文档。

**语法**

**express.Merge ( *Name* , *MergeTarget* , *DetectFormatChanges* , *UseFormattingFrom* , *AddToRecentFiles* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                  | 必选/可选 | 数据类型            | 说明                                                 |
|-----------------------|-----------|---------------------|------------------------------------------------------|
| *Name*                | 必选      | String              | 要合并的文档的路径和文件名。                         |
| *MergeTarget*         | 可选      | WdMergeTarget       | 指定放置最终合并的内容的位置。                       |
| *DetectFormatChanges* | 可选      | Boolean             | 指定是否标记格式差别。                               |
| *UseFormattingFrom*   | 可选      | WdUseFormattingFrom | 指定将哪篇文档用于设置合并文档的格式。               |
| *AddToRecentFiles*    | 可选      | Boolean             | 指定是否将 Name 参数中的文档添加到最近的文件列表中。 |

**示例**

``` JavaScript
/*本示例将在 Sales2.doc（活动文档）中所做的修改合并到 Sales1.doc。*/
function test() {
    if (Application.ActiveDocument.Name.indexOf("sales2.doc")) {
        Application.ActiveDocument.Merge("C:\\Docs\\Sales1.doc")
    }
}
```

### Document.Post

将指定的文档传送到 Microsoft Exchange 中的公用文件夹。

**语法**

**express.Post ()**

\*express是一个代表 Document 对象的变量。

**说明**

该方法显示 **“发送到 Exchange 文件夹”** 对话框，以便选择一个文件夹。

**示例**

``` JavaScript
/*本示例显示“发送到 Exchange 文件夹”对话框，以便将活动文档发送到公用文件夹。*/
Application.ActiveDocument.Post()
```

### Document.PresentIt

打开 WPP，并加载指定的 WPS 文档。

**语法**

**express.PresentIt ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例向 WPP 中发送名为“MyPresentation.doc”的文档的副本。*/
Application.Documents.Item("MyPresentation.doc").PresentIt()
```

### Document.PrintOut

打印指定文档的全部或部分内容。

**语法**

**express.PrintOut ( *Background* , *Append* , *Range* , *OutputFileName* , *From* , *To* , *Item* , *Copies* , *Pages* , *PageType* , *PrintToFile* , *Collate* , *FileName* , *ActivePrinterMacGX* , *ManualDuplexPrint* , *PrintZoomColumn* , *PrintZoomRow* , *PrintZoomPaperWidth* , *PrintZoomPaperHeight* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                 |
|------------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Background*           | 可选      | Variant  | 如果将该属性设置为 True，则 WPS 在打印文档时继续运行宏。                                                                                                                                                                                                                                             |
| *Append*               | 可选      | Variant  | 如果将该属性设置为 True，则将指定文档附加到 OutputFileName 参数所指定的文件名。如果将该属性设置为 False，则覆盖 OutputFileName 参数所指定文件的内容。                                                                                                                                                |
| *Range*                | 可选      | Variant  | 页码范围。可以是任意 WdPrintOutRange 常量。                                                                                                                                                                                                                                                          |
| *OutputFileName*       | 可选      | Variant  | 如果 PrintToFile 为 True，则该参数指定输出文件的路径和文件名。                                                                                                                                                                                                                                       |
| *From*                 | 可选      | Variant  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定起始页码。                                                                                                                                                                                                                                            |
| *To*                   | 可选      | Variant  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定结束页码。                                                                                                                                                                                                                                            |
| *Item*                 | 可选      | Variant  | 要打印的项目。可以是任意 WdPrintOutItem 常量。                                                                                                                                                                                                                                                       |
| *Copies*               | 可选      | Variant  | 要打印的份数。                                                                                                                                                                                                                                                                                       |
| *Pages*                | 可选      | Variant  | 要打印的页码和页码范围，中间用逗号分开。例如，“2, 6-10”表示打印第 2 页以及第 6 至第 10 页。                                                                                                                                                                                                          |
| *PageType*             | 可选      | Variant  | 要打印的页面类型。可以是任意 WdPrintOutPages 常量。                                                                                                                                                                                                                                                  |
| *PrintToFile*          | 可选      | Variant  | 如果该属性值为 True，则将打印指令发送到文件。请确保使用 OutputFileName 指定文件名。                                                                                                                                                                                                                  |
| *Collate*              | 可选      | Variant  | 在打印文档的多份副本时，如果该属性值为 True，则完成打印所有页面后再打印下一份副本。                                                                                                                                                                                                                  |
| *FileName*             | 可选      | Variant  | 要打印的文档的路径和文件名。如果省略该参数， WPS 将打印活动文档（仅应用于 Application 对象）。                                                                                                                                                                                                       |
| *ActivePrinterMacGX*   | 可选      | Variant  | 该参数仅适用于 WPS OfficeMacintosh Edition。有关该参数的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。                                                                                                                                                                            |
| *ManualDuplexPrint*    | 可选      | Variant  | 如果该属性值为 True，则在无双面打印组件的打印机上打印双面文档。如果该属性值为 True，则忽略 PrintBackground 和 PrintReverse 属性。使用 PrintOddPagesInAscendingOrder 和 PrintEvenPagesInAscendingOrder 属性可在手动双面打印时控制输出。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *PrintZoomColumn*      | 可选      | Variant  | 表示 WPS 在一页纸上水平放置的页数。可以是 1、2、3 或 4 页。与 PrintZoomRow 参数一同使用可在单张纸上打印多页文档。                                                                                                                                                                                    |
| *PrintZoomRow*         | 可选      | Variant  | 表示 WPS 在一页纸上垂直放置的页数。可以是 1、2 或 4 页。与 PrintZoomColumn 参数一同使用可在单张纸上打印多页文档。                                                                                                                                                                                    |
| *PrintZoomPaperWidth*  | 可选      | Variant  | WPS 将打印页面缩放到的宽度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。                                                                                                                                                                                                                             |
| *PrintZoomPaperHeight* | 可选      | Variant  | WPS 将打印页面缩放到的高度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。                                                                                                                                                                                                                             |

**示例**

``` JavaScript
/* 本示例打印活动文档的当前页面 */
Application.ActiveDocument.PrintOut(null, null, wdPrintCurrentPage)

/* 打印当前活动文档的第1-3页 */
Application.ActiveDocument.PrintOut(null,null,wdPrintFromTo,"","1","3")
```

### Document.PrintPreview

将视图切换到打印预览。

**语法**

**express.PrintPreview ()**

\*express是一个代表 Document 对象的变量。

**说明**

除了使用 **PrintPreview** 方法外，还可以将 **PrintPreview** 属性设置为 **True** 或 **False** ，前者切换到打印预览，后者从打印预览切回。还可以通过将 **View** 对象的 **Type** 属性设置为 **wdPrintPreview** 来更改视图。

**示例**

``` JavaScript
/* 切换到打印预览模式 */
Application.ActiveDocument.PrintPreview()
```

### Document.Protect

帮助保护指定的文档不被更改。保护文档后，用户只能进行有限的更改，例如添加注释，进行修订或填写表格。

**语法**

**express.Protect ( *Type* , *NoReset* , *Password* , *UseIRM* , *EnforceStyleLock* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                                                                        |
|--------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------|
| *Type*             | 必选      | Variant  | 指定文档的保护类型。 WdProtectionType 。                                                                                                    |
| *NoReset*          | 可选      | Variant  | False将表单字段重置为其默认值。如果指定的文档受保护，则为true以保留当前表单字段值。如果Type不是wdAllowOnlyFormFields，则将忽略NoReset参数。 |
| *Password*         | 可选      | Variant  | 从指定文档中删除保护所需的密码。 （请参阅下面的备注。）                                                                                     |
| *UseIRM*           | 可选      | Variant  | 指定在保护文档免受更改时是否使用信息权限管理（IRM）。                                                                                       |
| *EnforceStyleLock* | 可选      | Variant  | 指定是否在受保护的文档中实施格式限制。                                                                                                      |

### Document.Range

通过使用指定的开始和结束字符位置返回一个 **Range** 对象。

**语法**

**express.Range ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 本示例为活动文档中的前 10 个字符设置加粗格式 */
Application.ActiveDocument.Range(0, 10).Bold = true
```

### Document.Redo

重复最后一次撤消的操作（ **Undo** 方法的逆操作）。如果重复操作成功，则返回 **True** 。

**语法**

**express.Redo ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 将当前文档前两个字符加粗然后倾斜，然后撤销倾斜，撤销加粗，重做加粗，那么文字最终加粗，并没倾斜 */
function test() {
    let doc = Application.ActiveDocument;
    let rng = doc.Range(0,2);
    rng.Bold = true;
    rng.Italic = true;
    doc.Undo(2);
    doc.Redo(1);
}
```

### Document.RejectAllRevisions

拒绝指定文档的所有修订。

**语法**

**express.RejectAllRevisions ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 拒绝当前文档的所有修订 */
Application.ActiveDocument.RejectAllRevisions()
```

### Document.RejectAllRevisionsShown

拒绝显示在屏幕上的文档的所有修订。

**语法**

**express.RejectAllRevisionsShown ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 隐藏所有"authorname"做的修订，拒绝所有其他显示的修订 */
function test() {
    let rev = Application.ActiveWindow.View.Reviewers;
    rev.Item("authorname").Visible = false;
    Application.ActiveDocument.RejectAllRevisionsShown();
}
```

### Document.Reload

通过解析指向缓冲文档的超链接再下载该文档来重新加载该文档。

**语法**

**express.Reload ()**

\*express是一个代表 Document 对象的变量。

**说明**

该方法将同步重新加载文档，也就是说，在真正开始重新加载文档之前，可以执行此过程中 **Reload** 方法后紧跟的语句。因此，在宏中使用此方法将得到意想不到的结果。

**示例**

``` JavaScript
/*本示例打开并重新加载至本地内部网上地址为“main”的超链接。*/
function test() {
    Application.ActiveDocument.FollowHyperlink("http://main")
    Application.ActiveDocument.Reload()
}
```

### Document.ReloadAs

使用指定的编码重新加载一篇基于 HTML 文档的文档。

**语法**

**express.ReloadAs ( *Encoding* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型    | 说明                           |
|------------|-----------|-------------|--------------------------------|
| *Encoding* | 必选      | MsoEncoding | 指定重新加载文档时使用的编码。 |

**示例**

``` JavaScript
/*本示例用西里尔文编码重新加载当前文档。*/
Application.ActiveDocument.ReloadAs(msoEncodingCyrillic)
```

### Document.RemoveDocumentInformation

从文档中删除敏感性信息、属性、批注和其他元数据。

**语法**

**express.RemoveDocumentInformation ( *RemoveDocInfoType* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型            | 说明               |
|---------------------|-----------|---------------------|--------------------|
| *RemoveDocInfoType* | 必选      | WdRemoveDocInfoType | 指定要删除的内容。 |

### Document.RemoveLockedStyles

当文档中应用了格式设置限制时，清除锁定样式的文档。

**语法**

**express.RemoveLockedStyles ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*以下示例清除活动文档中的锁定样式。*/
Application.ActiveDocument.RemoveLockedStyles()
```

### Document.RemoveNumbers

删除指定文档中的编号或项目符号。

**语法**

**express.RemoveNumbers ( *NumberType* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型     | 说明               |
|--------------|-----------|--------------|--------------------|
| *NumberType* | 可选      | WdNumberType | 要删除的编号类型。 |

**示例**

``` JavaScript
/*本示例从头删除活动文档中已编号段落的编号。*/
Application.ActiveDocument.RemoveNumbers(wdNumberParagraph)
```

``` JavaScript
/*本示例删除 MyDocument.doc 中第三个列表的项目符号或编号。*/
function test() {
    if (Application.Documents.Item("MyDocument.doc").Lists.Count >= 3) {
        Application.Documents.Item("MyDocument.doc").Lists.Item(3).RemoveNumbers()
    }
}
```

### Document.RemoveTheme

删除当前文档中的活动 主题 。

**语法**

**express.RemoveTheme ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例将删除当前文档中的活动主题。*/
Application.ActiveDocument.RemoveTheme()
```

### Document.Repaginate

对整篇文档重新分页。

**语法**

**express.Repaginate ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*如果活动文档在上次保存后进行了更改，本示例将对其重新分页。*/
function test() {
    if (Application.ActiveDocument.Saved == false) {
        Application.ActiveDocument.Repaginate()
    }
}
```

``` JavaScript
/*本示例将对所有打开文档重新分页。*/
function test() {
    for (let i = 1; i <= Application.Documents.Count; i++) {
        Application.Documents.Item(i).Repaginate()
    }
}
```

### Document.Reply

打开一封新的电子邮件（ **“收件人”** 行已含发件人地址）以答复活动邮件。

**语法**

**express.Reply ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例打开一封新的电子邮件以答复活动邮件。*/
Application.MailMessage.Reply()
```

### Document.ReplyAll

打开一封新的电子邮件（ **“收件人”** 和 **“抄送”** 行已含发件人和所有其他收件人的地址）以答复活动邮件。

**语法**

**express.ReplyAll ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例打开一封新的电子邮件以答复活动邮件。*/
Application.MailMessage.ReplyAll()
```

### Document.ReplyWithChanges

为文档（该文档以被发送出去进行审阅）的作者发送电子邮件，通知他们审阅者已经审阅完文档。

**语法**

**express.ReplyWithChanges ( *ShowMessage* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                     |
|---------------|-----------|----------|----------------------------------------------------------------------------------------------------------|
| *ShowMessage* | 可选      | Variant  | 如果该属性值为 True，则发送前先显示邮件。如果属性值为 False，则自动发送邮件而不先显示它。默认值为 True。 |

**说明**

可用 **SendForReview** 方法开始对文档进行合作审阅。如果对不属于合作审阅阶段的文档执行 **ReplyWithChanges** 方法，则 WPS 会显示错误消息。

**示例**

``` JavaScript
/*本示例发送一条信息，通知作者审阅者已经完成审阅（而不先向审阅者显示该电子邮件）。本示例假定当前文档是合作审阅循环中的一部分。*/
function ReplyMsg() {
    Application.ActiveDocument.ReplyWithChanges(false)
}
```

### Document.ResetFormFields

清除文档中的所有窗体域，准备再次填充窗体。

**语法**

**express.ResetFormFields ()**

\*express是一个代表 Document 对象的变量。

**说明**

当文档中的域没有锁定时，可用 **ResetFormFields** 方法清除域。要在文档中的域处于锁定状态时清除域，请使用 **Protect** 方法。

**示例**

``` JavaScript
/*本示例清除活动文档中的域。本示例假定活动文档包含窗体域。*/
function ClearFormFields() {
    Application.ActiveDocument.ResetFormFields()
}
```

### Document.RunAutoMacro

运行存储于指定文档的自动宏。如果没有自动宏，则不做任何操作。

**语法**

**express.RunAutoMacro ( *Which* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明             |
|---------|-----------|--------------|------------------|
| *Which* | 必选      | WdAutoMacros | 要运行的自动宏。 |

**示例**

``` JavaScript
/*本示例运行活动文档中的 AutoOpen 宏。*/
Application.ActiveDocument.RunAutoMacro(wdAutoOpen)
```

### Document.RunLetterWizard

对指定的文档运行英文信函向导。

**语法**

**express.RunLetterWizard ( *LetterContent* , *WizardMode* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                         |
|-----------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *LetterContent* | 可选      | Variant  | 一个 LetterContent 对象。LetterContent 对象中的任何填充属性在“英文信函向导”对话框中显示为预填充元素。如果省略该参数，则 GetLetterContent 方法将自动用于从指定文档获取 LetterContent 对象。                                                                   |
| *WizardMode*    | 可选      | Variant  | Variant 如果该属性值为 True，则将“英文信函向导”对话框显示为一系列的步骤，并且提供“下一步”、“上一步”、“完成”按钮。如果该属性值为 False，则显示“英文信函向导”对话框就像从“工具”菜单上打开它一样（包含一个“确定”按钮和“取消”按钮的属性对话框）。默认值为 True。 |

**说明**

使用 **CreateLetterContent** 方法返回一个 **LetterContent** 对象，该对象给出各种信函内容属性。使用 **GetLetterContent** 方法返回一个基于指定文档内容的 **LetterContent** 对象。可以使用返回的 **LetterContent** 对象和 **RunLetterWizard** 方法来预先设置 **“英文信函向导”** 对话框中的元素。

**示例**

``` JavaScript
/*本示例创建一个新的 LetterContent 对象并设置该对象的若干属性，再使用 RunLetterWizard 方法运行信函向导。*/
function test() {
    let myContent = Application.ActiveDocument.GetLetterContent()
    myContent.Salutation = "Hello"
    myContent.SalutationType = wdSalutationOther
    myConent.SenderName = Application.UserName
    myContent.SenderInitials = Application.UserInitials
    Application.Documents.Add().RunLetterWizard(myContent, true)
}
```

``` JavaScript
/*下面的示例使用 CreateLetterContent 方法在活动文档中创建一个新的 LetterContent 对象，再使用该对象的 RunLetterWizard 方法。*/
function test() {
    let myLetter = Application.ActiveDocument.CreateLetterContent("July 31, 1999", false,
        +"C:\\WPSOffice\\Templates" + "\\Letters + Faxes\\Contemporary Letter.dot", wdFullBlock, true, wdLetterTop, InchesToPoints(1.5),
        +"Dave Edson", "436 SE Main St." + "\r" + "Bellevue, WA 98004", "Dear Dave,", wdSalutationInformal,
        + "", "", "", "End of year report", "", "", "", "Sincerely yours,", "", "", "", 0)
        Application.ActiveDocument.RunLetterWizard(myLetter)
}
```

### Document.Save

保存当前 文档。

**语法**

**express.Save ( *NoPrompt* , *OriginalFormat* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                      |
|------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------|
| *NoPrompt*       | 可选      | Variant  | 如果该属性值为 True，则 WPS 将自动保存所有文档。如果该属性值为 False， WPS 将提示用户保存自上次保存后已修改过的每篇文档。 |
| *OriginalFormat* | 可选      | Variant  | 指定文档的保存方式。可以是 WdOriginalFormat 常量之一。                                                                    |

**说明**

如果文档以前未保存过，则 **“另存为”** 对话框将提示用户键入文件名。

**示例**

``` JavaScript
/* 如果活动文档在上次保存后进行了更改，本示例将保存活动文档 */
function test() {
    if(Application.ActiveDocument.Saved == false) {
          Application.ActiveDocument.Save()
    }
}
```

### Document.SaveAs2

使用新的名称或格式保存指定的文档。此方法的一些参数与 **“另存为”** 对话框（ **“文件”** 选项卡）中的选项相对应。

**语法**

**express.SaveAs2 ( *FileName* , *FileFormat* , *LockComments* , *Password* , *AddToRecentFiles* , *WritePassword* , *ReadOnlyRecommended* , *EmbedTrueTypeFonts* , *SaveNativePictureFormat* , *SaveFormsData* , *SaveAsAOCELetter* , *Encoding* , *InsertLineBreaks* , *AllowSubstitutions* , *LineEnding* , *AddBiDiMarks* , *CompatibilityMode* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                      | 必选/可选 | 数据类型 | 说明                                                                                                                                                         |
|---------------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName*                | 可选      | Variant  | 文档的名称。默认值为当前文件夹和文件名。如果从未保存过文档，将使用默认名称（例如，Doc1.doc）。如果已经存在具有指定文件名的文档，将覆盖该文档，而不提示用户。 |
| *FileFormat*              | 可选      | Variant  | 文档的保存格式。可以是任何 WdSaveFormat 常量。若要以另一种格式保存文档，请为 FileConverter 对象的 SaveFormat 属性指定适当的值。                              |
| *LockComments*            | 可选      | Variant  | 如果该属性值为 True，则锁定文档备注。默认值为 False。                                                                                                        |
| *Password*                | 可选      | Variant  | 用于打开文档的密码字符串（请参阅下面的“说明”。）                                                                                                             |
| *AddToRecentFiles*        | 可选      | Variant  | 如果该属性值为 True，则将文档添加到“文件”菜单上最近使用的文件列表中。默认值为 True。                                                                         |
| *WritePassword*           | 可选      | Variant  | 用于将更改保存到文档的密码字符串（请参阅下面的“说明”。）                                                                                                     |
| *ReadOnlyRecommended*     | 可选      | Variant  | 如果该属性值为 True，则每次打开文档时，WPS 都将建议采用只读方式。默认值为 False。                                                                            |
| *EmbedTrueTypeFonts*      | 可选      | Variant  | 如果该属性值为 True，则 TrueType 字体随文档一起保存。如果省略，则 EmbedTrueTypeFonts 参数将假定为 EmbedTrueTypeFonts 属性的值。                              |
| *SaveNativePictureFormat* | 可选      | Variant  | 如果该属性值为 True，则对于从其他系统平台（如 Macintosh）导入的图形，只保存其 Microsoft Windows 版本。                                                       |
| *SaveFormsData*           | 可选      | Variant  | 如果该属性值为 True，则将用户在窗体中输入的数据保存为记录。                                                                                                  |
| *SaveAsAOCELetter*        | 可选      | Variant  | 如果文档包含附加的邮件，当该属性值为 True 时，将文档存为 AOCE 信函（同时保存邮件）。                                                                         |
| *Encoding*                | 可选      | Variant  | 用于另存为编码文本文件的文档的代码页或字符集。默认为系统代码页。不能将所有 MsoEncoding 常量用于该参数。                                                      |
| *InsertLineBreaks*        | 可选      | Variant  | 在将文档保存为文本文件时，如果该属性值为 True，则在每一行文本后插入换行符。                                                                                  |
| *AllowSubstitutions*      | 可选      | Variant  | 将文档保存为文本文件时，如果该属性值为 True，则 WPS 可以将一些符号替换为相似的文本。例如，将版权符号显示为 (c)。默认值为 False。                             |
| *LineEnding*              | 可选      | Variant  | WPS 在另存为文本文件的文档中标记换行符和分段符的方式。可为下列 WdLineEndingType 常量之一：wdCRLF（默认）或 wdCROnly。                                        |
| *AddBiDiMarks*            | 可选      | Variant  | 如果该属性值为 True，则为输出文件添加控制符，以保留原始文档中文本的双向版式。                                                                                |
| *CompatibilityMode*       | 可选      | Variant  | 打开文档时， WPS 2015 使用的兼容模式。WdCompatibilityMode 常量。                                                                                             |

**示例**

``` JavaScript
/* 下面的代码示例将活动文档另存为 Test.rtf，格式为 RTF*/
Application.ActiveDocument.SaveAs2("Text.rtf", wdFormatRTF)

/* 将文档保存为带密码 */
function test(strPWD) {
    Application.ActiveDocument.SaveAs2("", "", "", "", "", strPWD)
}
```

### Document.SaveAsQuickStyleSet

保存当前所用的快速样式组。

**语法**

**express.SaveAsQuickStyleSet ( *FileName* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                           |
|------------|-----------|----------|--------------------------------|
| *FileName* | 必选      | String   | 快速样式集文件的路径和文件名。 |

### Document.Select

选择指定文档的内容。

**语法**

**express.Select ()**

\*express是一个代表 Document 对象的变量。

**说明**

使用该方法后，可用 Selection 属性来处理选定项目。有关详细信息，请参阅 处理 Selection 对象。

**示例**

``` JavaScript
/* 选中当前文档所有文本 */
Application.ActiveDocument.Select()
```

### Document.SelectAllEditableRanges

选择指定用户或用户组有权修改的所有区域。

**语法**

**express.SelectAllEditableRanges ( *EditorID* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                                           |
|------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *EditorID* | 必选      | Variant  | 可以是代表用户电子邮件别名（如果在相同的域中）和电子邮件地址的 String 类型值，或者是代表用户组的 WdEditorType 常量。如果省略，则只选择所有用户有权修改的区域。 |

**示例**

``` JavaScript
/*以下示例选定当前用户有权修改的所有区域。*/
Application.ActiveDocument.SelectAllEditableRanges(wdEditorCurrent)
```

### Document.SelectContentControlsByTag

返回一个 **ContentControls** 集合，该集合代表文档中具有 *Tag* 参数中指定标记值的所有内容控件。只读。

**语法**

**express.SelectContentControlsByTag ( *Tag* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明                       |
|-------|-----------|----------|----------------------------|
| *Tag* | 必选      | String   | 要返回的内容控件的标记值。 |

**返回值**

ContentControls

### Document.SelectContentControlsByTitle

返回一个 **ContentControls** 集合，该集合代表文档中具有 *Title* 参数中指定标题的所有内容控件。只读。

**语法**

**express.SelectContentControlsByTitle ()**

\*express是一个代表 Document 对象的变量。

**返回值**

ContentControls

### Document.SelectLinkedControls

返回一个 **ContentControls** 集合，该集合代表文档中根据 Node 参数的指定，链接到文档 XML 数据存储区中特定自定义 XML 节点的所有内容控件。只读。

**语法**

**express.SelectLinkedControls ( *Node* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型      | 说明                                          |
|--------|-----------|---------------|-----------------------------------------------|
| *Node* | 必选      | CustomXMLNode | 与内容控件链接的文档数据存储区中的 XML 节点。 |

**返回值**

ContentControls

### Document.SelectNodes

返回一个 **XMLNodes** 集合，代表所有与 *XPath* 参数相匹配的节点并按其在文档或区域中显示的顺序排列。

**语法**

**express.SelectNodes ( *XPath* , *PrefixMapping* , *FastSearchSkippingTextNodes* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                          | 必选/可选 | 数据类型 | 说明                                                                                                                |
|-------------------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------|
| *XPath*                       | 必选      | String   | 有效的 XPath 字符串。有关 XPath 的详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 XPath 参考文档。     |
| *PrefixMapping*               | 可选      | Variant  | 提供用于搜索的架构中的前缀。如果 XPath 参数使用名称来搜索元素，则可使用 PrefixMapping 参数。                        |
| *FastSearchSkippingTextNodes* | 可选      | Boolean  | 如果该参数值为 True，则搜索指定节点时跳过所有文本节点。如果该参数值为 False，则搜索时包含文本节点。默认值为 False。 |

**返回值**

XMLNodes

**说明**

将 *FastSearchSkippingTextNodes* 参数设置为 **True** 会降低性能，因为 WPS 将根据节点中包含的文本对文档中所有节点进行搜索。

**示例**

``` JavaScript
/*以下示例返回活动文档中所有 book 元素的集合。*/
function test() {
    let strElement
    let strPrefix

    strElement = "/x:catalog/x:book"
    strPrefix = "xmlns:x=\"\"" + ActiveDocument.XMLSchemaReferences.Item(1).NamespaceURI + "\"\""

    let objElements = Application.ActiveDocument.SelectNodes(strElement, strPrefix)
}
```

### Document.SelectSingleNode

返回一个 **XMLNode** 对象，该对象代表画布与指定文档中 *XPath* 参数相匹配的第一个节点。

**语法**

**express.SelectSingleNode ( *XPath* , *PrefixMapping* , *FastSearchSkippingTextNodes* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                          | 必选/可选 | 数据类型 | 说明                                                                                                               |
|-------------------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------|
| *XPath*                       | 必选      | String   | 有效的 XPath 字符串。有关 XPath 的详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 XPath 参考文档。    |
| *PrefixMapping*               | 可选      | Variant  | 提供用于搜索的架构中的前缀。如果 XPath 参数使用名称来搜索元素，则可使用 PrefixMapping 参数。                       |
| *FastSearchSkippingTextNodes* | 可选      | Boolean  | 如果该参数值为 True，则搜索指定节点时跳过所有文本节点。如果该参数值为 False，则搜索时包含文本节点。默认值为 True。 |

**返回值**

XMLNode

**说明**

将 *FastSearchSkippingTextNodes* 参数设置为 **False** 会降低性能，因为 WPS 将根据节点中包含的文本对文档中所有节点进行搜索。

**示例**

``` JavaScript
/*以下示例返回在活动文档中找到的第一个 title 元素，该元素是 book 元素的子元素。*/
function test() {
    let strElement
    let strPrefix

    strElement = "/x:catalog/x:book"
    strPrefix = "xmlns:x=\"\"" + ActiveDocument.XMLSchemaReferences.Item(1).NamespaceURI + "\"\""

    let objElements = ActiveDocument.SelectSingleNode(strElement, strPrefix)
}
```

### Document.SelectUnlinkedControls

返回一个 **ContentControls** 集合，该集合代表文档中没有链接到文档 XML 数据存储区中 XML 节点的所有内容控件。只读。

**语法**

**express.SelectUnlinkedControls ( *Stream* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型      | 说明                                                                                                                             |
|----------|-----------|---------------|----------------------------------------------------------------------------------------------------------------------------------|
| *Stream* | 可选      | CustomXMLPart | 自定义 XML 部件引用。设置此参数会对返回的内容控件进行筛选，以只包括那些在其 XMLMapping 定义中引用了此 CustomXMLPart 的内容控件。 |

**返回值**

ContentControls

### Document.SendFax

将指定的文档作为传真发送，且无需用户交互。

**语法**

**express.SendFax ( *Address* , *Subject* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                               |
|-----------|-----------|----------|------------------------------------|
| *Address* | 必选      | String   | 收件人的传真号码。                 |
| *Subject* | 可选      | Variant  | 主题行的文字。字符数不可超过 255。 |

**示例**

``` JavaScript
/*本示例将活动文档作为传真发送。*/
Application.ActiveDocument.SendFax("12065551234",  "Important Fax")
```

### Document.SendFaxOverInternet

将文档发送给传真服务提供商，该提供商将该文档传真给一个或多个指定收件人。

**语法**

**express.SendFaxOverInternet ( *Recipients* , *Subject* , *ShowMessage* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                |
|---------------|-----------|----------|-----------------------------------------------------------------------------------------------------|
| *Recipients*  | 可选      | Variant  | 一个代表传真收件人的传真号码和电子邮件地址的 String 类型值。如果有多个收件人，则用分号隔开。        |
| *Subject*     | 可选      | Variant  | 一个代表传真文档的主题行的String 类型值。                                                           |
| *ShowMessage* | 可选      | Variant  | 如果该属性值为 True，则发送传真之前显示传真邮件。如果该属性值为 False，则发送传真而不显示传真邮件。 |

**说明**

使用 **SendFaxOverInternet** 方法需要在用户计算机上启用传真服务。如果未启用传真服务，则 **SendFaxOverInternet** 方法将会导致运行时错误。

可使用两种格式在 *Recipients* 参数中指定传真号码：一种是 recipientsfaxnumber@usersfaxprovider，另一种是 recipientsname@recipientsfaxnumber。可通过使用以下注册表路径访问用户的传真提供商信息：

``` JavaScript
HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\Services\Fax
```

使用此注册表位置的 FaxAddress 键值可确定用于某个用户的格式。如果此注册表项不存在，则没有可用的传真服务。

**示例**

``` JavaScript
/*以下示例将传真发送给传真服务提供商，该提供商将邮件传真给收件人。*/
Application.ActiveDocument.SendFaxOverInternet("14255550101@consolidatedmessenger.com", "For your review", true)
 
```

### Document.SendForReview

将文档通过电子邮件发送给指定收件人进行审阅。

**语法**

**express.SendForReview ( *Recipients* , *Subject* , *ShowMessage* , *IncludeAttachment* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                       |
|---------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Recipients*        | 可选      | Variant  | 列出邮件收件人的字符串。这些字符串可以是在电子邮件电话簿或完整电子邮件地址中未解析的名称和别名。用分号 (;) 分隔开多个收件人。如果留有空格并且 ShowMessage 为 False，则将收到错误消息，邮件也不会发送出去。 |
| *Subject*           | 可选      | Variant  | 邮件主题字符串。如果为空，则主题将为：请审阅“文件名”。                                                                                                                                                     |
| *ShowMessage*       | 可选      | Variant  | 一个代表执行该方法时是否显示邮件的 Boolean 值。默认值为 True。如果设置为 False，则邮件自动发送给接收者，而不先向发送者显示邮件。                                                                           |
| *IncludeAttachment* | 可选      | Variant  | 一个代表邮件是否应包含附件或到服务器位置的链接的 Boolean 值。默认值为 True。如果设置为 False，文档必须保存在共享位置中。                                                                                   |

**说明**

使用 **SendForReview** 方法可启动合作审阅流程。使用 **EndReview** 方法可结束审阅流程。

**示例**

``` JavaScript
/*本示例自动将当前文档作为电子邮件附件发送给指定收件人。*/
function WebReview() {
    Application.ActiveDocument.SendForReview("someone@example.com; amy jones", "Please review this document.", false, true)
}
```

### Document.SendMail

打开一个邮件窗口以通过 Microsoft Exchange 发送指定文档。

**语法**

**express.SendMail ()**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **SendMailAttach** 属性可控制文档是作为邮件窗口中的文本还是作为附件发送。

**示例**

``` JavaScript
/*本示例将活动文档作为邮件的附件发送。*/
function test()
{
    Application.Options.SendMailAttach = true
    Application.ActiveDocument.SendMail()
}
```

### Document.SendMailer

您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **SendMailer** 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。

**语法**

**express.SendMailer ()**

\*express是一个代表 Document 对象的变量。

### Document.SetCompatibilityMode

设置文档的兼容模式。

**语法**

**express.SetCompatibilityMode ( *Mode* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                             |
|--------|-----------|----------|----------------------------------------------------------------------------------|
| *Mode* | 必选      | Long     | 指定要靠近的 WPS 版本。使用 WdCompatibilityMode 枚举的常量作为该参数的一个参数。 |

**说明**

在 WPS 2015 中打开在以前版本的 WPS 中创建的文档时，将启用“兼容模式”。“兼容模式”确保在处理文档时无法使用 WPS 2015 中的任何新的或增强的功能，从而使用以前版本的 WPS 编辑该文档的用户将具有完全编辑能力。

**示例**

``` JavaScript
/*下面的示例代码将 WPS 2015 置于 WPS 2003 兼容模式。*/
Application.ActiveDocument.SetCompatibilityMode(wdWord2003)
```

### Document.SetDefaultTableStyle

指定用于文档中新建表的表样式。

**语法**

**express.SetDefaultTableStyle ( *Style* , *SetInTemplate* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                  |
|-----------------|-----------|----------|-------------------------------------------------------|
| *Style*         | 必选      | Variant  | 指定样式名称的字符串。                                |
| *SetInTemplate* | 必选      | Boolean  | 如果该属性值为 True，则在文档附加的模板中保存表样式。 |

**示例**

``` JavaScript
/*本示例检查活动文档中使用的默认表格样式是否是 Table Normal，如果是，将默认表格样式改为 TableStyle1。本示例假定存在名为 TableStyle1 的表格样式。*/
function TableDefaultStyle() {
    if (Application.ActiveDocument.DefaultTableStyle == "Table Normal") {
        Application.ActiveDocument.SetDefaultTableStyle("TableStyle1", true)
    }
}
```

### Document.SetLetterContent

将指定 **LetterContent** 对象的内容插入文档。

**语法**

**express.SetLetterContent ( *LetterContent* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型      | 说明                     |
|-----------------|-----------|---------------|--------------------------|
| *LetterContent* | 必选      | LetterContent | 包括不同信函内容的对象。 |

**说明**

该方法与 **RunLetterWizard** 方法类似，但前者不显示“英文信函向导”对话框。该方法基于 **LetterContent** 对象的内容，在指定文档中添加、删除信函内容或重新设置信函格式。

**示例**

``` JavaScript
/*本示例检索活动文档的“英文信函向导”元素，并更改信封用语文本；然后用 SetLetterContent 方法来更新活动文档，以便反映这些更改。*/
function test() {
    let myLetterContent = Application.ActiveDocument.GetLetterContent()
    myLetterContent.AttentionLine = "Greetings"
    Application.ActiveDocument.SetLetterContent(myLetterContent)
}
```

### Document.SetPasswordEncryptionOptions

设置 WPS 用于通过密码加密文档的选项

**语法**

**express.SetPasswordEncryptionOptions ( *PasswordEncryptionProvider* , *PasswordEncryptionAlgorithm* , *PasswordEncryptionKeyLength* , *PasswordEncryptionFileProperties* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称                               | 必选/可选 | 数据类型 | 说明                                                        |
|------------------------------------|-----------|----------|-------------------------------------------------------------|
| *PasswordEncryptionProvider*       | 必选      | String   | 加密提供程序的名称。                                        |
| *PasswordEncryptionAlgorithm*      | 必选      | String   | 加密算法的名称。 WPS 支持流式加密算法。                     |
| *PasswordEncryptionKeyLength*      | 必选      | Long     | 加密密钥长度。必须是从 40 开始的 8 的倍数。                 |
| *PasswordEncryptionFileProperties* | 可选      | Variant  | 如果该属性值为 True，则 WPS 对文件属性加密。默认值为 True。 |

**说明**

为了增强安全性，不要使用不可靠的加密 (XOR)（又称作“OfficeXor”）或“Office97/2000 Compatible”（又称作“OfficeStandard”）算法。

**示例**

``` JavaScript
/*如果使用的密码加密算法是“OfficeXor”或“OfficeStandard”，本示例将密码加密设置为更强的加密。*/
function PasswordSettings() {
    if (Application.ActiveDocument.PasswordEncryptionAlgorithm == "OfficeXor" || Application.ActiveDocument.PasswordEncryptionAlgorithm == "OfficeStandard") {
        Application.ActiveDocument.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)
    }
}
```

### Document.ToggleFormsDesign

在打开和关闭窗体设计模式之间切换。

**语法**

**express.ToggleFormsDesign ()**

\*express是一个代表 Document 对象的变量。

**说明**

当 WPS 处于窗体设计模式时，将显示 **“控件工具箱”** 工具栏。可以使用该工具箱插入各种 Microsoft ActiveX 控件，如命令按钮、滚动条和选项按钮。在窗体设计模式下，不会运行事件过程，并且单击一个嵌入的控件后，会出现控件的尺寸控点。

**示例**

``` JavaScript
/*如果当前没有显示“控件工具箱”工具栏，本示例将切换到窗体设计模式。*/
function test() {
    if (Application.CommandBars.Item("Control Toolbox").Visible == false) {
        Application.ActiveDocument.ToggleFormsDesign()
    }
}
```

### Document.TransformDocument

将 Extensible Stylesheet Language Transformation (XSLT) 文件应用于指定文档，并用结果替换该文档。

**语法**

**express.TransformDocument ( *Path* , *DataOnly* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                               |
|------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------|
| *Path*     | 必选      | String   | XSLT 使用的路径。                                                                                                                                  |
| *DataOnly* | 可选      | Boolean  | 如果该属性值为 True，则仅将转换应用于文档（不包括 WPS XML）中的数据。如果该属性值为 False，则将转换应用于整篇文档（包括 WPS XML）。默认值为 True。 |

**示例**

``` JavaScript
/*以下示例使用指定的 XSLT 文件转换活动文档。*/
Application.ActiveDocument.TransformDocument("c:\\schemas\\simplesample.xslt")
```

### Document.Undo

消最后一次操作或最后一系列操作，这些操作显示在 **“撤消”** 列表中。返回 **Boolean** 。 如果撤消操作成功，则返回 **True** 。

**语法**

**express.Undo ( *Times* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Times* | 可选      | Variant  | 要撤消的操作的数量。 |

**示例**

``` JavaScript
/* 加粗，倾斜后，再撤销一部，仅保留加粗 */
function test() {
    let rng = Application.ActiveDocument.Range(0,3);
    rng.Bold = true;
    rng.Italic = true;
    Application.ActiveDocument.Undo();
}
```

### Document.UndoClear

清除可对指定文档撤消的操作列表。

**语法**

**express.UndoClear ()**

\*express是一个代表 Document 对象的变量。

**说明**

该方法对应于单击 **“常用”** 工具栏上的 **“撤消”** 按钮旁边的箭头时出现的项目列表。在宏的末尾包括该方法可防止如“VBA-Selection.InsertAfter”等 Visual Basic 操作显示在 **“撤消”** 框中。

**示例**

``` JavaScript
/*本示例清除对活动文档撤消的操作列表。*/
Application.ActiveDocument.UndoClear()
```

### Document.Unprotect

清除对指定文档的保护。

**语法**

**express.Unprotect ( *Password* )**

\*express是一个代表 Document 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                         |
|------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------|
| *Password* | 可选      | Variant  | 用于保护文档的密码字符串。密码是区分大小写的。如果用户在使用一篇设置有密码的文档时没有提供正确的密码，就会显示一个对话框，提示用户输入密码。 |

**说明**

如果对文档没有加以保护，则该方法会导致出错。

## 安全性

请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

``` JavaScript
/*本示例使用 strPassword 变量的值作为密码，清除对活动文档的保护。*/
function test() {
    let strPassword
    if (Application.ActiveDocument.ProtectionType != wdNoProtection) {
        Application.ActiveDocument.Unprotect(strPassword)
    }
}
```

``` JavaScript
/*本示例清除对活动文档的保护。然后插入文本并对文档进行修订保护。*/
function test() {
    let strPassword
    let aDoc = Application.ActiveDocument
    if (aDoc.ProtectionType != wdNoProtection) {
        aDoc.Unprotect()
        Selection.InsertBefore("department six")
        aDoc.Protect(wdAllowOnlyRevisions, null, strPassword)
    }
}
```

### Document.UpdateStyles

将所有样式从附加模板复制到文档中，同时覆盖文档中所有现有同名的样式。

**语法**

**express.UpdateStyles ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例将相关模板中的样式复制到每一打开文档中，然后关闭每一文档。*/
function test() {
    for (let i = 1; i <= Application.Documents.Count; i++) {
        Application.Documents.Item(i).UpdateStyles()
        Application.Documents.Item(i).Close(wdSaveChanges)
    }
}
```

``` JavaScript
/*本示例更改活动文档所选用的模板中“标题 1”的样式设置。用 UpdateStyles 方法更新活动文档中的样式（包括“标题 1”样式）。*/
function test() {
    let aDoc = Application.ActiveDocument.AttachedTemplate.OpenAsDocument
    let font = aDoc.Styles.Item(wdStyleHeading1).Font
    font.Name = "Arial"
    font.Bold = false
    aDoc.Close(wdSaveChanges)
    Application.ActiveDocument.UpdateStyles()
}
```

### Document.UpdateSummaryProperties

更新 **“文件”** 菜单 **“属性”** 对话框中的关键词和备注文字，以反映为指定文档编写的自动摘要内容。

**语法**

**express.UpdateSummaryProperties ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例突出显示活动文档的要点，并且更新“文件”菜单“属性”对话框中的摘要信息。*/
function test() {
    Application.ActiveDocument.AutoSummarize(wd25Percent, wdSummaryModeHighlight)
    Application.ActiveDocument.UpdateSummaryProperties()
}
```

### Document.ViewCode

显示指定文档中所选 Microsoft ActiveX 控件的代码窗口。

**语法**

**express.ViewCode ()**

\*express是一个代表 Document 对象的变量。

**说明**

该方法只有在 WPS 之外才可用。

### Document.ViewPropertyBrowser

显示指定文档中所选 Microsoft ActiveX 控件的属性窗口。

**语法**

**express.ViewPropertyBrowser ()**

\*express是一个代表 Document 对象的变量。

**说明**

该方法只有在 WPS 之外才可用。

### Document.WebPagePreview

显示将当前文档保存为网页时的预览。

**语法**

**express.WebPagePreview ()**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例显示将文档保存为网页时的预览。*/
Application.ActiveDocument.WebPagePreview()
```

## 成员属性

### Document.ActiveTheme

返回指定文档的活动 主题 名称及该主题的格式选项。 **String** 类型，只读。

**语法**

**express.ActiveTheme**

\*express是一个代表 Document 对象的变量。

**说明**

如果该文档无活动主题，则 **ActiveTheme** 属性返回“none”。有关该属性返回值的详细解释，请参阅 **ApplyTheme** 方法的 *Name* 参数。该属性的返回值可能与主题的显示名称不对应。若要返回主题的显示名称，可使用 **ActiveThemeDisplayName** 属性。

**示例**

``` JavaScript
/*本示例实现的功能是：为当前文档应用主题并显示活动主题的名称及主题格式选项。*/
function CheckTheme() {
    Application.ActiveDocument.ApplyTheme("artsy 100")
    alert(Application.ActiveDocument.ActiveTheme)
}
```

### Document.ActiveThemeDisplayName

返回指定文档的活动 主题 的显示名称。 **String** 类型，只读。

**语法**

**express.ActiveThemeDisplayName**

\*express是一个代表 Document 对象的变量。

**说明**

如果该文档无活动主题，则 **ActiveThemeDisplayName** 属性返回“none”。主题的显示名称指显示在 **“主题”** 对话框中的名称。此名称可能与用来设置默认主题或为文档应用主题的字符串不对应。

**示例**

``` JavaScript
/*本示例返回当前文档活动主题的显示名称。*/
function CheckTheme() {
    Application.ActiveDocument.ApplyTheme("artsy 100")
    alert(Application.ActiveDocument.ActiveThemeDisplayName)
}
```

### Document.ActiveWindow

返回一个 **Window** 对象，该对象代表活动窗口（焦点所在的窗口）。只读。

**语法**

**express.ActiveWindow**

\*express是一个代表 Document 对象的变量。

**说明**

如果没有打开的窗口，则使用 **ActiveWindow** 属性将出现错误。

**示例**

``` JavaScript
/* 本示例显示活动窗口的标题文本 */
function CheckTheme() {
    Application.ActiveDocument.ActiveWindow.Caption
    Application.ActiveDocument.ActiveWindow.Split = true
}
```

### Document.ActiveWritingStyle

返回或设置指定文档中指定语言的写作风格。 **String** 类型，可读写。

**语法**

**express.ActiveWritingStyle**

\*express是一个代表 Document 对象的变量。

**说明**

**WritingStyleList** 属性返回一个可用写作风格名称的数组。

**示例**

``` JavaScript
/*本示例设置活动文档中法语、德语和美国英语的写作风格。若要运行本示例，必须安装法语、德语和美国英语的语法文件。*/
function test() {
    Application.ActiveDocument.ActiveWritingStyle(wdFrench).value = "Commercial"
    Application.ActiveDocument.ActiveWritingStyle(wdGerman).value = "Technisch/Wiss"
    Application.ActiveDocument.ActiveWritingStyle(wdEnglishUS).value = "Technical"
}
```

``` JavaScript
/*本示例返回选定内容的语言的写作风格。*/
function WhichLanguage() {
    let varLang = Application.Selection.LanguageID
    alert(Application.ActiveDocument.ActiveWritingStyle(varLang))
}
```

### Document.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Document 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Document.AttachedTemplate

返回一个 **Template** 对象，该对象代表与指定文档相关联的模板。可读/写 **Variant** 类型。

**语法**

**express.AttachedTemplate**

\*express是一个代表 Document 对象的变量。

**说明**

若要设置该属性，请指定模板名称或返回 **Template** 对象的表达式。

**示例**

``` JavaScript
/*本示例显示与活动文档相关联的模板名称和路径。*/
function test() {
    let myTemplate = Application.ActiveDocument.AttachedTemplate
    alert(myTemplate.Path + Application.PathSeparator + myTemplate.Name)
}
```

``` JavaScript
/*本示例在文档 1 的开始处插入图文场的内容（一个内置“自动图文集”词条）。*/
function test()
{
    let myRange = Application.Documents.Item(1).Range(0, 0)
    Applicatiion.Documents.Item(1).AttachedTemplate.AutoTextEntries.Item("Spike").Insert(myRange)
}
```

``` JavaScript
/*本示例为活动文档附加名为“Letter.dot”模板。*/
function test()
{
    Application.ActiveDocument.AttachedTemplate = "C:\\Templates\\Letter.dot"
}
```

### Document.AutoFormatOverride

返回或设置一个 **Boolean** 类型的值，该数值代表是否在已应用格式设置限制的文档中使用自动设置格式替代格式设置限制。

**语法**

**express.AutoFormatOverride**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*以下示例指定在受保护的文档中使用自动设置格式替代格式设置限制。*/
Application.ActiveDocument.AutoFormatOverride = true
```

### Document.AutoHyphenation

如果代表指定文档已启动自动断字功能，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoHyphenation**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例启动自动断字功能，断字区为 0.25 英寸。全为大写的单词不断字。*/
function test()
{
ActiveDocument.HyphenationZone = InchesToPoints(0.25)
ActiveDocument.HyphenateCaps = false
ActiveDocument.AutoHyphenation = true
}
```

### Document.Background

返回一个 **Shape** 对象，该对象代表指定文档的背景图像。只读。

**语法**

**express.Background**

\*express是一个代表 Document 对象的变量。

**说明**

只有在 Web 版式视图中才能看到背景。

**示例**

``` JavaScript
/*本示例将活动窗口的 Web 版式视图的背景色设置为浅灰色。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdWebView
let fill = ActiveDocument.Background.Fill
fill.Visible = true
fill.ForeColor.RGB = 192, 192, 192
}
```

``` JavaScript
/*本示例将 Web 版式视图的背景位图图像设置为 Bubbles.bmp。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdWebView
ActiveDocument.Background.Fill.UserPicture("C:\\Windows\\Bubbles.bmp")
}
```

### Document.Bibliography

返回一个 **Bibliography** 对象，代表文档中包含的书目参考。只读。

**语法**

**express.Bibliography**

\*express是一个代表 Document 对象的变量。

### Document.Bookmarks

返回一个 **Bookmarks** 集合，该集合代表文档中的所有书签。只读。

**语法**

**express.Bookmarks**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 返回第一个书签的开始结束位置 */
function test() {
    let myMark = Application.ActiveDocument.Bookmarks.Item(1)
    BookStart = myMark.Start
    BookEnd = myMark.End
    alert("start:" + BookStart + " end" + BookEnd)
}
```

### Document.BuiltInDocumentProperties

返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有内置的文档属性。

**语法**

**express.BuiltInDocumentProperties**

\*express是一个代表 Document 对象的变量。

**说明**

要返回单个代表特定内置文档属性的 **DocumentProperty** 对象，可使用 **BuiltinDocumentProperties** 属性。如果 WPS 没有为一个内置的文档属性定义一个值，则读取这个文档属性的 **Value** 属性时会产生一个错误。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

用 **CustomDocumentProperties** 属性返回自定义文档属性的集合。

**示例**

``` JavaScript
/*本示例在活动文档的末尾插入一个内置属性列表。*/
function ListProperties() {
    let rngDoc = ActiveDocument.Content
    rngDoc.Collapse(wdCollapseEnd)
    let proDoc = ActiveDocument.BuiltInDocumentProperties
    for(let i=1;i <= proDoc.Count;i++) {
        try {
            rngDoc.InsertParagraphAfter()
            rngDoc.InsertAfter(proDoc.Item(i).Name + "= ")
        }
        catch(exception) {
            rngDoc.InsertAfter(proDoc.Item(i).Value)
        }
    }
}
```

``` JavaScript
/*本示例显示活动文档中的单词数。*/
function DisplayTotalWords() {
    let intWords = ActiveDocument.BuiltInDocumentProperties.Item(wdPropertyWords).Value
    MsgBox("This document contains " + intWords + " words.")
}
```

### Document.Characters

返回一个 **Characters** 集合，该集合代表文档中的字符。只读。

**语法**

**express.Characters**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 返回文档中字符数，包括空格 */
Application.ActiveDocument.Characters.Count
```

### Document.ChildNodeSuggestions

返回一个 **XMLChildNodeSuggestions** 集合，该集合代表允许在文档中使用的元素列表。

**语法**

**express.ChildNodeSuggestions**

\*express是一个代表 Document 对象的变量。

**说明**

该属性返回所有附加架构的根元素。

**XMLChildNodeSuggestions** 集合中的每个 **XMLChildNodeSuggestion** 对象都是允许的、可能做为子元素的 XML 元素列表中的一项，该列表位于 **“XML 结构”** 任务窗格底部。

**示例**

``` JavaScript
/*下列示例循环遍历活动文档中选定的第一个元素的建议子元素，并在插入点插入所有允许的元素。*/
function GetChildNodeSuggestions() {
    let objNode = Selection.XMLParentNode
    let objSuggestion = objNode.ChildNodeSuggestions
    for(let i=1; i <= objSuggestion.Count; i++) {
        objSuggestion.Insert()
        Selection.MoveRight()
    }
}
```

### Document.ClickAndTypeParagraphStyle

返回或设置“即点即输”功能在指定文档中应用于文字的默认段落样式。 **Variant** 类型，可读写。

**语法**

**express.ClickAndTypeParagraphStyle**

\*express是一个代表 Document 对象的变量。

**说明**

要设置 **ClickAndTypeParagraphStyle** 属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。有关 **WdBuiltinStyle** 常量的列表，请参阅“Visual Basic 对象浏览器”。

如果将指定样式的 **InUse** 属性设为 **False** ，则会出现错误。

**示例**

``` JavaScript
/*本示例设置“即点即输”应用于纯文本文件的默认段落样式。*/
function test()
{
if(ActiveDocument.Styles.Item("Plain Text").InUse) {
    ActiveDocument.ClickAndTypeParagraphStyle = "Plain Text"
}
else {
    MsgBox("Sorry, this style is not in use yet.")
}
}
```

### Document.CoAuthoring

返回 CoAuthoring 对象，该对象提供文档中与共同创作相关的对象模型的入口点。只读。

**语法**

**express.CoAuthoring**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例返回活动文档中的所有合著者。*/
let allAuthors = ActiveDocument.CoAuthoring.Authors
```

### Document.CodeName

返回指定文档的代码名称。 **String** 类型，只读。

**语法**

**express.CodeName**

\*express是一个代表 Document 对象的变量。

**说明**

代码名称是保存文档的事件宏的模块名。该模块的默认名为“ThisDocument”，可在“工程”窗口中查看该模块。有关使用 Document 对象的事件的内容，请参阅 使用 Document 对象的事件。

**示例**

``` JavaScript
/*本示例返回活动文档的代码窗口名称。*/
MsgBox(ActiveDocument.CodeName)
```

### Document.CommandBars

返回一个 **CommandBars** 集合，该集合代表 WPS 中的菜单栏以及所有工具栏。

**语法**

**express.CommandBars**

\*express是一个代表 Document 对象的变量。

**说明**

在访问 **CommandBars** 集合之前，可用 **CustomizationContext** 属性设置模板或文档上下文。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例增大所有命令栏按钮并激活工具提示。*/
function test()
{
CommandBars.LargeButtons = true
CommandBars.DisplayTooltips = true
}
```

``` JavaScript
/*本示例在应用程序窗口底部显示“绘图”工具栏。*/
function test()
{
CommandBars.Item("Drawing").Visible = true
CommandBars.Item("Drawing").Position = msoBarBottom
}
```

``` JavaScript
/*本示例将“版本”命令按钮添至“常用”工具栏。*/
function test()
{
CustomizationContext = NormalTemplate
CommandBars.Item("Standard").Controls.Add(msoControlButton, 2522, null, 4)
}
```

### Document.Comments

返回一个 **Comments** 集合，该集合代表指定文档的所有批注。只读。

**语法**

**express.Comments**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 本示例将活动文档中的每一批注的作者姓名与“工具”菜单中“选项”对话框的“用户信息”选项卡的用户姓名作比较。如果两者不同,则设置备注内容为"test" */
function test() {
    let comm = Application.ActiveDocument.Comments
    for(let i=1;i <= comm.Count;i++) {
        if(comm.Item(i).Author != Application.UserName) {
            comm.Item(i).Range.Text = 'test'
        }
    }
}
```

### Document.CompatibilityMode

返回 **Number** 类型值，该值指定打开文档时 WPS 2015 使用的兼容模式。只读。

**语法**

**express.CompatibilityMode**

\*express是一个代表 Document 对象的变量。

**说明**

在 WPS 2015 中打开在以前版本的 WPS 中创建的文档时，将启用“兼容模式”。“兼容模式”确保在处理文档时无法使用 WPS 2015 中的任何新的或增强的功能，从而使用以前版本的 WPS 编辑该文档的用户将具有完全编辑能力。

**示例**

``` JavaScript
/* 获取文档的兼容文档 */
Application.ActiveDocument.CompatibilityMode
```

### Document.ConsecutiveHyphensLimit

返回或设置以连字符结束的连续行的最大数目。 **Long** 类型，可读写。

**语法**

**express.ConsecutiveHyphensLimit**

\*express是一个代表 Document 对象的变量。

**说明**

如果将 **ConsecutiveHyphensLimit** 属性设置为 0（零），则以连字符结束的连续行可以是任何数目。

**示例**

``` JavaScript
/*本示例为 MyReport.doc 设置自动断字功能并将以连字符结束的连续行的数目限制为 2 个。*/
function test()
{
Documents.Item("MyReport.doc").AutoHyphenation = true
Documents.Item("MyReport.doc").ConsecutiveHyphensLimit = 2
}
```

``` JavaScript
/*本示例不限制以连字符结束的连续行的数目。*/
ActiveDocument.ConsecutiveHyphensLimit = 0
```

### Document.Container

返回一个对象，该对象代表指定文档的容器应用程序。 **Object** 类型，只读。

**语法**

**express.Container**

\*express是一个代表 Document 对象的变量。

**说明**

如果指定的文档作为 OLE 对象嵌入到另一个应用程序中，则通过 **Container** 属性可以访问该文档的容器应用程序。如果一个 WPS 文档是作为 ActiveX 文档打开的（如在 WPS Office活页夹或 Internet Explorer 中打开一个 WPS 文档），则通过该属性也可访问容器应用程序中的对象模型。

**示例**

``` JavaScript
/*本示例显示活动文档第一个图形的容器应用程序名。若要让本示例正常运行，则该图形必须是一个 OLE 对象。*/
Msgbox(ActiveDocument.Shapes.Item(1).OLEFormat.Object.Container.Name)
```

### Document.Content

返回一个 **Range** 对象，该对象代表主文档 文章 。只读。

**语法**

**express.Content**

\*express是一个代表 Document 对象的变量。

**说明**

以下两个语句是等价的：

``` JavaScript
function test()
{
let mainStory
mainStory = ActiveDocument.Content
mainStory = ActiveDocument.StoryRanges(wdMainTextStory) 
}
```

**示例**

``` JavaScript
/*本示例将活动文档中的字体和字号分别改为 Arial 和 10 磅。*/
function test()
{
let myRange = ActiveDocument.Content
myRange.Font.Name = "Arial"
myRange.Font.Size = 10
}
```

``` JavaScript
/*本示例在名为“Changes.doc”的文档末尾插入文本。For Each...Next 语句用来判断此文档是否已打开。*/
function test()
{
for(let i=1;i <= Documents.Count;i++) {
    if(Documents.Item(i).Name.toLowerCase().indexOf("changes.doc") != -1) {
        let myRange = Documents.Item("Changes.doc").Content
        myRange.InsertAfter("the end.")
    }
}
}
```

### Document.ContentControls

返回一个 **ContentControls** 集合，该集合代表文档中的所有内容控件。只读。

**语法**

**express.ContentControls**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 下拉列表内容控件插入到活动文档中 */
let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
```

### Document.ContentTypeProperties

返回一个 **MetaProperties** 集合，该集合代表文档中所存储的元数据，如作者姓名、主题和公司。只读。

**语法**

**express.ContentTypeProperties**

\*express是一个代表 Document 对象的变量。

### Document.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Document 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Document.CurrentRsid

返回一个 **Long** 类型的值，该值代表 WPS 为文档中的更改所分配的随机数。只读。

**语法**

**express.CurrentRsid**

\*express是一个代表 Document 对象的变量。

**说明**

如果 **StoreRSIDOnSave** 属性为 **True** ，则每次保存文档时， WPS 生成一个随机数，应用程序使用该随机数来比较和合并文档。 WPS 将这些随机数保存在表格中并在每次保存后更新该表格。 **CurrentRsid** 属性将返回 WPS 分配给文档的上一个数。

### Document.CustomDocumentProperties

返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有自定义文档属性。

**语法**

**express.CustomDocumentProperties**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **BuiltInDocumentProperties** 属性可返回内置文档属性的集合。

**msoPropertyTypeString** 类型的属性的长度不能超过 255 个字符。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例在活动文档的末尾插入一个自定义内置属性列表。*/
function test()
{
let myRange = ActiveDocument.Content
myRange.Collapse(wdCollapseEnd)                                 
let prop = ActiveDocument.CustomDocumentProperties
for(let i=1;i <= prop.Count;i++) {
    myRange.InsertParagraphAfter()
    myRange.InsertAfter(prop.Item(i).Name + "= ")
    myRange.InsertAfter(prop.Item(i).Value)
}
}
```

``` JavaScript
/*本示例为 Sales.doc 添加一个自定义内置属性。*/
function test()
{
let thename = prompt("Please type your name", "Name")
Documents.Item("Sales.doc").CustomDocumentProperties.Add("YourName", false, msoPropertyTypeString, thename)
}
```

### Document.CustomXMLParts

返回一个 **CustomXMLParts** 集合，该集合代表 XML 数据存储区中的自定义 XML。只读。

**语法**

**express.CustomXMLParts**

\*express是一个代表 Document 对象的变量。

### Document.DefaultTableStyle

返回一个 **Variant** 类型的值，该值代表可以应用于文档中新创建的所有表格的表格样式。只读。

**语法**

**express.DefaultTableStyle**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例检查活动文档中使用的默认表格样式是否是“Table Normal”，如果是，将默认表格样式改为“TableStyle1”。本示例假定存在名为“TableStyle1”的表格样式。*/
function TableDefaultStyle() {
    if(ActiveDocument.DefaultTableStyle == "Table Normal") {
        ActiveDocument.SetDefaultTableStyle("TableStyle1", true)
    }
}
```

### Document.DefaultTabStop

以磅为单位返回或设置指定文档中的默认制表位间隔。 **Single** 类型，可读写。

**语法**

**express.DefaultTabStop**

\*express是一个代表 Document 对象的变量。

**示例**

------------------------------------------------------------------------

``` JavaScript
/*本示例将活动文档中的默认制表位设置为 1 英寸。其中 InchesToPoints 方法用来将英寸转化为磅值。*/
ActiveDocument.DefaultTabStop = InchesToPoints(1)
```

### Document.DefaultTargetFrame

返回或设置一个 **String** 类型的值，该值代表通过超链接可到达显示网页的浏览器框架。可读写。

**语法**

**express.DefaultTargetFrame**

\*express是一个代表 Document 对象的变量。

**说明**

当 **DefaultTargetFrame** 属性可以使用任意用户定义的字符串时，它有下列预定义的字符串：“\_top”、“\_blank”、“\_parent”和“\_self”。

**示例**

``` JavaScript
/*本示例的功能如下：当用户单击活动文档中的超链接时，WPS 打开一个新的空白浏览器窗口。*/
function DefaultFrame() {
    ActiveDocument.DefaultTargetFrame = "_blank"
}
```

### Document.DisableFeatures

如果为 **True** ，则禁用 **DisableFeaturesIntroducedAfter** 属性中指定的版本之后的所有功能。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.DisableFeatures**

\*express是一个代表 Document 对象的变量。

**说明**

**DisableFeatures** 属性只影响设置属性的文档。如果计划与使用 WPS 较早版本的用户共享文档，请使用该属性，这样就无需禁用其 WPS 版本不支持的文档功能。

**示例**

``` JavaScript
/*本示例在当前文档中禁用在 WPS for Windows 、7.0 和 7.0a 版本之后添加的功能，全局默认设置保持不变。*/
function FeaturesDisable() {

    //Checks whether features are disabled
    if(ActiveDocument.DisableFeatures == true) {

        //If they are, disables all features after  WPS for Windows
        ActiveDocument.DisableFeaturesIntroducedAfter = wd70
}
    else {

        //If not, turns on the disable features option and disables
        //all features introduced after  WPS for Windows
        ActiveDocument.DisableFeatures = true
        ActiveDocument.DisableFeaturesIntroducedAfter = wd70
}
}
```

### Document.DisableFeaturesIntroducedAfter

仅对该文档禁用指定的 WPS 版本后引入的所有功能。 **WdDisableFeaturesIntroducedAfter** 类型，可读写。

**语法**

**express.DisableFeaturesIntroducedAfter**

\*express是一个代表 Document 对象的变量。

**说明**

设置 **DisableFeaturesIntroducedAfter** 属性之前，必须先将 **DisableFeatures** 属性设为 **True** 。否则，该设置无效，并保持其默认值 WPS 97 for Windows。

**DisableFeaturesIntroducedAfter** 属性仅影响该属性所设置的文档。如果要设置应用程序的全局选项以禁用所有文档的功能。可使用 **DisableFeaturesIntroducedAfterByDefault** 属性。

**示例**

``` JavaScript
/*本示例仅对当前文档禁用 WPS for Windows 的 7.0 和 7.0a 版本之后引入的所有功能。全局默认设置保持不变。*/
function FeaturesDisable() {

    //Checks whether features are disabled
    if(ActiveDocument.DisableFeatures == true) {

        //If they are, disables all features after  WPS for Windows
        ActiveDocument.DisableFeaturesIntroducedAfter = wd70
}
    else {

        /*If not, turns on the disable features option and disables
        all features introduced after  WPS for Windows*/
        ActiveDocument.DisableFeatures = true
        ActiveDocument.DisableFeaturesIntroducedAfter = wd70
}
}
```

### Document.DocumentInspectors

返回一个 **DocumentInspectors** 集合，该集合使您能够查找隐藏的个人信息，如作者姓名、公司名称和修订日期。只读。

**语法**

**express.DocumentInspectors**

\*express是一个代表 Document 对象的变量。

### Document.DocumentLibraryVersions

返回一个 **DocumentLibraryVersions** 集合，该集合表示共享文档的版本集合，该共享文档启用了版本控制并存储在服务器的文档库中。

**语法**

**express.DocumentLibraryVersions**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*以下示例返回活动文档的版本集合。本示例假设活动文档启用了版本控制并存储在服务器的共享文档库中。*/
let objVersions = ActiveDocument.DocumentLibraryVersions
```

### Document.DocumentTheme

返回一个 **OfficeTheme** 对象，代表应用于文档的 WPS Office主题。只读。

**语法**

**express.DocumentTheme**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **ApplyDocumentTheme** 方法可应用 Office 主题。

### Document.DoNotEmbedSystemFonts

如果为 **True** ，则 WPS 不嵌入常规系统字体。 **Boolean** 类型，可读写。

**语法**

**express.DoNotEmbedSystemFonts**

\*express是一个代表 Document 对象的变量。

**说明**

如果用户使用中文系统创建文档，并希望该文档对于那些系统中没有该语言字体的用户也可读，可将 **Document** 属性设为 **False** 。例如，日语系统的用户可以选择在文档中嵌入字体，使日文文档在所有系统中可读。

**示例**

``` JavaScript
/*本示例在当前文档中嵌入所有字体。*/
function EmbedFonts() {
    if(ActiveDocument.EmbedTrueTypeFonts == false) {
        ActiveDocument.EmbedTrueTypeFonts = true
        ActiveDocument.DoNotEmbedSystemFonts = false
}
    else {
        ActiveDocument.DoNotEmbedSystemFonts = false
}
}
```

### Document.Email

返回一个 **Email** 对象，该对象包含当前文档的所有与电子邮件相关的属性。只读。

**语法**

**express.Email**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例将返回与当前电子邮件作者相关联的样式名。*/
MsgBox(ActiveDocument.Email.CurrentEmailAuthor.Style.NameLocal)
```

### Document.EmbedLinguisticData

如果为 **True** ，则 WPS 会嵌入语音和手写功能，以便数据可以转换回语音或手写。 **Boolean** 类型，可读写。

**语法**

**express.EmbedLinguisticData**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **EmbedLinguisticData** 属性还能保存东亚 IME 键击，以便使用 Windows 文本服务框架应用程序编程接口提高对文本服务数据（源自与 WPS Office连接的设备）的更正和控制能力。

**示例**

``` JavaScript
/*本示例将文档中可能出现的任何语音和手写输入嵌入活动文档中。*/
function EmbedSpeechHandwriting() {
    ActiveDocument.EmbedLinguisticData = true
}
```

### Document.EmbedTrueTypeFonts

如果在保存文档时，WPS 在文档中嵌入 TrueType 字体，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EmbedTrueTypeFonts**

\*express是一个代表 Document 对象的变量。

**说明**

通过嵌入 TrueType 字体，其他人可以使用创建文档时所用的字体来查看该文档。

**示例**

``` JavaScript
/*本示例设置 WPS 在保存文档时，在文档中自动嵌入 TrueType 字体，然后保存活动文档。*/
function test()
{
ActiveDocument.EmbedTrueTypeFonts = true
ActiveDocument.Save()
}
```

``` JavaScript
/*本示例返回“选项”对话框中“保存”选项卡上“保存选项”区域中“嵌入 TrueType 字体”复选框的当前状态。*/
let temp = ActiveDocument.EmbedTrueTypeFonts
```

### Document.EncryptionProvider

返回一个 **String** 类型的值，该值指定 WPS 在对文档进行加密时使用的算法加密提供程序的名称。可读写。

**语法**

**express.EncryptionProvider**

\*express是一个代表 Document 对象的变量。

### Document.Endnotes

返回一个 **Endnotes** 集合，该集合代表文档中的所有尾注。只读。

**语法**

**express.Endnotes**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例将活动文档中的尾注放在文档的结尾并将尾注引用编号设置为小写的罗马数字。*/
function test()
{
let endnotes = ActiveDocument.Endnotes
endnotes.Location = wdEndOfDocument
endnotes.NumberStyle = 2
}
```

### Document.EnforceStyle

返回或设置一个 **Boolean** 类型的值，该值代表是否在受保护的文档中实施格式设置限制。

**语法**

**express.EnforceStyle**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*以下示例启用活动文档中的格式设置限制。*/
ActiveDocument.EnforceStyle = true
```

### Document.Envelope

返回一个 **Envelope** 对象，该对象代表文档中的信封和信封功能。只读。

**语法**

**express.Envelope**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例将默认信封尺寸设为 C4（229 x 324 毫米）。*/
ActiveDocument.Envelope.DefaultSize = "C4"
```

``` JavaScript
/*如果文档中添加了一个信封，则本示例显示收信人地址，否则显示一个消息框。*/
function test()
{
try {
    let addr = ActiveDocument.Envelope.Address.Text
    MsgBox(addr, null, "Delivery Address")
}
catch(exception) {
    MsgBox("Add an envelope to the document")
}
}
```

``` JavaScript
/*本示例新建一篇文档，并添加有预定义收信人地址和寄信人地址的一个信封。*/
function test()
{
let addr = "Don Funk" + "\r" + "123 Skye St." + "\r" + "Our Town, WA  98040"
let retaddr = "Karin Gallagher" + "\r" + "123 Main" + "\r" + "Other Town, WA  98004"
Documents.Add().Envelope.Insert(null, addr, null, null, retaddr)
ActiveDocument.ActiveWindow.View.Type = wdPrintView
}
```

### Document.FarEastLineBreakLanguage

返回或设置一个 **WdFarEastLineBreakLanguageID** 类型的值，该值代表在指定文档或模板中对文本进行换行时使用的东亚语言。可读写。

**语法**

**express.FarEastLineBreakLanguage**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在当前文档中依照朝鲜语的规则换行。*/
ActiveDocument.FarEastLineBreakLanguage = wdLineBreakKorean
```

### Document.FarEastLineBreakLevel

返回或设置一个 **WdFarEastLineBreakLevel** 类型的值，该值代表指定文档的换行控制级别。可读写。

**语法**

**express.FarEastLineBreakLevel**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在活动文档中第一级首尾字符处换行。*/
ActiveDocument.FarEastLineBreakLevel = wdJustificationModeCompressKana
 
```

### Document.Fields

返回一个 **Fields** 集合，该集合代表文档中的所有域。只读。

**语法**

**express.Fields**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
Application.ActiveDocument.Fields.Update()
```

### Document.Final

返回或设置 **Boolean** 值，该值指示文档是否是最终的。可读/写。

**语法**

**express.Final**

\*express是一个代表 Document 对象的变量。

**说明**

**True** 将文档标记为最终版本，以通知收件人（如果有）文档是最终的，并将文档设置为只读。

### Document.Footnotes

返回一个 **Footnotes** 集合，该集合代表文档中的所有脚注。只读。

**语法**

**express.Footnotes**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 设置当前文档脚注的起始编号 */
Application.ActiveDocument.Footnotes.StartingNumber = 3
```

### Document.FormattingShowClear

如果为 **True** ，则 WPS 在 **“样式和格式”** 任务窗格中显示“清除格式”。 **Boolean** 类型，可读写。

**语法**

**express.FormattingShowClear**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例禁用样式列表中“清除格式”按钮的显示。*/
function ShowClearFormatting() {
    ActiveDocument.FormattingShowClear = false
    ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse
    ActiveDocument.FormattingShowFont = true
    ActiveDocument.FormattingShowNumbering = true
    ActiveDocument.FormattingShowParagraph = true
}
```

### Document.FormattingShowFilter

设置或返回一个 **WdShowFilter** 常量，代表 **“样式和格式”** 任务窗格中显示的样式和格式。 **Boolean** 类型，可读写

**语法**

**express.FormattingShowFilter**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例对格式进行筛选，在“样式和格式”任务窗格中仅显示活动文档中使用的格式。*/
function ShowClearFormatting() {
    ActiveDocument.FormattingShowClear = false
    ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse
    ActiveDocument.FormattingShowFont = true
    ActiveDocument.FormattingShowNumbering = true
    ActiveDocument.FormattingShowParagraph = true
}
```

### Document.FormattingShowFont

如果为 **True** ，则 WPS 显示 **“样式和格式”** 任务窗格中的字体格式。 **Boolean** 类型，可读写。

**语法**

**express.FormattingShowFont**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例启用“样式和格式”任务窗格中字体样式的显示功能。*/
function ShowClearFormatting() {
    ActiveDocument.FormattingShowClear = false
    ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse
    ActiveDocument.FormattingShowFont = true
    ActiveDocument.FormattingShowNumbering = true
    ActiveDocument.FormattingShowParagraph = true
}
```

### Document.FormattingShowNextLevel

返回或设置 **Boolean** 值，该值表示 WPS 在使用上一个标题级别时是否显示下一个标题级别。可读/写。

**语法**

**express.FormattingShowNextLevel**

\*express是一个代表 Document 对象的变量。

**说明**

此属性对应于 **“样式库选项”** 对话框中的 **“使用上一个级别时显示下一级标题”** 复选框。

### Document.FormattingShowNumbering

如果为 **True** ，则 WPS 在 **“样式和格式”** 任务窗格中显示编号格式。 **Boolean** 类型，可读写。

**语法**

**express.FormattingShowNumbering**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例启用“样式和格式”任务窗格中编号格式的显示功能。*/
function ShowClearFormatting() {
    ActiveDocument.FormattingShowClear = false
    ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse
    ActiveDocument.FormattingShowFont = true
    ActiveDocument.FormattingShowNumbering = true
    ActiveDocument.FormattingShowParagraph = true
}
```

### Document.FormattingShowParagraph

如果为 **True** ，则 WPS 在 **“样式和格式”** 任务窗格中显示段落格式。 **Boolean** 类型，可读写。

**语法**

**express.FormattingShowParagraph**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例启用“样式和格式”任务窗格中段落格式的显示功能。*/
function ShowClearFormatting() {
    ActiveDocument.FormattingShowClear = false
    ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse
    ActiveDocument.FormattingShowFont = true
    ActiveDocument.FormattingShowNumbering = true
    ActiveDocument.FormattingShowParagraph = true
}
```

### Document.FormattingShowUserStyleName

返回或设置 **Boolean** 值，该值表示是否显示用户定义的样式。可读/写。

**语法**

**express.FormattingShowUserStyleName**

\*express是一个代表 Document 对象的变量。

**说明**

此属性对应于 **“样式库选项”** 对话框中的 **“存在可选名称时隐藏内置名称”** 复选框。当存在可选的用户定义样式时，将 **FormattingShowUserStyleName** 属性设置为 **True** 会隐藏内置样式。例如，如果用户创建标题样式（如“标题 1”或“标题 2”）以利用 WPS 的其他内置功能（例如，目录），则用户定义的样式优先，并且在样式名称列表中不会显示以相似方式命名的内置样式。

### Document.FormFields

返回一个 **FormFields** 集合，该集合代表文档中的所有窗体域。只读。

**语法**

**express.FormFields**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例将名为“Text1”的窗体域的内容设置为“Name”。*/
ActiveDocument.FormFields.Item("Text1").Result = "Name"
```

### Document.FormsDesign

如果指定的文档处于窗体设计模式，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.FormsDesign**

\*express是一个代表 Document 对象的变量。

**说明**

如果在从 WPS 中运行的代码中使用 **FormsDesign** 属性，则本属性总是返回 **False** ，但如果通过自动功能运行，则可返回正确的值。例如，如果您在 ET 中运行示例，则处于设计模式时将返回 **True** 。

当 WPS 处于窗体设计模式时，将显示 **“控件工具箱”** 工具栏。使用该工具箱可以插入各种 ActiveX 控件，如命令按钮、滚动条和选项按钮。在窗体设计模式下，不会运行事件过程，并且单击一个嵌入的控件后，就会出现控件的尺寸控点。

**示例**

``` JavaScript
/*本示例显示一个消息框，以报告活动文档是否处于窗体设计模式。本示例总是返回 False。*/
MsgBox(ActiveDocument.FormsDesign)
```

### Document.FormsDesignFormsDesign

如果指定的文档处于窗体设计模式，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.FormsDesignFormsDesign**

\*express是一个代表 Document 对象的变量。

**说明**

如果在从 WPS 中运行的代码中使用 **FormsDesign** 属性，则本属性总是返回 **False** ，但如果通过自动功能运行，则可返回正确的值。例如，如果您在 ET 中运行示例，则处于设计模式时将返回 **True** 。

当 WPS 处于窗体设计模式时，将显示 **“控件工具箱”** 工具栏。使用该工具箱可以插入各种 ActiveX 控件，如命令按钮、滚动条和选项按钮。在窗体设计模式下，不会运行事件过程，并且单击一个嵌入的控件后，就会出现控件的尺寸控点。

**示例**

``` JavaScript
/*本示例显示一个消息框，以报告活动文档是否处于窗体设计模式。本示例总是返回 False。*/
MsgBox(ActiveDocument.FormsDesign)
```

### Document.Frames

返回一个 **Frames** 集合，该集合代表文档中的所有图文框。只读。

**语法**

**express.Frames**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例在选定内容的四周添加一个图文框，并将图文框对象返回给 myFrame 变量。*/
let myFrame = ActiveDocument.Frames.Add(Selection.Range)
```

### Document.Frameset

返回一个 **Frameset** 对象，该对象代表一个整体框架页或框架页的单个框架。只读。

**语法**

**express.Frameset**

\*express是一个代表 Document 对象的变量。

**说明**

有关创建框架页的详细信息，请参阅 创建框架页。

**示例**

``` JavaScript
/*本示例将指定框架页中框架边框的颜色设为茶色。*/
function test()
{
let frameset = ActiveWindow.Document.Frameset
frameset.FramesetBorderColor = wdColorTan
frameset.FramesetBorderWidth = 6
}
```

### Document.FullName

返回一个 **String** 类型的值，该值代表包括路径的文档的名称。只读。

**语法**

**express.FullName**

\*express是一个代表 Document 对象的变量。

**说明**

使用该属性与按顺序使用 **Path** 、 **PathSeparator** 和 **Name** 属性的作用相同。

**示例**

``` JavaScript
/*本示例显示活动文档的路径和文件名。*/
function DocName() {
    MsgBox(ActiveDocument.FullName)
}
```

### Document.GrammarChecked

如果已经检查了指定范围或文档的语法，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.GrammarChecked**

\*express是一个代表 Document 对象的变量。

**说明**

如果指定文档的部分或全部没有经过语法检查，则返回 **False** 。要重复检查文档的语法，可将 **GrammarChecked** 属性设为 **False** 。

**示例**

``` JavaScript
/*本示例判断是否对活动文档进行了语法检查。如果已检查，则显示文档的字数。如果尚未检查语法，则开始进行拼写和语法检查。*/
function test()
{
let myStat = ActiveDocument.ReadabilityStatistics
let passGram = ActiveDocument.GrammarChecked
if(passGram == true) {
    MsgBox(myStat.Item(1).Name + " - " + myStat.Item(1).Value)
}
else {
    ActiveDocument.CheckGrammar()
}
}
```

``` JavaScript
/*本示例设置活动文档的 GrammarChecked 属性为 False，然后再次进行语法检查。*/
function test()
{
ActiveDocument.GrammarChecked = false
ActiveDocument.CheckGrammar()
}
```

### Document.GrammaticalErrors

返回一个 **ProofreadingErrors** 集合，该集合代表指定的文档中有语法检查错误的句子。只读。

**语法**

**express.GrammaticalErrors**

\*express是一个代表 Document 对象的变量。

**说明**

在每个句子中可有多个错误。如果没有语法错误，由 **GrammaticalErrors** 属性返回的 **ProofreadingErrors** 集合的 **Count** 属性的返回值为 0。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例检查活动文档的语法错误。如果发现任何错误，则重新开始拼写和语法检查。*/
function test()
{
if(ActiveDocument.GrammaticalErrors.Count == 0) {
    MsgBox("There are no grammatical errors.")
}
else {
    ActiveDocument.CheckGrammar()
}
}
```

### Document.GridDistanceHorizontal

在指定文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置代表 WPS 所用的不可见的网格线之间的水平距离的 **Single** 类型，可读写。

**语法**

**express.GridDistanceHorizontal**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置网格线之间的水平和垂直间距，并为当前文档启用“对象与网格对齐”功能。*/
function test()
{
ActiveDocument.GridDistanceHorizontal = 9
ActiveDocument.GridDistanceVertical = 9
ActiveDocument.SnapToGrid = true
}
```

### Document.GridDistanceVertical

在指定的文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置代表 WPS 所用的不可见网格线之间的垂直距离的 **Single** 类型。可读写。

**语法**

**express.GridDistanceVertical**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置网格线之间的水平和垂直间距，并为当前文档启用“对象与网格对齐”功能。*/
function test()
{
ActiveDocument.GridDistanceHorizontal = 9
ActiveDocument.GridDistanceVertical = 9
ActiveDocument.SnapToGrid = true
}
```

### Document.GridOriginFromMargin

如果 WPS 将在页面左上角开始使用字符网格，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.GridOriginFromMargin**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 从活动文档页面的左上角开始使用字符网格。*/
ActiveDocument.GridOriginFromMargin = true
```

### Document.GridOriginHorizontal

返回或设置一个 **Single** 类型的值，该值代表不可见网格相对于页面左边的起点位置，当在指定的文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。可读写。

**语法**

**express.GridOriginHorizontal**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置网格的水平和垂直起点，以及网格线之间的水平和垂直距离，并为当前文档启用“对齐网格”功能。*/
function test()
{
let ment = ActiveDocument
ment.GridOriginHorizontal = 80
ment.GridOriginVertical = 90
ment.GridDistanceHorizontal = 9
ment.GridDistanceVertical = 9
ment.SnapToGrid = true
}
```

### Document.GridOriginVertical

返回或设置一个 **Single** 类型的值，该值代表不可见网格相对于页面顶边的起点位置，当在指定的文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。可读写。

**语法**

**express.GridOriginVertical**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置网格的水平和垂直起点，以及网格线之间的水平和垂直距离，并为当前文档启用“对象与网格对齐”功能。*/
function test()
{
let ment = ActiveDocument
ment.GridOriginHorizontal = 80
ment.GridOriginVertical = 90
ment.GridDistanceHorizontal = 9
ment.GridDistanceVertical = 9
ment.SnapToGrid = true
}
```

### Document.GridSpaceBetweenHorizontalLines

返回或设置 WPS 在页面视图中显示的水平字符网格的间隔。 **Long** 类型，可读写。

**语法**

**express.GridSpaceBetweenHorizontalLines**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 每隔四格显示一个水平字符网格线。*/
ActiveDocument.GridSpaceBetweenHorizontalLines = 5
```

### Document.GridSpaceBetweenVerticalLines

返回或设置 WPS 在页面视图中显示的垂直字符网格线间隔。 **Long** 类型，可读写。

**语法**

**express.GridSpaceBetweenVerticalLines**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 每隔一格显示一个垂直字符网格线。*/
ActiveDocument.GridSpaceBetweenVerticalLines = 2
```

### Document.HasMailer

您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **HasMailer** 属性的信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

**语法**

**express.HasMailer**

\*express是一个代表 Document 对象的变量。

### Document.HasPassword

如果需要密码才能打开指定文档，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.HasPassword**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例首先将“kittycat”设置为活动文档的密码，然后显示确认信息。*/
function test()
{
ActiveDocument.Password = "kittycat"
if(ActiveDocument.HasPassword == true) {
    MsgBox("The password is set.")
}
}
```

### Document.HasVBProject

返回一个 **Boolean** 类型的值，该值代表文档是否具有附加的 示例代码 项目。只读。

**语法**

**express.HasVBProject**

\*express是一个代表 Document 对象的变量。

**说明**

此属性非常有用，可用来以编程方式确定是否需要将文档保存为支持宏的文件格式。如果用其他格式保存，则文档中包含的宏和代码项目可能丢失。

### Document.HTMLDivisions

返回一个 **HTMLDivisions** 集合，该集合代表 Web 文档中的 HTML DIV 元素。

**语法**

**express.HTMLDivisions**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动文档中三个嵌套的区域的格式。本示例假定该活动文档为具有至少三个区域的 HTML 文档。*/
function FormatHTMLDivisions() {
    let htmldivisions = ActiveDocument.HTMLDivisions.Item(1)
    let htmldivisions1 = htmldivisions.HTMLDivisions.Item(1)
    let htmldivisions2 = htmldivisions1.HTMLDivisions.Item(1)
    let borderleft = htmldivisions.Borders.Item(wdBorderLeft)
    let borderright = htmldivisions.Borders.Item(wdBorderRight)
    let bordertop = htmldivisions1.Borders.Item(wdBorderTop)
    let borderbottom = htmldivisions1.Borders.Item(wdBorderBottom)
    borderleft.Color = wdColorRed
    borderleft.LineStyle = wdLineStyleSingle
    borderright.Color = wdColorRed
    borderright.LineStyle = wdLineStyleSingle
    htmldivisions1.LeftIndent = InchesToPoints(1)
    htmldivisions1.RightIndent = InchesToPoints(1)
    bordertop.Color = wdColorBlue
    bordertop.LineStyle = wdLineStyleDouble
    borderbottom.Color = wdColorBlue
    borderbottom.LineStyle = wdLineStyleDouble
    htmldivisions2.LeftIndent = InchesToPoints(1)
    htmldivisions2.RightIndent = InchesToPoints(1)
    htmldivisions2.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleDot
    htmldivisions2.Borders.Item(wdBorderRight).LineStyle = wdLineStyleDot
    htmldivisions2.Borders.Item(wdBorderTop).LineStyle = wdLineStyleDot
    htmldivisions2.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleDot
}
```

### Document.Hyperlinks

返回一个 **Hyperlinks** 集合，该集合代表指定文档中的所有超链接。只读。

**语法**

**express.Hyperlinks**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示 Home.doc 中第二个超链接的目标地址。*/
function test()
{
if(Documents.Item("Home.doc").Hyperlinks.Count >= 2) {
    MsgBox(Documents.Item("Home.doc").Hyperlinks.Item(2).Name)
}
}
```

``` JavaScript
/*本示例显示活动文档中地址包含“Microsoft”一词的每个超链接的名字。*/
function test()
{
let aHyperlink = ActiveDocument.Hyperlinks
for(let i = 1; i <= aHyperlink.Count; i++) {
    if((aHyperlink.Item(i).Address.toLowerCase().indexOf("microsoft") != -1) != 0) {
        MsgBox(aHyperlink.Item(i).Name)
}
}
}
```

### Document.HyphenateCaps

如果全是大写字母时可对其进行断字，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HyphenateCaps**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例自动为指定的文档断字，并允许对全是大写字母的单词进行断字。*/
function test()
{
ActiveDocument.AutoHyphenation = true
ActiveDocument.HyphenateCaps = true
}
```

### Document.HyphenationZone

返回或设置断字区的宽度（以磅为单位）。 **Long** 类型，可读写。

**语法**

**express.HyphenationZone**

\*express是一个代表 Document 对象的变量。

**说明**

断字区是 WPS 在一行最后一个单词与右边距间的保留的最大距离。

**示例**

``` JavaScript
/*本示例为 MyReport.doc 启用自动断字功能。并将断字区的宽度设置为 36 磅（0.5 英寸）。*/
function test()
{
Documents.Item("MyReport.doc").AutoHyphenation = true
Documents.Item("MyReport.doc").HyphenationZone = 36
}
```

``` JavaScript
/*本示例将断字区的宽度设置为 0.25 英寸（18 磅），并启用活动文档的手动断字功能。*/
function test()
{
ActiveDocument.HyphenationZone = InchesToPoints(0.25)
ActiveDocument.ManualHyphenation()
}
```

### Document.Indexes

返回一个 **Indexes** 集合，该集合代表指定文档的所有索引。只读。

**语法**

**express.Indexes**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例在活动文档的末尾添加索引。*/
function test()
{
let MyRange = ActiveDocument.Range(ActiveDocument.Content.End - 1, ActiveDocument.Content.End - 1)
ActiveDocument.Indexes.Add(MyRange, false, null, null, 1)
}
```

``` JavaScript
/*本示例插入关于选定文本的索引项。*/
function test()
{
if(Selection.Type == wdSelectionNormal) {
    ActiveDocument.Indexes.MarkEntry(Selection.Range, Selection.Range.Text)
}
}
```

### Document.InlineShapes

返回一个 **InlineShapes** 集合，该集合代表文档中的所有 **InlineShape** 对象。只读。

**语法**

**express.InlineShapes**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 本示例显示活动文档中形状和嵌入式形状的数目 */
Application.ActiveDocument.InlineShapes.Count
```

### Document.IsMasterDocument

如果指定的文档为一篇主文档，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.IsMasterDocument**

\*express是一个代表 Document 对象的变量。

**说明**

主文档包括一篇或多篇子文档。

**示例**

``` JavaScript
/*如果活动文档是一篇主文档，则本示例切换至主控文档视图，然后打开第一篇子文档。*/
function test()
{
if(ActiveDocument.IsMasterDocument == true) {
    ActiveDocument.ActiveWindow.View.Type = wdMasterView
    ActiveDocument.Subdocuments.Item(1).Open()
}
else {
    MsgBox("This document is not a master document.")
}
}
```

### Document.IsSubdocument

如果指定文档是主文档的子文档，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.IsSubdocument**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例判断 Sales.doc 是否为子文档，然后显示一条消息，表明 Sales.doc 的状态。*/
function test()
{
if(Documents.Item("Sales.doc").IsSubdocument == true) {
    MsgBox("Sales.doc is a subdocument.")
}
else {
    MsgBox("Sales.doc is not a subdocument.")
}
}
```

### Document.JustificationMode

返回或设置指定文档字符间距的调整量。可读/写 **WdJustificationMode** 类型。

**语法**

**express.JustificationMode**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在调整字符间距时只压缩标点符号。*/
ActiveDocument.JustificationMode = wdJustificationModeCompressKana
```

### Document.KerningByAlgorithm

如果 WPS 在指定文档中调整半角拉丁字符和标点符号的间距，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.KerningByAlgorithm**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 调整活动文档中的半角拉丁字符与标点符号的间距。*/
ActiveDocument.KerningByAlgorithm = true
```

### Document.Kind

返回或设置 WPS 在自动设置指定文档的格式时使用的格式类型。 **WdDocumentKind** 类型，可读写。

**语法**

**express.Kind**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例询问用户活动文档是否为电子邮件。如果是，则将此文档的格式设置为电子邮件格式。*/
function test()
{
let response = MsgBox("Is this document an e-mail message?", vbYesNo)
if(response == vbYes) {
    ActiveDocument.Kind = wdDocumentEmail
    ActiveDocument.Content.AutoFormat()
}
}
```

### Document.LanguageDetected

返回或设置一个指定 WPS 是否对指定文本的语言进行检测的值。 **Boolean** 类型，可读写。

**语法**

**express.LanguageDetected**

\*express是一个代表 Document 对象的变量。

**说明**

检查以前任何语言检测结果的 **LanguageID** 属性。

调用 **DetectLanguage** 方法时， **LanguageDetected** 属性会设置为 **True** 。要重新评估指定文本的语言，请将 **LanguageDetected** 属性设置为 **False** 。

**示例**

``` JavaScript
/*本示例检查活动文档，以确定其所用的语言类型并显示检查结果。*/
function test()
{
if(ActiveDocument.LanguageDetected == true) {
    x = MsgBox("This document has already " + "been checked. Do you want to check " + "it again?", vbYesNo)
    if(x == vbYes) {
        ActiveDocument.LanguageDetected = false
        ActiveDocument.DetectLanguage()
}
}
else {
    ActiveDocument.DetectLanguage()
}
if(ActiveDocument.Range.LanguageID == wdEnglishUS) {
    MsgBox("This is a U.S. English document.")
}
else {
    MsgBox("This is not a U.S. English document.")
}
}
```

### Document.ListParagraphs

返回一个 **ListParagraphs** 对象，该对象代表文档中的所有编号段落。只读。

**语法**

**express.ListParagraphs**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

``` JavaScript
/*此示例为第一个文档中的每个编号段落或有项目符号的段落添加黄色背景。*/
function test()
{
let numpar = Documents.Item(1).ListParagraphs
for(let i = 1; i <= numpar.Count; i++) {
    numpar.Item(i).Shading.BackgroundPatternColorIndex = wdYellow
}
}
```

### Document.Lists

返回一个 **Lists** 集合，该集合含有指定文档中的所有项目符号和编号列表。只读。

**语法**

**express.Lists**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例设置所选内容的格式为编号列表，然后在一个消息框中显示活动文档中列表数。*/
function test()
{
Selection.Range.ListFormat.ApplyListTemplate(ListGalleries.Item(wdNumberGallery).ListTemplates.Item(2))
MsgBox("This document has " + ActiveDocument.Lists.Count + " lists.")
}
```

``` JavaScript
/*本示例设置活动文档中第三个列表的格式为默认的项目符号列表格式。如果该列表已设置为项目符号列表格式，该示例清除此格式。*/
function test()
{
if(ActiveDocument.Lists.Count >= 3) {
    ActiveDocument.Lists.Item(3).Range.ListFormat.ApplyBulletDefault()
}
}
```

``` JavaScript
/*本示例在一个消息框中显示 MyLetter.doc 的每一列表的项目数。*/
function test()
{
let myDoc = Documents.Item("MyLetter.doc")
let i = myDoc.Lists.Count
for(let j = 1; j <= i; j++) {
    MsgBox("List " + i + " has " + myDoc.Lists.Item(j).CountNumberedItems() + " items.")
    i = i - 1
}
}
```

### Document.ListTemplates

返回一个 **ListTemplates** 集合，该集合代表指定文档中的所有列表格式。只读。

**语法**

**express.ListTemplates**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示活动文档中使用的列表模板数量。*/
MsgBox(ActiveDocument.ListTemplates.Count)
```

### Document.LockQuickStyleSet

返回或设置一个 **Boolean** ，它代表用户是否可以更改使用的快速样式集。可读写。

**语法**

**express.LockQuickStyleSet**

\*express是一个代表 Document 对象的变量。

### Document.LockTheme

返回或设置一个 **Boolean** 类型的值，该值代表用户是否可以更改文档主题。可读/写。

**语法**

**express.LockTheme**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **ApplyDocumentTheme** 方法可应用 WPS Office主题。使用 **DocumentTheme** 属性可返回应用于文档的 Office 主题。

### Document.MailEnvelope

返回一个 **MsoEnvelope** 对象，该对象代表文档的电子邮件标题。

**语法**

**express.MailEnvelope**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档的电子邮件标题设置批注。*/
function HeaderComments(){
    ActiveDocument.MailEnvelope.Introduction = "Please review this document and let me know " + "what you think.  I need your input by Friday." + "  Thanks."
}
```

### Document.Mailer

您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **Mailer** 属性的信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

**语法**

**express.Mailer**

\*express是一个代表 Document 对象的变量。

### Document.MailMerge

返回一个 **MailMerge** 对象，该对象代表指定文档的邮件合并功能。只读。

**语法**

**express.MailMerge**

\*express是一个代表 Document 对象的变量。

**说明**

不管指定文档是否是一个邮件合并主文档， **MailMerge** 对象都有效。用 **State** 属性可确定邮件合并操作当前的状态。

**示例**

``` JavaScript
/*如果活动文档是一个带有附加数据源的主文档，则本示例执行邮件合并。*/
function test()
{
let myMerge = ActiveDocument.MailMerge
if(myMerge.State == wdMainAndDataSource) {
    myMerge.Execute()
}
}
```

``` JavaScript
/*本示例将记录 1 到 4 与主文档合并，并将合并后的文档发送到打印机。*/
function test()
{
let mailmerge = ActiveDocument.MailMerge
    mailmerge.DataSource.FirstRecord = 1
    mailmerge.DataSource.LastRecord = 4
    mailmerge.Destination = wdSendToPrinter
    mailmerge.SuppressBlankLines = true
    mailmerge.Execute()

}
```

### Document.Name

返回指定文档的名称。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 返回当前文档的文件名 */
Application.ActiveDocument.FullName
```

### Document.NoLineBreakAfter

返回或设置 WPS 不在其后换行的首尾字符。可读/写 **String** 类型。

**语法**

**express.NoLineBreakAfter**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：在活动文档中，设定以“$”、“(”、“[”、“\”和“{”作为首尾字符时，WPS 不在其后面进行换行。*/
ActiveDocument.NoLineBreakAfter = "$([\{"
```

### Document.NoLineBreakBefore

返回或设置 WPS 不在其前换行的首尾字符。可读/写 **String** 类型。

**语法**

**express.NoLineBreakBefore**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档中设定，以“!”、“)”和“]”作为首尾字符时，WPS 不在这些首尾字符的前面进行换行。*/
ActiveDocument.NoLineBreakBefore = "!)]"
```

### Document.OMathBreakBin

返回或设置一个 **WdOMathBreakBin** 常量，该常量代表当公式跨越两行或多行时 WPS 用来放置二元运算符的位置。可读写。

**语法**

**express.OMathBreakBin**

\*express是一个代表 Document 对象的变量。

**说明**

如果公式在二元运算符（如加号、减号或乘号）处换行，则该运算符有三种不同的放置方式：换行符之前、换行符之后以及在换行符前后重复。

如果此属性设置为 **wdOMathBreakBinRepeat** ，则可以使用 **OMathBreakSub** 属性指定 WPS 如何处理位于换行符之前的减法运算符。

### Document.OMathBreakSub

返回或设置一个 **WdOMathBreakSub** 常量，该常量代表 WPS 如何处理位于换行符之前的减法运算符。可读写。

**语法**

**express.OMathBreakSub**

\*express是一个代表 Document 对象的变量。

**说明**

此属性仅在 **OMathBreakBin** 属性设置为 **wdOMathBreakBinRepeat** 时使用。当换行符刚好位于减法运算符之后，且文档设置为在下一行重复减法运算符时，由于负负得正的缘故，有时需要特殊对待减法运算符。有些作者选择将其中一个减号转变为加号，而有些作者则选择保留两个减号。

### Document.OMathFontName

返回或设置一个 **String** 类型的值，该值代表文档中用于显示公式的字体的名称。可读写。

**语法**

**express.OMathFontName**

\*express是一个代表 Document 对象的变量。

### Document.OMathIntSubSupLim

返回或设置一个 **Boolean** 类型的值，该值代表积分的默认极限位置。可读写。

**语法**

**express.OMathIntSubSupLim**

\*express是一个代表 Document 对象的变量。

### Document.OMathJc

返回或设置一个 **WdOMathJc** 常量，该常量代表一组公式的默认对齐方式：左对齐、右对齐、居中或作为一个组居中。可读写。

**语法**

**express.OMathJc**

\*express是一个代表 Document 对象的变量。

### Document.OMathLeftMargin

返回或设置一个代表公式的左边距的 **Single** 。可读写。

**语法**

**express.OMathLeftMargin**

\*express是一个代表 Document 对象的变量。

### Document.OMathNarySupSubLim

返回或设置一个 **Boolean** 类型的值，该值代表 *n* 元对象（而不是积分）的默认极限位置。可读写。

**语法**

**express.OMathNarySupSubLim**

\*express是一个代表 Document 对象的变量。

### Document.OMathRightMargin

返回或设置一个 **Single** 类型的值，该值代表公式的右边距。可读写。

**语法**

**express.OMathRightMargin**

\*express是一个代表 Document 对象的变量。

### Document.OMaths

返回一个 **OMaths** 集合，该集合代表指定区域内的 **OMath** 对象。只读。

**语法**

**express.OMaths**

\*express是一个代表 Document 对象的变量。

### Document.OMathSmallFrac

返回或设置一个 **Boolean** 类型的值，该值代表是否在文档所包含的公式中使用小型分数。可读写。

**语法**

**express.OMathSmallFrac**

\*express是一个代表 Document 对象的变量。

### Document.OMathWrap

返回或设置一个 **Single** 类型的值，该值代表换到新的一行的公式的第二行位置。可读写。

**语法**

**express.OMathWrap**

\*express是一个代表 Document 对象的变量。

**说明**

如果值为 -1，则指定公式的第二行为右对齐。如果为正值，则表示从左边距缩进。

### Document.OpenEncoding

返回用于打开指定文档的编码。 **MsoEncoding** 类型，只读。

**语法**

**express.OpenEncoding**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：测试当前文档是否使用编码 UTF7 打开。*/
function test()
{
if(ActiveDocument.OpenEncoding == msoEncodingUTF7){
    MsgBox("This is a UTF7-encoded text file!")
}
else{
    MsgBox("This is not a UTF7-encoded text file!")
}
}
```

### Document.OptimizeForWord97

如果 WPS 通过禁用所有不兼容的格式来优化当前文档，以便在 WPS 97 中查看，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.OptimizeForWord97**

\*express是一个代表 Document 对象的变量。

**说明**

若要为 WPS 97 按默认设置优化所有的新文档，可用 **OptimizeForWord97byDefault** 属性。

**示例**

``` JavaScript
/*本示例的功能如下：检查当前文档是否针对 WPS 97 进行过优化。如果没有，则询问用户是否进行优化。*/
function test()
{
if(ActiveDocument.OptimizeForWord97 == false) {
    let x = MsgBox("Is this document targeted at WPS 97 users?", jsYesNo)
    if (x == jsResultYes) {
        ActiveDocument.OptimizeForWord97 = true
    }
}
}
```

### Document.OriginalDocumentTitle

返回一个 **String** 类型的值，该值代表在运行“精确比较”文档比较功能之后原文档的文档标题。只读。

**语法**

**express.OriginalDocumentTitle**

\*express是一个代表 Document 对象的变量。

**说明**

要执行“精确比较”文档比较，请使用 **CompareDocuments** 方法。

### Document.PageSetup

返回一个与指定文档相关联的 **PageSetup** 对象。

**语法**

**express.PageSetup**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 本示例将活动文档的右边距设置为 72 磅（1 英寸） */
Application.ActiveDocument.PageSetup.RightMargin = Application.InchesToPoints(2)
```

### Document.Paragraphs

返回一个 **Paragraphs** 集合，该集合代表指定文档中的所有段落。只读。

**语法**

**express.Paragraphs**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 将当前文档首段设置为双倍行距 */
Application.ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceDouble
```

### Document.Parent

**语法**

**express.Parent**

\*express是一个代表 Document 对象的变量。

**说明**

返回一个 **Object** 类型值，该值代表指定 **Document** 对象的父对象。

### Document.Password

设置在打开文档必须使用的密码。 **String** 类型，只写。

**语法**

**express.Password**

\*express是一个代表 Document 对象的变量。

**说明**

**要点** ??请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

``` JavaScript
/*本示例首先打开文档 Earnings.doc 并为其设置密码，然后关闭该文档。*/
function test()
{
let myDoc = Documents.Open("C:\\My Documents\\Earnings.doc")
myDoc.Password = "strPassword"
myDoc.Close()
}
```

### Document.PasswordEncryptionAlgorithm

返回一个 **String** 类型的值，该值代表 WPS 用密码加密文档时所用的算法。只读。

**语法**

**express.PasswordEncryptionAlgorithm**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法指定 WPS 用密码加密文档时所用的算法。

**示例**

``` JavaScript
/*如果当前使用的密码加密算法是 WPS 97 for Windows 以前版本所使用的加密算法“OfficeXor”，则本示例将重新设置密码加密选项。*/
function PasswordSettings(){
    let ad = ActiveDocument
    if(ad.PasswordEncryptionAlgorithm == "OfficeXor"){
        ad.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)
    }
}
```

### Document.PasswordEncryptionFileProperties

如果 WPS 为密码保护的文档的文件属性加密，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.PasswordEncryptionFileProperties**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法指定 WPS 是否将密码保护的文档的文件属性加密。

**示例**

``` JavaScript
/*如果未对密码保护的文档的文件属性加密，本示例将设置密码加密选项。*/
function PasswordSettings(){
    let ad = ActiveDocument
    if(ad.PasswordEncryptionFileProperties == false){
        ad.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)
    }
}
```

### Document.PasswordEncryptionKeyLength

返回一个 **Long** 类型的值，该值代表 WPS 用密码加密文档时所用的密码键长。只读。

**语法**

**express.PasswordEncryptionKeyLength**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法指定 WPS 用密码加密文档时所用的密码键长。

**示例**

``` JavaScript
/*如果密码加密键长小于 40，则本示例将设置密码加密选项。*/
function PasswordSettings(){
    let ad = ActiveDocument
    if(ad.PasswordEncryptionKeyLength < 40){
        ad.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)
    }
}
```

### Document.PasswordEncryptionProvider

返回一个 **String** 类型的值，该值指定 WPS 用密码加密文档时所用的加密算法提供程序的名称。只读。

**语法**

**express.PasswordEncryptionProvider**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法指定 WPS 用密码加密文档时所用的加密算法提供程序的名称。

**示例**

``` JavaScript
/*如果当前使用的密码加密算法提供程序不是“Microsoft RSA SChannel Cryptographic Provider”，则本示例将设置密码加密选项。*/
function PasswordSettings(){
    let ad = ActiveDocument
    if(ad.PasswordEncryptionProvider != "Microsoft RSA SChannel Cryptographic Provider"){
        ad.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)
    }
}
```

### Document.Path

返回文档的磁盘或 Web 路径。只读 **String** 类型。

**语法**

**express.Path**

\*express是一个代表 Document 对象的变量。

**说明**

此路径不包括尾随字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符。使用 **Name** 属性可返回不带路径的文件名，而使用 **FullName** 属性可同时返回文件名和路径。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>可以使用 <strong>PathSeparator</strong> 属性建立 Web 地址，即使地址中包含正斜杠 (/)（ <strong>PathSeparator</strong> 属性默认使用反斜杠 (\)）。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

**示例**

``` JavaScript
/*本示例显示活动文档的路径和文件名。*/
MsgBox(ActiveDocument.Path + Application.PathSeparator + ActiveDocument.Name)
```

``` JavaScript
/*本示例将当前文件夹改为附加到活动文档的模板的路径。*/
ChDir ActiveDocument.AttachedTemplate.Path
```

### Document.Permission

返回一个 **Permission** 对象，该对象代表指定文档中的权限设置。

**语法**

**express.Permission**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*下例返回活动文档的权限设置。*/
let objPermission = ActiveDocument.Permission
```

### Document.PrintFormsData

如果 WPS 仅将相应联机窗体中的键入的数据打印在预打印窗体上，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintFormsData**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 仅打印联机窗体域的数据内容，然后开始打印活动文档。*/
function test()
{
ActiveDocument.PrintFormsData = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“选项”对话框“打印”选项卡上“只用于当前文档的选项”区域中“仅打印窗体域的内容”复选框当前的状态。*/
let temp = ActiveDocument.PrintFormsData
```

### Document.PrintFractionalWidths

如果设置指定文档的格式，使用分式间距来显示和打印字符，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintFractionalWidths**

\*express是一个代表 Document 对象的变量。

**说明**

**PrintFractionalWidths** 属性只适用于 Apple Macintosh 计算机。在 Windows 中， **PrintFractionalWidths** 属性始终返回 **False** 。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 的语言参考帮助。

### Document.PrintPostScriptOverText

如果为 **True** ，则使用 PostScript 打印机时，文档中的 PRINT 域指令（例如，PostScript 命令）将打印在文字和图形的上方。 **Boolean** 类型，可读写。

**语法**

**express.PrintPostScriptOverText**

\*express是一个代表 Document 对象的变量。

**说明**

**PrintPostScriptOverText** 属性控制是否将 Macintosh 文档中的 postscript 代码以转换的 WPS 形式打印。如果文档不包含 PRINT 域，则该属性无效。

**示例**

``` JavaScript
/*本示例设置 WPS 可在文字和图形的上方打印 PRINT 域指令，然后打印该活动文档。*/
function test()
{
ActiveDocument.PrintPostScriptOverText = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“选项”对话框“打印”选项卡上“打印选项”区域中的“在文本上方打印 postscript”复选框当前的状态。*/
let currSet = ActiveDocument.PrintPostScriptOverText
```

### Document.PrintRevisions

如果为 **True** ，则在打印文档的同时打印修订标记。如果返回 **False** ，则不打印修订标记（即打印接受修订后的状态）。 **Boolean** 类型，可读写。

**语法**

**express.PrintRevisions**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例在打印活动文档时不带有修订标记。*/
function test()
{
ActiveDocument.PrintRevisions = false
ActiveDocument.PrintOut()
}
```

### Document.Protect

返回或设置与指定传送名单相关联的文档的保护类型。 **WdProtectionType** 类型，可读写。

**语法**

**express.Protect**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例先指定用于活动文档（只允许批注）的保护类型，然后传送该文档。*/
function test()
{
ActiveDocument.HasRoutingSlip = true
let RoutingSlip = ActiveDocument.RoutingSlip
    RoutingSlip.Subject = "Status Doc"
    RoutingSlip.Protect = wdAllowOnlyComments
    RoutingSlip.AddRecipient("Kim Johnson")
ActiveDocument.Route()
}
```

### Document.ProtectionType

该属性返回指定文档的保护类型。可以是下列 **WdProtectionType** 常量之一： **wdAllowOnlyComments** 、 **wdAllowOnlyFormFields** 、 **wdAllowOnlyReading** 、 **wdAllowOnlyRevisions** 或 **wdNoProtection** 。

**语法**

**express.ProtectionType**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*如果没有保护活动文档，该示例保护文档，使审阅者只能插入批注而不改变文档内容。*/
function test()
{
if(ActiveDocument.ProtectionType == wdNoProtection) {
    ActiveDocument.Protect(wdAllowOnlyComments)
}
}
```

``` JavaScript
/*如果活动文档受到保护，则本示例解除文档的保护。*/
function test()
{
let doc = ActiveDocument
if(doc.ProtectionType != wdNoProtection) {
    doc.Unprotect()
}
}
```

### Document.ReadabilityStatistics

返回一个 **ReadabilityStatistics** 集合，该集合代表指定文档或区域的可读性统计信息。只读。

**语法**

**express.ReadabilityStatistics**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示文档 1 的每一种可读性统计信息及其值。*/
function test()
{
let rs = Documents.Item(1).ReadabilityStatistics
for(let i = 1; i <= rs.Count; i++) {
    MsgBox(rs.Item(i).Name + " - " + rs.Item(i).Value)
}
}
```

### Document.ReadingLayoutSizeX

设置或返回一个 **Long** 类型的值，该值表示当在阅读版式视图中显示文档并对其冻结了输入手写标记时文档中页面的宽度。

**语法**

**express.ReadingLayoutSizeX**

\*express是一个代表 Document 对象的变量。

**说明**

设置完 **ReadingLayoutSizeX** 和 **ReadingLayoutSizeY** 属性后，可使用 **ReadingModeLayoutFrozen** 属性以指定的高度和宽度显示页面。使用 **ReadingLayout** 属性可在阅读版式视图中显示文档。

**示例**

``` JavaScript
/*以下示例在阅读版式视图中显示活动文档，然后设置所显示页面的大小。*/
function test()
{
ActiveWindow.View.ReadingLayout = true
ActiveDocument.ReadingLayoutSizeX = 300
ActiveDocument.ReadingLayoutSizeY = 300
ActiveDocument.ReadingModeLayoutFrozen = true
}
```

### Document.ReadingLayoutSizeY

设置或返回一个 **Long** 类型的值，该值表示当在阅读版式视图中显示文档并对其冻结了输入手写标记时文档中页面的高度。

**语法**

**express.ReadingLayoutSizeY**

\*express是一个代表 Document 对象的变量。

**说明**

设置完 **ReadingLayoutSizeX** 和 **ReadingLayoutSizeY** 属性后，可使用 **ReadingModeLayoutFrozen** 属性以指定的高度和宽度显示页面。使用 **ReadingLayout** 属性可在阅读版式视图中显示文档。

**示例**

``` JavaScript
/*以下示例在阅读版式视图中显示活动文档，然后设置所显示页面的大小。*/
function test()
{
ActiveWindow.View.ReadingLayout = true
ActiveDocument.ReadingLayoutSizeX = 300
ActiveDocument.ReadingLayoutSizeY = 300
ActiveDocument.ReadingModeLayoutFrozen = true
}
```

### Document.ReadingModeLayoutFrozen

设置或返回一个 **Boolean** 类型的值，该值表示是否将阅读版式视图中显示的页面冻结为指定大小以向文档插入手写标记。

**语法**

**express.ReadingModeLayoutFrozen**

\*express是一个代表 Document 对象的变量。

**说明**

使用 **ReadingLayoutSizeX** 和 **ReadingLayoutSizeY** 属性可指定将阅读版式大小冻结以向文档插入手写标记时所显示的页面的大小。

**示例**

``` JavaScript
/*以下示例在阅读版式视图中显示活动文档，然后设置所显示的页面的大小。*/
function test()
{
ActiveWindow.View.ReadingLayout = true
ActiveDocument.ReadingLayoutSizeX = 300
ActiveDocument.ReadingLayoutSizeY = 300
ActiveDocument.ReadingModeLayoutFrozen = true
}
```

### Document.ReadOnly

如果对文档所作修改不能保存到原始文档中，则该属性为 **True** 。只读 **Boolean** 类型。

**语法**

**express.ReadOnly**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档不是只读文档时保存该文档。*/
function test()
{
if(ActiveDocument.ReadOnly == false) {
    ActiveDocument.Save()
}
}
```

### Document.ReadOnlyRecommended

如果无论用户何时打开文档，WPS 都显示一个消息框，提示以只读方式打开，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ReadOnlyRecommended**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例的功能如下：在打开文档时， WPS 建议以只读方式打开。*/
ActiveDocument.ReadOnlyRecommended = true
```

### Document.RemoveDateAndTime

设置或返回一个 **Boolean** 类型的值，该值指示文档是否存储修订的日期和时间元数据。

**语法**

**express.RemoveDateAndTime**

\*express是一个代表 Document 对象的变量。

**说明**

如果为 **True** ，则删除修订的日期和时间戳信息。如果为 **False** ，则不删除修订的日期和时间戳信息。将 **RemoveDateAndTime** 属性与 **RemovePersonalInformation** 属性结合使用，可帮助删除文档属性中的个人信息。

**示例**

``` JavaScript
/*以下示例从活动文档中删除个人信息，并从文档的所有修订中删除日期和时间信息。*/
function test()
{
ActiveDocument.RemovePersonalInformation = true
ActiveDocument.RemoveDateAndTime = true
}
```

### Document.RemovePersonalInformation

如果 WPS 在保存文档时从批注、修订和“属性”对话框中删除所有用户信息，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.RemovePersonalInformation**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设置当前文档，在用户下一次保存文档时删除所有的个人信息。*/
function RemovePersonalInfo() {
    ActiveDocument.RemovePersonalInformation = true
}
```

### Document.Research

返回一个 **Research** 对象，该对象代表文档的检索服务。只读。

**语法**

**express.Research**

\*express是一个代表 Document 对象的变量。

### Document.RevisedDocumentTitle

返回一个 **String** 类型的值，该值代表在运行“精确比较”文档比较功能之后修订的文档的文档标题。只读。

**语法**

**express.RevisedDocumentTitle**

\*express是一个代表 Document 对象的变量。

**说明**

要执行“精确比较”文档比较，请使用 **CompareDocuments** 方法。

### Document.Revisions

返回一个 **Revisions** 集合，该集合代表文档或区域中的修订。只读。

**语法**

**express.Revisions**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 本示例显示活动文档第一节中的修订数 */
Application.ActiveDocument.Sections.Item(1).Range.Revisions.Count
```

### Document.Saved

如果指定的文档或模板从上次保存后一直没有更改，则该属性值为 **True** 。如果关闭文档时，WPS 提示保存对文档所做的更改，则该属性值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.Saved**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档含有以前未保存的更改的情况下保存该文档。*/
function test()
{
if(ActiveDocument.Saved == false) {
    ActiveDocument.Save()
}
}
```

### Document.SaveEncoding

返回或设置保存文档时使用的编码。 **MsoEncoding** 类型，可读写。

**语法**

**express.SaveEncoding**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：指定用公历编码保存当前文档。*/
ActiveDocument.SaveEncoding = msoEncodingWestern
```

### Document.SaveFormat

返回指定文档或文件转换器的文件格式。只读 **Long** 类型。

**语法**

**express.SaveFormat**

\*express是一个代表 Document 对象的变量。

**说明**

**SaveFormat** 属性将是一个指定外部文件转换器的唯一数字或一个 **WdSaveFormat** 常量。

如果对 **SaveAs** 方法的 *FileFormat* 参数使用 **SaveFormat** 属性的值，可以将文档保存为没有对应 **WdSaveFormat** 常量的某种文件格式。

**示例**

``` JavaScript
/*如果活动文档为 RTF 格式的文档，本示例将其另存为 WPS 文档。*/
function test()
{
if(ActiveDocument.SaveFormat == wdFormatRTF) {
    ActiveDocument.SaveAs(null, wdFormatDocument)
}
}
```

### Document.SaveFormsData

如果 WPS 将输入到一个窗体中的数据作为以制表位分隔的数据库记录保存，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SaveFormsData**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 只保存输入到一个窗体中的数据。*/
ActiveDocument.SaveFormsData = true
```

``` JavaScript
/*本示例返回“选项”对话框“保存”选项卡上“保存选项”区域中“仅保存窗体域内容”复选框当前的状态。*/
let temp = ActiveDocument.SaveFormsData
```

### Document.SaveSubsetFonts

如果 WPS 将嵌入的 TrueType 字体子集和文档一起保存，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SaveSubsetFonts**

\*express是一个代表 Document 对象的变量。

**说明**

如果在一个文档中使用的 TrueType 字体少于 32 个字符，则 WPS 将在文档中嵌入 TrueType 字体子集（只包含所用的字体）；如果使用的 TrueType 字体多于 32 个字符，则 WPS 在文档中嵌入整个 TrueType 字体。

**示例**

``` JavaScript
/*本示例设置一个名为“MyDoc”的文档只保存嵌入的 TrueType 字体子集（如果使用很少的字符），然后再保存“MyDoc”。*/
function test()
{
let md = Documents.Item("MyDoc")
    md.EmbedTrueTypeFonts = true
    md.SaveSubsetFonts = true
    md.Save()

}
```

### Document.Scripts

返回一个 **Scripts** 集合，该集合代表指定对象中的 HTML 脚本的集合。

**语法**

**express.Scripts**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例显示活动文档的第一个 Script 对象中的文本。*/
Debug.Print(ActiveDocument.Scripts.Item(1).ScriptText)
```

### Document.Sections

返回一个 **Section** 集合，该集合代表指定文档中的节。只读。

**语法**

**express.Sections**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 本示例设置活动文档中所有节的页面方向 */
function test() {
    let sec = Application.ActiveDocument.Sections
    for(let i = 1; i<= sec.Count; i++) {
        sec.Item(i).PageSetup.Orientation = wdOrientLandscape
    }
}
```

### Document.Sentences

返回一个 **Sentences** 集合，该集合代表文档中的所有句子。只读。

**语法**

**express.Sentences**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例复制活动文档中的第一个句子。*/
ActiveDocument.Sentences.Item(1).Copy()
```

``` JavaScript
/*本示例删除活动文档中的最后一个句子。*/
ActiveDocument.Sentences.Last.Delete()
```

### Document.ServerPolicy

返回一个 **ServerPolicy** 对象，该对象代表为运行 SharePoint Server 2007 的服务器上存储的文档指定的策略。只读。

**语法**

**express.ServerPolicy**

\*express是一个代表 Document 对象的变量。

### Document.Shapes

返回一个 **Shapes** 集合，该集合代表指定文档中的所有 **Shape** 对象。只读。

**语法**

**express.Shapes**

\*express是一个代表 Document 对象的变量。

**说明**

该集合可以包含绘图、形状、图片、OLE 对象、ActiveX 控件、文本对象和标注。有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**Shapes** 属性应用于文档时将返回文档主体部分的所有 **Shape** 对象，不包括页眉和页脚。

**示例**

``` JavaScript
/* 本示例新建一篇文档，为其添加一个矩形 */
Application.Documents.Add().Shapes.AddShape(msoShapeRectangle, 5, 25, 100, 50)
```

### Document.ShowGrammaticalErrors

如果指定文档中的语法错误用绿色波浪线标记，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowGrammaticalErrors**

\*express是一个代表 Document 对象的变量。

**说明**

只有将 **CheckGrammarAsYouType** 属性设为 **True** ，才能查看文档中的语法错误。

**示例**

``` JavaScript
/*本示例对 WPS 进行设置，使其在您键入时检查活动文档中的语法错误并显示发现的所有错误。*/
function test()
{
Options.CheckGrammarAsYouType = true
ActiveDocument.ShowGrammaticalErrors = true
}
```

### Document.ShowRevisions

如果在屏幕上显示对指定文档的修订，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowRevisions**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例对活动文档进行设置，以便能标记对该文档的修改，并在屏幕上显示修订标记。*/
function test()
{
ActiveDocument.TrackRevisions = true
ActiveDocument.ShowRevisions = true

}
```

### Document.ShowSpellingErrors

如果 WPS 为文档中的拼写错误添加下划线，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowSpellingErrors**

\*express是一个代表 Document 对象的变量。

**说明**

必须将 **CheckSpellingAsYouType** 属性设置为 **True** ，才能查看文档中的拼写错误。

**示例**

``` JavaScript
/*本示例设置 WPS 隐藏活动文档中的拼写错误标记，该标记为红色波浪线。*/
ActiveDocument.ShowSpellingErrors = false
```

``` JavaScript
/*本示例显示活动文档中的拼写错误标记。*/
function test()
{
Options.CheckSpellingAsYouType = true
ActiveDocument.ShowSpellingErrors = true
}
```

``` JavaScript
/*本示例返回“选项”对话框中“拼写和语法”选项卡的“拼写”区域中的“隐藏文档中的拼写错误”复选框的当前状态。*/
let temp = ActiveDocument.ShowSpellingErrors
```

### Document.Signatures

返回一个 **SignatureSet** 集合，该集合代表文档的数字签名。

**语法**

**express.Signatures**

\*express是一个代表 Document 对象的变量。

**说明**

若要对 WPS 文档进行数字签名，并验证文档中的其他签名，您需要 Microsoft CryptoAPI 和唯一的数字签名证书。CryptoAPI 随着 Microsoft Internet Explorer 4.01 或更高版本一起安装。可从证书颁发机构获得数字签名证书。

**示例**

``` JavaScript
/*本示例显示“签名”对话框，可在其中为文档添加数字签名。*/
function AddSignature() {
    ActiveDocument.Signatures.Add()
}
```

### Document.SmartDocument

返回一个 **SmartDocument** 对象，该对象代表智能文档解决方案的设置。

**语法**

**express.SmartDocument**

\*express是一个代表 Document 对象的变量。

**说明**

有关智能文档的详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document Software Development Kit。

**示例**

``` JavaScript
/*下列示例显示一个对话框，其中包含用于活动文档的有效 XML 扩展包列表。*/
ActiveDocument.SmartDocument.PickSolution()
```

### Document.SmartTagsAsXMLProps

如果该属性的值为 **True** ，当将包含智能标记的文档保存为 HTML 格式时，WPS 会创建一个包含智能标记信息的 XML 页眉。 **Boolean** 类型，可读写。

**语法**

**express.SmartTagsAsXMLProps**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*如果活动文档被保存为 HTML 格式，本示例可以在 XML 页眉中保存智能标记信息。*/
function SaveXMLForSmartTags() {
    ActiveDocument.SmartTagsAsXMLProps = true
}
```

### Document.SnapToGrid

如果在指定文档中绘制、移动自选图形/东亚字符或者调整它们的大小时，自选图形/东亚字符自动与不可见的网格线对齐，则该属性为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.SnapToGrid**

\*express是一个代表 Document 对象的变量。

**说明**

绘制、移动自选图形或调整其大小时，按下 Alt 可以暂时使该设置无效。

**示例**

``` JavaScript
/*本示例设置 WPS ，使当前文档中的东亚字符自动与不可见的网格线对齐。*/
ActiveDocument.SnapToGrid = true
```

### Document.SnapToShapes

如果 WPS 在指定文档中使自选图形或东亚字符自动与不可见网格线对齐，这些网格线穿过其他自选图形或东亚字符的水平和垂直边缘，则该属性为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.SnapToShapes**

\*express是一个代表 Document 对象的变量。

**说明**

该属性为每个自选图形创建附加的不可见网格线。 **SnapToShapes** 属性与 **SnapToGrid** 属性独立工作。

**示例**

``` JavaScript
/*本示例设置 WPS ，使当前文档中的中文字符自动与不可见的网格线对齐，这些网格线穿过其他中文字符的水平和垂直边缘。*/
ActiveDocument.SnapToShapes = true
```

### Document.SpellingChecked

如果已对指定的区域或文档完成拼写检查，则该属性值为 **True** 。如果所有或部分区域或文档尚未进行拼写检查，则该属性值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.SpellingChecked**

\*express是一个代表 Document 对象的变量。

**说明**

若要重新检查某一区域或文档的拼写，请将 **SpellingChecked** 属性设置为 **False** 。

若要查看该区域或文档中是否含有拼写错误，请使用 **SpellingErrors** 属性。

**示例**

``` JavaScript
/*本示例将“MyDocument.doc”的 SpellingChecked 属性设置为 False，然后对该文档进行另一轮拼写检查。*/
function test()
{
Documents.Item("MyDocument.doc").SpellingChecked = false
Documents.Item("MyDocument.doc").CheckSpelling(null,false)
}
```

### Document.SpellingErrors

返回一个 **ProofreadingErrors** 集合，该集合代表在指定文档或区域中标识为拼写错误的单词。只读。

**语法**

**express.SpellingErrors**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例检查活动文档的拼写错误，并显示找到的错误数目。*/
function test()
{
let myErr = ActiveDocument.SpellingErrors.Count
if(myErr == 0) {
    MsgBox ("No spelling errors found.")
}
else{
    MsgBox(myErr + " spelling errors found.")
}
}
```

### Document.StoryRanges

返回一个 **StoryRanges** 集合，该集合代表指定文档中的所有的文字部分。只读。

**语法**

**express.StoryRanges**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例检查 StoryRanges 集合中的每一个元素，判定 wdPrimaryFooterStory 是否为 StoryRanges 集合的一部分。*/
function test()
{
let aStory = ActiveDocument.StoryRanges
for(let i = 1; i <= aStory.Count; i++) {
    if(aStory.Item(i).StoryType == wdEvenPagesFooterStory) {
        MsgBox("Document includes an even page footer")
    }
}
}
```

``` JavaScript
/*本示例将文字添至文档页眉部分，然后显示该文字。*/
function test()
{
ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Range.Text = "Header text"
MsgBox(ActiveDocument.StoryRanges.Item(wdPrimaryHeaderStory).Text)
}
```

### Document.Styles

本示例返回一个指定文档的 **Styles** 集合。只读。

**语法**

**express.Styles**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例将活动文档中所有以“Chapter”开始的段落样式设置为“标题 1”。*/
function test()
{
let para = ActiveDocument.Paragraphs
for(let i = 1; i <= para.Count; i++) {
    if(para.Item(i).Range.Words.Item(1).Text == "Chapter ") {
       para.Item(i).Style = ActiveDocument.Styles.Item(wdStyleHeading1)
    }
}
}
```

``` JavaScript
/*本示例打开活动文档所用的模板，并修改“标题 1”样式，然后保存模板，更新活动文档中所有的样式。*/
function test()
{
let tempDoc = ActiveDocument.AttachedTemplate.OpenAsDocument()
let font = tempDoc.Styles.Item(wdStyleHeading1).Font
    font.Name = "Arial"
    font.Size = 16
tempDoc.Close(wdSaveChanges)
ActiveDocument.UpdateStyles()
}
```

### Document.StyleSortMethod

返回或设置一个 **WdStyleSort** 常量，该常量代表要在对 **“样式”** 任务窗格中的样式进行排序时使用的排序方法。可读写。

**语法**

**express.StyleSortMethod**

\*express是一个代表 Document 对象的变量。

### Document.Subdocuments

返回一个 **Subdocuments** 集合，该集合代表指定文档中的所有子文档。只读。

**语法**

**express.Subdocuments**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示嵌入在活动文档中的子文档的数量。*/
MsgBox(ActiveDocument.Subdocuments.Count)
```

``` JavaScript
/*本示例显示活动文档中每篇子文档的路径和文件名。*/
function test()
{
let subdoc = ActiveDocument.Subdocuments
for(let i = 1; i <= subdoc.Count; i++) {
    if(subdoc.Item(i).HasFile == true) {
        MsgBox(subdoc.Item(i).Path + Application.PathSeparator + subdoc.Item(i).Name)
    }
    else {
        MsgBox("This subdocument has not been saved.")
    }   
}
}
```

### Document.Sync

返回一个 **Sync** 对象，该对象提供对构成“文档工作区”的文档的方法和属性的访问。

**语法**

**express.Sync**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*如果活动文档是“文档工作区”中的共享文档，则以下示例显示活动文档最后一个修改者的姓名。*/
function test()
{
let strLastUser
let eStatus = ActiveDocument.Sync.Status

if(eStatus == msoSyncStatusLatest) {
    strLastUser = ActiveDocument.Sync.WorkspaceLastChangedBy
    MsgBox("You have the most up-to-date copy." + "This file was last modified by " + strLastUser)
}
}
```

### Document.Tables

返回一个 **Table** 集合，该集合代表指定文档中的所有表格。只读。

**语法**

**express.Tables**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 本示例在活动文档第一个表格的第一列中插入数字和文本 */
function test() {
    let num = 90
    let acell = Application.ActiveDocument.Tables.Item(1).Columns.Item(1).Cells
    for(let i = 1; i <= acell.Count; i++) {
        acell.Item(i).Range.Text = num + " Sales"
        num = num + 1
    }
}
```

### Document.TablesOfAuthorities

返回一个 **TableOfAuthorities** 集合，该集合代表指定文档中的引文目录。只读。

**语法**

**express.TablesOfAuthorities**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例在文件 Sales.doc 的起始处添加一个引文目录。在该引文目录中汇编了所有类别的引文。*/
function test()
{
let myRange = Documents.Item("Sales.doc").Range(0, 0)
Documents.Item("Sales.doc").TablesOfAuthorities.Add(myRange, 0, null, true, null, null, null, null, null, true)
}
```

``` JavaScript
/*本示例将更新活动文档中的所有引文目录。*/
function test()
{
let myTOA = ActiveDocument.TablesOfAuthorities
for(let i = 1; i <= myTOA.Count; i++) {
    myTOA.Item(i).Update()
}
}
```

### Document.TablesOfAuthoritiesCategories

返回一个 **TablesOfAuthoritiesCategories** 集合，该集合代表指定文档中可用的引文目录类别。只读。

**语法**

**express.TablesOfAuthoritiesCategories**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例更改活动文档的引文目录类别列表中的八个元素的名称。*/
ActiveDocument.TablesOfAuthoritiesCategories.Item(8).Name = "Other case"
```

``` JavaScript
/*本示例在最后一个引文目录类别的名称发生了变化的情况下，显示该类别的名称。*/
function test()
{
let last = ActiveDocument.TablesOfAuthoritiesCategories.Count
if(ActiveDocument.TablesOfAuthoritiesCategories.Item(last).Name != "16") {
    MsgBox(ActiveDocument.TablesOfAuthoritiesCategories.Item(last).Name)
}
}
```

### Document.TablesOfContents

返回一个 **TablesOfContents** 集合，该集合代表指定文档中的目录。只读。

**语法**

**express.TablesOfContents**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例在 Sales.doc 文档的开始处添加目录。目录从 TC 域收集项目文本。*/
function test()
{
let myRange = Documents.Item("Sales.doc").Range(0, 0)
Documents.Item("Sales.doc").TablesOfContents.Add(myRange, false, null, null, true)
 


}
```

``` JavaScript
/*本示例更新活动文档的目录中的各项目页码。*/
function test()
{
let myTOC = ActiveDocument.TablesOfContents
for(let i = 1; i <= myTOC.Count; i++) {
    myTOC.Item(i).UpdatePageNumbers()
}
}
```

### Document.TablesOfFigures

返回一个 **TablesOfFigures** 集合，该集合代表指定文档中的图表目录。只读。

**语法**

**express.TablesOfFigures**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例在活动文档的插入点位置添加图表目录。*/
function test()
{
Selection.Collapse(wdCollapseStart)
ActiveDocument.TablesOfFigures.Add(Selection.Range, wdCaptionFigure)
}
```

``` JavaScript
/*本示例更新 Report.doc 文档中第一个图表目录的内容。*/
Documents.Item("Report.doc").TablesOfFigures.Item(1).Update()
```

### Document.TextEncoding

返回或设置代码页或字符集，WPS 应用该代码页或字符集将文档保存为编码文本文件。 **MsoEncoding** 类型，可读写。

**语法**

**express.TextEncoding**

\*express是一个代表 Document 对象的变量。

**说明**

**TextEncoding** 属性用不同于 HTML 编码的方式设置文本编码，用户可以用 **Encoding** 属性设置 HTML 编码。如果要对保存为文本文件的所有文档设置文本编码，可使用 **DefaultTextEncoding** 属性。

**示例**

``` JavaScript
/*如果把活动文档保存为文本文件，本示例将设置活动文档的文本编码为日语。*/
function EncodeText() {
    ActiveDocument.TextEncoding = msoEncodingJapaneseShiftJIS
}
```

### Document.TextLineEnding

返回或设置一个 **WdLineEndingType** 常量，该常量代表 WPS 在保存为文本文件的文档中标记直线和段落分隔符的方式。可读写。

**语法**

**express.TextLineEnding**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*将活动文档保存为文本文件时，本示例将为该活动文档中的直线或段落分隔符输入回车。*/
function LineEndings() {
    ActiveDocument.TextLineEnding = wdCROnly
}
```

### Document.TrackFormatting

返回或设置 **Boolean** 值，该值表示在打开更改跟踪功能时是否跟踪格式更改。可读/写。

**语法**

**express.TrackFormatting**

\*express是一个代表 Document 对象的变量。

### Document.TrackMoves

返回或设置 **Boolean** 值，该值表示在打开跟踪更改功能时是否对移动文本进行标记。可读/写。

**语法**

**express.TrackMoves**

\*express是一个代表 Document 对象的变量。

**说明**

默认情况下，如果打开跟踪更改功能，则移动文本将标记为删除和插入。如果 **TrackMoves** 为 **True** ，则移动文本将标记为移动，并以删除线格式来标记原始文本。此属性对应于 **“跟踪更改选项”** 对话框中的 **“跟踪移动”** 复选框。

### Document.TrackRevisions

如果标记对指定文档的修改，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.TrackRevisions**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例对活动文档进行设置，以便能标记对该文档的修改，并在屏幕上显示修订标记。*/
function test()
{
ActiveDocument.TrackRevisions = true
ActiveDocument.ShowRevisions = true
}
```

``` JavaScript
/*本示例将在没有启用修订标记的情况下插入文本。*/
function test()
{
if(ActiveDocument.TrackRevisions == false) {
    Selection.InsertBefore("new text")
}
}
```

### Document.Type

返回文档的类型（模板或文档）。只读 **WdDocumentType** 类型。

**语法**

**express.Type**

\*express是一个代表 Document 对象的变量。

### Document.UpdateStylesOnOpen

如果在每次打开指定文档时会依照其附属的模板更新文档样式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UpdateStylesOnOpen**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例将启用所有打开文档的更新文档样式选项，然后关闭这些文档。当重新打开这些文档时，对该文档的附属模板中样式的修改将自动显示在文档中。*/
function test()
{
let doc = Documents
for(let i = 1; i <= doc.Count; i++) {
    doc.Item(i).UpdateStylesOnOpen = true
    doc.Item(i).Close(wdSaveChanges)
}
}
```

``` JavaScript
/*本示例将禁用更新文档样式的选项，这样，对文档 Report.doc 的附属模板的样式进行的修改就不会反映在该文档中。*/
function test()
{
Documents.Item("Report.doc").UpdateStylesOnOpen = false
}
```

### Document.UseMathDefaults

返回或设置一个代表在创建新的公式时是否使用默认的数学设置的 **Boolean** 。可读写。

**语法**

**express.UseMathDefaults**

\*express是一个代表 Document 对象的变量。

### Document.UserControl

如果文档是由用户创建或打开的，则该属性为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.UserControl**

\*express是一个代表 Document 对象的变量。

**说明**

如果文档是由其他 WPS Office应用程序用 **Open** 方法或 Visual Basic **CreateObject** 或 **GetObject** 命令以编程方式创建或打开的，则该属性返回 **False** 。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>如果 WPS 对用户可见，或者您在 WPS 代码模块内部调用 <strong>UserControl</strong> 属性，则该属性将始终返回 <strong>True</strong> 。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

**示例**

``` JavaScript
/*该示例显示活动文档的 UserControl 属性的状态。本示例仅在从装载 WPS 对象库的其他 Office 应用程序运行的情况下有效。*/
function test()
{
Set wd = New Kwps.application
Set wdDoc = _
    wd.Documents.Open("C:\My Documents\doc1.doc")
If wdDoc.UserControl = True Then
    MsgBox "This document was created or opened by the user."
Else
    MsgBox "This document was created programmatically."
End If
}
```

### Document.Variables

返回一个 **Variables** 集合，该集合代表指定文档所存的变量。只读。

**语法**

**express.Variables**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/* 本示例为活动文档添加一个名为“Value1”的文档变量 */
Application.ActiveDocument.Variables.Add('num', 3)
```

### Document.VBASigned

如果指定文档的 示例代码 (VBA) 项目具有数字签名，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.VBASigned**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例加载一个名为“Temp.doc”的文档，并测试该文档是否具有数字签名。如果没有数字签名，本示例将显示警告信息。*/
function test()
{
Documents.Open("C:\\My Documents\\Temp.doc")
if(ActiveDocument.VBASigned == false) {
    MsgBox("Warning! This document has not been digitally signed.", vbCritical, "Digital Signature Warning")
}
}
```

### Document.VBProject

返回指定模板或文档的 **VBProject** 对象。

**语法**

**express.VBProject**

\*express是一个代表 Document 对象的变量。

**说明**

使用该属性可访问代码模块和用户窗体。

要在对象浏览器中查看 **VBProject** 对象，必须在“Visual Basic 编辑器”的 **“引用”** 对话框中选定 **“示例代码 Extensibility”** 复选框（在 **“工具”** 菜单上）。

**示例**

``` JavaScript
/*本示例显示活动文档的 Visual Basic 项目的名称。*/
function test()
{
Set currProj = ActiveDocument.VBProject
MsgBox currProj.Name
}
```

``` JavaScript
/*本示例将一个标准代码模块添至活动文档，并将其重命名为“MyModule”。*/
function test()
{
Set newModule = ActiveDocument.VBProject.VBComponents _
    .Add(vbext_ct_StdModule)
NewModule.Name = "MyModule"
}
```

### Document.WebOptions

将文档保存为网页或打开网页时，该属性将返回 **WebOptions** 对象，该对象包含 WPS 所用的文档属性。只读。

**语法**

**express.WebOptions**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*本示例指定将活动文档的项目保存为网页时，使用级联样式表和西文文档编码。*/
function test()
{
let objWO = ActiveDocument.WebOptions
objWO.RelyOnCSS = true
objWO.Encoding = msoEncodingWestern
}
```

### Document.Windows

返回一个 **Windows** 集合，该集合代表指定文档的所有窗口。只读。

**语法**

**express.Windows**

\*express是一个代表 Document 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示 NewWindow 方法运行前后活动文档的窗口数目。*/
function test()
{
MsgBox(ActiveDocument.Windows.Count + " window(s)", ActiveDocument.Name)
ActiveDocument.ActiveWindow.NewWindow()
MsgBox(ActiveDocument.Windows.Count + " windows", ActiveDocument.Name)
}
```

### Document.WordOpenXML

返回一个 **String** 类型的值，该值代表文档的 WPS Open XML 内容的 Flat XML 格式 。只读。

**语法**

**express.WordOpenXML**

\*express是一个代表 Document 对象的变量。

### Document.Words

返回一个 **Words** 集合，该集合代表文档中的所有字词。只读。

**语法**

**express.Words**

\*express是一个代表 Document 对象的变量。

**说明**

文档中的标点符号和段落标记也包括在 **Words** 集合中。

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示所选内容的字数。段落标记、部分单词和标点符号都统计在内。*/
MsgBox("There are " + Selection.Words.Count + " words.")
```

``` JavaScript
/*本示例遍历 myRange 中的单词（从活动文档的开始之所选内容的末尾），并删除该区域中出现的任何“Franklin”（包含尾部空格）。*/
function test()
{
let myRange = ActiveDocument.Range(0, Selection.End)        
let aWord = myRange.Words                   
for(let i = 1; i <= aWord.Count; i++) {
    if(aWord.Item(i).Text == "Franklin ") {
        aWord.Item(i).Delete()
}
}
}
```

### Document.WritePassword

该属性设置一个保存对指定文档所做的修改时所需的密码。 **String** 类型，只写。

**语法**

**express.WritePassword**

\*express是一个代表 Document 对象的变量。

**说明**

**要点** ??请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

``` JavaScript
/*如果活动文档没有设置密码以防修改文档，则本示例将“secret”设置为该文档的写权限密码。*/
function test()
{
let myDoc = ActiveDocument
if(myDoc.WriteReserved == false) {
    myDoc.WritePassword = "secret"
}
}
```

### Document.WriteReserved

如果指定文档设有写权限密码，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.WriteReserved**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*如果活动文档设有写权限密码，则本示例显示一条消息。*/
function test()
{
if(ActiveDocument.WriteReserved == true) {
    MsgBox("Changes cannot be made to this document.")
}
}
```

### Document.XMLHideNamespaces

返回一个 **Boolean** 类型的值，该值表示是否隐藏 XML 命名空间，该命名空间位于“XML 结构”任务窗格的元素列表中。

**语法**

**express.XMLHideNamespaces**

\*express是一个代表 Document 对象的变量。

**说明**

如果为 **True** ，则显示元素，并在元素名称右侧显示该元素 XML 架构命名空间。如果为 **False** ，则不显示 XML 架构命名空间。

当包含相似或相同元素名称的多个架构附加到同一文档时，将 **XMLHideNamespaces** 属性值设置为 **False** 可能会很有帮助。

**示例**

``` JavaScript
/*下列示例中， WPS 不显示活动文档中的 XML 架构命名空间。*/
ActiveDocument.XMLHideNamespaces = false
```

### Document.XMLNodes

返回一个 **XMLNodes** 集合，该集合代表文档中的所有 XML 元素。

**语法**

**express.XMLNodes**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*以下示例返回活动文档中的第一个 XML 节点。*/
let objNode = ActiveDocument.XMLNodes.Item(1)
```

### Document.XMLSaveDataOnly

设置或返回一个 **Boolean** 类型的值，保存文档时包含 XML 标记还是仅保存文本。

**语法**

**express.XMLSaveDataOnly**

\*express是一个代表 Document 对象的变量。

**说明**

如果为 **True** ，则表明 WPS 在保存文档时仅包含自定义 XML 标记及其相关内容。如果为 **False** ，则表明 WPS 在保存文档时将包含全部 XML 标记。

**示例**

``` JavaScript
/*下列示例指定活动文档仅与自定义 XML 标记及相关内容一起保存。*/
ActiveDocument.XMLSaveDataOnly = true
```

### Document.XMLSaveThroughXSLT

设置或返回一个 **String** 类型的值，指定用户保存文档时，Extensible Stylesheet Language Transformation (XSLT) 要应用的路径和文件名。

**语法**

**express.XMLSaveThroughXSLT**

\*express是一个代表 Document 对象的变量。

**说明**

仅当将 **XMLUseXSLTWhenSaving** 属性值设置为 **True** 时， **XMLSaveThroughXSLT** 属性才可用。如果将 **XMLUseXSLTWhenSaving** 属性值设置为 **False** ，则 WPS 会忽略 **XMLSaveThroughXSLT** 属性。

**示例**

``` JavaScript
/*下列示例指定 WPS 保存活动文档时应使用 XSLT，并指定了要使用的 XSLT。*/
function test()
{
ActiveDocument.XMLUseXSLTWhenSaving = true
ActiveDocument.XMLSaveThroughXSLT = "C:\\schemas\\book.xsl"
}
```

### Document.XMLSchemaReferences

返回一个 XMLSchemaReferences 集合，代表附加到文档的架构。

**语法**

**express.XMLSchemaReferences**

\*express是一个代表 Document 对象的变量。

**说明**

**示例**

``` JavaScript
/*下列示例重新加载附加到活动文档的第一个架构。*/
function test()
{
let objSchema = ActiveDocument.XMLSchemaReferences.Item(1)
objSchema.Reload()
}
```

### Document.XMLSchemaViolations

返回一个 XMLNodes 集合，代表文档中所有存在验证错误的节点。

**语法**

**express.XMLSchemaViolations**

\*express是一个代表 Document 对象的变量。

**示例**

``` JavaScript
/*下列示例创建一个对活动文档中存在验证错误的 XML 元素的引用。*/
let objNodes = ActiveDocument.XMLSchemaViolations
```

### Document.XMLShowAdvancedErrors

返回或设置一个 **Boolean** 类型的值，该值表示是由内置的 WPS 错误消息还是由包括在 Office 中的 Microsoft XML Core Services (MSXML) 5.0 组件来生成错误消息文本。

**语法**

**express.XMLShowAdvancedErrors**

\*express是一个代表 Document 对象的变量。

**说明**

使用 MSXML 5.0 组件的高级错误信息可提供更具体的错误信息。当 **XMLShowAdvancedErrors** 属性为 **True** 时，通过 XML Core Services 可提供大约 500 条错误信息。

WPS 大约有 50 条内置的常规架构错误。当 **XMLShowAdvancedErrors** 属性为 **False** 时， WPS 对于 XML 文档中生成的错误使用内置的错误信息。

**示例**

``` JavaScript
/*以下示例在活动文档中启用高级错误消息。*/
ActiveDocument.XMLShowAdvancedErrors = true
```

### Document.XMLUseXSLTWhenSaving

返回一个 **Boolean** 类型的值，代表是否通过 Extensible Stylesheet Language Transformation (XSLT) 保存文档。如果通过 XSLT 保存文档，则该值为 **True** 。

**语法**

**express.XMLUseXSLTWhenSaving**

\*express是一个代表 Document 对象的变量。

**说明**

在将 XMLUseXSLTWhenSaving 属性值设置为 **True** 时，使用 **XMLSaveThroughXSLT** 属性可指定要使用的 XSLT 的路径和文件名。

**示例**

``` JavaScript
/*下列示例指定 WPS 保存活动文档时应使用 XSLT，并指定了要使用的 XSLT。*/
function test()
{
ActiveDocument.XMLUseXSLTWhenSaving = true
ActiveDocument.XMLSaveThroughXSLT = "C:\\schemas\\book.xslt"
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Document](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Fields 对象

## 简介

Field 对象的集合，这些对象代表所选内容、范围或文档中的所有域。

## 方法

| 序号 | 名称                                       | 说明                                                                        |
|:----:|:-------------------------------------------|:----------------------------------------------------------------------------|
|  1   | [Add](#Fields.Add)                         | 将 Field 对象添加到 Fields 集合。返回指定区域内的 Field 对象。              |
|  2   | [Item](#Fields.Item)                       | 返回集合中的单个 Field 对象。                                               |
|  3   | [ToggleShowCodes](#Fields.ToggleShowCodes) | 在域代码与域结果之间切换域的显示。可以使用 ShowCodes 属性控制单个域的显示。 |
|  4   | [Unlink](#Fields.Unlink)                   | 将 Fields 集合中的所有域替换为它们的最新结果。                              |
|  5   | [Update](#Fields.Update)                   | 更新域对象的结果。返回 Number 类型。                                        |
|  6   | [UpdateSource](#Fields.UpdateSource)       | 将对 INCLUDETEXT 域所做的更改保存回源文档。                                 |

## 属性

| 序号 | 名称                               | 说明                                                                         |
|:----:|:-----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Fields.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Fields.Count)             | 返回一个 Long 类型的值，该值代表集合中域的数量。只读。                       |
|  3   | [Creator](#Fields.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Locked](#Fields.Locked)           | 如果 Fields 集合中的所有域都已被锁定，则该属性值为 True 。可读写 Long 类型。 |
|  5   | [Parent](#Fields.Parent)           | 返回一个 Object 类型值，该值代表指定 Fields 对象的父对象。                   |

## 成员方法

### Fields.Add

将 **Field** 对象添加到 **Fields** 集合。返回指定区域内的 **Field** 对象。

**语法**

**express.Add ( *Range* , *Type* , *Text* , *PreserveFormatting* )**

\*express是一个代表 Fields 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                         |
|----------------------|-----------|----------|----------------------------------------------------------------------------------------------|
| *Range*              | 必选      | Range    | 需要添加域的区域。如果该区域未折叠，那么域将替换该区域。                                     |
| *Type*               | 可选      | Variant  | 可以是任意 WdFieldType 常量。有关有效的常量列表，请参阅“对象浏览器”。默认值为 wdFieldEmpty。 |
| *Text*               | 可选      | Variant  | 域所需的其他文本。例如，如果要给域指定一个开关，可将其添加到此处。                           |
| *PreserveFormatting* | 可选      | Variant  | 如果该属性值为 True，则更新时保留域所应用的格式。                                            |

**说明**

不能用 **Fields** 集合的 **Add** 方法来插入诸如 **wdFieldOCX** 和 **wdFieldFormCheckBox** 之类的域，而应使用诸如 **FormFields** 集合的 **AddOLEControl** 方法和 **Add** 方法。

**示例**

``` JavaScript
/* 选区开始插入一个日期域 */
function test() {
    Application.Selection.Collapse(wdCollapseStart)
    let myField = Application.ActiveDocument.Fields.Add(Application.Selection.Range,wdFieldDate)
}
```

### Fields.Item

返回集合中的单个 **Field** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Fields 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                           |
|---------|-----------|----------|----------------------------------------------------------------|
| *Index* | 可选      | Number   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

### Fields.ToggleShowCodes

在域代码与域结果之间切换域的显示。可以使用 **ShowCodes** 属性控制单个域的显示。

**语法**

**express.ToggleShowCodes ()**

\*express是一个代表 Fields 对象的变量。

**示例**

``` JavaScript
/*本示例打开或关闭所选内容中域的显示（相当于按 Shift+F9）。*/
Selection.Fields.ToggleShowCodes()
```

``` JavaScript
/*本示例打开或关闭活动文档中域的显示（相当于按 Alt+F9）。*/
ActiveDocument.Fields.ToggleShowCodes()
```

### Fields.Unlink

将 **Fields** 集合中的所有域替换为它们的最新结果。

**语法**

**express.Unlink ()**

\*express是一个代表 Fields 对象的变量。

**说明**

如果断开一个域的链接，则当前结果会转换为文字或图形，并且无法再自动更新。注意，某些域（例如，XE（索引项）域和 SEQ（序列）域）不能断开链接。

**示例**

``` JavaScript
/*以下示例更新并断开活动文档第一节中的所有域的链接。*/
function test()
{
ActiveDocument.Sections.Item(1).Range.Fields.Update()
ActiveDocument.Sections.Item(1).Range.Fields.Unlink()
}
```

### Fields.Update

更新域对象的结果。返回 Number 类型。

**语法**

**express.Update ()**

\*express是一个代表 Fields 对象的变量。

**说明**

如果更新域时没有出错，则返回 0（零）；或者返回一个 **Long** 类型的值，该值代表第一个出错域的索引。

**示例**

``` JavaScript
/* 更新文档第一个域，并且弹框提示是否成功 */
function test() {
    if( Application.ActiveDocument.Fields.Update() == 0) {
        alert("Update Successful")
    }
    else {
        alert( "Field " +  Application.ActiveDocument.Fields.Update() + " has an error")
    }
}
```

### Fields.UpdateSource

将对 INCLUDETEXT 域所做的更改保存回源文档。

**语法**

**express.UpdateSource ()**

\*express是一个代表 Fields 对象的变量。

**说明**

源文档必须是 WPS 文档格式。

## 成员属性

### Fields.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Fields 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Fields.Count

返回一个 **Long** 类型的值，该值代表集合中域的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Fields 对象的变量。

### Fields.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Fields 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Fields.Locked

如果 **Fields** 集合中的所有域都已被锁定，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.Locked**

\*express是一个代表 Fields 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** （集合中的部分域已被锁定，而其他域未锁定）。

**示例**

``` JavaScript
/*以下示例锁定所选内容中的所有域。*/
Selection.Fields.Locked = true
```

``` JavaScript
/*如果活动文档中的部分域已被锁定，则以下示例将显示一条消息。*/
function test()
{
let theFields = ActiveDocument.Fields
if( theFields.Locked == wdUndefined) {
    MsgBox( "Some fields are locked")
}
else if ( theFields.Locked == false) { 
    MsgBox ("No fields are locked")
}
else if( theFields.Locked == true) {
    MsgBox ("All fields are locked")
}
}
```

### Fields.Parent

返回一个 **Object** 类型值，该值代表指定 **Fields** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Fields 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Fields](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# OLEFormat 对象

## 简介

代表 OLE 对象、ActiveX 控件或域的 OLE（而不是链接）特性。

## 说明

使用形状、内嵌形状或域的 OLEFormat 属性可返回 OLEFormat 对象。以下示例显示活动文档中的第一个形状的类类型。

``` JavaScript
MsgBox(ActiveDocument.Shapes.Item(1).OLEFormat.ClassType)
```

并非所有类型的形状、内嵌形状和域都具有 OLE 功能。使用 Shape 和 InlineShape 对象的 Type 属性可确定指定形状或内嵌形状的类别。 Field 对象的 Type 属性返回域的类型。

您可以使用 Activate 、 Edit 、 Open 和 DoVerb 方法自动处理 OLE 对象。

使用 Object 属性可返回一个代表 ActiveX 控件或 OLE 对象的对象。借助于该对象，您可以使用容器应用程序或 ActiveX 控件的属性和方法。

## 方法

| 序号 | 名称                                | 说明                                                                                                          |
|:----:|:------------------------------------|:--------------------------------------------------------------------------------------------------------------|
|  1   | [Activate](#OLEFormat.Activate)     | 激活指定的 OLEFormat 对象。                                                                                   |
|  2   | [ActivateAs](#OLEFormat.ActivateAs) | 设置确定用于激活指定 OLE 对象的默认应用程序的 Windows 注册表值。                                              |
|  3   | [ConvertTo](#OLEFormat.ConvertTo)   | 将指定的 OLE 对象转换为另一种类型，以便能在另一种服务器应用程序中编辑该对象，或者改变对象在文档中的显示方式。 |
|  4   | [DoVerb](#OLEFormat.DoVerb)         | 请求 OLE 对象执行一个有效动作，以激活其内容。                                                                 |
|  5   | [Edit](#OLEFormat.Edit)             | 在创建指定 OLE 对象的应用程序中打开该对象，以便进行编辑。                                                     |
|  6   | [Open](#OLEFormat.Open)             | 打开指定的 OLEFormat 对象。                                                                                   |

## 属性

| 序号 | 名称                                                                | 说明                                                                                                                                                                                                     |
|:----:|:--------------------------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#OLEFormat.Application)                               | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                                                           |
|  2   | [ClassType](#OLEFormat.ClassType)                                   | 返回或设置指定的 OLE 对象、图片或域的类型。 String 类型，可读写。                                                                                                                                        |
|  3   | [Creator](#OLEFormat.Creator)                                       | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                                                             |
|  4   | [DisplayAsIcon](#OLEFormat.DisplayAsIcon)                           | 如果指定对象显示为图标，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                      |
|  5   | [IconIndex](#OLEFormat.IconIndex)                                   | 返回或设置 DisplayAsIcon 属性为 True 时所用的图标。 Long 类型，可读写。                                                                                                                                  |
|  6   | [IconLabel](#OLEFormat.IconLabel)                                   | 返回或设置 OLE 对象的图标下面显示的文字。 String 类型，可读写。                                                                                                                                          |
|  7   | [IconName](#OLEFormat.IconName)                                     | 返回或设置存储 OLE 对象的图标的程序文件。 String 类型，可读写。                                                                                                                                          |
|  8   | [IconPath](#OLEFormat.IconPath)                                     | 返回存储 OLE 对象的图标的文件的路径。 String 类型，只读。                                                                                                                                                |
|  9   | [Label](#OLEFormat.Label)                                           | 返回用于标识源文件中被链接的部分的字符串。 String 类型，只读。                                                                                                                                           |
|  10  | [Object](#OLEFormat.Object)                                         | 返回一个 Object 类型的值，该值代表指定的 OLE 对象的最上层接口。                                                                                                                                          |
|  11  | [Parent](#OLEFormat.Parent)                                         | 返回一个 Object 类型值，该值代表指定 OLEFormat 对象的父对象。                                                                                                                                            |
|  12  | [PreserveFormattingOnUpdate](#OLEFormat.PreserveFormattingOnUpdate) | 如果为 True ，则将 WPS 中的格式设置保存到链接的 OLE 对象中，例如链接到 ET 电子表格的表格。 Boolean 类型，可读写。                                                                                        |
|  13  | [ProgID](#OLEFormat.ProgID)                                         | 返回指定 OLE 对象的 程序标识符 (ProgID) ?（程序标识符：窗体 OLEServerName.ObjectName 中的标识符（例如，Excel.Sheet 或 PowerPoint.Slide），由 Windows 注册表用来唯一标识一个对象。） 。只读 String 类型。 |

## 成员方法

### OLEFormat.Activate

激活指定的 **OLEFormat** 对象。

**语法**

**express.Activate ()**

\*express是一个代表 OLEFormat 对象的变量。

### OLEFormat.ActivateAs

设置确定用于激活指定 OLE 对象的默认应用程序的 Windows 注册表值。

**语法**

**express.ActivateAs ( *ClassType* )**

\*express是一个代表 OLEFormat 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                               |
|-------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ClassType* | 必选      | String   | 在其中打开 OLE 对象的应用程序的名称。要查看 OLE 对象可被激活为的对象类型列表，请单击该对象，然后打开“转换”对话框。将一个对象作为内嵌形状插入，然后查看域代码，就可以找到 ClassType 字符串。对象的类类型后带有“EMBED”或“LINK”一词。 |

**示例**

``` JavaScript
/*本示例实现的功能是：设置在 ET 中打开活动文档的第一个浮动图形，然后激活该图形。为使本示例能够运行，此图形必须是可以在 ET 中打开的 OLE 对象。*/
function test()
{
ActiveDocument.Shapes.Item(1).OLEFormat.ActivateAs("Excel.Sheet")
ActiveDocument.Shapes.Item(1).OLEFormat.Activate()
}
```

### OLEFormat.ConvertTo

将指定的 OLE 对象转换为另一种类型，以便能在另一种服务器应用程序中编辑该对象，或者改变对象在文档中的显示方式。

**语法**

**express.ConvertTo ( *ClassType* , *DisplayAsIcon* , *IconFileName* , *IconIndex* , *IconLabel* )**

\*express是一个代表 OLEFormat 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                               |
|-----------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ClassType*     | 可选      | Variant  | 用于激活 OLE 对象的应用程序的名称。在“对象”对话框的“新建”选项卡上的“对象类型”框中可以看到可用应用程序列表。将一个对象作为内嵌形状插入，然后查看域代码，就可以找到 ClassType 字符串。对象的类类型后带有“EMBED”或“LINK”一词。                                                                                                        |
| *DisplayAsIcon* | 可选      | Variant  | 如果该属性值为 True，则将 OLE 对象显示为图标。默认值为 False。                                                                                                                                                                                                                                                                     |
| *IconFileName*  | 可选      | Variant  | 包含将要显示的图标的文件。                                                                                                                                                                                                                                                                                                         |
| *IconIndex*     | 可选      | Variant  | IconFileName 内的图标的索引编号。指定文件中图标的顺序与在选定“显示为图标”复选框后“更改图标”对话框中（可通过“插入对象”对话框访问）图标显示的顺序对应。文件中的第一个图标的索引编号为 0（零）。如果具有给定索引编号的图标在 IconFileName 所指定的文件中不存在，则使用索引编号为 1 的图标（即文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选      | Variant  | 显示在图标下面的标签（题注）。                                                                                                                                                                                                                                                                                                     |

**示例**

``` JavaScript
/*本示例创建一篇文档，然后插入一篇嵌入的包含文字的文档。再将嵌入的文档转换成 WPS 图片。*/
function test()
{
Documents.Add()
let objEmbedded = ActiveDocument.Shapes.AddOLEObject("Word.Document")
    objEmbedded.Activate()

Selection.TypeText("Test")
objEmbedded.OLEFormat.OLEFormat.ConvertTo("Word.Picture")
}
```

### OLEFormat.DoVerb

请求 OLE 对象执行一个有效动作，以激活其内容。

**语法**

**express.DoVerb ( *VerbIndex* )**

\*express是一个代表 OLEFormat 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                           |
|-------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------|
| *VerbIndex* | 可选      | Variant  | OLE 对象应执行的动作。如果省略该参数，则发送默认动作。如果 OLE 对象不支持被请求的动作，将发生错误。可以是任意 WdOLEVerb 常量。 |

**说明**

每个 OLE 对象都支持附加于该对象的一组动作。

**示例**

``` JavaScript
/*本示例将活动文档第一个浮动 OLE 对象的默认动作发送给服务器。*/
ActiveDocument.Shapes.Item(1).OLEFormat.DoVerb()
```

### OLEFormat.Edit

在创建指定 OLE 对象的应用程序中打开该对象，以便进行编辑。

**语法**

**express.Edit ()**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*本示例打开（为了进行编辑）活动文档中第一个嵌入式 OLE 对象（定义为图形）。*/
function test()
{
let shapesAll = ActiveDocument.Shapes
if(shapesAll.Count >= 1){
    if(shapesAll.Item(1).Type == msoEmbeddedOLEObject){
        shapesAll.Item(1).OLEFormat.Edit()
    }
}
}
```

``` JavaScript
/*本示例打开（为了进行编辑）活动文档中第一个链接的 OLE 对象（定义为嵌入式图形）。*/
function test()
{
let colIS = ActiveDocument.InlineShapes
if(colIS.Count >= 1){
    if(colIS.Item(1).Type == wdInlineShapeLinkedOLEObject){
        colIS.Item(1).OLEFormat.Edit()
    }
}
}
```

### OLEFormat.Open

打开指定的 **OLEFormat** 对象。

**语法**

**express.Open ()**

\*express是一个代表 OLEFormat 对象的变量。

## 成员属性

### OLEFormat.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### OLEFormat.ClassType

返回或设置指定的 OLE 对象、图片或域的类型。 **String** 类型，可读写。

**语法**

**express.ClassType**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

对于除 DDE 链接之外的其他链接对象，本属性是只读的。

在 **“插入”** 菜单上 **“对象”** 对话框中 **“新建”** 选项卡上的 **“对象类型”** 框中可以看到可用应用程序的列表。以嵌入式图形形式插入一个对象，然后查看域代码，可以找到 **ClassType** 字符串。对象的类类型后面都带有单词“EMBED”或“LINK”。

**示例**

``` JavaScript
/*本示例遍历活动文档中所有浮动图形，并将所有链接的 ET 工作表设置为可自动更新。*/
function test()
{
for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){
    if(ActiveDocument.Shapes.Item(i).Type == msoLinkedOLEObject){
        if(ActiveDocument.Shapes.Item(i).OLEFormat.ClassType == "Excel.Sheet"){
            ActiveDocument.Shapes.Item(i).LinkFormat.AutoUpdate = true
        }
    }
}
}
```

### OLEFormat.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### OLEFormat.DisplayAsIcon

如果指定对象显示为图标，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayAsIcon**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*本示例显示一个消息框，框内包含活动文档中每个显示为图标的浮动图形名称。*/
function test()
{
for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){
    if(ActiveDocument.Shapes.Item(i).OLEFormat.DisplayAsIcon == true){
       MsgBox(shapeLoop.Name + " is displayed as an icon.")
    }
}
}
```

``` JavaScript
/*本示例将 ET 工作表作为链接的 OLE 对象插入活动文档，然后将该对象的显示更改为图标方式。*/
function test()
{
let objNew = ActiveDocument.Shapes.AddOLEObject(null, "C:\\Program Files\\Microsoft Office\\Office\\Samples\\samples.xls", true)
ActiveDocument.Shapes.Item(1).OLEFormat.DisplayAsIcon = true
}
```

### OLEFormat.IconIndex

返回或设置 **DisplayAsIcon** 属性为 **True** 时所用的图标。 **Long** 类型，可读写。

**语法**

**express.IconIndex**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

零 (0) 对应第一个图标，1 对应第二个图标，依此类推。如果省略该参数，则使用第一个（默认）图标。

**示例**

``` JavaScript
/*本示例在消息框中返回显示为图标的第一个选定图形的图标索引号。*/
function test()
{
if(Selection.ShapeRange.Count >= 1){
    let olefTemp = Selection.ShapeRange.Item(1).OLEFormat
    if(olefTemp.DisplayAsIcon == true){
        MsgBox(olefTemp.IconIndex)
    }
}
}
```

### OLEFormat.IconLabel

返回或设置 OLE 对象的图标下面显示的文字。 **String** 类型，可读写。

**语法**

**express.IconLabel**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*本示例更改选定内容中第一个图形的图标下面的文字。*/
function test()
{
if(Selection.ShapeRange.Count >= 1){
    let olefTemp = Selection.ShapeRange.Item(1).OLEFormat
    olefTemp.DisplayAsIcon = true
    olefTemp.IconLabel = "My Icon"
}
}
```

### OLEFormat.IconName

返回或设置存储 OLE 对象的图标的程序文件。 **String** 类型，可读写。

**语法**

**express.IconName**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*本示例将所选内容中第一个图形显示为图标，并将图标下面的文字设为该图标的文件名。*/
function test()
{
if(Selection.ShapeRange.Count >= 1){
    let olefTemp = Selection.ShapeRange.Item(1).OLEFormat
    olefTemp.DisplayAsIcon = true
    olefTemp.IconLabel = olefTemp.IconName
}
}
```

### OLEFormat.IconPath

返回存储 OLE 对象的图标的文件的路径。 **String** 类型，只读。

**语法**

**express.IconPath**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*本示例显示每个嵌入式 OLE 对象的路径，这些 OLE 对象在活动文档中显示为图标。*/
function test()
{
for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){
    if(ActiveDocument.Shapes.Item(i).Type == msoEmbeddedOLEObject){
        if(ActiveDocument.Shapes.Item(i).OLEFormat.DisplayAsIcon == true){
            MsgBox(shapeLoop.OLEFormat.IconPath)
        }
    }
}
}
```

### OLEFormat.Label

返回用于标识源文件中被链接的部分的字符串。 **String** 类型，只读。

**语法**

**express.Label**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

如果源文件是 ET 工作簿，并且 OLE 对象只包含工作表中少量的单元格，则 **Label** 属性可能返回“Workbook1!R3C1:R4C2”。

该属性仅对链接为 OLE 对象的图形、嵌入式图形或域有效。

**示例**

``` JavaScript
/*本示例返回活动文档第一个域的标签。*/
MsgBox(ActiveDocument.Fields.Item(1).OLEFormat.Label)
```

### OLEFormat.Object

返回一个 **Object** 类型的值，该值代表指定的 OLE 对象的最上层接口。

**语法**

**express.Object**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

使用该属性可以访问 ActiveX 控件或应用程序（创建 OLE 对象的应用程序）的属性和方法。为使该属性生效，OLE 对象必须支持“OLE 自动化”。

**示例**

``` JavaScript
/*本示例设置活动文档中第一个图形的值。为使本示例能够运行，第一个图形必须是 ActiveX 控件（例如，复选框或选项按钮）。*/
function test()
{
let myOA = ActiveDocument.Shapes.Item(1).OLEFormat
myOA.Activate()
myOA.Object.Value = true
}
```

``` JavaScript
/*本示例向活动文档中添加一个新的 ActiveX 控件，然后激活新添的选项按钮，并设置该按钮的部分属性。*/
function test()
{
let myOB = ActiveDocument.Shapes.AddOLEControl("Forms.OptionButton.1")
let myOA = myOB.OLEFormat
    myOA.Activate()

let myObj = myOA.Object
    myObj.Value = false
    myObj.Caption = "My Caption"
    myObj.AutoSize = true
}
```

### OLEFormat.Parent

返回一个 **Object** 类型值，该值代表指定 **OLEFormat** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 OLEFormat 对象的变量。

### OLEFormat.PreserveFormattingOnUpdate

如果为 **True** ，则将 WPS 中的格式设置保存到链接的 OLE 对象中，例如链接到 ET 电子表格的表格。 **Boolean** 类型，可读写。

**语法**

**express.PreserveFormattingOnUpdate**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

如果将 **PreserveFormattingOnUpdate** 设为 **True** ，则更新对象时，保留在 WPS 中对该对象进行的格式更改。 WPS 仅更新链接对象中的内容。

**示例**

``` JavaScript
/*假定该文档中的第一个图形是链接的 OLE 对象，本示例保留当前文档中第一个图形的格式。*/
ActiveDocument.Shapes.Item(1).OLEFormat.PreserveFormattingOnUpdate = true
```

### OLEFormat.ProgID

返回指定 OLE 对象的 程序标识符 (ProgID) ?（程序标识符：窗体 OLEServerName.ObjectName 中的标识符（例如，Excel.Sheet 或 PowerPoint.Slide），由 Windows 注册表用来唯一标识一个对象。） 。只读 **String** 类型。

**语法**

**express.ProgID**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

**ProgID** 属性和 **ClassType** 属性（默认情况下）会返回相同的字符串。但是，您可以更改 DDE 链接的 **ClassType** 属性。

**安全性** ??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

有关程序标识符的信息，请参阅 OLE 程序标识符。

**示例**

``` JavaScript
/*本示例遍历活动文档中的所有浮动形状，并将所有链接的 ET 工作表设置为自动更新。*/
function test()
{
for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){
    if(ActiveDocument.Shapes.Item(i).Type == msoLinkedOLEObject){
        if(ActiveDocument.Shapes.Item(i).OLEFormat.ProgID == "Excel.Sheet"){
           ActiveDocument.Shapes.Item(i).LinkFormat.AutoUpdate = true
        }
    }
}
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/OLEFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PageNumbers 对象

## 方法

| 序号 | 名称                      | 说明                                                                 |
|:----:|:--------------------------|:---------------------------------------------------------------------|
|  1   | [Add](#PageNumbers.Add)   | 返回一个 PageNumber 对象，该对象代表添加到节中的页眉或页脚中的页码。 |
|  2   | [Item](#PageNumbers.Item) | 返回集合中的单个 PageNumber 对象。                                   |

## 属性

| 序号 | 名称                        | 说明                                                   |
|:----:|:----------------------------|:-------------------------------------------------------|
|  1   | [Count](#PageNumbers.Count) | 返回一个 Long 类型的值，该值代表集合中的页码数。只读。 |

## 成员方法

### PageNumbers.Add

返回一个 **PageNumber** 对象，该对象代表添加到节中的页眉或页脚中的页码。

**语法**

**express.Add ( *PageNumberAlignment* , *FirstPage* )**

\*express是一个代表 PageNumbers 对象的变量。

**参数**

| 名称                  | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                        |
|-----------------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *PageNumberAlignment* | 可选      | Variant  | 可以是任意 WdPageNumberAlignment 常量。                                                                                                                                                                     |
| *FirstPage*           | 可选      | Variant  | 如果该属性值为 False，则文档中的第一页的页眉和页脚与后面各页的页眉和页脚都不相同；如果将 FirstPage 设置为 False，则在第一页不添加页码；如果省略该参数，则由 DifferentFirstPageHeaderFooter 属性控制该设置。 |

**说明**

如果将 **HeaderFooter** 对象的 **LinkToPrevious** 属性设置为 **True** ，则文档中各节的页码都将连续编号。

**示例**

``` JavaScript
/*本示例为活动文档第一节的主页脚添加页码*/
function test() {
 let Sec = Application.ActiveDocument.Sections(1)
     Sec.Footers(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberLeft,true)
}

/*本示例在活动文档的页眉中创建页码并进行格式设置。*/
function test() {
    let myPgNum = Application.ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberCenter,true)
    myPgNum.Select()
    Selection.Range.Italic = true
    Selection.Range.Bold = true
}
```

### PageNumbers.Item

返回集合中的单个 **PageNumber** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 PageNumbers 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

## 成员属性

### PageNumbers.Count

返回一个 **Long** 类型的值，该值代表集合中的页码数。只读。

**语法**

**express.Count**

\*express是一个代表 PageNumbers 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/PageNumbers](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Rows 对象

## 简介

[Row]() 对象的集合，这些对象代表指定的选定内容、区域或表格中的表格行

## 方法

| 序号 | 名称                                       | 说明                                                                    |
|:----:|:-------------------------------------------|:------------------------------------------------------------------------|
|  1   | [Add](#Rows.Add)                           | 返回一个 Row 对象，该对象代表添加到表中的行。                           |
|  2   | [ConvertToText](#Rows.ConvertToText)       | 将表格中的行转换为文本并返回一个 Range 对象，该对象代表带分隔符的文本。 |
|  3   | [Delete](#Rows.Delete)                     | 删除指定的表格行                                                        |
|  4   | [DistributeHeight](#Rows.DistributeHeight) | 将指定行或单元格的高度调整为相等。                                      |
|  5   | [Item](#Rows.Item)                         | 返回集合中的单个 Row 对象。                                             |
|  6   | [Select](#Rows.Select)                     | 选择表格中行的集合。                                                    |
|  7   | [SetHeight](#Rows.SetHeight)               | 设置表格行的高度。                                                      |
|  8   | [SetLeftIndent](#Rows.SetLeftIndent)       | 设置表格中一行或多行的缩进。                                            |

## 属性

| 序号 | 名称                                                           | 说明                                                                                               |
|:----:|:---------------------------------------------------------------|:---------------------------------------------------------------------------------------------------|
|  1   | [Alignment](#Rows.Alignment)                                   | 返回或设置一个 WdRowAlignment 常量，该常量代表指定行的对齐方式。可读写。                           |
|  2   | [AllowBreakAcrossPage](#Rows.AllowBreakAcrossPage)             | 如果允许分页符拆分表格中一行或多行中的文本，则该属性值为 True 。可读写 Long 类型。                 |
|  3   | [AllowOverlap](#Rows.AllowOverlap)                             | 返回或设置一个值，该值指定是否允许指定行与其他行重叠。                                             |
|  4   | [Application](#Rows.Application)                               | 返回一个代表 WPS 应用程序的 Application 对象。                                                     |
|  5   | [Borders](#Rows.Borders)                                       | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                              |
|  6   | [Count](#Rows.Count)                                           | 返回一个 Long 类型的值，该值代表集合中的行数。只读。                                               |
|  7   | [Creator](#Rows.Creator)                                       | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                       |
|  8   | [DistanceBottom](#Rows.DistanceBottom)                         | 返回或设置文档文字与指定表格下边缘之间的距离（以磅为单位）。 Single 类型，可读写。                 |
|  9   | [DistanceLeft](#Rows.DistanceLeft)                             | 返回或设置文档文字与指定表格的左边缘之间的距离（以磅为单位）。 Single 类型，可读写。               |
|  10  | [DistanceRight](#Rows.DistanceRight)                           | 返回或设置文档文字与指定表格右边缘之间的距离（以磅为单位）。 Single 类型，可读写。                 |
|  11  | [DistanceTop](#Rows.DistanceTop)                               | 返回或设置文档文字与指定表格的上边缘之间的距离（以磅为单位）。 Single 类型，可读写。               |
|  12  | [First](#Rows.First)                                           | 返回一个 Row 对象，该对象代表 Rows 集合中的第一个项目。                                            |
|  13  | [HeadingFormat](#Rows.HeadingFormat)                           | 如果将指定一行或数行的格式设置为表格标题，则该属性值为 True 。 Long 类型，可读写。                 |
|  14  | [Height](#Rows.Height)                                         | 返回或设置表格中指定的行的高度。Single 类型，可读写。                                              |
|  15  | [HeightRule](#Rows.HeightRule)                                 | 返回或设置确定指定单元格或行高度的规则。 WdRowHeightRule 类型，可读写。                            |
|  16  | [HorizontalPosition](#Rows.HorizontalPosition)                 | 返回或设置由 RelativeHorizontalPosition 属性指定的项目与行边缘之间的水平距离。可读写 Single 类型。 |
|  17  | [Last](#Rows.Last)                                             | 将 Rows 集合中的最后一项作为 Row 对象返回。                                                        |
|  18  | [LeftIndent](#Rows.LeftIndent)                                 | 返回或设置一个 Single 类型的值，该值代表指定表格行的左缩进值（以磅为单位）。可读写。               |
|  19  | [NestingLevel](#Rows.NestingLevel)                             | 返回指定表格行的嵌套层。只读 Long 类型。                                                           |
|  20  | [Parent](#Rows.Parent)                                         | 返回一个 Object 类型值，该值代表指定 Rows 对象的父对象。                                           |
|  21  | [RelativeHorizontalPosition](#Rows.RelativeHorizontalPosition) | 指定一组行的相对水平位置。可读写 WdRelativeHorizontalPosition 类型。                               |
|  22  | [RelativeVerticalPosition](#Rows.RelativeVerticalPosition)     | 指定一组行的相对垂直位置。可读写 WdRelativeVerticalPosition 类型。                                 |
|  23  | [Shading](#Rows.Shading)                                       | 返回一个 Shading 对象，该对象代表指定对象的底纹格式。                                              |
|  24  | [SpaceBetweenColumns](#Rows.SpaceBetweenColumns)               | 返回或设置指定的一行或多行中相邻列的文本之间的距离（以磅为单位）。可读写 Single 类型。             |
|  25  | [TableDirection](#Rows.TableDirection)                         | 返回或设置 WPS 对指定的表格或行中的单元格进行排序的方向。可读写 WdTableDirection 类型。            |
|  26  | [VerticalPosition](#Rows.VerticalPosition)                     | 返回或设置由 RelativeVerticalPosition 属性指定的项目与行边缘之间的垂直距离。可读写 Single 类型。   |
|  27  | [WrapAroundText](#Rows.WrapAroundText)                         | 返回或设置文本是否环绕指定行。 Long 类型，可读写。                                                 |

## 成员方法

### Rows.Add

返回一个 **Row** 对象，该对象代表添加到表中的行。

**语法**

**express.Add ()**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
//本示例在选定内容的第一行之前插入一个新行。
function test()
{
    if (Application.Selection.Information(wdWithInTable))
        Application.Selection.Rows.Add(Selection.Rows.Item(1))
}

//本示例在第一张表中添加一行，然后将文本 Cell 插入该行。
function test()
{   
    let tblNew = Application.ActiveDocument.Tables.Item(1)
    let rowNew = tblNew.Rows.Add(tblNew.Rows.Item(1))
    for(i = 1; i <= rowNew.Cells.Count; ++i)
    {
        let celTable = rowNew.Cells.Item(i)
        celTable.Range.InsertAfter("Cell " + i)
    }
}
```

### Rows.ConvertToText

将表格中的行转换为文本并返回一个 **Range** 对象，该对象代表带分隔符的文本。

**语法**

**express.ConvertToText ( *Separator* , *NestedTables* )**

\*express是一个代表 Rows 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                      |
|----------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Separator*    | 可选      | Variant  | 用以分隔被转换列的字符（被转换行由段落标记分隔）。可以是下列任何 WdTableFieldSeparator 常量：wdSeparateByCommas、wdSeparateByDefaultListSeparator、wdSeparateByParagraphs 或 wdSeparateByTabs（默认值）。 |
| *NestedTables* | 可选      | Variant  | 如果要将嵌套的表格转换为文本，则为 True。如果 Separator 不是 wdSeparateByParagraphs，则将忽略此参数。默认值为 True。                                                                                      |

### Rows.Delete

删除指定的表格行

**语法**

**express.Delete ()**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
//删除选中区域的表格行
Application.Selection.Rows.Delete()
```

### Rows.DistributeHeight

将指定行或单元格的高度调整为相等。

**语法**

**express.DistributeHeight ()**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档的第一张表格的高度调整为相等。*/
ActiveDocument.Tables.Item(1).Rows.DistributeHeight()
```

``` JavaScript
/*本示例将第一张表格的前三行高度调整为相等。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1) {
    let rngTemp = ActiveDocument.Range(ActiveDocument.Tables.Item(1).Rows.Item(1).Range.Start, ActiveDocument.Tables.Item(1).Rows.Item(3).Range.End)
    rngTemp.Rows.DistributeHeight()
}
}
```

### Rows.Item

返回集合中的单个 **Row** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Rows 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

**返回值**

Row

### Rows.Select

选择表格中行的集合。

**语法**

**express.Select ()**

\*express是一个代表 Rows 对象的变量。

**说明**

使用此方法后，可使用 **Selection** 对象处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

### Rows.SetHeight

设置表格行的高度。

**语法**

**express.SetHeight ( *RowHeight* , *HeightRule* )**

\*express是一个代表 Rows 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型        | 说明                           |
|--------------|-----------|-----------------|--------------------------------|
| *RowHeight*  | 必选      | Single          | 一行或多行的高度，以磅为单位。 |
| *HeightRule* | 必选      | WdRowHeightRule | 用于确定指定行的高度的规则。   |

**示例**

``` JavaScript
/*以下示例创建一个表格，然后将表格中所有行的行高设置为 0.5 英寸（36 磅）。*/
function test()
{
let newDoc = Documents.Add()
let aTable = newDoc.Tables.Add(Selection.Range, 3, 3)
aTable.Rows.SetHeight(InchesToPoints(0.5), wdRowHeightExactly)
}
```

### Rows.SetLeftIndent

设置表格中一行或多行的缩进。

**语法**

**express.SetLeftIndent ( *LeftIndent* , *RulerStyle* )**

\*express是一个代表 Rows 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型     | 说明                                                                                                                                                                        |
|--------------|-----------|--------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *LeftIndent* | 必选      | Single       | 指定的一行或多行的当前左边缘与所需左边缘之间的距离（以磅为单位）。                                                                                                          |
| *RulerStyle* | 必选      | WdRulerStyle | 控制当左缩进发生变化时 WPS 调整表格的方式。WdRulerStyle 行为适用于左对齐的表格。居中和右对齐表格的 WdRulerStyle 行为不可预测；在这种情况下，请谨慎使用 SetLeftIndent 方法。 |

**示例**

``` JavaScript
/*以下示例在新文档中创建一个表格，并将表格中的所有行缩进 0.5 英寸（36 磅）。在更改左缩进时，单元格宽度将调整以保持表格的右边缘。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
tableNew.Rows.SetLeftIndent(InchesToPoints(0.5), wdAdjustSameWidth)
}
```

``` JavaScript
/*本示例将活动文档中“表格 1”的第一行缩进 18 磅，并通过缩小第一列的宽度来保持表格的右边界位置。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1) {
    ActiveDocument.Tables.Item(1).Rows.SetLeftIndent(18, wdAdjustFirstColumn)
}
}
```

## 成员属性

### Rows.Alignment

返回或设置一个 **WdRowAlignment** 常量，该常量代表指定行的对齐方式。可读写。

**语法**

**express.Alignment**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中第一张表格的所有行都设置为居中对齐。*/
function CenterRows() {
    ActiveDocument.Tables.Item(1).Rows.Alignment = wdAlignRowCenter
}
```

### Rows.AllowBreakAcrossPage

如果允许分页符拆分表格中一行或多行中的文本，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.AllowBreakAcrossPage**

\*express是一个代表 Rows 对象的变量。

**说明**

该属性可以是 **True** 、 **False** 或 **wdUndefined** （仅允许拆分部分指定文本）。

**示例**

``` JavaScript
/*以下示例新建一篇包含 5x5 表格的文档，并防止在分页时拆分表格的第三行。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 5, 5)

tableNew.Rows.Item(3).AllowBreakAcrossPages = false
}
```

``` JavaScript
/*本示例确定分页时是否可以拆分当前表格中的行。如果插入点不在表格中，则显示一个消息框。*/
function test()
{
Selection.Collapse(wdCollapseStart) 
if(!Selection.Tables.Count) {
    MsgBox("The insertion point is not in a table.")
}
else {
    let lngAllowBreak = Selection.Rows.AllowBreakAcrossPages
}
}
```

### Rows.AllowOverlap

返回或设置一个值，该值指定是否允许指定行与其他行重叠。

**语法**

**express.AllowOverlap**

\*express是一个代表 Rows 对象的变量。

**说明**

如果指定行包含重叠行和非重叠行，则该属性返回 **wdUndefined** 。可以设置为 **True** 或 **False** 。 **Long** 类型，可读写。如果将 **AllowOverlap** 设置为 **True** ，会将 **WrapAroundText** 也设置为 **True** ；如果将 **WrapAroundText** 设置为 **False** ，则会将 **AllowOverlap** 也设置为 **False** 。

因为 HTML 不支持表格或图形的重叠，所以在 Web 版式视图中将忽略 **AllowOverlap** 。

**示例**

``` JavaScript
/*本示例指定文字环绕选定的表格，并且此表不能与其他被文字环绕的表格重叠。*/
function test()
{
Selection.Rows.WrapAroundText = true
Selection.Rows.AllowOverlap = false
}
```

``` JavaScript
/*本示例指定活动文档中的第一个图形可与其他图形重叠。*/
ActiveDocument.Shapes.Item(1).WrapFormat.AllowOverlap = true
```

### Rows.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Rows 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Rows.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Rows 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/**/
function test()
{
let myTable = ActiveDocument.Tables.Item(1)
let myBorders = myTable.Rows.Borders
myBorders.InsideLineStyle = wdLineStyleSingle
myBorders.OutsideLineStyle = wdLineStyleDouble
}
```

### Rows.Count

返回一个 **Long** 类型的值，该值代表集合中的行数。只读。

**语法**

**express.Count**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
//获取激活文档的表格1的行数
Application.ActiveDocument.Tables.Item(1).Rows.Count
```

### Rows.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Rows 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Rows.DistanceBottom

返回或设置文档文字与指定表格下边缘之间的距离（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.DistanceBottom**

\*express是一个代表 Rows 对象的变量。

**说明**

如果 **WrapAroundText** 为 **False** ，则该属性将不起作用。

**示例**

``` JavaScript
/*本示例设定文字环绕在活动文档中第一个表格的周围，并且环绕文字与该表格各边的间距为 20 磅。*/
function test()
{
let myRows = ActiveDocument.Tables.Item(1).Rows
myRows.WrapAroundText = true
myRows.DistanceLeft = 20
myRows.DistanceRight = 20
myRows.DistanceTop = 20
myRows.DistanceBottom = 20
}
```

### Rows.DistanceLeft

返回或设置文档文字与指定表格的左边缘之间的距离（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.DistanceLeft**

\*express是一个代表 Rows 对象的变量。

**说明**

如果 **WrapAroundText** 为 **False** ，则该属性将不起作用。

**示例**

``` JavaScript
/*本示例设定文字环绕在活动文档中第一个表格的周围，并且环绕文字与该表格各边的间距为 20 磅。*/
function test()
{
let myRows = ActiveDocument.Tables.Item(1).Rows
myRows.WrapAroundText = true
myRows.DistanceLeft = 20
myRows.DistanceRight = 20
myRows.DistanceTop = 20
myRows.DistanceBottom = 20
}
```

### Rows.DistanceRight

返回或设置文档文字与指定表格右边缘之间的距离（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.DistanceRight**

\*express是一个代表 Rows 对象的变量。

**说明**

如果 **WrapAroundText** 为 **False** ，则该属性将不起作用。

**示例**

``` JavaScript
/*本示例设定文字环绕在活动文档中第一个表格的周围，并且环绕文字与该表格各边的间距为 20 磅。*/
function test()
{
let myRows = ActiveDocument.Tables.Item(1).Rows
myRows.WrapAroundText = true
myRows.DistanceLeft = 20
myRows.DistanceRight = 20
myRows.DistanceTop = 20
myRows.DistanceBottom = 20
}
```

### Rows.DistanceTop

返回或设置文档文字与指定表格的上边缘之间的距离（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.DistanceTop**

\*express是一个代表 Rows 对象的变量。

**说明**

如果 **WrapAroundText** 为 **False** ，则该属性将不起作用。

**示例**

``` JavaScript
/*本示例设定文字环绕在活动文档中第一个表格的周围，并且环绕文字与该表格各边的间距为 20 磅。*/
function test()
{
let myRows = ActiveDocument.Tables.Item(1).Rows
myRows.WrapAroundText = true
myRows.DistanceLeft = 20
myRows.DistanceRight = 20
myRows.DistanceTop = 20
myRows.DistanceBottom = 20
}
```

### Rows.First

返回一个 **Row** 对象，该对象代表 **Rows** 集合中的第一个项目。

**语法**

**express.First**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
/*本示例对活动文档第一个表格的第一行应用底纹和下边框格式。*/
function test()
{
ActiveDocument.Tables.Item(1).Borders.Enable = false
let myFirst = ActiveDocument.Tables.Item(1).Rows.First
myFirst.Shading.Texture = wdTexture10Percent
myFirst.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleSingle
}
```

### Rows.HeadingFormat

如果将指定一行或数行的格式设置为表格标题，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.HeadingFormat**

\*express是一个代表 Rows 对象的变量。

**说明**

当表格不止一页时，可将多行的格式重复设置为表格标题。属性值可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*本示例在活动文档的开始处创建一个 5x5 表格，然后将表格第一行的格式设置为表格标题。*/
function test()
{
let rngTemp = ActiveDocument.Range(0, 0)
let tableNew = ActiveDocument.Tables.Add(rngTemp, 5, 5)
tableNew.Rows.HeadingFormat = true
}
```

``` JavaScript
/*本示例判定是否将插入点所在行的格式设置为表格标题。*/
function test()
{
if(Selection.Information(wdWithInTable)) {
    if(Selection.Rows.HeadingFormat) {
        MsgBox("The current row is a table heading")
    }
}
else {
    MsgBox("The insertion point is not in a table.")
}
}
```

### Rows.Height

返回或设置表格中指定的行的高度。Single 类型，可读写。

**语法**

**express.Height**

\*express是一个代表 Rows 对象的变量。

**说明**

如果指定行的 **HeightRule** 属性为 **wdRowHeightAuto** ，则 **Height** 返回 **wdUndefined** ；设置 **Height** 属性会将 **HeightRule** 设置为 **wdRowHeightAtLeast** 。

**示例**

``` JavaScript
/*本示例将活动文档第一张表格的行高度设置为至少 20 磅。*/
ActiveDocument.Tables.Item(1).Rows.Height = 20
```

### Rows.HeightRule

返回或设置确定指定单元格或行高度的规则。 **WdRowHeightRule** 类型，可读写。

**语法**

**express.HeightRule**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
/*本示例将确定所选行高度的规则设为按行中最高的单元格自动调整。*/
function test()
{
if(Selection.Information(wdWithInTable)) {
    Selection.Rows.HeightRule = wdRowHeightAuto
}
else {
    MsgBox("The insertion point is not in a table.")
}
}
```

### Rows.HorizontalPosition

返回或设置由 **RelativeHorizontalPosition** 属性指定的项目与行边缘之间的水平距离。可读写 **Single** 类型。

**语法**

**express.HorizontalPosition**

\*express是一个代表 Rows 对象的变量。

**说明**

该属性可以是一个以磅为单位表示度量的数字，也可以是 **WdTablePosition** 常量之一。如果 **WrapAroundText** 属性为 **False** ，则该属性无效。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个表格与右边距水平对齐。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1) {
    let myRows = ActiveDocument.Tables.Item(1).Rows
    myRows.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
    myRows.HorizontalPosition = wdTableRight
}
}
```

### Rows.Last

将 **Rows** 集合中的最后一项作为 **Row** 对象返回。

**语法**

**express.Last**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
/*以下示例删除活动文档中第一个表格中的最后一行。*/
ActiveDocument.Tables.Item(1).Rows.Last.Cells.Delete()
```

### Rows.LeftIndent

返回或设置一个 **Single** 类型的值，该值代表指定表格行的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
/*以下示例为活动文档中第一个表格的各行设置左缩进。*/
ActiveDocument.Tables.Item(1).Rows.LeftIndent = InchesToPoints(1)
```

### Rows.NestingLevel

返回指定表格行的嵌套层。只读 **Long** 类型。

**语法**

**express.NestingLevel**

\*express是一个代表 Rows 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

**示例**

``` JavaScript
/*以下示例新建一个文档，创建一个三层嵌套表格，并在每张表格的第一个单元格中填入该表格所在的嵌套层数。*/
function test()
{
Documents.Add()
ActiveDocument.Tables.Add(Selection.Range, 3, 3, wdWord9TableBehavior, wdAutoFitContent)

let myRange1 = ActiveDocument.Tables.Item(1).Range
myRange1.Copy()
myRange1.Cells.Item(1).Range.Text = myRange1.Cells.Item(1).NestingLevel
myRange1.Cells.Item(5).Range.PasteAsNestedTable()
    
let myRange2 = myRange1.Cells.Item(5).Tables.Item(1).Range
myRange2.Cells.Item(1).Range.Text = myRange2.Cells.Item(1).NestingLevel
myRange2.Cells.Item(5).Range.PasteAsNestedTable()

let myRange3 = myRange2.Cells.Item(5).Tables.Item(1).Range
myRange3.Cells.Item(1).Range.Text = myRange3.Cells.Item(1).NestingLevel
}
```

### Rows.Parent

返回一个 **Object** 类型值，该值代表指定 **Rows** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Rows 对象的变量。

### Rows.RelativeHorizontalPosition

指定一组行的相对水平位置。可读写 **WdRelativeHorizontalPosition** 类型。

**语法**

**express.RelativeHorizontalPosition**

\*express是一个代表 Rows 对象的变量。

### Rows.RelativeVerticalPosition

指定一组行的相对垂直位置。可读写 **WdRelativeVerticalPosition** 类型。

**语法**

**express.RelativeVerticalPosition**

\*express是一个代表 Rows 对象的变量。

### Rows.Shading

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
/*以下示例对所选内容的第一段应用黄色底纹。*/
function test()
{
let myShading = Selection.Paragraphs.Item(1).Shading
myShading.Texture = wdTexture12Pt5Percent
myShading.BackgroundPatternColorIndex = wdYellow
myShading.ForegroundPatternColorIndex = wdBlack
}
```

``` JavaScript
/*以下示例将横线纹理应用于表格 1 的第一行。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1) {
    let myShading = ActiveDocument.Tables.Item(1).Rows.Item(1).Shading
    myShading.Texture = wdTextureHorizontal
}
}
```

``` JavaScript
/*以下示例将活动文档中第一个单词的底纹设置为 10%。*/
ActiveDocument.Words.Item(1).Shading.Texture = wdTexture10Percent
```

### Rows.SpaceBetweenColumns

返回或设置指定的一行或多行中相邻列的文本之间的距离（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.SpaceBetweenColumns**

\*express是一个代表 Rows 对象的变量。

**示例**

``` JavaScript
/*以下示例返回所选表格行中各列之间的距离（以磅为单位）。*/
function test()
{
if(Selection.Information(wdWithInTable)) {
    MsgBox(Selection.Rows.SpaceBetweenColumns)
}
}
```

### Rows.TableDirection

返回或设置 WPS 对指定的表格或行中的单元格进行排序的方向。可读写 **WdTableDirection** 类型。

**语法**

**express.TableDirection**

\*express是一个代表 Rows 对象的变量。

**说明**

如果将 **TableDirection** 属性设为 **wdTableDirectionLtr** ，则选定的行按照最左侧的第一列进行排列。如果将 **TableDirection** 属性设为 **wdTableDirectionRtl** ，则选定的行按照最右侧的第一列进行排列。

**示例**

``` JavaScript
/*以下示例设置 WPS 从右向左对选定行中的单元格进行排序。*/
Selection.Rows.TableDirection = wdTableDirectionRtl
```

### Rows.VerticalPosition

返回或设置由 **RelativeVerticalPosition** 属性指定的项目与行边缘之间的垂直距离。可读写 **Single** 类型。

**语法**

**express.VerticalPosition**

\*express是一个代表 Rows 对象的变量。

**说明**

该属性可以是一个以磅为单位表示度量的数字，也可以是任何有效的 **WdFramePosition** 常量。

**示例**

``` JavaScript
/*以下示例将活动文档中的第一个表格与页面顶端垂直对齐。*/
function test()
{
let myTable = ActiveDocument.Tables.Item(1).Rows
myTable.RelativeVerticalPosition = wdRelativeVerticalPositionPage
myTable.VerticalPosition = wdTableTop
}
```

### Rows.WrapAroundText

返回或设置文本是否环绕指定行。 **Long** 类型，可读写。

**语法**

**express.WrapAroundText**

\*express是一个代表 Rows 对象的变量。

**说明**

如果设置仅环绕部分指定行，则返回 **wdUndefined** 。可设置为 **True** 或 **False** 。如果将 **WrapAroundText** 属性设置为 **False** ，则 **AllowOverlap** 属性也会设置为 **False** 。如果将 **AllowOverlap** 属性设置为 **True** ，则 **WrapAroundText** 属性也会设置为 **True** 。

**示例**

``` JavaScript
/*本示例设置 WPS ，使文本环绕在文档中第一张表格的周围。*/
ActiveDocument.Tables.Item(1).Rows.WrapAroundText = true
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Rows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ShadowFormat 对象

## 简介

代表形状的阴影格式。

## 说明

使用 Shadow 属性可返回一个 ShadowFormat 对象。以下示例向活动文档添加一个带阴影的矩形。半透明的蓝色阴影超出矩形右边 10 磅，高出矩形 10 磅。

``` JavaScript
function test() {
    let shps = Application.ActiveDocument.Shapes;
    let shp = shps.AddShape(msoShapeRectangle,50,50,100,200);
    let shd = shp.Shadow;
    shd.ForeColor.RGB=0x7f0000;
    shd.OffsetX = 10;
    shd.OffsetY = 10;
    shd.Size = 100;
    shd.Transparency = 0.5;
    shd.Visible = true;
}
```

## 属性

| 序号 | 名称                                             | 说明                                                                                                                       |
|:----:|:-------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------|
|  1   | [Blur](#ShadowFormat.Blur)                       | 返回或设置一个 Number 类型的值，该值代表阴影格式的模糊级别。可读写。                                                       |
|  2   | [ForeColor](#ShadowFormat.ForeColor)             | 返回或设置一个 ColorFormat 对象，该对象代表填充、线条或阴影的前景色。可读写。                                              |
|  3   | [OffsetX](#ShadowFormat.OffsetX)                 | 返回或设置阴影根据指定图形的水平偏移量（以磅为单位）。正值表示阴影向图形右侧偏移，负值表示向左偏移。 Number 类型，可读写。 |
|  4   | [OffsetY](#ShadowFormat.OffsetY)                 | 返回或设置阴影相对指定图形的垂直偏移量（以磅为单位）。 Number 类型，可读写。                                               |
|  5   | [RotateWithShape](#ShadowFormat.RotateWithShape) | 返回或设置一个 MsoTriState ，代表在旋转形象时是否旋转阴影。可读写。                                                        |
|  6   | [Size](#ShadowFormat.Size)                       | 返回或设置一个 Number 类型的值，该值代表阴影的宽度。可读写。                                                               |
|  7   | [Transparency](#ShadowFormat.Transparency)       | 返回或设置指定阴影的透明度，其值为从 0.0（不透明）到 1.0（全透明）。可读写 Number 类型。                                   |

## 成员属性

### ShadowFormat.Blur

返回或设置一个 **Number** 类型的值，该值代表阴影格式的模糊级别。可读写。

**语法**

**express.Blur**

\*express是一个代表 ShadowFormat 对象的变量。

### ShadowFormat.ForeColor

返回或设置一个 **ColorFormat** 对象，该对象代表填充、线条或阴影的前景色。可读写。

**语法**

**express.ForeColor**

\*express是一个代表 ShadowFormat 对象的变量。

### ShadowFormat.OffsetX

返回或设置阴影根据指定图形的水平偏移量（以磅为单位）。正值表示阴影向图形右侧偏移，负值表示向左偏移。 **Number** 类型，可读写。

**语法**

**express.OffsetX**

\*express是一个代表 ShadowFormat 对象的变量。

**说明**

如果不指明绝对位置又要从当前位置水平或垂直地微移阴影，请使用 **[IncrementOffsetX]()** 或 **[IncrementOffsetY]()** 方法。

**示例**

``` JavaScript
function test() {
    let shps = Application.ActiveDocument.Shapes;
    let shp = shps.AddShape(msoShapeRectangle,50,50,100,200);
    let shd = shp.Shadow;
    shd.ForeColor.RGB=0x7f0000;
    shd.OffsetX = 10;
    shd.OffsetY = 10;
    shd.Size = 100;
    shd.Transparency = 0.5;
    shd.Visible = true;
}
```

### ShadowFormat.OffsetY

返回或设置阴影相对指定图形的垂直偏移量（以磅为单位）。 **Number** 类型，可读写。

**语法**

**express.OffsetY**

\*express是一个代表 ShadowFormat 对象的变量。

**说明**

正值代表阴影向图形下偏移，负值代表向上偏移。如果不指明绝对位置又要从当前位置水平或垂直地微移阴影，请使用 **[IncrementOffsetX]()** 或 **[IncrementOffsetY]()** 方法。

**示例**

``` JavaScript
function test() {
    let shps = Application.ActiveDocument.Shapes;
    let shp = shps.AddShape(msoShapeRectangle,50,50,100,200);
    let shd = shp.Shadow;
    shd.ForeColor.RGB=0x7f0000;
    shd.OffsetX = 10;
    shd.OffsetY = 10;
    shd.Size = 100;
    shd.Transparency = 0.5;
    shd.Visible = true;
}
```

### ShadowFormat.RotateWithShape

返回或设置一个 **MsoTriState** ，代表在旋转形象时是否旋转阴影。可读写。

**语法**

**express.RotateWithShape**

\*express是一个代表 ShadowFormat 对象的变量。

### ShadowFormat.Size

返回或设置一个 **Number** 类型的值，该值代表阴影的宽度。可读写。

**语法**

**express.Size**

\*express是一个代表 ShadowFormat 对象的变量。

**示例**

``` JavaScript
/* 设置阴影尺寸为10 */
Application.ActiveDocument.Shapes.Item(1).Shadow.Size = 10
```

### ShadowFormat.Transparency

返回或设置指定阴影的透明度，其值为从 0.0（不透明）到 1.0（全透明）。可读写 **Number** 类型。

**语法**

**express.Transparency**

\*express是一个代表 ShadowFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ShadowFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TextFrame 对象

## 简介

代表 Shape 对象中的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;background-color:#FFFFFF;">文本框</span> 。 TextFrame 对象包含文本框架中的文本以及用于控制文本框架的边距和方向的属性。

## 说明

使用 TextFrame 属性可返回形状的 TextFrame 对象。 TextRange 属性可返回一个 Range 对象，该对象代表指定文本框架中的文本范围。以下示例在活动文档中第一个形状的文本框架中添加文本。

``` JavaScript
Application.ActiveDocument.Shapes.Item(1).TextFrame.TextRange.Text = "My Text"
```

| 注释                                                                                                                                      |
|-------------------------------------------------------------------------------------------------------------------------------------------|
| 有些形状不支持附加文本（如线条、任意多边形、图片和 OLE 对象）。如果试图返回或设置用于控制这些对象的文本框架中文本的属性，那么将发生错误。 |

使用 HasText 属性可确定文本框架中是否包含文本，如以下示例所示。

``` JavaScript
for(let i = 1; i <= Application.ActiveDocument.Shapes.Count; i++){
    if(Application.ActiveDocument.Shapes.Item(i).TextFrame.HasText){
        alert(Application.ActiveDocument.Shapes.Item(i).TextFrame.TextRange.Text)
    }
}
```

文本框架可以链接在一起，以使一个形状的文本框架中的文本排入另一个形状的文本框架。使用 Next 和 Previous 属性可链接文本框架。以下示例创建一个文本框（一个带有文本框架的矩形），并在其中添加文本。然后创建另一个文本框，并将两个文本框架链接在一起，以使第一个文本框架中的文本排入第二个文本框架。

``` JavaScript
let myTB1 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 72, 72, 72, 36)
    myTB1.TextFrame.TextRange.Text = "This is some text. This is some more text."

let myTB2 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 144, 144, 72, 36)
    myTB1.TextFrame.Next = myTB2.TextFrame
```

可以使用 ContainingRange 属性返回 Range 对象，该对象代表在链接的文本框架之间排列的整个文字部分。以下示例检查文本框 3 中的文本和链接到文本框 3 的其他所有文本的拼写。

``` JavaScript
let myStory = Application.ActiveDocument.Shapes.Item("TextBox 3").TextFrame.ContainingRange
myStory.CheckSpelling()
```

## 方法

| 序号 | 名称                                            | 说明                                                                  |
|:----:|:------------------------------------------------|:----------------------------------------------------------------------|
|  1   | [BreakForwardLink](#TextFrame.BreakForwardLink) | 在链接存在的情况下，断开指定文本框的前向链接。                        |
|  2   | [DeleteText](#TextFrame.DeleteText)             | 删除文本框中的文本以及该文本的所有关联属性，包括字体属性。            |
|  3   | [ValidLinkTarget](#TextFrame.ValidLinkTarget)   | 确定一个形状的文本框是否可以链接到另一个形状的文本框。 返回 Boolean值 |

## 属性

| 序号 | 名称                                            | 说明                                                                                                  |
|:----:|:------------------------------------------------|:------------------------------------------------------------------------------------------------------|
|  1   | [Application](#TextFrame.Application)           | 返回一个代表 WPS 应用程序的 Application 对象。                                                        |
|  2   | [AutoSize](#TextFrame.AutoSize)                 | 返回或设置一个 Long 类型的值，该值代表是否自动调整文本框的大小。可读写。                              |
|  3   | [Column](#TextFrame.Column)                     | 返回代表指定文本框架的列的 Column 对象。只读。                                                        |
|  4   | [ContainingRange](#TextFrame.ContainingRange)   | 返回一个 Range 对象，该对象代表一系列带有链接文本框架（包括指定文本框架）的图形的整个文字部分。只读。 |
|  5   | [Creator](#TextFrame.Creator)                   | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                          |
|  6   | [HasText](#TextFrame.HasText)                   | 如果指定的图形有相关的文字，则该属性值为 True 。 Boolean 类型，只读。                                 |
|  7   | [HorizontalAnchor](#TextFrame.HorizontalAnchor) | 返回或设置文本框架中文本的水平对齐方式。可读/写 MsoHorizontalAnchor 。                                |
|  8   | [MarginBottom](#TextFrame.MarginBottom)         | 以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读写。 Single 类型。          |
|  9   | [MarginLeft](#TextFrame.MarginLeft)             | 以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。 Single 类型。      |
|  10  | [MarginRight](#TextFrame.MarginRight)           | 以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。 Single 类型。      |
|  11  | [MarginTop](#TextFrame.MarginTop)               | 以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。 Single 类型。        |
|  12  | [Next](#TextFrame.Next)                         | 返回一个 TextFrame 对象，该对象代表形状集合中的下一个文本框架。只读。                                 |
|  13  | [NoTextRotation](#TextFrame.NoTextRotation)     | 如果在旋转形状时不应旋转文本框架中的文本，则该属性值为 True。可读/写 MsoTriState。                    |
|  14  | [Orientation](#TextFrame.Orientation)           | 返回或设置图文框内的文字方向。可读写 MsoTextOrientation 类型。                                        |
|  15  | [Overflowing](#TextFrame.Overflowing)           | 如果指定文本框架中的文本溢出该框架，则该属性值为 True 。 Boolean 类型，只读。                         |
|  16  | [Parent](#TextFrame.Parent)                     | 返回一个 Shape 对象，该对象代表文本框架的父形状。                                                     |
|  17  | [PathFormat](#TextFrame.PathFormat)             | 返回或设置指定文本框架的路径类型。可读/写 MsoPathType 。                                              |
|  18  | [Previous](#TextFrame.Previous)                 | 返回一个 TextFrame 对象，该对象代表形状集合中的上一个文本框架。只读。                                 |
|  19  | [TextRange](#TextFrame.TextRange)               | 返回一个 Range 对象，该对象代表指定文本框架中的文本。                                                 |
|  20  | [ThreeD](#TextFrame.ThreeD)                     | 返回 ThreeDFormat 对象，该对象包含指定文本框架的三维效果格式属性。只读。                              |
|  21  | [VerticalAnchor](#TextFrame.VerticalAnchor)     | 返回或设置一个 MsoVerticalAnchor 常量，该常量代表形状内文本的垂直对齐方式。可读写。                   |
|  22  | [WarpFormat](#TextFrame.WarpFormat)             | 返回或设置指定文本框架的环绕格式（文本如何环绕）。可读/写 MsoWarpFormat 。                            |
|  23  | [WordWrap](#TextFrame.WordWrap)                 | 如果 WPS 在指定文本框架的西文单词中间断字换行，则该属性值为 True 。 Long 类型，可读写。               |

## 成员方法

### TextFrame.BreakForwardLink

在链接存在的情况下，断开指定文本框的前向链接。

**语法**

**express.BreakForwardLink ()**

\*express是一个代表 TextFrame 对象的变量。

**说明**

如果将此方法应用于图形，且该图形位于一串与图文框相链接的图形的中间位置，则会使链接断开，而形成两组链接的图形。但所有文本都将保留在第一组链接的图形中。

**示例**

``` JavaScript
/*本示例实现的功能是：新建一个文档，在文档中添加三个前后链接的文本框，然后断开第二个文本框后面的链接。*/
function test() {
    Documents.Add()
    let shapeTextBox1 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, InchesToPoints(1.5), InchesToPoints(0.5), InchesToPoints(1), InchesToPoints(0.5))
    shapeTextBox1.TextFrame.TextRange.Text = "This is some text. This is some more text. This is even more text."

    let shapeTextBox2 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, InchesToPoints(1.5), InchesToPoints(1.5), InchesToPoints(1), InchesToPoints(0.5))
    let shapeTextBox3 = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, InchesToPoints(1.5), InchesToPoints(2.5), InchesToPoints(1), InchesToPoints(0.5))

    shapeTextBox1.TextFrame.Next = shapeTextBox2.TextFrame
    shapeTextBox2.TextFrame.Next = shapeTextBox3.TextFrame

    alert("TextBoxes 1, 2, and 3 are linked.")

    shapeTextBox2.TextFrame.BreakForwardLink()
}
```

### TextFrame.DeleteText

删除文本框中的文本以及该文本的所有关联属性，包括字体属性。

**语法**

**express.DeleteText ()**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例从活动文档的第一个形状中删除文本（如果该形状包含文本）。*/
function DeleteText_Example(){
    Application.ActiveDocument.Shapes.Item(1).TextFrame.DeleteText()
}
```

### TextFrame.ValidLinkTarget

确定一个形状的文本框是否可以链接到另一个形状的文本框。 返回 Boolean值

**语法**

**express.ValidLinkTarget ( *TargetTextFrame* )**

\*express是一个代表 TextFrame 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型  | 说明                                        |
|-------------------|-----------|-----------|---------------------------------------------|
| *TargetTextFrame* | 必选      | TextFrame | expression 返回的文本框要链接的目标文本框。 |

**说明**

如果 *TargetTextFrame* 是有效目标，则该方法返回 **True** ；如果 *TargetTextFrame* 已包含文字或已链接，或者该形状不支持附加文字，则返回 **False** 。

**示例**

``` JavaScript
/*本示例实现的功能是：检查活动文档第一和第二个图形的文本框，以判断其是否能相互链接。如果能，则链接两个文本框。*/
let textFrame1 = Application.ActiveDocument.Shapes.Item(1).TextFrame
let textFrame2 = Application.ActiveDocument.Shapes.Item(2).TextFrame

if(textFrame1.ValidLinkTarget(textFrame2) == true){
    textFrame1.Next = textFrame2
}
```

## 成员属性

### TextFrame.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 TextFrame 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### TextFrame.AutoSize

返回或设置一个 **Long** 类型的值，该值代表是否自动调整文本框的大小。可读写。

**语法**

**express.AutoSize**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.Column

返回代表指定文本框架的列的 Column 对象。只读。

**语法**

**express.Column**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
function Column_Example(){
    Application.ActiveDocument.Shapes.Item(1).TextFrame.Column.Number = 2
}
```

### TextFrame.ContainingRange

返回一个 Range 对象，该对象代表一系列带有链接文本框架（包括指定文本框架）的图形的整个文字部分。只读。

**语法**

**express.ContainingRange**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例对 TextBox 1 及所有链接至 TextBox 1 的文本框架的文本进行拼写检查。*/
let rngStory = Application.ActiveDocument.Shapes.Item("TextBox 1").TextFrame.ContainingRange
rngStory.CheckSpelling()
```

### TextFrame.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TextFrame 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### TextFrame.HasText

如果指定的图形有相关的文字，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.HasText**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*如果活动文档中的第二个图形包含文字，并且文字超出了图文框时，本示例显示一条消息。*/
let docActive = Application.ActiveDocument

if(docActive.Shapes.Item(2).TextFrame.HasText == true){
    if(docActive.Shapes.Item(2).TextFrame.Overflowing == true){
        MsgBox("Text overflows the frame.")
    }
}
```

### TextFrame.HorizontalAnchor

返回或设置文本框架中文本的水平对齐方式。可读/写 **MsoHorizontalAnchor** 。

**语法**

**express.HorizontalAnchor**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例演示如何将活动文档中第一个形状的对齐方式设置为靠上居中。*/
function HorizontalAnchor_Example(){
    Application.ActiveDocument.Shapes.Item(1).TextFrame.HorizontalAnchor = msoAnchorCenter
    Application.ActiveDocument.Shapes.Item(1).TextFrame.VerticalAnchor = msoAnchorTop
}
```

### TextFrame.MarginBottom

以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读写。 **Single** 类型。

**语法**

**express.MarginBottom**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
function test() {
    let myShape = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

    myShape.TextRange.Text = "Here is some test text"
    myShape.MarginBottom = 100
    myShape.MarginLeft = 100
    myShape.MarginRight = 50
    myShape.MarginTop = 20
}
```

### TextFrame.MarginLeft

以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。 **Single** 类型。

**语法**

**express.MarginLeft**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myShape = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

myShape.TextRange.Text = "Here is some test text"
myShape.MarginBottom = 100
myShape.MarginLeft = 100
myShape.MarginRight = 50
myShape.MarginTop = 20
```

### TextFrame.MarginRight

以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。 **Single** 类型。

**语法**

**express.MarginRight**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myShape = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

myShape.TextRange.Text = "Here is some test text"
myShape.MarginBottom = 100
myShape.MarginLeft = 100
myShape.MarginRight = 50
myShape.MarginTop = 20
```

### TextFrame.MarginTop

以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。 **Single** 类型。

**语法**

**express.MarginTop**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myShape = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

myShape.TextRange.Text = "Here is some test text"
myShape.MarginBottom = 100
myShape.MarginLeft = 100
myShape.MarginRight = 50
myShape.MarginTop = 20
```

### TextFrame.Next

返回一个 **TextFrame** 对象，该对象代表形状集合中的下一个文本框架。只读。

**语法**

**express.Next**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.NoTextRotation

如果在旋转形状时不应旋转文本框架中的文本，则该属性值为 True。可读/写 MsoTriState。

**语法**

**express.NoTextRotation**

\*express是一个代表 TextFrame 对象的变量。

**说明**

仅可使用 **MsoTriState** 常量 **msoFalse** 和 **msoTrue** 来设置该属性。

### TextFrame.Orientation

返回或设置图文框内的文字方向。可读写 **MsoTextOrientation** 类型。

**语法**

**express.Orientation**

\*express是一个代表 TextFrame 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **MsoTextOrientation** 常量可能不可用。

您可以设置文本框架或者恰好位于文本框架内的范围或所选内容的方向。有关文本框架和文本框之间的区别的信息，请参阅 **TextFrame** 对象。

**示例**

``` JavaScript
/*以下示例新建一篇文档，在文档中插入文字，接着使用这些文字创建一个文本框，然后设置该文本框架的方向，使其中的文字向上倾斜。*/
function test() {
    let myDoc = Application.Documents.Add()
    Application.Selection.TypeText("This is some text.")
    myDoc.Content.Select()
    Application.Selection.CreateTextbox()
    myDoc.Shapes.Item(1).TextFrame.Orientation = msoTextOrientationUpward
}
```

### TextFrame.Overflowing

如果指定文本框架中的文本溢出该框架，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.Overflowing**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例检查 MyTextBox 中的文本是否超出其文本框架。如果超出，本示例将添加另一个文本框架并链接这两个文本框，以使文本排入第二个文本框架。*/
let myTBox = Application.ActiveDocument.Shapes.Item("MyTextBox")

if(myTBox.TextFrame.Overflowing == true){
    let nextTBox = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 72, 72, 100, 200)
    myTBox.TextFrame.Next = nextTBox.TextFrame
}
```

### TextFrame.Parent

返回一个 **Shape** 对象，该对象代表文本框架的父形状。

**语法**

**express.Parent**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.PathFormat

返回或设置指定文本框架的路径类型。可读/写 **MsoPathType** 。

**语法**

**express.PathFormat**

\*express是一个代表 TextFrame 对象的变量。

**说明**

**PathFormat** 属性的值可以是下列 **MsoPathType** 常量之一：

- msoPathType1

- msoPathType2

- msoPathType3

- msoPathType4

- msoPathTypeMixed

- msoPathTypeNone

无法设置值 **msoPathTypeMixed** 。如果设置值 **msoPathTypeNone** ，则会删除所有现有路径。

### TextFrame.Previous

返回一个 **TextFrame** 对象，该对象代表形状集合中的上一个文本框架。只读。

**语法**

**express.Previous**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.TextRange

返回一个 Range 对象，该对象代表指定文本框架中的文本。

**语法**

**express.TextRange**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档中添加一个文本框，然后在文本框中添加文字。*/
let myTBox = Application.ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 200)
myTBox.TextFrame.TextRange.Text = "Test Box"

/*本示例在活动文档的文本框 1 中添加文字。*/
Application.ActiveDocument.Shapes.Item("TextBox 1").TextFrame.TextRange.InsertAfter("New Text")

/*本示例返回活动文档的文本框 1 中的文字，并显示在消息框中。*/
alert(ActiveDocument.Shapes.Item("TextBox 1").TextFrame.TextRange.Text)
```

### TextFrame.ThreeD

返回 **ThreeDFormat** 对象，该对象包含指定文本框架的三维效果格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.VerticalAnchor

返回或设置一个 **MsoVerticalAnchor** 常量，该常量代表形状内文本的垂直对齐方式。可读写。

**语法**

**express.VerticalAnchor**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.WarpFormat

返回或设置指定文本框架的环绕格式（文本如何环绕）。可读/写 **MsoWarpFormat** 。

**语法**

**express.WarpFormat**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例演示如何设置活动文档中第一个形状的环绕格式。*/
Application.ActiveDocument.Shapes.Item(1).TextFrame.WarpFormat = msoWarpFormat15
```

### TextFrame.WordWrap

如果 WPS 在指定文本框架的西文单词中间断字换行，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.WordWrap**

\*express是一个代表 TextFrame 对象的变量。

**说明**

如果该属性对指定文本框架中的某些指定文本设置为 **True** 而对其他文本设置为 false，则该属性返回 **wdUndefined** 。

| 注释                                                               |
|--------------------------------------------------------------------|
| 根据您所选择或安装的语言支持（例如，美国英语），该属性可能不可用。 |

> WPS 开发文档： [WPS 基础接口/文字 API 参考/TextFrame](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Window 对象

## 简介

代表一个窗口。许多文档特征（如滚动条和标尺）实际上是窗口的属性。

## 说明

Window 对象是 Windows 集合的一个成员。 Application 对象的 Windows 集合包含应用程序中的所有窗口，而 Document 对象的 Windows 集合仅包含显示指定文档的窗口。

使用 Windows ( *Index* ) 可返回一个 Window 对象，其中 *Index* 为窗口名称或索引号。以下示例将 Document1 窗口最大化。

``` JavaScript
Application.Windows.Item("Document1").WindowState = wdWindowStateMaximize
```

索引号是 “窗口” 菜单中窗口名称左侧的数字。以下示例显示 Windows 集合中第一个窗口的标题。

``` JavaScript
alert(Windows.Item(1).Caption)
```

使用 Add 方法或 NewWindow 方法可向 Windows 集合添加新窗口。下列每个语句都为活动窗口中的文档创建一个新窗口。

``` JavaScript
Application.ActiveDocument.ActiveWindow.NewWindow()
Application.Windows.Add()
```

对同一文档打开多个窗口时，窗口标题中将出现一个冒号 (:) 和一个数字。

在将视图切换到打印预览时，会创建一个新窗口。关闭打印预览时，该窗口从 Windows 集合中删除。

## 方法

| 序号 | 名称                                                     | 说明                                                                   |
|:----:|:---------------------------------------------------------|:-----------------------------------------------------------------------|
|  1   | [Activate](#Window.Activate)                             | 激活指定窗口。                                                         |
|  2   | [Close](#Window.Close)                                   | 关闭指定的窗口。                                                       |
|  3   | [GetPoint](#Window.GetPoint)                             |                                                                        |
|  4   | [LargeScroll](#Window.LargeScroll)                       | 将窗口或窗格滚动指定的屏数。                                           |
|  5   | [NewWindow](#Window.NewWindow)                           | 打开一个新窗口，该窗口与指定窗口显示同一个文档。返回一个 Window 对象。 |
|  6   | [PageScroll](#Window.PageScroll)                         | 使指定的窗格或窗口逐页滚动。                                           |
|  7   | [PrintOut](#Window.PrintOut)                             | 打印指定窗口中显示的全部或部分文档。                                   |
|  8   | [RangeFromPoint](#Window.RangeFromPoint)                 | 返回 Range 或 Shape 对象，该对象位于由屏幕位置坐标对指定的位置。       |
|  9   | [ScrollIntoView](#Window.ScrollIntoView)                 | 滚动文档窗口，以便在文档窗口显示指定的区域或图形。                     |
|  10  | [SetFocus](#Window.SetFocus)                             | 在指定的文档窗口为电子邮件设置焦点。                                   |
|  11  | [SmallScroll](#Window.SmallScroll)                       | 将窗口或窗格滚动指定的行数。                                           |
|  12  | [ToggleRibbon](#Window.ToggleRibbon)                     | 显示或隐藏功能区。                                                     |
|  13  | [ToggleShowAllReviewers](#Window.ToggleShowAllReviewers) | 显示或隐藏包含备注和修订的文档中的所有备注。                           |

## 属性

| 序号 | 名称                                                             | 说明                                                                                                                |
|:----:|:-----------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------|
|  1   | [Active](#Window.Active)                                         | 如果指定窗口处于活动状态，则该属性值为 True 。 Boolean 类型，只读。                                                 |
|  2   | [ActivePane](#Window.ActivePane)                                 | 返回 Pane 对象，该对象代表指定窗口的活动窗格。只读。                                                                |
|  3   | [Application](#Window.Application)                               | 返回一个代表 WPS Office WPS 应用程序的 Application 对象。                                                           |
|  4   | [Caption](#Window.Caption)                                       | 返回或设置显示在文档或应用程序窗口标题栏中的窗口题注文字。 String 类型，可读写。                                    |
|  5   | [Creator](#Window.Creator)                                       | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                        |
|  6   | [DisplayHorizontalScrollBar](#Window.DisplayHorizontalScrollBar) | 如果在指定的窗口中显示水平滚动条，则该属性值为 True 。 Boolean 类型，可读写。                                       |
|  7   | [DisplayLeftScrollBar](#Window.DisplayLeftScrollBar)             | 如果垂直滚动条显示在文档窗口左侧，则该属性值为 True 。 Boolean 类型，可读写。                                       |
|  8   | [DisplayRightRuler](#Window.DisplayRightRuler)                   | 如果垂直标尺显示在页面视图中文档窗口右侧，则该属性值为 True 。 Boolean 类型，可读写。                               |
|  9   | [DisplayRulers](#Window.DisplayRulers)                           | 如果在指定的窗口或窗格中显示标尺，则该属性值为 True 。 Boolean 类型，可读写。                                       |
|  10  | [DisplayScreenTips](#Window.DisplayScreenTips)                   | 如果批注、脚注、尾注和超链接以提示形式显示，则该属性值为 True 。可读/写 Boolean 类型。                              |
|  11  | [DisplayVerticalRuler](#Window.DisplayVerticalRuler)             | 如果在指定的窗口或窗格中显示垂直标尺，则该属性值为 True 。 Boolean 类型，可读写。                                   |
|  12  | [Document](#Window.Document)                                     | 返回一个 Document 对象，该对象与指定窗格、窗口或选定内容相关。只读。                                                |
|  13  | [DocumentMap](#Window.DocumentMap)                               | 如果显示文档结构图，则该属性值为 True 。 Boolean 类型，可读写。                                                     |
|  14  | [EnvelopeVisible](#Window.EnvelopeVisible)                       | 如果电子邮件的页眉在文档窗口是可见的，则该属性值为 True 。默认值为 False 。 Boolean 类型，可读写。                  |
|  15  | [Height](#Window.Height)                                         | 返回或设置窗口的高度。Long 类型，可读写。                                                                           |
|  16  | [HorizontalPercentScrolled](#Window.HorizontalPercentScrolled)   | 返回或设置水平滚动条位置（以文档宽度的百分比表示）。可读写 Long 类型。                                              |
|  17  | [IMEMode](#Window.IMEMode)                                       | 返回或设置日文输入法编辑器 (IME) 的默认启动模式。 WdIMEMode 类型，可读写。                                          |
|  18  | [Index](#Window.Index)                                           | 返回一个 Long 类型的值，该值代表项目在集合中的位置。只读。                                                          |
|  19  | [Left](#Window.Left)                                             | 返回或设置一个 Long 类型的值，该值代表指定窗口的水平位置，以磅为单位。可读写。                                      |
|  20  | [Next](#Window.Next)                                             | 返回打开的文档窗口集合中的下一个文档窗口。只读。                                                                    |
|  21  | [Panes](#Window.Panes)                                           | 返回一个 Panes 集合，该集合代表指定窗口中所有的窗格。                                                               |
|  22  | [Parent](#Window.Parent)                                         | 返回一个 Object 类型值，该值代表指定 Window 对象的父对象。                                                          |
|  23  | [Previous](#Window.Previous)                                     | 返回打开的文档窗口集合中的上一个文档窗口。只读。                                                                    |
|  24  | [Selection](#Window.Selection)                                   | 返回一个 Selection 对象，该对象代表一个选定范围或插入点。只读。                                                     |
|  25  | [ShowSourceDocuments](#Window.ShowSourceDocuments)               | 返回或设置一个 WdShowSourceDocuments 常量，该常量代表 WPS 在比较和合并过程之后显示源文档的方式。可读写。            |
|  26  | [Split](#Window.Split)                                           | 如果将窗口拆分为多个窗格，则该属性值为 True 。 Boolean 类型，可读写。                                               |
|  27  | [SplitVertical](#Window.SplitVertical)                           | 返回或设置指定窗口的垂直拆分百分比。 Long 类型，可读写。                                                            |
|  28  | [StyleAreaWidth](#Window.StyleAreaWidth)                         | 该属性返回或设置样式区的宽度，以磅为单位。 Single 类型，可读写。                                                    |
|  29  | [Thumbnails](#Window.Thumbnails)                                 | 设置或返回一个 Boolean 类型的值，代表是否沿 WPS 文档窗口的左侧显示文档页面的缩略图。                                |
|  30  | [Top](#Window.Top)                                               | 返回或设置指定文档窗口的垂直位置（以磅为单位）。可读写 Long 类型。                                                  |
|  31  | [Type](#Window.Type)                                             | 返回窗口的类型。只读 WdWindowType 类型。                                                                            |
|  32  | [UsableHeight](#Window.UsableHeight)                             | 返回指定的文档窗口中活动工作区的高度（以磅为单位）。 Long 类型，只读。                                              |
|  33  | [UsableWidth](#Window.UsableWidth)                               | 返回指定的文档窗口中活动工作区的宽度（以磅为单位）。 Long 类型，只读。                                              |
|  34  | [VerticalPercentScrolled](#Window.VerticalPercentScrolled)       | 返回或设置垂直滚动条的位置（以文档长度的百分比表示）。可读写 Long 类型。                                            |
|  35  | [View](#Window.View)                                             | 返回一个 View 对象，该对象代表指定窗口或窗格的视图。                                                                |
|  36  | [Visible](#Window.Visible)                                       | 如果指定对象可见，则该属性值为 True 。 Boolean 类型，可读写。                                                       |
|  37  | [Width](#Window.Width)                                           | 返回或设置指定文档窗口的宽度（以磅为单位）。可读写 Long 类型。                                                      |
|  38  | [WindowNumber](#Window.WindowNumber)                             | 该属性返回在指定窗口中显示的文档窗口序号。例如，如果窗口标题为“Sales.doc:2”，则该属性返回数值 2。 Long 类型，只读。 |
|  39  | [WindowState](#Window.WindowState)                               | 返回或设置指定文档窗口或任务窗口的状态。 WdWindowState 类型，可读写。                                               |

## 成员方法

### Window.Activate

激活指定窗口。

**语法**

**express.Activate ()**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例激活 Windows 集合中的下一个窗口。*/
function NextWindow() {
    //Two or more documents must be open for this statement to execute.
    ActiveDocument.ActiveWindow.Next.Activate()
}
```

### Window.Close

关闭指定的窗口。

**语法**

**express.Close ( *SaveChanges* , *RouteDocument* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                |
|-----------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------|
| *SaveChanges*   | 可选      | Variant  | 指定保存文档的操作。可以是下列 WdSaveOptions 常量之一：wdDoNotSaveChanges、wdPromptToSaveChanges 或 wdSaveChanges。 |
| *RouteDocument* | 可选      | Variant  | 如果该参数值为 True，则将文档传送给下一个收件人。如果文档没有附加的传送名单，则忽略此参数。                         |

**示例**

``` JavaScript
/*以下示例关闭活动文档的活动窗口并保存文档。*/
ActiveDocument.ActiveWindow.Close(wdSaveChanges)
```

### Window.GetPoint

**语法**

**express.GetPoint ( *ScreenPixelsLeft* , *ScreenPixelsTop* , *ScreenPixelsWidth* , *ScreenPixelsHeight* , *obj* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                      |
|----------------------|-----------|----------|-------------------------------------------|
| *ScreenPixelsLeft*   | 必选      | Long     | 需要 WPS 为对象的左边界值返回的变量名。   |
| *ScreenPixelsTop*    | 必选      | Long     | 需要 WPS 为对象的顶部边界值返回的变量名。 |
| *ScreenPixelsWidth*  | 必选      | Long     | 需要 WPS 为对象的宽度返回的变量名。       |
| *ScreenPixelsHeight* | 必选      | Long     | 需要 WPS 为对象的高度值返回的变量名。     |
| *obj*                | 必选      | Object   | Range 或 Shape 对象。                     |

**说明**

如果在屏幕上看不到整个区域或图形，则导致出错。

**示例**

``` JavaScript
/*本示例检查当前选定对象并返回其屏幕坐标。*/
function test()
{
let pLeft=0.1
let pTop=0.1
let pWidth=0.1
let pHeight=0.1
ActiveWindow.GetPoint(pLeft, pTop, pWidth, pHeight, Selection.Range)
MsgBox("Left = " + pLeft + "\n" +
        "Top = " + pTop + "\n" +
        "Width = " + pWidth + "\n" +
        "Height = " + pHeight)
}
```

### Window.LargeScroll

将窗口或窗格滚动指定的屏数。

**语法**

**express.LargeScroll ( *Down* , *Up* , *ToRight* , *ToLeft* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                 |
|-----------|-----------|----------|----------------------|
| *Down*    | 可选      | Variant  | 向下滚动窗口的屏数。 |
| *Up*      | 可选      | Variant  | 向上滚动窗口的屏数。 |
| *ToRight* | 可选      | Variant  | 向右滚动窗口的屏数。 |
| *ToLeft*  | 可选      | Variant  | 向左滚动窗口的屏数。 |

**说明**

此方法等效于在水平和垂直滚动条上单击滚动块以外的位置。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动屏数为这两个参数的差。例如，如果 *Down* 为 2， *Up* 为 4，则窗口向上滚动两屏。同样，如果指定了 *ToLeft* 和 *ToRight* 两个参数，则窗口滚动屏数也为这两个参数的差。

这四个参数都可以取负值。如果没有指定参数，窗口将向下滚动一屏。

**示例**

``` JavaScript
/*以下示例将活动窗口下滚一屏。*/
ActiveDocument.ActiveWindow.LargeScroll(1)
```

``` JavaScript
/*以下示例拆分活动窗口，并且上滚两屏，右滚一屏。*/
function test()
{
ActiveDocument.ActiveWindow.Split = true
ActiveDocument.ActiveWindow.LargeScroll(null, 2, 1)
}
```

### Window.NewWindow

打开一个新窗口，该窗口与指定窗口显示同一个文档。返回一个 **Window** 对象。

**语法**

**express.NewWindow ()**

\*express是一个代表 Window 对象的变量。

**返回值**

Window

**说明**

当多个窗口打开同一个文档时，将在窗口标题中显示一个冒号 (:) 和一个数字。下列两条指令在功能上是等效的。

``` JavaScript
function test()
{
let myWindow = ActiveDocument.ActiveWindow.NewWindow()
let myWindow = NewWindow()
}
```

**示例**

``` JavaScript
/*以下示例在为 Document1 打开一个新窗口之前和之后各发布一条消息，以指示现存窗口的数目。*/
function test()
{
MsgBox(Windows.Count + " windows open")
Windows.Item("Document1").NewWindow()
MsgBox(Windows.Count + " windows open")
}
```

### Window.PageScroll

使指定的窗格或窗口逐页滚动。

**语法**

**express.PageScroll ( *Down* , *Up* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                               |
|--------|-----------|----------|----------------------------------------------------|
| *Down* | 可选      | Variant  | 向下滚动的页数。如果省略此参数，则参数值设置为 1。 |
| *Up*   | 可选      | Variant  | 向上滚动的页数。                                   |

**说明**

**PageScroll** 方法仅用于页面视图或 Web 版式视图。此方法不影响插入点的位置。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动页数为这两个参数的差。例如，如果 *Down* 为 2， *Up* 为 4，则窗口向上滚动两页。

**示例**

``` JavaScript
/*以下示例使活动窗口向下滚动三页。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.PageScroll(3)
}
```

``` JavaScript
/*以下示例使活动窗口向下滚动一页。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.PageScroll()
}
```

### Window.PrintOut

打印指定窗口中显示的全部或部分文档。

**语法**

**express.PrintOut ( *Background* , *Append* , *Range* , *OutputFileName* , *From* , *To* , *Item* , *Copies* , *Pages* , *PageType* , *PrintToFile* , *Collate* , *FileName* , *ActivePrinterMacGX* , *ManualDuplexPrint* , *PrintZoomColumn* , *PrintZoomRow* , *PrintZoomPaperWidth* , *PrintZoomPaperHeight* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                 |
|------------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Background*           | 可选      | Variant  | 如果将该属性设置为 True，则 WPS 在打印文档时继续运行宏。                                                                                                                                                                                                                                             |
| *Append*               | 可选      | Variant  | 如果将该属性设置为 True，则将指定文档附加到 OutputFileName 参数所指定的文件名。如果将该属性设置为 False，则覆盖 OutputFileName 参数所指定文件的内容。                                                                                                                                                |
| *Range*                | 可选      | Variant  | 页码范围。可以是任意 WdPrintOutRange 常量。                                                                                                                                                                                                                                                          |
| *OutputFileName*       | 可选      | Variant  | 如果 PrintToFile 为 True，则该参数指定输出文件的路径和文件名。                                                                                                                                                                                                                                       |
| *From*                 | 可选      | Variant  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定起始页码。                                                                                                                                                                                                                                            |
| *To*                   | 可选      | Variant  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定结束页码。                                                                                                                                                                                                                                            |
| *Item*                 | 可选      | Variant  | 要打印的项目。可以是任意 WdPrintOutItem 常量。                                                                                                                                                                                                                                                       |
| *Copies*               | 可选      | Variant  | 要打印的份数。                                                                                                                                                                                                                                                                                       |
| *Pages*                | 可选      | Variant  | 要打印的页码和页码范围，中间用逗号分开。例如，“2, 6-10”表示打印第 2 页以及第 6 至第 10 页。                                                                                                                                                                                                          |
| *PageType*             | 可选      | Variant  | 要打印的页面类型。可以是任意 WdPrintOutPages 常量。                                                                                                                                                                                                                                                  |
| *PrintToFile*          | 可选      | Variant  | 如果该属性值为 True，则将打印指令发送到文件。请确保使用 OutputFileName 指定文件名。                                                                                                                                                                                                                  |
| *Collate*              | 可选      | Variant  | 在打印文档的多份副本时，如果该属性值为 True，则完成打印所有页面后再打印下一份副本。                                                                                                                                                                                                                  |
| *FileName*             | 可选      | Variant  | 要打印的文档的路径和文件名。如果省略该参数， WPS 将打印活动文档（仅应用于 Application 对象）。                                                                                                                                                                                                       |
| *ActivePrinterMacGX*   | 可选      | Variant  | 该参数仅适用于 WPS OfficeMacintosh Edition。有关该参数的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。                                                                                                                                                                            |
| *ManualDuplexPrint*    | 可选      | Variant  | 如果该属性值为 True，则在无双面打印组件的打印机上打印双面文档。如果该属性值为 True，则忽略 PrintBackground 和 PrintReverse 属性。使用 PrintOddPagesInAscendingOrder 和 PrintEvenPagesInAscendingOrder 属性可在手动双面打印时控制输出。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *PrintZoomColumn*      | 可选      | Variant  | 表示 WPS 在一页纸上水平放置的页数。可以是 1、2、3 或 4 页。与 PrintZoomRow 参数一同使用可在单张纸上打印多页文档。                                                                                                                                                                                    |
| *PrintZoomRow*         | 可选      | Variant  | 表示 WPS 在一页纸上垂直放置的页数。可以是 1、2 或 4 页。与 PrintZoomColumn 参数一同使用可在单张纸上打印多页文档。                                                                                                                                                                                    |
| *PrintZoomPaperWidth*  | 可选      | Variant  | WPS 将打印页面缩放到的宽度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。                                                                                                                                                                                                                             |
| *PrintZoomPaperHeight* | 可选      | Variant  | WPS 将打印页面缩放到的高度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。                                                                                                                                                                                                                             |

**示例**

``` JavaScript
/*本示例打印活动文档的当前页面。*/
ActiveDocument.PrintOut(null,null,wdPrintCurrentPage)
```

``` JavaScript
/*本示例打印当前文件夹中的所有文档。Dir 函数用于返回所有扩展名为“.doc”的文件名。*/
function test()
{
adoc = Dir("*.DOC")
Do While adoc <> ""
    Application.PrintOut FileName:=adoc
    adoc = Dir()
Loop
}
```

``` JavaScript
/*本示例打印活动窗口中文档的前三页。*/
ActiveDocument.ActiveWindow.PrintOut(null,null,wdPrintFromTo,null,"1","3")
```

``` JavaScript
/*本示例打印活动文档中的备注。*/
function test()
{
if(ActiveDocument.Comments.Count >= 1){
    ActiveDocument.PrintOut(null,null,null,null,null,null,wdPrintComments)
}
}
```

``` JavaScript
/*本示例将打印活动文档，每张纸上打印六页文档。*/
ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,3,2)
```

``` JavaScript
/*本示例按实际尺寸的 75% 打印活动文档。*/
ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,0.75 * (8.5 * 1440),0.75 * (11 * 1440))
```

### Window.RangeFromPoint

返回 **Range** 或 **Shape** 对象，该对象位于由屏幕位置坐标对指定的位置。

**语法**

**express.RangeFromPoint ( *x* , *y* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                                             |
|------|-----------|----------|--------------------------------------------------|
| *x*  | 必选      | Long     | 从屏幕左边缘到该点的水平距离（以像素为单位）。   |
| *y*  | 必选      | Long     | 从屏幕顶部边缘到该点的垂直距离（以像素为单位）。 |

**返回值**

Object

**说明**

如果在坐标对指定的位置中没有区域或图形，则该方法返回 **Nothing** 。

**示例**

``` JavaScript
/*本示例可实现的功能是：新建一个文档，并在其中添加一个五角星。然后便获取该图形的屏幕位置并计算其中心位置。用这些坐标（本示例使用 RangeFromPoint 方法）来返回一个到该图形的引用并且改变其填充颜色。*/
function test()
{
let pLeft = 0.1
let pTop = 0.1
let pWidth = 0.1
let pHeight = 0.1
let newShape
let newDoc = Documents.Add()

newDoc.Shapes.AddShape(msoShape5pointStar, 288, 100, 100, 72)
newDoc.ActiveWindow.GetPoint(pLeft, pTop, pWidth, pHeight, newDoc.Shapes.Item(1))
newShape = newDoc.ActiveWindow.RangeFromPoint(pLeft + pWidth * 0.5, pTop + pHeight * 0.5)
newShape.Fill.ForeColor.RGB = (80, 160, 130)
}
```

### Window.ScrollIntoView

滚动文档窗口，以便在文档窗口显示指定的区域或图形。

**语法**

**express.ScrollIntoView ( *Obj* , *Start* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                      |
|---------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------|
| *Obj*   | 必选      | Object   | Range 或 Shape 对象。                                                                                                                     |
| *Start* | 可选      | Boolean  | 如果该区域或形状的左上角出现在文档窗口的左上角，则为 True 。如果该区域或形状的右下角出现在文档窗口的右下角，则为 False 。默认值为 True 。 |

**说明**

如果该区域或图形大于文档窗口， *Start* 参数可以指定该区域或图形的显示部分或获得初始焦点的部分。此方法无法在大纲视图中使用。

**示例**

``` JavaScript
/*本示例滚动活动文档，以便可以在文档窗口看到选定内容*/
Application.ActiveWindow.ScrollIntoView(Selection.Range, true)
```

### Window.SetFocus

在指定的文档窗口为电子邮件设置焦点。

**语法**

**express.SetFocus ()**

\*express是一个代表 Window 对象的变量。

**说明**

如果该文档不是电子邮件，则此方法将不起作用。

**示例**

``` JavaScript
/*本示例实现的功能是：显示电子邮件的页眉，并给邮件正文设置焦点。*/
function test()
{
ActiveWindow.EnvelopeVisible = true
ActiveWindow.SetFocus()
}
```

### Window.SmallScroll

将窗口或窗格滚动指定的行数。

**语法**

**express.SmallScroll ( *Down* , *Up* , *ToRight* , *ToLeft* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                             |
|-----------|-----------|----------|----------------------------------------------------------------------------------|
| *Down*    | 可选      | Variant  | 向下滚动窗口的行数。一“行”对应于单击一次垂直滚动条上的向下滚动箭头所滚动的距离。 |
| *Up*      | 可选      | Variant  | 向上滚动窗口的行数。一“行”对应于单击一次垂直滚动条上的向上滚动箭头所滚动的距离。 |
| *ToRight* | 可选      | Variant  | 向右滚动窗口的行数。一“行”对应于单击一次水平滚动条上的向右滚动箭头所滚动的距离。 |
| *ToLeft*  | 可选      | Variant  | 向左滚动窗口的行数。一“行”对应于单击一次水平滚动条上的向左滚动箭头所滚动的距离。 |

**说明**

此方法等效于单击水平和垂直滚动条上的滚动箭头。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动行数为这两个参数的差。例如，如果 *Down* 为 3， *Up* 为 6，则窗口将向上滚动三行。同样，如果指定了 *ToLeft* 和 *ToRight* 两个参数，则窗口滚动行数也为这两个参数的差。

这四个参数都可以取负值。如果没有指定参数，窗口将下滚一行。

**示例**

``` JavaScript
/*以下示例将活动窗口下滚一行。*/
ActiveDocument.ActiveWindow.SmallScroll(1)
```

``` JavaScript
/*以下示例拆分活动窗口，并上滚及左滚窗口。*/
function test()
{
ActiveDocument.ActiveWindow.Split = true
ActiveDocument.ActiveWindow.SmallScroll(null, 5, null, 5)
}
```

### Window.ToggleRibbon

显示或隐藏功能区。

**语法**

**express.ToggleRibbon ()**

\*express是一个代表 Window 对象的变量。

**说明**

如果功能区可见，则 **TobbleRibbon** 方法隐藏功能区。如果功能区隐藏，则 **ToggleRibbon** 方法显示功能区。

### Window.ToggleShowAllReviewers

显示或隐藏包含备注和修订的文档中的所有备注。

**语法**

**express.ToggleShowAllReviewers ()**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动文档中的所有备注。本示例假设文档中存在备注。*/
function test()
{
Application.ActiveWindow.ToggleShowAllReviewers()
}
```

## 成员属性

### Window.Active

如果指定窗口处于活动状态，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.Active**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*如果 Windows 集合中的第一个窗口当前不处于活动状态，则本示例将激活该窗口。*/
function ActiveWin() {
    if(Windows.Item(1).Active == fasle){
        Windows.Item(1).Activate()
    }
}
```

### Window.ActivePane

返回 **Pane** 对象，该对象代表指定窗口的活动窗格。只读。

**语法**

**express.ActivePane**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例拆分活动窗口，然后激活活动窗格的下一个窗格。*/
function test()
{
let win = ActiveDocument.ActiveWindow
win.Split = true
win.ActivePane.Next.Activate()
MsgBox("Pane " + win.ActivePane.Index + " is active")
}
```

``` JavaScript
/*本示例激活第一个窗口，并显示活动窗格中的制表符。*/
function test()
{
Application.Windows.Item(1).Activate()
Application.Windows.Item(1).ActivePane.View.ShowTabs = true
}
```

### Window.Application

返回一个代表 WPS Office WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Window 对象的变量。

**说明**

可以通过 Visual Basic 的 **CreateObject** 和 **GetObject** 函数从 示例代码 (VBA) 项目访问 OLE 自动化对象。

### Window.Caption

返回或设置显示在文档或应用程序窗口标题栏中的窗口题注文字。 **String** 类型，可读写。

**语法**

**express.Caption**

\*express是一个代表 Window 对象的变量。

**说明**

要将应用程序窗口的题注改为默认文本，请将该属性设置为空字符串（""）。

**示例**

``` JavaScript
/*本示例显示 Windows 集合中每个窗口的题注*/
for(let i=1;i<=Windows.Count;i++){
    MsgBox(Windows.Item(i).Caption, "Window" + i + " Caption")
}

/*本示例重新设置应用程序窗口的题注。*/
Application.Caption = ""

/*本示例将活动窗口的题注设置为与活动文档的名称相同。*/
Application.ActiveDocument.ActiveWindow.Caption = ActiveDocument.FullName

/*本示例改变 WPS 应用程序窗口的题注，使其包括用户名。*/
Application.Caption = UserName + "'s copy of Word"

/*本示例插入一个“表格”题注，然后将第一个图表目录的题注改为“表格”。*/
Application.Selection.Collapse(wdCollapseStart)
Application.Selection.Range.InsertCaption("Table")
if(Application.ActiveDocument.TablesOfFigures.Count >= 1){
    Application.ActiveDocument.TablesOfFigures.Item(1).Caption = "Table"
}
```

### Window.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Window 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Window.DisplayHorizontalScrollBar

如果在指定的窗口中显示水平滚动条，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayHorizontalScrollBar**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例在活动窗口中显示水平和垂直滚动条。*/
function test()
{
ActiveDocument.ActiveWindow.DisplayHorizontalScrollBar = true
ActiveDocument.ActiveWindow.DisplayVerticalScrollBar = true
}
```

``` JavaScript
/*本示例切换 Document1 窗口中的水平滚动条。*/
function test()
{
let winTemp = Windows.Item("Document1")
winTemp.DisplayHorizontalScrollBar = !winTemp.DisplayHorizontalScrollBar
}
```

### Window.DisplayLeftScrollBar

如果垂直滚动条显示在文档窗口左侧，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayLeftScrollBar**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例在活动窗口左侧显示垂直滚动条。*/
ActiveWindow.DisplayLeftScrollBar = true
```

### Window.DisplayRightRuler

如果垂直标尺显示在页面视图中文档窗口右侧，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayRightRuler**

\*express是一个代表 Window 对象的变量。

**说明**

有关在 WPS 中使用从右向左语言的详细信息，请参阅 WPS 从右向左的语言功能。

**示例**

``` JavaScript
/*本示例将活动窗口设为页面视图，并在窗口右侧显示垂直标尺。*/
function test()
{
ActiveWindow.View = wdPrintView
ActiveWindow.DisplayRightRuler = true
}
```

### Window.DisplayRulers

如果在指定的窗口或窗格中显示标尺，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayRulers**

\*express是一个代表 Window 对象的变量。

**说明**

该属性与 **“视图”** 菜单中的“标尺” 命令等效。如果 **DisplayRulers** 为 **False** ，则无论 **DisplayVerticalRuler** 属性的状态如何，都不显示水平和垂直标尺。

**示例**

``` JavaScript
/*本示例切换活动窗口中的标尺显示。*/
ActiveDocument.ActiveWindow.DisplayRulers = !ActiveDocument.ActiveWindow.DisplayRulers
```

``` JavaScript
/*本示例将窗口切换到页面视图并显示水平和垂直标尺。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.DisplayVerticalRuler = true
ActiveDocument.ActiveWindow.DisplayRulers = true
}
```

### Window.DisplayScreenTips

如果批注、脚注、尾注和超链接以提示形式显示，则该属性值为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.DisplayScreenTips**

\*express是一个代表 Window 对象的变量。

**说明**

突出显示那些标记有批注的文本。

**示例**

``` JavaScript
/*本示例使 WPS 将批注、脚注和尾注显示为提示。而且突出显示那些标记有批注的文本。*/
Application.DisplayScreenTips = true
```

``` JavaScript
/*本示例返回“ WPS 选项”对话框中“显示”选项卡上“页面显示选项”部分中“悬停时显示文档工具提示”复选框的当前状态。*/
let temp = Application.DisplayScreenTips
```

### Window.DisplayVerticalRuler

如果在指定的窗口或窗格中显示垂直标尺，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayVerticalRuler**

\*express是一个代表 Window 对象的变量。

**说明**

垂直标尺只出现在页面视图中，且 **DisplayRulers** 属性必须设置为 **True** 。

**示例**

``` JavaScript
/*本示例将 Windows 集合中的各个窗口切换到页面视图并显示水平和垂直标尺。*/
function test()
{
let winLoop 
for(let i=1;i<=Windows.Count;i++){
    winLoop = Windows.Item(i)
    winLoop.View.Type = wdPrintView
    winLoop.DisplayRulers = true
    winLoop.DisplayVerticalRuler = true
}
}
```

``` JavaScript
/*本示例隐藏活动窗口的水平和垂直标尺。*/
function test()
{
ActiveDocument.ActiveWindow.DisplayVerticalRuler = false
ActiveDocument.ActiveWindow.DisplayRulers = false
}
```

### Window.Document

返回一个 **Document** 对象，该对象与指定窗格、窗口或选定内容相关。只读。

**语法**

**express.Document**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*该示例将 myDoc 设置为与活动窗口相关的文档。然后焦点切换至下一个窗口，并加以拆分。再使用 Activate 方法切换回原文档。*/
function test()
{
let myDoc = Application.ActiveWindow.Document
if(Windows.Count >= 2) { 
    Application.ActiveWindow.Next.Activate()
    Application.ActiveWindow.Split = true
    myDoc.Activate()
}
}
```

### Window.DocumentMap

如果显示文档结构图，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DocumentMap**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例切换活动窗口的文档结构图的显示。*/
ActiveDocument.ActiveWindow.DocumentMap = !ActiveDocument.ActiveWindow.DocumentMap
```

``` JavaScript
/*本示例显示 Sales.doc 窗口的文档结构图。*/
function test()
{
let docSales = Documents.Open("C:\\Documents\\Sales.doc")
docSales.ActiveWindow.DocumentMap = true
}
```

### Window.EnvelopeVisible

如果电子邮件的页眉在文档窗口是可见的，则该属性值为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.EnvelopeVisible**

\*express是一个代表 Window 对象的变量。

**说明**

如果文档不是电子邮件，则该属性将不起作用。

**示例**

``` JavaScript
/*本示例将显示电子邮件的页眉。*/
ActiveWindow.EnvelopeVisible = true
```

### Window.Height

返回或设置窗口的高度。Long 类型，可读写。

**语法**

**express.Height**

\*express是一个代表 Window 对象的变量。

**说明**

当窗口处于最小化或最大化时，不能设置该属性。使用 **Application** 对象的 **UsableHeight** 属性可确定窗口的最大尺寸。使用 **WindowState** 属性可确定窗口状态。

**示例**

``` JavaScript
/*本示例对活动窗口的高度加以调整，以铺满应用程序的窗口区域。*/
function test()
{
ActiveDocument.ActiveWindow.WindowState = wdWindowStateNormal
ActiveDocument.ActiveWindow.Height = Application.UsableHeight
}
```

### Window.HorizontalPercentScrolled

返回或设置水平滚动条位置（以文档宽度的百分比表示）。可读写 **Long** 类型。

**语法**

**express.HorizontalPercentScrolled**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动窗口水平滚动的百分比。*/
MsgBox(ActiveDocument.ActiveWindow.HorizontalPercentScrolled + "%")
```

### Window.IMEMode

返回或设置日文输入法编辑器 (IME) 的默认启动模式。 **WdIMEMode** 类型，可读写。

**语法**

**express.IMEMode**

\*express是一个代表 Window 对象的变量。

### Window.Index

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Window 对象的变量。

### Window.Left

返回或设置一个 **Long** 类型的值，该值代表指定窗口的水平位置，以磅为单位。可读写。

**语法**

**express.Left**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动窗口的水平位置设置为 100 磅。*/
function test()
{
let win = ActiveDocument.ActiveWindow
win.WindowState = wdWindowStateNormal
win.Left = 100
win.Top = 0
}
```

### Window.Next

返回打开的文档窗口集合中的下一个文档窗口。只读。

**语法**

**express.Next**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*以下示例激活下一个窗口。*/
function test()
{
if(Windows.Count > 1){
    ActiveDocument.ActiveWindow.Next.Activate()
}
}
```

### Window.Panes

返回一个 **Panes** 集合，该集合代表指定窗口中所有的窗格。

**语法**

**express.Panes**

\*express是一个代表 Window 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例将活动窗口拆分为两个窗格。*/
function test()
{
if(ActiveDocument.ActiveWindow.Panes.Count == 1) {
    ActiveDocument.ActiveWindow.Panes.Add()
}
}
```

``` JavaScript
/*本示例激活 Document2 窗口中的第一个窗格。*/
Windows.Item("Document2").Panes.Item(1).Activate()
```

### Window.Parent

返回一个 **Object** 类型值，该值代表指定 **Window** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Window 对象的变量。

### Window.Previous

返回打开的文档窗口集合中的上一个文档窗口。只读。

**语法**

**express.Previous**

\*express是一个代表 Window 对象的变量。

### Window.Selection

返回一个 **Selection** 对象，该对象代表一个选定范围或插入点。只读。

**语法**

**express.Selection**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*以下示例将第一个窗口的所选内容复制到下一个窗口。*/
function test()
{
if(Windows.Count >= 2) {
    Windows.Item(1).Selection.Copy()
    Windows.Item(1).Next.Activate()
    Selection.Paste()
}
}
```

### Window.ShowSourceDocuments

返回或设置一个 **WdShowSourceDocuments** 常量，该常量代表 WPS 在比较和合并过程之后显示源文档的方式。可读写。

**语法**

**express.ShowSourceDocuments**

\*express是一个代表 Window 对象的变量。

### Window.Split

如果将窗口拆分为多个窗格，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Split**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口拆分为大小相等的两个窗格。*/
ActiveDocument.ActiveWindow.Split = true
```

``` JavaScript
/*如果拆分 Document1 窗口，则本示例关闭活动窗格。*/
function test()
{
if(Windows.Item("Document1").Split == true) {
    Windows.Item("Document1").ActivePane.Close()
}
}
```

### Window.SplitVertical

返回或设置指定窗口的垂直拆分百分比。 **Long** 类型，可读写。

**语法**

**express.SplitVertical**

\*express是一个代表 Window 对象的变量。

**说明**

若要取消拆分，可将该属性设置为零 (0) 或将 **Split** 属性设置为 **False** 。

**示例**

``` JavaScript
/*本示例拆分活动窗口，使顶部窗格占窗口的 70%。*/
ActiveDocument.ActiveWindow.SplitVertical = 70
```

``` JavaScript
/*本示例垂直均分 Document1 窗口。*/
Windows.Item("Document1").SplitVertical = 50
```

### Window.StyleAreaWidth

该属性返回或设置样式区的宽度，以磅为单位。 **Single** 类型，可读写。

**语法**

**express.StyleAreaWidth**

\*express是一个代表 Window 对象的变量。

**说明**

如果 **StyleAreaWidth** 属性值大于 0（零），则样式名显示在文本的左侧。在页面视图或 Web 版式视图中不显示样式区。

**示例**

``` JavaScript
/*本示例将活动窗口切换至普通视图，并设置样式区宽度为 1 英寸。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdNormalView
ActiveDocument.ActiveWindow.StyleAreaWidth = InchesToPoints(1)
}
```

### Window.Thumbnails

设置或返回一个 **Boolean** 类型的值，代表是否沿 WPS 文档窗口的左侧显示文档页面的缩略图。

**语法**

**express.Thumbnails**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*下列示例显示活动文档中页面的缩略图。*/
ActiveDocument.ActiveWindow.Thumbnails = true
```

### Window.Top

返回或设置指定文档窗口的垂直位置（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Top**

\*express是一个代表 Window 对象的变量。

### Window.Type

返回窗口的类型。只读 **WdWindowType** 类型。

**语法**

**express.Type**

\*express是一个代表 Window 对象的变量。

### Window.UsableHeight

返回指定的文档窗口中活动工作区的高度（以磅为单位）。 **Long** 类型，只读。

**语法**

**express.UsableHeight**

\*express是一个代表 Window 对象的变量。

**说明**

如果在文档窗口中看不见任何工作区，则 **UsableHeight** 返回 1。实际可用的高度可用 **UsableHeight** 值减去 1 来确定。

**示例**

``` JavaScript
/*以下示例显示活动文档窗口中工作区的大小。*/
function test()
{
let win = ActiveDocument.ActiveWindow

MsgBox("Working area height = " + win.UsableHeight + "\n" + 
    "Working area width = " + win.UsableWidth)
}
```

### Window.UsableWidth

返回指定的文档窗口中活动工作区的宽度（以磅为单位）。 **Long** 类型，只读。

**语法**

**express.UsableWidth**

\*express是一个代表 Window 对象的变量。

**说明**

如果在文档窗口中看不见任何工作区，则 **UsableWidth** 返回 1。要确定实际可用的高度，可用 **UsableWidth** 值减去 1。

**示例**

``` JavaScript
/*以下示例显示活动文档窗口中工作区的大小。*/
function test()
{
let win = ActiveDocument.ActiveWindow

MsgBox("Working area height = " + win.UsableHeight + "\n" + 
    "Working area width = " + win.UsableWidth)
}
```

### Window.VerticalPercentScrolled

返回或设置垂直滚动条的位置（以文档长度的百分比表示）。可读写 **Long** 类型。

**语法**

**express.VerticalPercentScrolled**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动窗口垂直滚动的百分比。*/
MsgBox(ActiveDocument.ActiveWindow.VerticalPercentScrolled + "%")
```

``` JavaScript
/*以下示例将活动窗口垂直滚动 10%。*/
function test()
{
let aWindow = ActiveDocument.ActiveWindow
aWindow.VerticalPercentScrolled = aWindow.VerticalPercentScrolled + 10
}
```

### Window.View

返回一个 **View** 对象，该对象代表指定窗口或窗格的视图。

**语法**

**express.View**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动窗口切换为全屏显示。*/
ActiveDocument.ActiveWindow.View.FullScreen = true
```

``` JavaScript
/*以下示例设置 Windows 集合中所有窗口的视图选项。*/
function test()
{
let myWindow

for(let i=1; i<=Windows.Count; i++) {
    myWindow = Windows.Item(i)
    myWindow.View.ShowTabs = true
    myWindow.View.ShowParagraphs = true
    myWindow.View.Type = wdNormalView
}
}
```

### Window.Visible

如果指定对象可见，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Visible**

\*express是一个代表 Window 对象的变量。

**说明**

如果 **Visible** 属性为 **False** ，则某些方法和属性可能无法使用。

### Window.Width

返回或设置指定文档窗口的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

\*express是一个代表 Window 对象的变量。

### Window.WindowNumber

该属性返回在指定窗口中显示的文档窗口序号。例如，如果窗口标题为“Sales.doc:2”，则该属性返回数值 2。 **Long** 类型，只读。

**语法**

**express.WindowNumber**

\*express是一个代表 Window 对象的变量。

**说明**

使用 属性可返回 **Windows** 集合中指定窗口的序号。

**示例**

``` JavaScript
/*本示例检索活动窗口的窗口序号，新建一个窗口，然后激活前一个窗口。*/
function WinNum(){
    let lwindowNum

    lwindowNum = ActiveDocument.ActiveWindow.WindowNumber
    NewWindow()
    ActiveDocument.Windows.Item(lwindowNum).Activate()
}
```

### Window.WindowState

返回或设置指定文档窗口或任务窗口的状态。 **WdWindowState** 类型，可读写。

**语法**

**express.WindowState**

\*express是一个代表 Window 对象的变量。

**说明**

**wdWindowStateNormal** 常量表示没有最大化或最小化的窗口。无法设置不活动窗口的状态。在设置窗口状态之前可用 **Activate** 方法激活窗口。

**示例**

``` JavaScript
/*以下示例将没有最大化或最小化的活动窗口最大化。*/
function test()
{
if(ActiveDocument.ActiveWindow.WindowState == wdWindowStateNormal) {
    ActiveDocument.ActiveWindow.WindowState = wdWindowStateMaximize
}
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Window](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# NamedSlideShows 对象

## 简介

演示文稿中所有 NamedSlideShow 对象的集合。每个 NamedSlideShow 对象代表一个自定义幻灯片放映。

## 方法

| 序号 | 名称                          | 说明                                                                                                                             |
|:----:|:------------------------------|:---------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Add](#NamedSlideShows.Add)   | 创建一个新的命名幻灯片放映，并将其添加到指定演示文稿的命名幻灯片放映集合中。返回一个代表新命名幻灯片放映的 NamedSlideShow 对象。 |
|  2   | [Item](#NamedSlideShows.Item) | 从指定集合中返回单个对象。                                                                                                       |

## 属性

| 序号 | 名称                                        | 说明                                                    |
|:----:|:--------------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#NamedSlideShows.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#NamedSlideShows.Count)             | 返回指定集合中的对象数目。                              |
|  3   | [Parent](#NamedSlideShows.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### NamedSlideShows.Add

创建一个新的命名幻灯片放映，并将其添加到指定演示文稿的命名幻灯片放映集合中。返回一个代表新命名幻灯片放映的 **NamedSlideShow** 对象。

**语法**

**express.Add ( *Name* , *safeArrayOfSlideIDs* )**

\*express是一个代表 NamedSlideShows 对象的变量。

**参数**

| 名称                  | 必选/可选 | 数据类型 | 说明                                              |
|-----------------------|-----------|----------|---------------------------------------------------|
| *Name*                | 必选      | String   | 幻灯片放映的名称。                                |
| *safeArrayOfSlideIDs* | 必选      | Variant  | 包含将在幻灯片放映中显示的幻灯片的唯一幻灯片 ID。 |

**返回值**

NamedSlideShow

**说明**

添加命名幻灯片放映时指定的名称就是用作 **Run** 方法的参数的名称，以便运行命名幻灯片放映。

### NamedSlideShows.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 NamedSlideShows 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

NamedSlideShow

## 成员属性

### NamedSlideShows.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 NamedSlideShows 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres){
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### NamedSlideShows.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 NamedSlideShows 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口。*/
function test(){
　　　　let win = Application.Windows
   　　　　 for(let i = 2; i <= win.Count; i++){
     　　　　   win.Item(i).Close()
  　　　　  }
}
```

### NamedSlideShows.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 NamedSlideShows 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let sha = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
　　　　sha.TextRange.Text = "Test text"
　　　　sha.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/NamedSlideShows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Presentation 对象

## 简介

代表一个WPP 演示文稿。

## 说明

Presentation 对象是 Presentations 集合的成员。 Presentations 集合包含代表WPP 中打开的演示文稿的所有 Presentation 对象。

以下示例说明如何执行下列操作：

- 返回通过名称或索引号指定的演示文稿
- 返回活动窗口中的演示文稿
- 返回指定的任一文档窗口或幻灯片放映窗口中的演示文稿

## 示例：

使用 Presentations ( *index* ) 返回单个 Presentation 对象，其中 *index* 是演示文稿的名称或索引号。演示文稿的名称就是其文件名，是否带文件扩展名均可，但不包含路径。以下示例在“Sample Presentation”的开头添加一张幻灯片。

``` JavaScript
Application.Presentations.Item("Sample Presentation").Slides.Add(1, 1)
```

请注意，如果同时打开了多个具有相同名称的演示文稿，则会返回集合中具有指定名称的第一个演示文稿。

使用 ActivePresentation 属性可返回活动窗口中的演示文稿。以下示例将保存活动演示文稿。

``` JavaScript
Application.ActivePresentation.Save()
```

使用 Presentation 属性可返回指定文档窗口或幻灯片放映窗口中的演示文稿。以下示例将显示运行于第一个幻灯片放映窗口中的幻灯片放映的名称。

``` JavaScript
alert(Application.SlideShowWindows.Item(1).Presentation.Name)
```

## 方法

| 序号 | 名称                                                                       | 说明                                                                                                                                                                      |
|:----:|:---------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddTitleMaster](#Presentation.AddTitleMaster)                             | 为指定的演示文稿添加一个标题母版。返回一个代表该标题母版的 Master 对象。如果演示文稿已有标题母版，则会出错。                                                              |
|  2   | [AddToFavorites](#Presentation.AddToFavorites)                             | 在 Program Files 文件夹的 Favorites 文件夹中，添加一个快捷方式，以代表指定演示文稿的当前选择内容（对于 Presentation 对象）或指定超链接的目标文档（对于 Hyperlink 对象）。 |
|  3   | [ApplyTemplate](#Presentation.ApplyTemplate)                               | 将设计模板应用于指定的演示文稿。                                                                                                                                          |
|  4   | [ApplyTheme](#Presentation.ApplyTheme)                                     | 将主题或设计模板应用于指定的演示文稿。                                                                                                                                    |
|  5   | [CheckIn](#Presentation.CheckIn)                                           | 从本地计算机将演示文稿返回到服务器，并且设置本地文件为只读，以便使得其不能在本地编辑。                                                                                    |
|  6   | [CheckIn12](#Presentation.CheckIn12)                                       | 从本地计算机将演示文稿返回到服务器，并且设置本地文件为只读，以便使得其不能在本地编辑。                                                                                    |
|  7   | [Close](#Presentation.Close)                                               | 关闭指定的文档窗口、演示文稿或打开的任意多边形                                                                                                                            |
|  8   | [Convert](#Presentation.Convert)                                           | 将图示转换为另一种类型。                                                                                                                                                  |
|  9   | [Export](#Presentation.Export)                                             | 使用指定的图形筛选器导出演示文稿中的每张幻灯片，并将导出的文件保存在指定的文件夹中。                                                                                      |
|  10  | [ExportAsFixedFormat](#Presentation.ExportAsFixedFormat)                   | 将 演示文稿的副本发布为固定格式的文件（PDF 或 XPS）。                                                                                                                     |
|  11  | [FollowHyperlink](#Presentation.FollowHyperlink)                           | 如果缓存文档已被下载，本方法将其显示出来。否则，该方法将识别超链接，下载目标文档并将其显示在适当的应用程序中。                                                            |
|  12  | [GetWorkflowTasks](#Presentation.GetWorkflowTasks)                         | 返回 WorkflowTasks 集合。 可读写。                                                                                                                                        |
|  13  | [GetWorkflowTemplates](#Presentation.GetWorkflowTemplates)                 | 返回 WorkflowTemplates 集合。 可读写。                                                                                                                                    |
|  14  | [LockServerFile](#Presentation.LockServerFile)                             | 锁定服务器上的演示文稿，以阻止对其进行的修改。                                                                                                                            |
|  15  | [NewWindow](#Presentation.NewWindow)                                       | 打开一个包含指定演示文稿的新窗口。返回一个代表新窗口的 DocumentWindow 对象。                                                                                              |
|  16  | [PrintOut](#Presentation.PrintOut)                                         | 打印指定演示文稿。                                                                                                                                                        |
|  17  | [PublishSlides](#Presentation.PublishSlides)                               | 通过加载的任意演示文稿创建包含幻灯片的 Web 演示文稿（HTML 格式）。可以在 Web 浏览器中查看发布的演示文稿。                                                                 |
|  18  | [ReloadAs](#Presentation.ReloadAs)                                         | 基于指定的 HTML 文档编码重新加载演示文稿。                                                                                                                                |
|  19  | [RemoveDocumentInformation](#Presentation.RemoveDocumentInformation)       | 删除 演示文稿中的文档信息，如个人信息、批注和文档属性。                                                                                                                   |
|  20  | [Save](#Presentation.Save)                                                 | 保存指定的演示文稿。                                                                                                                                                      |
|  21  | [SaveAs](#Presentation.SaveAs)                                             | 保存未保存过的演示文稿，或以其他文件名保存已保存过的演示文稿。                                                                                                            |
|  22  | [SaveCopyAs](#Presentation.SaveCopyAs)                                     | 在不修改原文件的情况下将指定演示文稿的副本保存至另一文件。                                                                                                                |
|  23  | [SendFaxOverInternet](#Presentation.SendFaxOverInternet)                   | 向指定的收件人以传真形式发送演示文稿。                                                                                                                                    |
|  24  | [SetPasswordEncryptionOptions](#Presentation.SetPasswordEncryptionOptions) | 设置 WPP 通过密码加密演示文稿时使用的选项。                                                                                                                               |
|  25  | [UpdateLinks](#Presentation.UpdateLinks)                                   | 更新指定演示文稿中链接的 OLE 对象。                                                                                                                                       |
|  26  | [WebPagePreview](#Presentation.WebPagePreview)                             | 在当前 Web 浏览器中预览演示文稿。                                                                                                                                         |

## 属性

| 序号 | 名称                                                                               | 说明                                                                                                                                                                                                        |
|:----:|:-----------------------------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Presentation.Application)                                           | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                     |
|  2   | [BuiltInDocumentProperties](#Presentation.BuiltInDocumentProperties)               | 返回一个 DocumentProperties 集合，该集合代表指定演示文稿的所有内置文档属性。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。                                                                   |
|  3   | [ColorSchemes](#Presentation.ColorSchemes)                                         | 返回一个 ColorSchemes 集合，该集合代表指定演示文稿中的配色方案。只读。                                                                                                                                      |
|  4   | [CommandBars](#Presentation.CommandBars)                                           | 返回一个 CommandBars 集合，该集合代表主机容器应用程序和 WPP 中合并的命令栏集。仅当容器是 DocObject 服务器（如 活页夹）并且 WPP 作为 OLE 服务器时，该属性返回有效对象。只读。                                |
|  5   | [Container](#Presentation.Container)                                               | 返回包含指定嵌入演示文稿的对象。只读 Object 类型。如果容器不支持 OLE 自动化，或指定的演示文稿不是嵌入在“活页夹”文件中，则该属性无效。                                                                       |
|  6   | [ContentTypeProperties](#Presentation.ContentTypeProperties)                       | 返回 MetaProperties 集合，该集合描述演示文稿中存储的元数据。只读。                                                                                                                                          |
|  7   | [CustomDocumentProperties](#Presentation.CustomDocumentProperties)                 | 返回一个 DocumentProperties 集合，该集合代表指定演示文稿的所有自定义文档属性。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。                                                                 |
|  8   | [CustomerData](#Presentation.CustomerData)                                         | 返回一个 CustomerData 对象。只读。                                                                                                                                                                          |
|  9   | [CustomXMLParts](#Presentation.CustomXMLParts)                                     | 返回一个 CustomXMLParts 对象。                                                                                                                                                                              |
|  10  | [DefaultLanguageID](#Presentation.DefaultLanguageID)                               | 返回或设置演示文稿的默认语言。在您设置一个演示文稿的 DefaultLanguageID 属性时，同时也为所有后续新演示文稿设置了该属性。可读/写 MsoLanguageID 类型。                                                         |
|  11  | [DefaultShape](#Presentation.DefaultShape)                                         | 返回一个 Shape 对象，该对象代表演示文稿的默认形状。只读。                                                                                                                                                   |
|  12  | [Designs](#Presentation.Designs)                                                   | 返回一个 Designs 对象，该对象代表一个设计集合。                                                                                                                                                             |
|  13  | [DisplayComments](#Presentation.DisplayComments)                                   | 决定是否在指定的演示文稿中显示批注。可读/写 MsoTriState 类型。                                                                                                                                              |
|  14  | [DocumentInspectors](#Presentation.DocumentInspectors)                             | 返回 DocumentInspectors 集合。只读。                                                                                                                                                                        |
|  15  | [DocumentLibraryVersions](#Presentation.DocumentLibraryVersions)                   | 返回一个 DocumentLibraryVersions 集合，该集合代表共享演示文稿的版本集合，该共享演示文稿启用了版本控制功能，并存储在服务器上的文档库中。                                                                     |
|  16  | [EnvelopeVisible](#Presentation.EnvelopeVisible)                                   | 决定电子邮件的邮件头在文档窗口中是否可见。可读/写 MsoTriState 类型。                                                                                                                                        |
|  17  | [ExtraColors](#Presentation.ExtraColors)                                           | 返回一个 ExtraColors 对象，该对象代表指定演示文稿中可用的其他颜色。只读。                                                                                                                                   |
|  18  | [FarEastLineBreakLanguage](#Presentation.FarEastLineBreakLanguage)                 | 返回或设置启用了换行控制选项时用于确定所用换行级别的语言。可读/写 MsoFarEastLineBreakLanguageID 类型。                                                                                                      |
|  19  | [FarEastLineBreakLevel](#Presentation.FarEastLineBreakLevel)                       | 返回或设置基于亚洲字符级别的换行符。可读/写 Long 类型。可读/写 PpFarEastLineBreakLevel 类型。                                                                                                               |
|  20  | [Final](#Presentation.Final)                                                       | 确定是否将演示文稿标记为最终版本（只读）。可读/写。                                                                                                                                                         |
|  21  | [Fonts](#Presentation.Fonts)                                                       | 返回一个 Fonts 集合，该集合代表用于指定演示文稿的所有字体。只读。                                                                                                                                           |
|  22  | [FullName](#Presentation.FullName)                                                 | 返回指定的加载宏或保存的演示文稿的名称，包括路径、当前文件系统分隔符和文件扩展名。只读。 String 类型。                                                                                                      |
|  23  | [GridDistance](#Presentation.GridDistance)                                         | 设置或返回一个 Single 类型的值，该值代表网格线之间的距离。可读/写。                                                                                                                                         |
|  24  | [HandoutMaster](#Presentation.HandoutMaster)                                       | 返回一个代表讲义母版的 Master 对象。只读。                                                                                                                                                                  |
|  25  | [HasTitleMaster](#Presentation.HasTitleMaster)                                     | 如果指定的演示文稿具有标题母版，则为 MsoTrue 。只读。 MsoTriState 类型。                                                                                                                                    |
|  26  | [HasVBProject](#Presentation.HasVBProject)                                         | 返回一个值，该值指示当前演示文稿是否包含 示例代码 (VBA) 项目。只读                                                                                                                                          |
|  27  | [LayoutDirection](#Presentation.LayoutDirection)                                   | 为用户界面返回或设置版式方向。可为以下值之一，默认值取决于已选择或安装的语言支持。可读/写 PpDirection 类型。                                                                                                |
|  28  | [Name](#Presentation.Name)                                                         | 演示文稿的名称包括文件扩展名（对于已注册的文件类型），但不包括其路径。不能使用该属性来设置名称。如果需要更改名称，请使用 [SaveAs](#Presentation.SaveAs) 方法将演示文稿另存为其他名称。 String 类型，只读 。 |
|  29  | [NoLineBreakAfter](#Presentation.NoLineBreakAfter)                                 | 返回或设置前置字符。可读/写 String 类型。                                                                                                                                                                   |
|  30  | [NoLineBreakBefore](#Presentation.NoLineBreakBefore)                               | 返回或设置后置标点字符。可读/写 String 类型。                                                                                                                                                               |
|  31  | [NotesMaster](#Presentation.NotesMaster)                                           | 返回一个代表备注母版的 Master 对象。只读 。                                                                                                                                                                 |
|  32  | [PageSetup](#Presentation.PageSetup)                                               | 返回一个 PageSetup 对象，该对象的属性控制指定演示文稿的幻灯片的页面设置属性。只读。                                                                                                                         |
|  33  | [Parent](#Presentation.Parent)                                                     | 返回指定对象的父对象。                                                                                                                                                                                      |
|  34  | [Password](#Presentation.Password)                                                 | 返回或设置一个 String 类型的值，该值代表要打开指定演示文稿必须提供的密码。可读/写。                                                                                                                         |
|  35  | [PasswordEncryptionAlgorithm](#Presentation.PasswordEncryptionAlgorithm)           | 返回一个 String 类型的值，该值代表 WPP 通过密码加密文档所使用的算法。只读。                                                                                                                                 |
|  36  | [PasswordEncryptionFileProperties](#Presentation.PasswordEncryptionFileProperties) | 如果 WPP 对使用密码保护的文档的文件属性进行加密，则返回 MsoTrue 属性值。只读 MsoTriState 类型。                                                                                                             |
|  37  | [PasswordEncryptionKeyLength](#Presentation.PasswordEncryptionKeyLength)           | 返回一个 Long 类型的值，该值表示 WPP 通过密码加密文档时所使用算法的密钥长度。只读。                                                                                                                         |
|  38  | [PasswordEncryptionProvider](#Presentation.PasswordEncryptionProvider)             | 返回一个 String 类型的值，该值指定 WPP 通过密码加密文档时所使用的加密算法提供商的名称。只读。                                                                                                               |
|  39  | [Path](#Presentation.Path)                                                         | 返回 String ，该值代表到指定的 Presentation 对象的路径， 只读。                                                                                                                                             |
|  40  | [Permission](#Presentation.Permission)                                             | 返回一个 Permission 对象，该对象可用于限制对活动演示文稿的权限，还可用于返回或设置特定权限设置。只读。                                                                                                      |
|  41  | [PrintOptions](#Presentation.PrintOptions)                                         | 返回一个 PrintOptions 对象，该对象代表与指定演示文稿共同保存的打印选项。只读。                                                                                                                              |
|  42  | [PublishObjects](#Presentation.PublishObjects)                                     | 返回一个 PublishObjects 集合，该集合代表已完全或部分加载的演示文稿的集合，这些演示文稿可作为 HTML 进行发布。只读。                                                                                          |
|  43  | [ReadOnly](#Presentation.ReadOnly)                                                 | 返回指定的演示文稿是否为只读。只读 MsoTriState 类型。                                                                                                                                                       |
|  44  | [RemovePersonalInformation](#Presentation.RemovePersonalInformation)               | 如果该属性值为 MsoTrue ，则 WPP 在保存演示文稿时从批注和修订中删除所有用户信息。 MsoTriState 类型，可读/写。                                                                                                |
|  45  | [Research](#Presentation.Research)                                                 | 返回一个 Research 对象，利用该对象可访问信息检索服务功能。只读。                                                                                                                                            |
|  46  | [Saved](#Presentation.Saved)                                                       | 决定演示文稿自上次保存以来是否进行过更改。可读/写 MsoTriState 类型。                                                                                                                                        |
|  47  | [ServerPolicy](#Presentation.ServerPolicy)                                         | 返回一个ServerPolicy 对象。只读。                                                                                                                                                                           |
|  48  | [SharedWorkspace](#Presentation.SharedWorkspace)                                   | 返回一个 SharedWorkspace 对象，该对象代表指定演示文稿所在的文档工作区。只读。                                                                                                                               |
|  49  | [Signatures](#Presentation.Signatures)                                             | 返回一个 SignatureSet 对象，该对象代表数字签名的集合。                                                                                                                                                      |
|  50  | [SlideMaster](#Presentation.SlideMaster)                                           | 返回一个 Master 对象，该对象代表幻灯片母版。                                                                                                                                                                |
|  51  | [Slides](#Presentation.Slides)                                                     | 返回一个 Slides 集合，该集合代表指定演示文稿中的所有幻灯片。只读。                                                                                                                                          |
|  52  | [SlideShowSettings](#Presentation.SlideShowSettings)                               | 返回一个 SlideShowSettings 对象，该对象代表指定演示文稿的幻灯片放映设置。只读。                                                                                                                             |
|  53  | [SlideShowWindow](#Presentation.SlideShowWindow)                                   | 返回一个 SlideShowWindow 对象，该对象代表正在运行的指定演示文稿所在的幻灯片放映窗口。只读。                                                                                                                 |
|  54  | [SnapToGrid](#Presentation.SnapToGrid)                                             | 该属性值为 MsoTrue 时，在指定的演示文稿中将形状与网格线对齐。可读/写 MsoTriState 类型。                                                                                                                     |
|  55  | [Sync](#Presentation.Sync)                                                         | 返回一个 Sync 对象，使用该对象可对存储在 Windows SharePoint Services 共享工作区中的共享演示文稿的本地副本和服务器副本的同步进行管理。只读。                                                                 |
|  56  | [Tags](#Presentation.Tags)                                                         | 返回一个代表指定对象的标签的 Tags 对象。只读。                                                                                                                                                              |
|  57  | [TemplateName](#Presentation.TemplateName)                                         | 返回与指定演示文稿相关的设计模板的名称。只读 String 类型。                                                                                                                                                  |
|  58  | [TitleMaster](#Presentation.TitleMaster)                                           | 返回 Master 对象，该对象代表指定演示文稿的标题母版。如果演示文稿没有标题母版，则会产生一个错误。                                                                                                            |
|  59  | [WebOptions](#Presentation.WebOptions)                                             | 返回一个 WebOptions 对象，该对象中包含演示文稿级别的属性，当以网页的方式发布或保存整个或部分演示文稿，或打开某个网页时，WPP 将使用这些属性。只读。                                                          |
|  60  | [Windows](#Presentation.Windows)                                                   | 返回一个 DocumentWindows 集合，该集合代表与指定演示文稿相关联的所有文档窗口。该属性不返回与该演示文稿相关联的任何幻灯片放映窗口。只读。                                                                     |
|  61  | [WritePassword](#Presentation.WritePassword)                                       | 设置或返回一个 String 类型的值，该值代表保存对指定文档所做的更改时使用的密码。可读/写。                                                                                                                     |

## 成员方法

### Presentation.AddTitleMaster

为指定的演示文稿添加一个标题母版。返回一个代表该标题母版的 **Master** 对象。如果演示文稿已有标题母版，则会出错。

**语法**

**express.AddTitleMaster ()**

\*express是一个代表 Presentation 对象的变量。

**返回值**

Master

**示例**

``` JavaScript
//以下示例中，如果当前演示文稿没有标题母版，则为其添加一个标题母版。
function test() {
    if (Application.ActivePresentation.HasTitleMaster == false) {
        Application.ActivePresentation.AddTitleMaster()
    }
}
```

### Presentation.AddToFavorites

在 Program Files 文件夹的 Favorites 文件夹中，添加一个快捷方式，以代表指定演示文稿的当前选择内容（对于 **Presentation** 对象）或指定超链接的目标文档（对于 **Hyperlink** 对象）。

**语法**

**express.AddToFavorites ()**

\*express是一个代表 Presentation 对象的变量。

**说明**

如果文档具有显示名称，则快捷方式的名称是该显示名称；否则，快捷方式名称由 HLINK.DLL 决定。

**示例**

``` JavaScript
//本示例在 Windows 程序文件夹的 Favorites 中添加一个指向当前演示文稿的超链接。
Application.ActivePresentation.AddToFavorites()
```

### Presentation.ApplyTemplate

将设计模板应用于指定的演示文稿。

**语法**

**express.ApplyTemplate ( *FileName* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                 |
|------------|-----------|----------|----------------------|
| *FileName* | 必选      | String   | 指定设计模板的名称。 |

**说明**

如果在字符串中引用未安装的演示文稿设计模板，则将产生运行时错误。无论 **Application.** **FeatureInstall** 属性设置是什么，该模板都不自动安装。要对当前未安装的模板使用 **ApplyTemplate** 方法，必须先安装附加设计模板。为此，请通过运行 WPS Office安装程序（单击 Windows 控制面板中的 **“添加/删除程序”** 图标）安装WPP 的附加设计模板。

**示例**

``` JavaScript
/*本示例将“Professional”设计模板应用于当前演示文稿。*/
Application.ActivePresentation.ApplyTemplate("c:\\program files\\office\\templates" + "\\presentation designs\\professional.pot")
```

### Presentation.ApplyTheme

将主题或设计模板应用于指定的演示文稿。

**语法**

**express.ApplyTheme ( *themeName* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                            |
|-------------|-----------|----------|---------------------------------------------------------------------------------|
| *themeName* | 必选      | String   | 应用于 Presentation 对象的主题文件 (.thmx) 或设计模板文件 (.pot) 的路径和名称。 |

**示例**

``` JavaScript
/*以下示例将保存的主题应用于当前演示文稿：*/
Application.ActivePresentation.ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.thmx")
```

### Presentation.CheckIn

从本地计算机将演示文稿返回到服务器，并且设置本地文件为只读，以便使得其不能在本地编辑。

**语法**

**express.CheckIn ( *SaveChanges* , *Comments* , *MakePublic* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                                                                                             |
|---------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *SaveChanges* | 可选      | Boolean  | 如果为 True，则将演示文稿保存到服务器位置。默认值为 False。                                                                                                                      |
| *Comments*    | 可选      | Boolean  | 正被签入的演示文稿的修订批注（仅在 SaveChanges 等于 True 时应用）。                                                                                                              |
| *MakePublic*  | 可选      | Variant  | 如果为 True，则允许用户在签入之后发布演示文稿。这样会提交文档使之进入审批流程，最终可将一种演示文稿版本发布给对演示文稿具有只读权限的用户（仅在 SaveChanges 等于 True 时应用）。 |

**说明**

若要利用 WPP 中的协作功能，演示文稿必须保存在 SharePoint Portal Server 上。

**示例**

本示例检查服务器是否可以签入指定的演示文稿；如果可以，它将关闭此演示文稿并且将其签回到服务器中。

``` JavaScript
function CheckInPresentation(strPresentation){
    if(Application.Presentations.Item(strPresentation).CanCheckIn == true ){
        Application.Presentations.Item(strPresentation).CheckIn()
        alert(strPresentation + " has been checked in.")
    }
    else{
        alert(strPresentation + " cannot be checked in " + "at this time.  Please try again later.")
    }
}
```

若要调用上述子程序，请使用下列子程序并用在上面备注部分中提到的服务器上的实际文件替换“http://servername/workspace/report.ppt”文件名。

``` JavaScript
function CheckInPresentation(){
    CheckInPresentation("http://servername/workspace/report.ppt")
}
```

### Presentation.CheckIn12

从本地计算机将演示文稿返回到服务器，并且设置本地文件为只读，以便使得其不能在本地编辑。

**语法**

**express.CheckIn12 ( *SaveChanges* , *Comments* , *MakePublic* , *VersionType* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                                                                                             |
|---------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *SaveChanges* | 可选      | Boolean  | 如果为 True，则将演示文稿保存到服务器位置。默认值为 False。                                                                                                                      |
| *Comments*    | 可选      | Variant  | 正被签入的演示文稿的修订批注（仅在 SaveChanges 等于 True 时应用）。                                                                                                              |
| *MakePublic*  | 可选      | Variant  | 如果为 True，则允许用户在签入之后发布演示文稿。这样会提交文档使之进入审批流程，最终可将一种演示文稿版本发布给对演示文稿具有只读权限的用户（仅在 SaveChanges 等于 True 时应用）。 |
| *VersionType* | 可选      | Variant  | 演示文稿的版本号。                                                                                                                                                               |

**说明**

要利用 WPP 中内置的协作功能，演示文稿必须存储在 SharePoint Portal Server 上。

**示例**

本示例检查服务器是否可以签入指定的演示文稿；如果可以，它将关闭此演示文稿并且将其签回到服务器中。

``` JavaScript
function CheckInPresentation(strPresentation){
    if(Application.Presentations.Item(strPresentation).CanCheckIn == true ){
        Application.Presentations.Item(strPresentation).CheckIn()
        alert(strPresentation + " has been checked in.")
    }
    else{
        alert(strPresentation + " cannot be checked in " + "at this time.  Please try again later.")
    }
}
```

若要调用上述子程序，请使用下列子程序并用在上面备注部分中提到的服务器上的实际文件替换“http://servername/workspace/report.ppt”文件名。

``` JavaScript
function CheckInPresentation(){
    CheckInPresentation("http://servername/workspace/report.ppt")
}
```

### Presentation.Close

关闭指定的文档窗口、演示文稿或打开的任意多边形

**语法**

**express.Close ()**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用此方法，WPP关闭打开的演示文档，并且不提示用户保存所做的工作。为防止工作丢失，您应在使用 **Close** 方法之前使用 **Save** 或 **SaveAs** 方法。

**示例**

``` JavaScript
/*本示例在不保存更改的情况下关闭 Pres1.ppt。*/
function test()
{
    let myprtion = Application.Presentations.Item("pres1.ppt")
    myprtion.Saved = true
    myprtion.Close()
}

/*本示例关闭所有打开的演示文稿。*/
function test()
{
    for(let i = Application.Presentations.Count; i >= 1; i--)
    {
        Application.Presentations.Item(i).Close()
    }
}
```

### Presentation.Convert

将图示转换为另一种类型。

**语法**

**express.Convert ( *Type* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型       | 说明                 |
|--------|-----------|----------------|----------------------|
| *Type* | 必选      | MsoDiagramType | 要转换为的图示类型。 |

**说明**

如果目标图示的 **Type** 属性是组织结构图 ( **msoDiagramTypeOrgChart** )，此方法将产生错误。

|                                                             |
|-------------------------------------------------------------|
| **MsoDiagramType** 可以是下列 **MsoDiagramType** 常数之一。 |
| **msoDiagramCycle**                                         |
| **msoDiagramMixed**                                         |
| **msoDiagramOrgChart**                                      |
| **msoDiagramPyramid**                                       |
| **msoDiagramRadial**                                        |
| **msoDiagramTarget**                                        |
| **msoDiagramVenn**                                          |

### Presentation.Export

使用指定的图形筛选器导出演示文稿中的每张幻灯片，并将导出的文件保存在指定的文件夹中。

**语法**

**express.Export ( *Path* , *FilterName* , *ScaleWidth* , *ScaleHeight* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                        |
|---------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Path*        | 必选      | String   | 要用来保存导出的幻灯片的文件夹的路径。可以包括完整路径；如果不包括完整路径，WPP 就会在当前文件夹中为导出的幻灯片创建一个子文件夹。                                                                                          |
| *FilterName*  | 必选      | String   | 要导出幻灯片的图形格式。指定的图形格式必须在 Windows 注册表中注册导出筛选器。可以指定注册的扩展名或注册的筛选器名称。WPP将首先在注册表中搜索匹配的扩展名。如果没有找到与指定字符串匹配的扩展名，WPP将查找匹配的筛选器名称。 |
| *ScaleWidth*  | 可选      | Number   | 导出幻灯片的宽度（以像素为单位）。                                                                                                                                                                                          |
| *ScaleHeight* | 可选      | Number   | 导出幻灯片的高度（以像素为单位）。                                                                                                                                                                                          |

**说明**

导出一个演示文稿不会将演示文稿的 **[Saved](#Presentation.Saved)** 属性设置为 **True** 。

WPP使用指定的图形筛选器保存演示文稿中每张单独的幻灯片。导出并保存到磁盘的幻灯片的名称由WPP 决定。通常情况下，这些文件保存为 Slide1.wmf、Slide2.wmf 等。保存文件的路径在 **Path** 参数中指定。

**示例**

``` JavaScript
/*本示例将活动演示文稿保存为 WPP 演示文稿，然后将其中的每个幻灯片导出为移植网络图形 (PNG) 文件并保存在 Current Work 文件夹中*/
/*本示例还将每个导出的幻灯片的高度和宽度均设为 100 像素。*/
function test()
{
    let mypretion = Application.ActivePresentation
    mypretion.SaveAs("c:\\Current Work\\Annual Sales", Application.Enum.ppSaveAsPresentation)
    mypretion.Export("c:\\Current Work", "png", 100, 100)
}
```

### Presentation.ExportAsFixedFormat

将 演示文稿的副本发布为固定格式的文件（PDF 或 XPS）。

**语法**

**express.ExportAsFixedFormat ( *Path* , *FixedFormatType* , *Intent* , *FrameSlides* , *HandoutOrder* , *OutputType* , *PrintHiddenSlides* , *PrintRange* , *RangeType* , *SlideShowName* , *IncludeDocProperties* , *KeepIRMSettings* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型            | 说明                                       |
|------------------------|-----------|---------------------|--------------------------------------------|
| *Path*                 | 必选      | String              | 导出的路径。                               |
| *FixedFormatType*      | 必选      | PpFixedFormatType   | 幻灯片应导出为的格式。                     |
| *Intent*               | 可选      | PpFixedFormatIntent | 导出的目的。                               |
| *FrameSlides*          | 可选      | MsoTriState         | 要导出的幻灯片。                           |
| *HandoutOrder*         | 可选      | PpPrintHandoutOrder | 打印讲义的顺序。                           |
| *OutputType*           | 可选      | PpPrintOutputType   | 讲义的类型。                               |
| *PrintHiddenSlides*    | 可选      | MsoTriState         | 打印隐藏的幻灯片。                         |
| *PrintRange*           | 可选      | PrintRange          | 幻灯片范围。                               |
| *RangeType*            | 可选      | PpPrintRangeType    | 幻灯片范围的类型。                         |
| *SlideShowName*        | 可选      | String              | 幻灯片的名称。                             |
| *IncludeDocProperties* | 可选      | Boolean             | 如果还应导出文档属性，则该参数值为 True。  |
| *KeepIRMSettings*      | 可选      | Boolean             | 如果还应导出 IRM 设置，则该参数值为 True。 |

### Presentation.FollowHyperlink

如果缓存文档已被下载，本方法将其显示出来。否则，该方法将识别超链接，下载目标文档并将其显示在适当的应用程序中。

**语法**

**express.FollowHyperlink ( *Address* , *SubAddress* , *NewWindow* , *AddHistory* , *ExtraInfo* , *Method* , *HeaderInfo* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型           | 说明                                                                                                                                                                                                                                |
|--------------|-----------|--------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Address*    | 必选      | String             | 目标文档的地址。                                                                                                                                                                                                                    |
| *SubAddress* | 可选      | String             | 在目标文档中的位置。此参数默认为一个空字符串。                                                                                                                                                                                      |
| *NewWindow*  | 可选      | Boolean            | 值为 True 时，在新窗口中打开目标应用程序。默认值为 False。                                                                                                                                                                          |
| *AddHistory* | 可选      | Boolean            | 值为 True 时，将链接添加到当天的历史记录文件夹中。                                                                                                                                                                                  |
| *ExtraInfo*  | 可选      | String             | 指定 HTTP 信息的字符串或字节数组。例如，该参数可用于指定图像映射的坐标或窗体的内容等，还可用于指定 FAT 文件名。Method 参数将决定如何处理此附加信息。                                                                                |
| *Method*     | 可选      | MsoExtraInfoMethod | 指定发布或附加 ExtraInfo 的方式。                                                                                                                                                                                                   |
| *HeaderInfo* | 可选      | String             | 指定 HTTP 请求的标头信息的字符串。默认值为一个空字符串。可以通过使用下列语法将几个标头行组合为单个字符串："string1" & vbCr & "string2"。所指定的字符串自动转换为 ANSI 字符。请注意，HeaderInfo 参数可能会覆盖默认的 HTTP 标头字段。 |

**说明**

*MsoExtraInfoMethod* 可为以下 **MsoExtraInfoMethod** 常量之一。

| 常量              | 说明                                                    |
|-------------------|---------------------------------------------------------|
| **msoMethodGet**  | 默认值。 *ExtraInfo* 是追加到该地址的一个 **String** 。 |
| **msoMethodPost** | *ExtraInfo* 以 **String** 或字节数组形式发布。          |

**示例**

``` JavaScript
//本示例将 example.microsoft.com 中的文档加载到一个新窗口，并将其添加到“历史记录”文件夹。
Application.ActivePresentation.FollowHyperlink("http://example.microsoft.com",undefined,true,true)
```

### Presentation.GetWorkflowTasks

返回 WorkflowTasks 集合。 可读写。

**语法**

**express.GetWorkflowTasks ()**

\*express是一个代表 Presentation 对象的变量。

**返回值**

WorkFlowTasks

### Presentation.GetWorkflowTemplates

返回 WorkflowTemplates 集合。 可读写。

**语法**

**express.GetWorkflowTemplates ()**

\*express是一个代表 Presentation 对象的变量。

**返回值**

WordFlowTemplates

### Presentation.LockServerFile

锁定服务器上的演示文稿，以阻止对其进行的修改。

**语法**

**express.LockServerFile ()**

\*express是一个代表 Presentation 对象的变量。

### Presentation.NewWindow

打开一个包含指定演示文稿的新窗口。返回一个代表新窗口的 **DocumentWindow** 对象。

**语法**

**express.NewWindow ()**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*本示例创建一个与当前窗口有相同内容的新窗口（同时激活新窗口）并返回到原窗口。*/
function test()
{
    let oldW = Application.ActiveWindow
    let newW = oldW.NewWindow()
    oldW.Activate()
}
```

### Presentation.PrintOut

打印指定演示文稿。

**语法**

**express.PrintOut ( *From* , *To* , *PrintToFile* , *Copies* , *Collate* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型    | 说明                                                                                                                                                         |
|---------------|-----------|-------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *From*        | 可选      | Number      | 要打印的第一页的页码。如果忽略此参数，将从演示文稿的开头开始打印。指定 To 和 From 参数会设置 PrintRanges 对象的内容，并设置演示文稿的 RangeType 属性的值。   |
| *To*          | 可选      | Number      | 要打印的最后一页的页码。如果忽略此参数，将一直打印到演示文稿的末尾。指定 To 和 From 参数会设置 PrintRanges 对象的内容，并设置演示文稿的 RangeType 属性的值。 |
| *PrintToFile* | 可选      | String      | 要输出到的文件名。如果指定此参数，则文件将输出到某个文件，而不是发送到打印机。如果忽略此参数，文件将发送至打印机。                                           |
| *Copies*      | 可选      | Number      | 要打印的份数。如果忽略此参数，将只打印一份。指定此参数会设置 NumberOfCopies 属性的值。                                                                       |
| *Collate*     | 可选      | MsoTriState | 如果忽略此参数，将逐份打印多份。指定此参数会设置 Collate 属性的值。                                                                                          |

**说明**

|                                                                        |
|------------------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。                          |
| **msoCTrue**                                                           |
| **msoFalse**                                                           |
| **msoTriStateMixed**                                                   |
| **msoTriStateToggle**                                                  |
| **msoTrue** ：先打印该演示文稿的完整副本，然后再打印下一副本的第一页。 |

**示例**

``` JavaScript
/*本示例打印每个幻灯片（从当前演示文稿的第二个幻灯片到第五个幻灯片）的两个不分页复制的副本（不管是可见的还是隐藏的）。*/
function test()
{
    let mypertion = Application.ActivePresentation
    mypertion.PrintOptions.PrintHiddenSlides = true
    mypertion.PrintOut(2,5,"test",2,msoFalse)
}

/*本示例将当前演示文稿中的所有幻灯片的单个副本输出到文件 Testprnt.prn。*/
Application.ActivePresentation.PrintOut(undefined,undefined,"TestPrnt")
```

### Presentation.PublishSlides

通过加载的任意演示文稿创建包含幻灯片的 Web 演示文稿（HTML 格式）。可以在 Web 浏览器中查看发布的演示文稿。

**语法**

**express.PublishSlides ( *SlideLibraryUrl* , *Overwrite* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型 | 说明                                      |
|-------------------|-----------|----------|-------------------------------------------|
| *SlideLibraryUrl* | 必选      | String   | 幻灯片库的 URL。                          |
| *Overwrite*       | 可选      | Boolean  | 如果覆盖原始演示文稿，则该参数值为 True。 |

### Presentation.ReloadAs

基于指定的 HTML 文档编码重新加载演示文稿。

**语法**

**express.ReloadAs ( *cp* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称 | 必选/可选 | 数据类型    | 说明                           |
|------|-----------|-------------|--------------------------------|
| *cp* | 必选      | MsoEncoding | 重新加载网页时使用的文档编码。 |

**说明**

|                                               |
|-----------------------------------------------|
| MsoEncoding 可以是下列 MsoEncoding 常量之一。 |
| **msoEncodingArabicAutoDetect**               |
| **msoEncodingAutoDetect**                     |
| **msoEncodingCyrillicAutoDetect**             |
| **msoEncodingGreekAutoDetect**                |
| **msoEncodingJapaneseAutoDetect**             |
| **msoEncodingKoreanAutoDetect**               |
| **msoEncodingSimplifiedChineseAutoDetect**    |
| **msoEncodingTraditionalChineseAutoDetect**   |

**示例**

``` JavaScript
//本示例使用西方编码重新加载当前演示文稿。
Application.ActivePresentation.ReloadAs(Application.Enum.msoEncodingWestern)
```

### Presentation.RemoveDocumentInformation

删除 演示文稿中的文档信息，如个人信息、批注和文档属性。

**语法**

**express.RemoveDocumentInformation ( *Type* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型            | 说明                 |
|--------|-----------|---------------------|----------------------|
| *Type* | 必选      | PpRemoveDocInfoType | 要删除的信息的类型。 |

### Presentation.Save

保存指定的演示文稿。

**语法**

**express.Save ()**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用 **[SaveAs](#Presentation.SaveAs)** 方法可保存尚未保存的演示文稿。若要判断演示文稿是否已保存，请测试 **[FullName](#Presentation.FullName)** 或 **[Path](#Presentation.Path)** 属性是否非空。如果磁盘上已存在与指定演示文稿同名的文件，将覆盖该文件且不显示警告信息。

若要将演示文稿标记为已保存，但不将其写入磁盘，请将 **Saved** 属性设为 **True** 。

**示例**

``` JavaScript
/*如果当前演示文稿在上次保存后进行了更改，则本示例保存当前演示文稿。*/
function test()
{
    let tion = Application.ActivePresentation
    if(tion.Saved == false && tion.Path != "") 
    {
        tion.Save()
    }
}
```

### Presentation.SaveAs

保存未保存过的演示文稿，或以其他文件名保存已保存过的演示文稿。

**语法**

**express.SaveAs ( *Filename* , *FileFormat* , *EmbedFonts* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型         | 说明                                                                                                    |
|--------------|-----------|------------------|---------------------------------------------------------------------------------------------------------|
| *Filename*   | 必选      | String           | 指定文件的保存名称。如果不指定完整路径，WPP会将该文件保存于当前文件夹中。                               |
| *FileFormat* | 可选      | PpSaveAsFileType | 指定文件的保存格式。如果省略该参数，则文件将以WPP 当前版本中的演示文稿格式保存 (ppSaveAsPresentation)。 |
| *EmbedFonts* | 可选      | MsoTriState      | 指定WPP 是否将 TrueType 字体嵌入保存的演示文稿中。                                                      |

**说明**

|                                                         |
|---------------------------------------------------------|
| PpSaveAsFileType 可以是下列 PpSaveAsFileType 常量之一。 |
| ppSaveAsHTMLv3                                          |
| ppSaveAsAddIn                                           |
| ppSaveAsBMP                                             |
| ppSaveAsDefault                                         |
| ppSaveAsGIF                                             |
| ppSaveAsHTML                                            |
| ppSaveAsHTMLDual                                        |
| ppSaveAsJPG                                             |
| ppSaveAsMetaFile                                        |
| ppSaveAsPNG                                             |
| ppSaveAsPowerPoint3                                     |
| ppSaveAsPowerPoint4                                     |
| ppSaveAsPowerPoint4FarEast                              |
| ppSaveAsPowerPoint7                                     |
| ppSaveAsPresentation 默认值。                           |
| ppSaveAsRTF                                             |
| ppSaveAsShow                                            |
| ppSaveAsTemplate                                        |
| ppSaveAsTIF                                             |
| ppSaveAsWebArchive                                      |

|                                                          |
|----------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。            |
| **msoCTrue**                                             |
| **msoFalse**                                             |
| **msoTriStateMixed** 默认值。                            |
| **msoTriStateToggle**                                    |
| **msoTrue** WPP 将 TrueType 字体嵌入到保存的演示文稿中。 |

**示例**

``` JavaScript
/*本示例以文件名“New Format Copy.ppt”保存当前演示文稿的一个副本。*/
/*默认情况下，该副本以当前版本的WPP 演示文稿格式保存，然后以文件名“Old Format Copy”另存为WPP 4.0 格式文件。*/
function test()
{
    let tion= Application.ActivePresentation
    tion.SaveCopyAs("New Format Copy")
    tion.SaveAs("Old Format Copy", ppSaveAsPowerPoint4)
}
```

### Presentation.SaveCopyAs

在不修改原文件的情况下将指定演示文稿的副本保存至另一文件。

**语法**

**express.SaveCopyAs ( *FileName* , *FileFormat* , *EmbedTrueTypeFonts* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型         | 说明                                                                                 |
|----------------------|-----------|------------------|--------------------------------------------------------------------------------------|
| *FileName*           | 必选      | String           | 指定文件的保存名称。如果不指定完整路径，WPP会将该文件保存于当前文件夹中。 FileFormat |
| *FileFormat*         | 可选      | PpSaveAsFileType | 文件格式。                                                                           |
| *EmbedTrueTypeFonts* | 可选      | MsoTriState      | 指定是否嵌入 TrueType 字体。                                                         |

**说明**

|                                 |
|---------------------------------|
| FileFormat 可以是下列常量之一。 |
| ppSaveAsHTMLv3                  |
| ppSaveAsAddIn                   |
| ppSaveAsBMP                     |
| ppSaveAsDefault 默认值          |
| ppSaveAsGIF                     |
| ppSaveAsHTML                    |
| ppSaveAsHTMLDual                |
| ppSaveAsJPG                     |
| ppSaveAsMetaFile                |
| ppSaveAsPNG                     |
| ppSaveAsPowerPoint3             |
| ppSaveAsPowerPoint4             |
| ppSaveAsPowerPoint4FarEast      |
| ppSaveAsPowerPoint7             |
| ppSaveAsPresentation            |
| ppSaveAsRTF                     |
| ppSaveAsShow                    |
| ppSaveAsTemplate                |
| ppSaveAsTIF                     |
| ppSaveAsWebArchive              |

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed** 默认值                   |
| **msoTriStateToggle**                         |
| **msoTrue**                                   |

**示例**

``` JavaScript
/*本示例以文件名“New Format Copy.ppt”保存当前演示文稿的一个副本。*/
/*默认情况下，该副本以当前版本的WPP 演示文稿格式保存，然后以文件名“Old Format Copy”另存为WPP 4.0 格式文件。*/
function test()
{
    let tion= Application.ActivePresentation
    tion.SaveCopyAs("New Format Copy")
    tion.SaveAs("Old Format Copy", ppSaveAsPowerPoint4)
}
```

### Presentation.SendFaxOverInternet

向指定的收件人以传真形式发送演示文稿。

**语法**

**express.SendFaxOverInternet ( *Recipients* , *Subject* , *ShowMessage* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                         |
|---------------|-----------|----------|----------------------------------------------------------------------------------------------|
| *Recipients*  | 可选      | Variant  | 一个代表传真收件人的传真号码和电子邮件地址的 String 类型值。如果有多个收件人，则用分号隔开。 |
| *Subject*     | 可选      | Variant  | 一个代表传真演示文稿的主题行的 String 类型值。                                               |
| *ShowMessage* | 可选      | Variant  | 如果值为 True，则发送传真之前显示传真邮件。如果值为 False，则发送传真而不显示传真邮件。      |

**说明**

要使用 **SendFaxOverInternet** 方法，需要在用户的计算机上启用传真服务。

用于指定 **Recipients** 参数中传真号码的格式可为 *recipientsfaxnumber* @ *usersfaxprovider* 或 *recipientsname* @ *recipientsfaxnumber* 。可使用以下注册表路径访问用户的传真提供商信息：

` HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\Services\Fax `

使用上面的注册表路径下的 FaxAddress 键值确定用户可使用的格式。

**示例**

``` JavaScript
//以下示例将传真发送给传真服务提供商，该提供商将邮件传真给收件人。
Application.ActivePresentation.SendFaxOverInternet("14255550101@consolidatedmessenger.com","For your review", true)
```

### Presentation.SetPasswordEncryptionOptions

设置 WPP 通过密码加密演示文稿时使用的选项。

**语法**

**express.SetPasswordEncryptionOptions ( *PasswordEncryptionProvider* , *PasswordEncryptionAlgorithm* , *PasswordEncryptionKeyLength* , *PasswordEncryptionFileProperties* )**

\*express是一个代表 Presentation 对象的变量。

**参数**

| 名称                               | 必选/可选 | 数据类型    | 说明                                           |
|------------------------------------|-----------|-------------|------------------------------------------------|
| *PasswordEncryptionProvider*       | 必选      | String      | 加密提供程序的名称。                           |
| *PasswordEncryptionAlgorithm*      | 必选      | String      | 加密算法的名称。WPP支持流加密算法。            |
| *PasswordEncryptionKeyLength*      | 必选      | Long        | 加密密钥的长度。必须为 8 的倍数，且不小于 40。 |
| *PasswordEncryptionFileProperties* | 必选      | MsoTriState | 如果值为 MsoTrue，则WPP 会对文件属性进行加密。 |

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue** 不适用于此方法。                 |
| **msoFalse**                                  |
| **msoTriStateMixed** 不适用于此方法。         |
| **msoTriStateToggle** 不适用于此方法。        |
| **msoTrue**                                   |

**示例**

``` JavaScript
//本示例中，如果使用密码保护的文档的文件属性没有加密，则对密码加密选项进行设置。
function PasswordSettings() {
    let tion = Application.ActivePresentation
    if (tion.PasswordEncryptionFileProperties == msoFalse) {
        tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider", "RC4", 56, true)
    }
}
```

### Presentation.UpdateLinks

更新指定演示文稿中链接的 OLE 对象。

**语法**

**express.UpdateLinks ()**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*本示例更新当前演示文稿中的所有 OLE 链接。*/
Application.ActivePresentation.UpdateLinks()
```

### Presentation.WebPagePreview

在当前 Web 浏览器中预览演示文稿。

**语法**

**express.WebPagePreview ()**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//本示例将第二篇演示文稿作为网页进行预览。
Application.Presentations.Item(2).WebPagePreview()
```

## 成员属性

### Presentation.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*以下示例中在活动的演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test() 
{
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test()
{
    for(let i=1; i<= Application.ActivePresentation.Slides.Item(1).Shapes.Count; i++) 
    {
        let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes.Item(i)
        if(shpOle.Type == msoLinkedOLEObject) 
        {
            alert(shpOle.OLEFormat.Application.Name)
        }
    }
}
```

### Presentation.BuiltInDocumentProperties

返回一个 **DocumentProperties** 集合，该集合代表指定演示文稿的所有内置文档属性。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**语法**

**express.BuiltInDocumentProperties**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用 **CustomDocumentProperties** 属性返回自定义文档属性集合。

**示例**

``` JavaScript
//本示例显示活动演示文稿的所有内置文档属性的名称。
function test() {
    let bidpList
    for (let i = 1; i <= Application.ActivePresentation.BuiltInDocumentProperties.Count; i++) {
        let p = Application.ActivePresentation.BuiltInDocumentProperties.Item(i)
        bidpList = bidpList + p.Name + "\r"
    }
    alert(bidpList)
}

//如果当前演示文稿的作者是“Jake Jarmel”，则本示例将该演示文稿的内置属性设置为“Category”。
function test() {
    let ties = Application.ActivePresentation.BuiltInDocumentProperties
    if (ties.Item("author").Value == "Jake Jarmel") {
        ties.Item("category").Value = "Creative Writing"
    }
}
```

### Presentation.ColorSchemes

返回一个 **ColorSchemes** 集合，该集合代表指定演示文稿中的配色方案。只读。

**语法**

**express.ColorSchemes**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
////本示例为返回活动演示文稿配色方案。
Application.ActivePresentation.ColorSchemes
```

### Presentation.CommandBars

返回一个 **CommandBars** 集合，该集合代表主机容器应用程序和 WPP 中合并的命令栏集。仅当容器是 DocObject 服务器（如 活页夹）并且 WPP 作为 OLE 服务器时，该属性返回有效对象。只读。

**语法**

**express.CommandBars**

\*express是一个代表 Presentation 对象的变量。

### Presentation.Container

返回包含指定嵌入演示文稿的对象。只读 **Object** 类型。如果容器不支持 OLE 自动化，或指定的演示文稿不是嵌入在“活页夹”文件中，则该属性无效。

**语法**

**express.Container**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//本示例将隐藏包含嵌入活动演示文稿的“活页夹”文件的第二部分。演示文稿的 Container 属性将返回一个 Section 对象，该 Parent 对象的 Section 属性将返回一个 Binder 对象。
Application.ActivePresentation.Container.Parent.Sections.Item(2).Visible = false
```

### Presentation.ContentTypeProperties

返回 MetaProperties 集合，该集合描述演示文稿中存储的元数据。只读。

**语法**

**express.ContentTypeProperties**

\*express是一个代表 Presentation 对象的变量。

### Presentation.CustomDocumentProperties

返回一个 **DocumentProperties** 集合，该集合代表指定演示文稿的所有自定义文档属性。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**语法**

**express.CustomDocumentProperties**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用 **BuiltInDocumentProperties** 属性可以返回内置文档属性集合。

**示例**

``` JavaScript
//本示例为活动演示文稿添加一个名为“Complete”的静态自定义属性。
Application.ActivePresentation.CustomDocumentProperties.Add("Complete",false,msoPropertyTypeBoolean,false)//如果自定义属性“Complete”的值为 True，本示例将显示当前演示文稿。
function test() {
    let tion = Application.ActivePresentation
    if (tion.CustomDocumentProperties.Item("complete")) {
        tion.PrintOut()
    }
}
```

### Presentation.CustomerData

返回一个 CustomerData 对象。只读。

**语法**

**express.CustomerData**

\*express是一个代表 Presentation 对象的变量。

### Presentation.CustomXMLParts

返回一个 **CustomXMLParts** 对象。

**语法**

**express.CustomXMLParts**

\*express是一个代表 Presentation 对象的变量。

### Presentation.DefaultLanguageID

返回或设置演示文稿的默认语言。在您设置一个演示文稿的 **DefaultLanguageID** 属性时，同时也为所有后续新演示文稿设置了该属性。可读/写 **MsoLanguageID** 类型。

**语法**

**express.DefaultLanguageID**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| MsoLanguageID 可以是下列 MsoLanguageID 常量之一。 |
| **msoLanguageIDAfrikaans**                        |
| **msoLanguageIDAlbanian**                         |
| **msoLanguageIDAmharic**                          |
| **msoLanguageIDArabic**                           |
| **msoLanguageIDArabicAlgeria**                    |
| **msoLanguageIDArabicBahrain**                    |
| **msoLanguageIDArabicEgypt**                      |
| **msoLanguageIDArabicIraq**                       |
| **msoLanguageIDArabicJordan**                     |
| **msoLanguageIDArabicKuwait**                     |
| **msoLanguageIDArabicLebanon**                    |
| **msoLanguageIDArabicLibya**                      |
| **msoLanguageIDArabicMorocco**                    |
| **msoLanguageIDArabicOman**                       |
| **msoLanguageIDArabicQatar**                      |
| **msoLanguageIDArabicSyria**                      |
| **msoLanguageIDArabicTunisia**                    |
| **msoLanguageIDArabicUAE**                        |
| **msoLanguageIDArabicYemen**                      |
| **msoLanguageIDArmenian**                         |
| **msoLanguageIDAssamese**                         |
| **msoLanguageIDAzeriCyrillic**                    |
| **msoLanguageIDAzeriLatin**                       |
| **msoLanguageIDBasque**                           |
| **msoLanguageIDBelgianDutch**                     |
| **msoLanguageIDBelgianFrench**                    |
| **msoLanguageIDBengali**                          |
| **msoLanguageIDBrazilianPortuguese**              |
| **msoLanguageIDBulgarian**                        |
| **msoLanguageIDBurmese**                          |
| **msoLanguageIDByelorussian**                     |
| **msoLanguageIDCatalan**                          |
| **msoLanguageIDCherokee**                         |
| **msoLanguageIDChineseHongKong**                  |
| **msoLanguageIDChineseMacao**                     |
| **msoLanguageIDChineseSingapore**                 |
| **msoLanguageIDCroatian**                         |
| **msoLanguageIDCzech**                            |
| **msoLanguageIDDanish**                           |
| **msoLanguageIDDutch**                            |
| **msoLanguageIDEnglishAUS**                       |
| **msoLanguageIDEnglishBelize**                    |
| **msoLanguageIDEnglishCanadian**                  |
| **msoLanguageIDEnglishCaribbean**                 |
| **msoLanguageIDEnglishIreland**                   |
| **msoLanguageIDEnglishJamaica**                   |
| **msoLanguageIDEnglishNewZealand**                |
| **msoLanguageIDEnglishPhilippines**               |
| **msoLanguageIDEnglishSouthAfrica**               |
| **msoLanguageIDEnglishTrinidad**                  |
| **msoLanguageIDEnglishUK**                        |
| **msoLanguageIDEnglishUS**                        |
| **msoLanguageIDEnglishZimbabwe**                  |
| **msoLanguageIDEstonian**                         |
| **msoLanguageIDFaeroese**                         |
| **msoLanguageIDFarsi**                            |
| **msoLanguageIDFinnish**                          |
| **msoLanguageIDFrench**                           |
| **msoLanguageIDFrenchCameroon**                   |
| **msoLanguageIDFrenchCanadian**                   |
| **msoLanguageIDFrenchCotedIvoire**                |
| **msoLanguageIDFrenchLuxembourg**                 |
| **msoLanguageIDFrenchMali**                       |
| **msoLanguageIDFrenchMonaco**                     |
| **msoLanguageIDFrenchReunion**                    |
| **msoLanguageIDFrenchSenegal**                    |
| **msoLanguageIDFrenchWestIndies**                 |
| **msoLanguageIDFrenchZaire**                      |
| **msoLanguageIDFrisianNetherlands**               |
| **msoLanguageIDGaelicIreland**                    |
| **msoLanguageIDGaelicScotland**                   |
| **msoLanguageIDGalician**                         |
| **msoLanguageIDGeorgian**                         |
| **msoLanguageIDGerman**                           |
| **msoLanguageIDGermanAustria**                    |
| **msoLanguageIDGermanLiechtenstein**              |
| **msoLanguageIDGermanLuxembourg**                 |
| **msoLanguageIDGreek**                            |
| **msoLanguageIDGujarati**                         |
| **msoLanguageIDHebrew**                           |
| **msoLanguageIDHindi**                            |
| **msoLanguageIDHungarian**                        |
| **msoLanguageIDIcelandic**                        |
| **msoLanguageIDIndonesian**                       |
| **msoLanguageIDInuktitut**                        |
| **msoLanguageIDItalian**                          |
| **msoLanguageIDJapanese**                         |
| **msoLanguageIDKannada**                          |
| **msoLanguageIDKashmiri**                         |
| **msoLanguageIDKazakh**                           |
| **msoLanguageIDKhmer**                            |
| **msoLanguageIDKirghiz**                          |
| **msoLanguageIDKonkani**                          |
| **msoLanguageIDKorean**                           |
| **msoLanguageIDLao**                              |
| **msoLanguageIDLatvian**                          |
| **msoLanguageIDLithuanian**                       |
| **msoLanguageIDMacedonian**                       |
| **msoLanguageIDMalayalam**                        |
| **msoLanguageIDMalayBruneiDarussalam**            |
| **msoLanguageIDMalaysian**                        |
| **msoLanguageIDMaltese**                          |
| **msoLanguageIDManipuri**                         |
| **msoLanguageIDMarathi**                          |
| **msoLanguageIDMexicanSpanish**                   |
| **msoLanguageIDMixed**                            |
| **msoLanguageIDMongolian**                        |
| **msoLanguageIDNepali**                           |
| **msoLanguageIDNone**                             |
| **msoLanguageIDNoProofing**                       |
| **msoLanguageIDNorwegianBokmol**                  |
| **msoLanguageIDNorwegianNynorsk**                 |
| **msoLanguageIDOriya**                            |
| **msoLanguageIDPolish**                           |
| **msoLanguageIDPunjabi**                          |
| **msoLanguageIDRhaetoRomanic**                    |
| **msoLanguageIDRomanian**                         |
| **msoLanguageIDRomanianMoldova**                  |
| **msoLanguageIDRussian**                          |
| **msoLanguageIDRussianMoldova**                   |
| **msoLanguageIDSamiLappish**                      |
| **msoLanguageIDSanskrit**                         |
| **msoLanguageIDSerbianCyrillic**                  |
| **msoLanguageIDSerbianLatin**                     |
| **msoLanguageIDSesotho**                          |
| **msoLanguageIDSimplifiedChinese**                |
| **msoLanguageIDSindhi**                           |
| **msoLanguageIDSlovak**                           |
| **msoLanguageIDSlovenian**                        |
| **msoLanguageIDSorbian**                          |
| **msoLanguageIDSpanish**                          |
| **msoLanguageIDSpanishArgentina**                 |
| **msoLanguageIDSpanishBolivia**                   |
| **msoLanguageIDSpanishChile**                     |
| **msoLanguageIDSpanishColombia**                  |
| **msoLanguageIDSpanishCostaRica**                 |
| **msoLanguageIDSpanishDominicanRepublic**         |
| **msoLanguageIDSpanishEcuador**                   |
| **msoLanguageIDSpanishElSalvador**                |
| **msoLanguageIDSpanishGuatemala**                 |
| **msoLanguageIDSpanishHonduras**                  |
| **msoLanguageIDSpanishModernSort**                |
| **msoLanguageIDSpanishNicaragua**                 |
| **msoLanguageIDSpanishPanama**                    |
| **msoLanguageIDSpanishParaguay**                  |
| **msoLanguageIDSpanishPeru**                      |
| **msoLanguageIDSpanishPuertoRico**                |
| **msoLanguageIDSpanishUruguay**                   |
| **msoLanguageIDSpanishVenezuela**                 |
| **msoLanguageIDSutu**                             |
| **msoLanguageIDSwahili**                          |
| **msoLanguageIDSwedish**                          |
| **msoLanguageIDSwedishFinland**                   |
| **msoLanguageIDSwissFrench**                      |
| **msoLanguageIDSwissGerman**                      |
| **msoLanguageIDSwissItalian**                     |
| **msoLanguageIDTajik**                            |
| **msoLanguageIDTamil**                            |
| **msoLanguageIDTatar**                            |
| **msoLanguageIDTelugu**                           |
| **msoLanguageIDThai**                             |
| **msoLanguageIDTibetan**                          |
| **msoLanguageIDTraditionalChinese**               |
| **msoLanguageIDTsonga**                           |
| **msoLanguageIDTswana**                           |
| **msoLanguageIDTurkish**                          |
| **msoLanguageIDTurkmen**                          |
| **msoLanguageIDUkrainian**                        |
| **msoLanguageIDUrdu**                             |
| **msoLanguageIDUzbekCyrillic**                    |
| **msoLanguageIDUzbekLatin**                       |
| **msoLanguageIDVenda**                            |
| **msoLanguageIDVietnamese**                       |
| **msoLanguageIDWelsh**                            |
| **msoLanguageIDXhosa**                            |
| **msoLanguageIDZulu**                             |
| **msoLanguageIDPortuguese**                       |

**示例**

``` JavaScript
//本示例将活动演示文稿以及随后的所有新演示文稿的默认语言设置为 German。
Application.ActivePresentation.DefaultLanguageID = msoLanguageIDGerman
```

### Presentation.DefaultShape

返回一个 **Shape** 对象，该对象代表演示文稿的默认形状。只读。

**语法**

**express.DefaultShape**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//本示例为活动演示文稿的第一张幻灯片添加一个形状，将演示文稿中形状的默认填充颜色设为红色。然后添加另一个形状，默认填充颜色将自动应用于此形状。
function test() {
    let tion = Application.ActivePresentation
    let sld1Shapes = tion.Slides.Item(1).Shapes
    sld1Shapes.AddShape(msoShape16pointStar, 20, 20, 100, 100)
    tion.DefaultShape.Fill.ForeColor.RGB = (255, 0, 0)
    sld1Shapes.AddShape(msoShape16pointStar, 150, 20, 100, 100)
}
```

### Presentation.Designs

返回一个 **Designs** 对象，该对象代表一个设计集合。

**语法**

**express.Designs**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*以下示例对于当前演示文稿中的每个设计都会显示一条消息。*/
function test()
{
    let pre = Application.ActivePresentation
    for(let i=1; i<=pre.Designs.Count; i++) 
    {
        alert("The design name is " + pre.Designs.Item(pre.Designs.Item(i).Index).Name)
    }
}
```

### Presentation.DisplayComments

决定是否在指定的演示文稿中显示批注。可读/写 **MsoTriState** 类型。

**语法**

**express.DisplayComments**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 在指定的演示文稿中显示批注。      |

**示例**

``` JavaScript
//本示例隐藏活动演示文稿中的批注。
Application.ActivePresentation.DisplayComments = msoFalse
```

### Presentation.DocumentInspectors

返回 DocumentInspectors 集合。只读。

**语法**

**express.DocumentInspectors**

\*express是一个代表 Presentation 对象的变量。

### Presentation.DocumentLibraryVersions

返回一个 **DocumentLibraryVersions** 集合，该集合代表共享演示文稿的版本集合，该共享演示文稿启用了版本控制功能，并存储在服务器上的文档库中。

**语法**

**express.DocumentLibraryVersions**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//下面的示例返回活动演示文稿的版本集合。本示例假定活动演示文稿已启用了版本控制功能，并存储在服务器上的共享文档库中。
let objVersions = Application.ActivePresentation.DocumentLibraryVersions
```

### Presentation.EnvelopeVisible

决定电子邮件的邮件头在文档窗口中是否可见。可读/写 **MsoTriState** 类型。

**语法**

**express.EnvelopeVisible**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                                |
|------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。  |
| **msoCTrue**                                   |
| **msoFalse** 默认值。                          |
| **msoTriStateMixed**                           |
| **msoTriStateToggle**                          |
| **msoTrue** 电子邮件的邮件头在文档窗口中可见。 |

**示例**

``` JavaScript
//本示例显示电子邮件的邮件头。
Application.ActivePresentation.EnvelopeVisible = msoTrue
```

### Presentation.ExtraColors

返回一个 **ExtraColors** 对象，该对象代表指定演示文稿中可用的其他颜色。只读。

**语法**

**express.ExtraColors**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//以下示例在当前演示文稿的第一张幻灯片中添加一个矩形，并将其前景填充颜色设为第一种其他颜色。演示文稿中必须已定义至少一种其他颜色，否则本示例将无效。
function test() {
    let tion = Application.ActivePresentation
    let rect = tion.Slides.Item(1).Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)
    rect.Fill.ForeColor.RGB = tion.ExtraColors.Item(1)
}
```

### Presentation.FarEastLineBreakLanguage

返回或设置启用了换行控制选项时用于确定所用换行级别的语言。可读/写 **MsoFarEastLineBreakLanguageID** 类型。

**语法**

**express.FarEastLineBreakLanguage**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                                                                   |
|-----------------------------------------------------------------------------------|
| MsoFarEastLineBreakLanguageID 可以是下列 MsoFarEastLineBreakLanguageID 常量之一。 |
| **MsoFarEastLineBreakLanguageJapanese**                                           |
| **MsoFarEastLineBreakLanguageKorean**                                             |
| **MsoFarEastLineBreakLanguageSimplifiedChinese**                                  |
| **MsoFarEastLineBreakLanguageTraditionalChinese**                                 |

**示例**

``` JavaScript
//以下示例将换行语言设置为日语。
Application.ActivePresentation.FarEastLineBreakLanguage = MsoFarEastLineBreakLanguageJapanese
```

### Presentation.FarEastLineBreakLevel

返回或设置基于亚洲字符级别的换行符。可读/写 **Long** 类型。可读/写 **PpFarEastLineBreakLevel** 类型。

**语法**

**express.FarEastLineBreakLevel**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                                                       |
|-----------------------------------------------------------------------|
| PpFarEastLineBreakLevel 可以是下列 PpFarEastLineBreakLevel 常量之一。 |
| **ppFarEastLineBreakLevelCustom**                                     |
| **ppFarEastLineBreakLevelNormal**                                     |
| **ppFarEastLineBreakLevelStrict**                                     |

**示例**

``` JavaScript
//本示例将换行控制设置为使用第一级首尾字符。
Application.ActivePresentation.FarEastLineBreakLevel = ppFarEastLineBreakLevelNormal
```

### Presentation.Final

确定是否将演示文稿标记为最终版本（只读）。可读/写。

**语法**

**express.Final**

\*express是一个代表 Presentation 对象的变量。

**说明**

返回值 Boolean

### Presentation.Fonts

返回一个 **Fonts** 集合，该集合代表用于指定演示文稿的所有字体。只读。

**语法**

**express.Fonts**

\*express是一个代表 Presentation 对象的变量。

### Presentation.FullName

返回指定的加载宏或保存的演示文稿的名称，包括路径、当前文件系统分隔符和文件扩展名。只读。 **String** 类型。

**语法**

**express.FullName**

\*express是一个代表 Presentation 对象的变量。

**说明**

此属性与 **Path** 属性等效，附加上当前文件系统分隔符和 **Name** 属性。

**示例**

``` JavaScript
/*以下示例显示当前演示文稿的路径和文件名（假设演示文稿已保存）。*/
alert(Application.ActivePresentation.FullName)
```

### Presentation.GridDistance

设置或返回一个 **Single** 类型的值，该值代表网格线之间的距离。可读/写。

**语法**

**express.GridDistance**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//本示例显示网格线，然后指定网格线之间的距离并启用对齐网格的设置。
function SetGridLines() {
    Application.DisplayGridLines = msoTrue
    let tion = Application.ActivePresentation
    tion.GridDistance = 18
    tion.SnapToGrid = msoTrue
}
```

### Presentation.HandoutMaster

返回一个代表讲义母版的 **Master** 对象。只读。

**语法**

**express.HandoutMaster**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//本示例设置活动演示文稿的讲义母版的背景图案。
Application.ActivePresentation.HandoutMaster.Background.Fill.Patterned(Application.Enum.msoPatternDarkHorizontal)
```

### Presentation.HasTitleMaster

如果指定的演示文稿具有标题母版，则为 **MsoTrue** 。只读。 **MsoTriState** 类型。

**语法**

**express.HasTitleMaster**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** ：指定的演示文稿具有标题母版。    |

**示例**

``` JavaScript
//以下示例中，如果当前演示文稿没有标题母版，则为其添加一个标题母版。
function test() {
    let tion = Application.ActivePresentation
    if (!tion.HasTitleMaster) {
        tion.AddTitleMaster()
    }
}
```

### Presentation.HasVBProject

返回一个值，该值指示当前演示文稿是否包含 示例代码 (VBA) 项目。只读

**语法**

**express.HasVBProject**

\*express是一个代表 Presentation 对象的变量。

### Presentation.LayoutDirection

为用户界面返回或设置版式方向。可为以下值之一，默认值取决于已选择或安装的语言支持。可读/写 **PpDirection** 类型。

**语法**

**express.LayoutDirection**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| PpDirection 可以是下列 PpDirection 常量之一。 |
| **ppDirectionLeftToRight**                    |
| **ppDirectionMixed**                          |
| **ppDirectionRightToLeft**                    |

**示例**

``` JavaScript
/*本示例将版式方向设置为从右向左。*/
Application.ActivePresentation.LayoutDirection = ppDirectionRightToLeft
```

### Presentation.Name

演示文稿的名称包括文件扩展名（对于已注册的文件类型），但不包括其路径。不能使用该属性来设置名称。如果需要更改名称，请使用 **[SaveAs](#Presentation.SaveAs)** 方法将演示文稿另存为其他名称。 **String** 类型，只读 。

**语法**

**express.Name**

\*express是一个代表 Presentation 对象的变量。

### Presentation.NoLineBreakAfter

返回或设置前置字符。可读/写 **String** 类型。

**语法**

**express.NoLineBreakAfter**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//本示例将“$”、“(”、“[”、“\”和“{”设置为前置字符。
function test() {
    let tion = Application.ActivePresentation
    tion.FarEastLineBreakLevel = ppFarEastLineBreakLevelCustom
    tion.NoLineBreakAfter = "$([\{"
}
```

### Presentation.NoLineBreakBefore

返回或设置后置标点字符。可读/写 **String** 类型。

**语法**

**express.NoLineBreakBefore**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//本示例将“!”、“)”和“]”设置为后置字符。
function test() {
    let tion = Application.ActivePresentation
    tion.FarEastLineBreakLevel = ppFarEastLineBreakLevelCustom
    tion.NoLineBreakBefore = "!)]"
}
```

### Presentation.NotesMaster

返回一个代表备注母版的 **Master** 对象。只读 。

**语法**

**express.NotesMaster**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动演示文稿的备注母版的页眉和页脚文。*/
function test()
{
    let ster = Application.ActivePresentation.NotesMaster.HeadersFooters
    ster.Header.Text = "Employee Guidelines"
    ster.Footer.Text = "Volcano Coffee"
}
```

### Presentation.PageSetup

返回一个 **PageSetup** 对象，该对象的属性控制指定演示文稿的幻灯片的页面设置属性。只读。

**语法**

**express.PageSetup**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*以下示例为名为“Pres1”的演示文稿设置幻灯片的大小和方向。*/
function test()
{
    let setup = Application.Presentations.Item("pres1").PageSetup
    setup.SlideSize = ppSlideSize35MM
    setup.SlideOrientation = msoOrientationHorizontal
}
```

### Presentation.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

### Presentation.Password

返回或设置一个 **String** 类型的值，该值代表要打开指定演示文稿必须提供的密码。可读/写。

**语法**

**express.Password**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*本示例打开 Earnings.ppt，为其设置密码，然后关闭该演示文稿。*/
function test()
{
    let pen = Application.Presentations.Open("C:\\My Documents\\Earnings.ppt")
    pen.Password = "complexstrPWD"
    pen.Save()
    pen.Close()
}
```

### Presentation.PasswordEncryptionAlgorithm

返回一个 **String** 类型的值，该值代表 WPP 通过密码加密文档所使用的算法。只读。

**语法**

**express.PasswordEncryptionAlgorithm**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法可指定WPP 通过密码加密文档所使用的算法。

**示例**

``` JavaScript
//本示例中，如果使用的密码加密算法不是 RC4，则对密码加密选项进行设置。
function PasswordSettings() {
    let tion = Application.ActivePresentation
    if(tion.PasswordEncryptionAlgorithm!="RC4") {
        tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider","RC4",56,true)
    }
}
```

### Presentation.PasswordEncryptionFileProperties

如果 WPP 对使用密码保护的文档的文件属性进行加密，则返回 **MsoTrue** 属性值。只读 **MsoTriState** 类型。

**语法**

**express.PasswordEncryptionFileProperties**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法可指定WPP 通过密码加密文档所使用的算法。

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**                                   |

**示例**

``` JavaScript
//本示例中，如果使用密码保护的文档的文件属性没有加密，则对密码加密选项进行设置。
function PasswordSettings() {
    let tion = Application.ActivePresentation
    if(tion.PasswordEncryptionFileProperties == msoFalse) {
        tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider","RC4",56,true)
    }
}
```

### Presentation.PasswordEncryptionKeyLength

返回一个 **Long** 类型的值，该值表示 WPP 通过密码加密文档时所使用算法的密钥长度。只读。

**语法**

**express.PasswordEncryptionKeyLength**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法可指定WPP 通过密码加密文档所使用的算法。

**示例**

``` JavaScript
//本示例中，如果密码加密密钥的长度小于 40，则对密码加密选项进行设置。
function PasswordSettings() {
    let tion = Application.ActivePresentation
    if(tion.PasswordEncryptionKeyLength < 40) {
        tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider","RC4",56,true)
    }
}
```

### Presentation.PasswordEncryptionProvider

返回一个 **String** 类型的值，该值指定 WPP 通过密码加密文档时所使用的加密算法提供商的名称。只读。

**语法**

**express.PasswordEncryptionProvider**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法可指定WPP 通过密码加密文档所使用的算法。

**示例**

``` JavaScript
//本示例中，如果使用的密码加密算法的提供商不是 RSA SChannel Cryptographic Provider，则对密码加密选项进行设置。
function PasswordSettings() {
    let tion = Application.ActivePresentation
    if(tion.PasswordEncryptionProvider != " RSA SChannel Cryptographic Provider") {
        tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider","RC4",56,true)
    }
}
```

### Presentation.Path

返回 **String** ，该值代表到指定的 **Presentation** 对象的路径， 只读。

**语法**

**express.Path**

\*express是一个代表 Presentation 对象的变量。

**说明**

如果使用此属性返回一个尚未保存的演示文稿的路径，则会得到一个空字符串。

路径中不包括最后的反斜线 (\\ 或指定对象的名称。使用 **Presentation** 对象的 **Name** 属性可返回不带路径的文件名，使用 **FullName** 属性则可返回带路径的文件名。

### Presentation.Permission

返回一个 **Permission** 对象，该对象可用于限制对活动演示文稿的权限，还可用于返回或设置特定权限设置。只读。

**语法**

**express.Permission**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用 **Permission** 对象可限制对活动文档的权限，还可返回或设置特定权限设置。

使用 **Enabled** 属性可决定是否限制对活动文档的权限。使用 **Count** 属性可返回拥有权限的用户数量，使用 **RemoveAll** 方法可重置现有的所有权限。

**DocumentAuthor** 、 **EnableTrustedBrowser** 、 **RequestPermissionURL** 和 **StoreLicenses** 属性提供有关权限设置的其他信息。

**Permission** 对象分配对 **UserPermission** 对象集合的访问权限。使用 **UserPermission** 对象可将特定的权限集合与各个用户关联。当某些通过用户界面（如 **msoPermissionPrint** ）分配的权限应用到所有用户时，可使用 **UserPermission** 对象为每个用户分配权限，并为每个用户指定权限的到期日期。

“信息版权管理”支持管理权限策略的使用，这些策略列出用户和组及其文档权限。使用 **ApplyPolicy** 方法可应用权限策略，使用 **PermissionFromPolicy** 、 **PolicyName** 和 **PolicyDescription** 属性可返回策略信息。

无论是否限制了对活动文档的权限，都可使用 **Permission** 。活动文档没有受限权限时， **Presentation** 对象的 **Permission** 属性不返回 **Nothing** 。使用 **Enabled** 属性可决定文档是否有受限权限。

**示例**

``` JavaScript
//下面的示例创建一个新的演示文稿，并向电子邮件地址为“someone@example.com”的用户分配对新建演示文稿的读取权限。该示例将显示所有者和新用户的权限。
function AddUserPermissions() {
    let myPres = Application.Presentations.Add(msoTrue)
    let myPer = myPres.Permission
    myPer.Enabled = true
    let NewOwnerPer = myPer.Add("someone@example.com", msoPermissionRead )
    alert(myPer.Item(1).UserId + " " + Str(myPer.Item(1).Permission))
    alert(myPer.Item(2).UserId + " " + Str(myPer.Item(2).Permission))
}
```

### Presentation.PrintOptions

返回一个 **PrintOptions** 对象，该对象代表与指定演示文稿共同保存的打印选项。只读。

**语法**

**express.PrintOptions**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//以下示例打印当前演示文稿中的隐藏幻灯片，并调整幻灯片的尺寸以适应纸张大小。
function test() {
    let tion = Application.ActivePresentation
    let ons = tion.PrintOptions
    ons.PrintHiddenSlides = true
    ons.FitToPage = true
    tion.PrintOut()
}
```

### Presentation.PublishObjects

返回一个 **PublishObjects** 集合，该集合代表已完全或部分加载的演示文稿的集合，这些演示文稿可作为 HTML 进行发布。只读。

**语法**

**express.PublishObjects**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//本示例将当前演示文稿的第三张幻灯片到第五张幻灯片发布为 HTML 格式，并将发布的幻灯片命名为“Mallard.htm”。
function test() {
    let ject = Application.ActivePresentation.PublishObjects.Item(1)
    ject.FileName = "C:\\Test\\Mallard.htm"
    ject.SourceType = ppPublishSlideRange
    ject.RangeStart = 3
    ject.RangeEnd = 5
    ject.Publish()
}
```

### Presentation.ReadOnly

返回指定的演示文稿是否为只读。只读 **MsoTriState** 类型。

**语法**

**express.ReadOnly**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 指定的演示文稿为只读。            |

**示例**

``` JavaScript
/*如果活动演示文稿为只读，本示例将其保存为文件“Newfile.ppt”。*/
function test()
{
    let pre = Application.ActivePresentation
    if(pre.ReadOnly) 
    {
        pre.SaveAs("newfile")
    }
}
```

### Presentation.RemovePersonalInformation

如果该属性值为 **MsoTrue** ，则 WPP 在保存演示文稿时从批注和修订中删除所有用户信息。 **MsoTriState** 类型，可读/写。

**语法**

**express.RemovePersonalInformation**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                                       |
|-------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue** 不应用于此属性。                         |
| **msoFalse** 批注、修订和个人信息保留在演示文稿中。   |
| **msoTriStateMixed** 不应用于该属性。                 |
| **msoTriStateToggle** 不应用于此属性。                |
| **msoTrue** 保存演示文稿时删除批注、修订和个人信息。  |

**示例**

``` JavaScript
//以下示例设置在用户下次保存当前演示文稿时删除所有个人信息。
function RemovePersonalInfo() {
    Application.ActivePresentation.RemovePersonalInformation = msoTrue
}
```

### Presentation.Research

返回一个 Research 对象，利用该对象可访问信息检索服务功能。只读。

**语法**

**express.Research**

\*express是一个代表 Presentation 对象的变量。

### Presentation.Saved

决定演示文稿自上次保存以来是否进行过更改。可读/写 **MsoTriState** 类型。

**语法**

**express.Saved**

\*express是一个代表 Presentation 对象的变量。

**说明**

如果将已修改演示文稿的 **Saved** 属性设置为 **msoTrue** ，则关闭演示文稿时不会提示用户保存所做的更改，并将丢失自上次保存后进行的所有更改。

|                                                    |
|----------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。      |
| **msoCTrue**                                       |
| **msoFalse**                                       |
| **msoTriStateMixed**                               |
| **msoTriStateToggle**                              |
| **msoTrue** 演示文稿自上次保存以来没有进行过更改。 |

**示例**

``` JavaScript
/*如果当前演示文稿在上次保存后进行了更改，则以下示例保存当前演示文稿。*/
function test()
{
    let pre = Application.ActivePresentation
    if(pre.Saved == false && pre.Path != "") 
    {
        pre.Save()
    }
}
```

### Presentation.ServerPolicy

返回一个ServerPolicy 对象。只读。

**语法**

**express.ServerPolicy**

\*express是一个代表 Presentation 对象的变量。

### Presentation.SharedWorkspace

返回一个 **SharedWorkspace** 对象，该对象代表指定演示文稿所在的文档工作区。只读。

**语法**

**express.SharedWorkspace**

\*express是一个代表 Presentation 对象的变量。

**说明**

使用 **SharedWorkspace** 对象可以将当前的 WPP 演示文稿添加到服务器上的 Windows SharePoint Services 文档工作区网站中，以便利用工作区的协作功能，或者断开文档与工作区的连接或从工作区删除文档。使用 **SharedWorkspace** 对象的集合可以管理与共享文档关联的文件、文件夹、链接、成员和任务。

无论文档是否存储在工作区中，均可以使用 **SharedWorkspace** 对象模型。当文档不共享时， **Presentation** 对象的 **SharedWorkspace** 属性不返回 **Nothing** 。使用 **SharedWorkspace** 对象的 **Connected** 属性可以确定当前演示文稿是否实际保存和连接到共享工作区。

用户必须有相应的权限才能使用 **SharedWorkspace** 对象层次结构中的对象、属性和方法。

使用 **SharedWorkspaceFiles** 集合（通过 **SharedWorkspace** 对象的 **Files** 属性访问）可以管理保存在共享工作区中的演示文稿和文件。

使用 **SharedWorkspaceFolders** 集合（通过 **SharedWorkspace** 对象的 **Folders** 属性访问）可以管理共享工作区主文档库文件夹中的子文件夹。

使用 **SharedWorkspaceLinks** 集合（通过 **SharedWorkspace** 对象的 **Links** 属性访问）可以管理附加文档和信息的链接，这些文档和信息是共享工作区中文档的协作成员所关心的。

使用 **SharedWorkspaceMembers** 集合（通过 **SharedWorkspace** 对象的 **Members** 属性访问）可以管理具有共享工作区参与权限的用户和具有工作区中保存的共享文档的协作权限的用户。

使用 **SharedWorkspaceTasks** 集合（通过 **SharedWorkspace** 对象的 **Tasks** 属性访问）可以管理分配给共享工作区中文档的协作成员的任务。

使用 **CreateNew** 方法可以新建文档工作区，并将活动文档添加到工作区中。使用 **Name** 和 **URL** 属性可以返回有关工作区的信息。

**SharedWorkspace** 对象通过服务器使用对象和属性的本地缓存。开发人员可能需要在执行某些操作之前更新此缓存，也可能需要将经过缓存的属性更改重新保存到服务器。使用 **SharedWorkspace** 对象的 **Refresh** 方法可以通过服务器刷新本地缓存，使用 **LastRefreshed** 属性可以确定上次刷新操作发生的时间。在本地修改了 **SharedWorkspaceLink** 和 **SharedWorkspaceTask** 对象的属性之后，使用这两类对象的 **Save** 方法可以将所做更改上载到服务器。

使用 **Disconnect** 方法可以断开活动文档的本地副本与共享工作区的连接，同时，使工作区中的共享副本保持原样。使用 **RemoveDocument** 方法可以将共享文档从共享工作区中彻底删除。

用户必须有相应的权限才能使用 **SharedWorkspace** 对象层次结构中的对象、属性和方法。向 **SharedWorkspaceMembers** 集合添加成员时，可以使用 *Role* 参数指定特定于每个工作区成员的权限集。

使用 **SharedWorkspace** 对象模型时可能会创建以下条件，即 **SharedWorkspace** 对象缓存与活动文档的 **“共享工作区”** 窗格中显示的用户界面不同步。例如，如果在 **“共享工作区”** 窗格打开的同时， **CreateNew** 方法以编程方式将活动文档添加到新的工作区中，则 **“共享工作区”** 窗格将继续显示 **“新建”** 按钮。在这种情况下，如果用户在 **“共享工作区”** 窗格进行的选择已不再有效，就会产生错误，此时，执行刷新操作以使显示与当前文档状态和共享工作区数据同步。

此外， **Presentation** 对象还具有返回 **Sync** 对象的 **Sync** 属性。使用 **Sync** 对象及其属性和方法可以管理共享文档的本地副本和服务器副本之间的同步。

**示例**

``` JavaScript
//下面的示例返回一个对存储活动文档的文档工作区的引用。本示例假定活动文档属于文档工作区。
let objWorkspace = Application.ActivePresentation.SharedWorkspace
```

### Presentation.Signatures

返回一个 **SignatureSet** 对象，该对象代表数字签名的集合。

**语法**

**express.Signatures**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//下面的代码行显示数字签名的个数。
function DisplayNumberOfSignatures() {
    alert("Number of digital signatures: " + Application.ActivePresentation.Signatures.Count)
}
```

### Presentation.SlideMaster

返回一个 Master 对象，该对象代表幻灯片母版。

**语法**

**express.SlideMaster**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*以下示例设置当前演示文稿的幻灯片母版的背景图案。*/
Application.ActivePresentation.SlideMaster.Background.Fill.PresetTextured(Application.Enum.msoTextureGreenMarble)
```

### Presentation.Slides

返回一个 **Slides** 集合，该集合代表指定演示文稿中的所有幻灯片。只读。

**语法**

**express.Slides**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*本示例向活动演示文稿添加一张幻灯片。*/
Application.ActivePresentation.Slides.Add(1, Application.Enum.ppLayoutTitle)
```

### Presentation.SlideShowSettings

返回一个 **SlideShowSettings** 对象，该对象代表指定演示文稿的幻灯片放映设置。只读。

**语法**

**express.SlideShowSettings**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*本示例启动一个预设使用扬声器的幻灯片放映。运行该幻灯片放映时关闭动画和旁白。*/
function test()
{
    let ting = Application.ActivePresentation.SlideShowSettings
    ting.ShowType = ppShowTypeSpeaker
    ting.ShowWithNarration = false
    ting.ShowWithAnimation = false
    ting.Run()
}
```

### Presentation.SlideShowWindow

返回一个 **SlideShowWindow** 对象，该对象代表正在运行的指定演示文稿所在的幻灯片放映窗口。只读。

**语法**

**express.SlideShowWindow**

\*express是一个代表 Presentation 对象的变量。

### Presentation.SnapToGrid

该属性值为 **MsoTrue** 时，在指定的演示文稿中将形状与网格线对齐。可读/写 **MsoTriState** 类型。

**语法**

**express.SnapToGrid**

\*express是一个代表 Presentation 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**                                   |

**示例**

``` JavaScript
//本示例切换在活动演示文稿中形状与网格线的对齐设置。
function ToggleSnapToGrid() {
    let tion = Application.ActivePresentation
    if(tion.SnapToGrid == msoTrue) {
        tion.SnapToGrid = msoFalse
    }
    else {
        tion.SnapToGrid = msoTrue
    }
}
```

### Presentation.Sync

返回一个 **Sync** 对象，使用该对象可对存储在 Windows SharePoint Services 共享工作区中的共享演示文稿的本地副本和服务器副本的同步进行管理。只读。

**语法**

**express.Sync**

\*express是一个代表 Presentation 对象的变量。

**说明**

**Sync** 对象的 **Status** 属性返回有关同步的当前状态的重要信息。使用 **GetUpdate** 方法可刷新同步状态。使用 **LastSyncTime** 、 **ErrorType** 和 **WorkspaceLastChangedBy** 属性可返回其他信息。

有关共享演示文稿的本地副本和服务器副本之间可存在的差异和冲突的详细信息，请参阅 **Status** 属性。

使用 **PutUpdate** 方法可将本地更改保存到服务器。未进行任何本地更改时，关闭并重新打开文档可从服务器检索最新版本。使用 **ResolveConflict** 方法可解决本地副本和服务器副本之间的差异，或者使用 **OpenVersion** 方法在当前打开文档的本地版本的情况下再打开其他版本。

**Sync** 对象的 **GetUpdate** 、 **PutUpdate** 和 **ResolveConflict** 方法不返回状态代码，因为它们以异步方式完成任务。 **Sync** 对象通过单个事件，即 **Application** 对象的 **PresentationSync** 事件，提供重要的状态信息。

上述 **Sync** 事件返回一个 **msoSyncEventType** 类型的值。

|                                                             |
|-------------------------------------------------------------|
| MsoSyncEventType 可以是下列 **msoSyncEventType** 常量之一。 |
| **msoSyncEventDownloadInitiated** (0)                       |
| **msoSyncEventDownloadSucceeded** (1)                       |
| **msoSyncEventDownloadFailed** (2)                          |
| **msoSyncEventUploadInitiated** (3)                         |
| **msoSyncEventUploadSucceeded** (4)                         |
| **msoSyncEventUploadFailed** (5)                            |
| **msoSyncEventDownloadNoChange** (6)                        |
| **msoSyncEventOffline** (7)                                 |

无论活动文档上启用或禁用共享和同步，都可使用 **Sync** 对象模型。未共享活动文档或未启用同步时， **Presentation** 对象的 **Sync** 属性不返回 **Nothing** 。使用 **Status** 属性可决定是否共享文档以及是否启用同步。

并不是所有文档同步问题都会导致可捕获的运行时错误。在使用了 **Sync** 对象的方法后，最好对 **Status** 属性进行检查。如果 **Status** 属性为 **msoSyncStatusError** ，请检查 **ErrorType** 属性以获取有关所发生的错误类型的其他信息。

在许多情况下，解决错误条件的最好方法是调用 **GetUpdate** 方法。例如，如果对 **PutUpdate** 的调用导致错误条件的产生，则对 **GetUpdate** 的调用会将状态重置为 **msoSyncStatusLocalChanges** 。

**示例**

``` JavaScript
//下面的示例显示最后一个修改活动演示文稿的人的名称（如果活动演示文稿是文档工作区中的共享文档）。
let eStatus = Application.ActivePresentation.Sync.Status
if(eStatus == msoSyncStatusLatest) {
    strLastUser = Application.ActivePresentation.Sync.WorkspaceLastChangedBy
    alert("You have the most up-to-date copy.This file was last modified by " + strLastUser)
}
```

### Presentation.Tags

返回一个代表指定对象的标签的 **Tags** 对象。只读。

**语法**

**express.Tags**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//以下示例向当前演示文稿的第一张幻灯片中添加名为“REGION”和“PRIORITY”的标签。
function test() {
    let tag = Application.ActivePresentation.Slides.Item(1).Tags
    tag.Add("Region", "East")   //Adds "Region" tag with value "East"
    tag.Add("Priority", "Low")    //Adds "Priority" tag with value "Low"
}
```

### Presentation.TemplateName

返回与指定演示文稿相关的设计模板的名称。只读 **String** 类型。

**语法**

**express.TemplateName**

\*express是一个代表 Presentation 对象的变量。

**说明**

返回的字符串包含 MS-DOS 文件扩展名（对于已注册的文件类型），但不包括完整路径。

**示例**

``` JavaScript
/*以下示例中，如果设计模板“Professional.pot”尚未应用于演示文稿“Pres1.ppt”，则将该模板应用于该演示文稿。*/
function test()
{
    let tion = Application.Presentations.Item("Pres1.ppt")
    if(tion.TemplateName != "Professional.pot") 
    {
        tion.ApplyTemplate("C:\\program files\\office\\templates\\presentation designs\\Professional.pot")
    }
}
```

### Presentation.TitleMaster

返回 Master 对象，该对象代表指定演示文稿的标题母版。如果演示文稿没有标题母版，则会产生一个错误。

**语法**

**express.TitleMaster**

\*express是一个代表 Presentation 对象的变量。

**说明**

可以使用 **AddTitleMaster** 方法为演示文稿添加一个标题母版。

**示例**

``` JavaScript
/*以下示例中，如果当前演示文稿有一个标题母版，则为该标题母版设置页脚文本。*/
function test()
{
    let tion = Application.ActivePresentation
    if(tion.HasTitleMaster) 
    {
        tion.TitleMaster.HeadersFooters.Footer.Text = "Introduction"
    }
}
```

### Presentation.WebOptions

返回一个 **WebOptions** 对象，该对象中包含演示文稿级别的属性，当以网页的方式发布或保存整个或部分演示文稿，或打开某个网页时，WPP 将使用这些属性。只读。

**语法**

**express.WebOptions**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
//本示例的功能：当以网页保存或发布当前演示文稿时，允许使用可移植网络图形 (PNG)，并将大纲窗格的文字颜色设置为白色，将大纲窗格和幻灯片窗格的背景色设置为黑色。
function test() {
    let tion = Application.ActivePresentation.WebOptions
    tion.FrameColors = ppFrameColorsWhiteTextOnBlack
    tion.AllowPNG = true
}
```

### Presentation.Windows

返回一个 **DocumentWindows** 集合，该集合代表与指定演示文稿相关联的所有文档窗口。该属性不返回与该演示文稿相关联的任何幻灯片放映窗口。只读。

**语法**

**express.Windows**

\*express是一个代表 Presentation 对象的变量。

### Presentation.WritePassword

设置或返回一个 **String** 类型的值，该值代表保存对指定文档所做的更改时使用的密码。可读/写。

**语法**

**express.WritePassword**

\*express是一个代表 Presentation 对象的变量。

**示例**

``` JavaScript
/*本示例设置保存对活动演示文稿所做的更改时使用的密码。*/
Application.ActivePresentation.WritePassword = "complexstrPWD"
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Presentation](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SlideRange 对象

## 简介

代表备注页或幻灯片范围的集合，该范围是一组幻灯片，少则仅包含一个，多则包含演示文稿中的所有幻灯片。

## 说明

您可以使用任何所需的幻灯片（从演示文稿的所有幻灯片中进行选择，或从选定内容的所有幻灯片中进行选择）构造幻灯片范围。例如，可以构造一个 SlideRange 集合，其中包含演示文稿中的前三张幻灯片、演示文稿中所有选定的幻灯片，或演示文稿中所有的标题幻灯片。

如同在用户界面中选中多个幻灯片并通过命令同时操作它们一样，通过建立一个 SlideRange 集合并对其使用属性和方法，可以在编程中同时操作多个幻灯片。如同用户界面中用于单张幻灯片的命令不能用于多张幻灯片一样，某些应用于单独 SlideRange 对象或只包含一张幻灯片的 SlideRange 集合的属性和方法不能用于包含多张幻灯片的 SlideRange 集合。一般情况下，如果选中多张幻灯片时，某些操作无法手动完成（例如返回某一幻灯片中的单个形状），则编程时也不能对包含多张幻灯片的 SlideRange 集合进行该操作。

对于用户界面中可用于一张或多张选中幻灯片的操作（例如复制幻灯片到剪贴板或设置幻灯片背景填充），相应的属性和方法也可用于包含多张幻灯片的 SlideRange 集合。下面是如何对多张幻灯片使用这些属性和方法的一些指导。

- 对 SlideRange 集合应用某方法等效于将该范围中所有 Slide 对象作为一个组来应用此方法。
- 设置 SlideRange 集合的属性值等效于逐一设置该范围中每个幻灯片的属性值（对于枚举类型的属性，设置“Mixed”值无效）。
- 如果 SlideRange 集合中的所有幻灯片对于某个返回枚举类型的属性具有相同的值，则该属性将返回该集合中单张幻灯片的属性值。如果该集合中的幻灯片对于该属性具有不同的值，则该属性将返回“Mixed”值。
- 如果 SlideRange 集合中的所有幻灯片对于某个返回简单数据类型（如 Long 、 Single 或 String ）的属性具有相同的值，则该属性将返回该集合中单张幻灯片的该属性值。如果该集合中的幻灯片对于该属性具有不同的值，则该属性将返回 -2 或产生错误。例如，对于包含多张幻灯片的 SlideRange 对象使用 Name 属性将产生错误，因为每张幻灯片的 Name 属性都有不同的值。
- 幻灯片的某些格式属性不是通过直接应用于 SlideRange 集合的属性和方法来设置，而是通过应用于 SlideRange 集合中的对象（如 ColorScheme 对象）的属性和方法来设置。如果所含对象代表用户界面中可用于多个对象的操作，则可以从包含多张幻灯片的 SlideRange 集合返回该对象，且其属性和方法遵循前述规则。例如，可以使用 ColorScheme 属性返回 ColorScheme 对象，该对象代表用于指定 SlideRange 集合中所有幻灯片的配色方案。设置此 ColorScheme 对象的属性也将设置 SlideRange 集合中所有单张幻灯片的 ColorScheme 对象的相应属性。

以下示例说明如何执行下列操作：

- 返回一组通过名称或索引号指定的幻灯片
- 返回演示文稿中全部或部分选定的幻灯片
- 返回备注页
- 将属性和方法应用于某个幻灯片范围

备注：

对 SlideRange 对象代表的备注页使用下列属性和方法无效： Copy 方法、 Cut 方法、 Delete 方法、 Duplicate 方法、 HeadersFooters 属性、 Hyperlinks 属性、 Layout 属性、 PrintSteps 属性、 SlideShowTransition 属性。

## 方法

| 序号 | 名称                                                       | 说明                                                                                                                                                                         |
|:----:|:-----------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ApplyTemplate](#SlideRange.ApplyTemplate)                 | 将设计模板应用于指定的演示文稿。                                                                                                                                             |
|  2   | [ApplyTheme](#SlideRange.ApplyTheme)                       | 对指定的幻灯片范围应用主题或设计模板。                                                                                                                                       |
|  3   | [ApplyThemeColorScheme](#SlideRange.ApplyThemeColorScheme) | 对指定的幻灯片范围应用配色方案。                                                                                                                                             |
|  4   | [Copy](#SlideRange.Copy)                                   | 将指定对象复制到剪贴板。                                                                                                                                                     |
|  5   | [Cut](#SlideRange.Cut)                                     | 删除指定对象并将其放到剪贴板中。                                                                                                                                             |
|  6   | [Delete](#SlideRange.Delete)                               | 删除指定的对象。                                                                                                                                                             |
|  7   | [Duplicate](#SlideRange.Duplicate)                         | 创建指定的 Slide 或 SlideRange 对象的副本，将新的幻灯片或幻灯片范围添加到 Slides 集合中最初指定的幻灯片或幻灯片范围之后，然后返回代表复制幻灯片的 Slide 或 SlideRange 对象。 |
|  8   | [Export](#SlideRange.Export)                               | 使用指定的图形筛选器导出幻灯片或幻灯片范围，并将导出的文件以指定的文件名保存。                                                                                               |
|  9   | [Item](#SlideRange.Item)                                   | 从指定集合中返回单个对象。                                                                                                                                                   |
|  10  | [MoveTo](#SlideRange.MoveTo)                               | 将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。                                                                                             |
|  11  | [Select](#SlideRange.Select)                               | 选择指定的对象。                                                                                                                                                             |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                                                                                                                                                                                                                                                                                   |
|:----:|:-------------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SlideRange.Application)                 | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                                                                                                                                |
|  2   | [Background](#SlideRange.Background)                   | 返回代表幻灯片背景的 ShapeRange 对象。                                                                                                                                                                                                                                                                                                                                                 |
|  3   | [ColorScheme](#SlideRange.ColorScheme)                 | 返回或设置 ColorScheme 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。                                                                                                                                                                                                                                                                                         |
|  4   | [Comments](#SlideRange.Comments)                       | 返回一个 Comments 对象，该对象代表批注的集合。                                                                                                                                                                                                                                                                                                                                         |
|  5   | [Count](#SlideRange.Count)                             | 返回指定集合中的对象数目。                                                                                                                                                                                                                                                                                                                                                             |
|  6   | [Design](#SlideRange.Design)                           | 返回代表设计的 Design 对象。                                                                                                                                                                                                                                                                                                                                                           |
|  7   | [HeadersFooters](#SlideRange.HeadersFooters)           | 返回一个 HeadersFooters 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。                                                                                                                                                                                                                                                           |
|  8   | [Hyperlinks](#SlideRange.Hyperlinks)                   | 返回一个代表指定幻灯片中的所有超链接的 Hyperlinks 集合。只读。                                                                                                                                                                                                                                                                                                                         |
|  9   | [Layout](#SlideRange.Layout)                           | 返回或设置一个 PpSlideLayout 常量，该常量代表幻灯片版式。可读/写。                                                                                                                                                                                                                                                                                                                     |
|  10  | [Master](#SlideRange.Master)                           | 返回一个 Master 对象，该对象代表幻灯片母版。只读。                                                                                                                                                                                                                                                                                                                                     |
|  11  | [Name](#SlideRange.Name)                               | 向一个演示文稿插入幻灯片时，Powerpoint 自动为幻灯片分配 Sliden 形式的名称，其中 n 是一个整数，代表幻灯片在演示文稿中创建的顺序。例如，在某个演示文稿中插入的第一张幻灯片被自动命名为 Slide1。如果将幻灯片从一个演示文稿复制到另一个演示文稿，则该幻灯片会失去它在第一个演示文稿中的名称并在第二个演示文稿中被自动分配新名称。幻灯片范围必须至少包含一张幻灯片。 String 类型，可读/写。 |
|  12  | [NotesPage](#SlideRange.NotesPage)                     | 返回一个 SlideRange 对象，该对象代表指定幻灯片或幻灯片区域的备注页。只读。                                                                                                                                                                                                                                                                                                             |
|  13  | [Shapes](#SlideRange.Shapes)                           | 返回一个 Shapes 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。                                                                                                                                     |
|  14  | [SlideID](#SlideRange.SlideID)                         | 返回指定幻灯片的唯一 ID 号。只读。 Long 类型。                                                                                                                                                                                                                                                                                                                                         |
|  15  | [SlideIndex](#SlideRange.SlideIndex)                   | 返回 Slides 集合内指定幻灯片的索引号。只读。 Long 类型。                                                                                                                                                                                                                                                                                                                               |
|  16  | [SlideNumber](#SlideRange.SlideNumber)                 | 返回幻灯片编号。只读 Integer 类型。                                                                                                                                                                                                                                                                                                                                                    |
|  17  | [SlideShowTransition](#SlideRange.SlideShowTransition) | 返回一个 SlideShowTransition 对象，该对象代表指定幻灯片切换的特殊效果。只读。                                                                                                                                                                                                                                                                                                          |
|  18  | [ThemeColorScheme](#SlideRange.ThemeColorScheme)       | 返回一个 ThemeColorScheme 对象，该对象代表与指定幻灯片区域相关联的配色方案。                                                                                                                                                                                                                                                                                                           |
|  19  | [TimeLine](#SlideRange.TimeLine)                       | 返回 TimeLine 对象，该对象代表幻灯片的动画时间线。                                                                                                                                                                                                                                                                                                                                     |

## 成员方法

### SlideRange.ApplyTemplate

将设计模板应用于指定的演示文稿。

**语法**

**express.ApplyTemplate ( *FileName* )**

\*express是一个代表 SlideRange 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                 |
|------------|-----------|----------|----------------------|
| *FileName* | 必选      | String   | 指定设计模板的名称。 |

**说明**

如果在字符串中引用未安装的演示文稿设计模板，则将产生运行时错误。无论 FeatureInstall 属性设置是什么，该模板都不自动安装。要对当前未安装的模板使用 ApplyTemplate 方法，必须先安装附加设计模板。为此，请通过运行 WPS Office安装程序（单击 Windows 控制面板中的“添加/删除程序”图标）安装WPP 的附加设计模板。

**示例**

``` JavaScript
/*本示例将“Professional”设计模板应用于当前演示文稿*/
Application.ActivePresentation.ApplyTemplate("c:\\program files\\office\\templates" + "\\presentation designs\\professional.pot")
```

### SlideRange.ApplyTheme

对指定的幻灯片范围应用主题或设计模板。

**语法**

**express.ApplyTheme ( *themeName* )**

\*express是一个代表 SlideRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                          |
|-------------|-----------|----------|-------------------------------------------------------------------------------|
| *themeName* | 必选      | String   | 应用于 SlideRange 对象的主题文件 (.thmx) 或设计模板文件 (.pot) 的路径和名称。 |

**示例**

``` JavaScript
/*以下示例对指定的幻灯片范围应用保存的主题*/
Application.ActivePresentation.Slides.Range([1, 3]).ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.thmx")

/*以下示例对指定的幻灯片范围应用保存的设计模板*/
Application.ActivePresentation.Slides.Range([1, 3]).ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.pot")
```

### SlideRange.ApplyThemeColorScheme

对指定的幻灯片范围应用配色方案。

**语法**

**express.ApplyThemeColorScheme ( *themeColorSchemeName* )**

\*express是一个代表 SlideRange 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型 | 说明                                         |
|------------------------|-----------|----------|----------------------------------------------|
| *themeColorSchemeName* | 必选      | String   | 应用于幻灯片范围的配色方案文件的路径和名称。 |

### SlideRange.Copy

将指定对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 SlideRange 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

``` JavaScript
/*本示例在第一个窗口中拷贝选定幻灯片到剪贴板*/
Application.ActivePresentation.Windows.Item(1).Selection.Copy()

/*本示例将当前演示文稿中第一张幻灯片的第一个和第二个形状复制到剪贴板，然后将该副本粘贴到第二张灯片*/
function test(){
    let slides2 = Application.ActivePresentation
    slides2.Slides.Item(1).Shapes.Range([1, 2]).Copy()
    slides2.Slides.Item(2).Shapes.Paste()
}

/*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()
```

### SlideRange.Cut

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut ()**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中*/
Application.ActivePresentation.Windows.Item(1).Selection.Cut()

/*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中*/
Application.ActivePresentation.Slides.Item(1).Cut()

/*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut()
```

### SlideRange.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 SlideRange 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**示例**

``` JavaScript
/*此示例删除选中的图形*/
Application.ActivePresentation.Windows.Item(1).Selection.Delete()
```

### SlideRange.Duplicate

创建指定的 **Slide** 或 **SlideRange** 对象的副本，将新的幻灯片或幻灯片范围添加到 **Slides** 集合中最初指定的幻灯片或幻灯片范围之后，然后返回代表复制幻灯片的 **Slide** 或 **SlideRange** 对象。

**语法**

**express.Duplicate ()**

\*express是一个代表 SlideRange 对象的变量。

### SlideRange.Export

使用指定的图形筛选器导出幻灯片或幻灯片范围，并将导出的文件以指定的文件名保存。

**语法**

**express.Export ( *FileName* , *FilerName* , *ScaleWidth* , *ScaleHeight* )**

\*express是一个代表 SlideRange 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                        |
|---------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName*    | 必选      | String   | 将导出并保存到磁盘的文件的名称。可以包括完整路径；如果不包括完整路径，WPP 就会在当前文件夹中创建一个文件。                                                                                                                  |
| *FilerName*   | 必选      | String   | 要导出幻灯片的图形格式。指定的图形格式必须在 Windows 注册表中注册导出筛选器。可以指定注册的扩展名或注册的筛选器名称。WPP将首先在注册表中搜索匹配的扩展名。如果没有找到与指定字符串匹配的扩展名，WPP将查找匹配的筛选器名称。 |
| *ScaleWidth*  | 可选      | Long     | 导出幻灯片的宽度（以像素为单位）。                                                                                                                                                                                          |
| *ScaleHeight* | 可选      | Long     | 导出幻灯片的高度（以像素为单位）。                                                                                                                                                                                          |

**说明**

导出一个演示文稿不会将演示文稿的 Saved 属性设置为 True。 WPP使用指定的图形筛选器保存每张单独的幻灯片。导出并保存到磁盘的幻灯片的名称由WPP 决定。通常情况下，这些文件使用下列命名规范保存：Slide1.wmf、Slide2.wmf。保存文件的路径在 FileName 参数中指定。

### SlideRange.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SlideRange 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

**示例**

``` JavaScript
/*此示例返回当前窗口的第二张幻灯片*/
Application.ActivePresentation.Slides.Item(2)
```

### SlideRange.MoveTo

将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。

**语法**

**express.MoveTo ( *toPos* )**

\*express是一个代表 SlideRange 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                     |
|---------|-----------|----------|--------------------------|
| *toPos* | 必选      | Long     | 动画效果要移动到的索引。 |

**示例**

``` JavaScript
/*本示例在指定形状的动画效果集合中将一个动画效果移动到第二个动画效果的位置上*/
function test() {
    let sldFirst = Application.ActivePresentation.Slides.Item(1)
    let shpFirst = sldFirst.Shapes.Item(1)
    let effAdd = sldFirst.TimeLine.MainSequence.AddEffect(shpFirst, msoAnimEffectBlinds)
    effAdd.MoveTo(2)
}

/*本示例将当前演示文稿中的第二张幻灯片移动到第一张幻灯片的位置上*/
Application.ActivePresentation.Slides.Item(2).MoveTo(1)
```

### SlideRange.Select

选择指定的对象。

**语法**

**express.Select ()**

\*express是一个代表 SlideRange 对象的变量。

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片。

**示例**

``` JavaScript
/*本示例选择活动演示文稿中第一张幻灯片标题的前五个字符*/
Application.ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.Characters(1, 5).Select()

/*本示例选择活动演示文稿中的第一张幻灯片*/
Application.ActivePresentation.Slides.Item(1).Select()

/*本示例选取一个已添加到演示文稿的新幻灯片中的表格。该表格具有三行和三列*/
Application.Presentations.Add().Slides.Add(1, ppLayoutBlank).Shapes.AddTable(3, 3).Select()
```

## 成员属性

### SlideRange.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/
function test(pptPres) {
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}


/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/
function test(){
    for(let shpOle = 1; shpOle <= Application.ActivePresentation.Slides.Item(1).Shapes.Count; shpOle++) {
        if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(shpOle).Type == msoLinkedOLEObject) {
            MsgBox(Application.ActivePresentation.Slides.Item(1).Shapes.Item(shpOle).OLEFormat.Application.Name)
        }
    }
}
```

### SlideRange.Background

返回代表幻灯片背景的 ShapeRange 对象。

**语法**

**express.Background**

\*express是一个代表 SlideRange 对象的变量。

**说明**

如果要使用 Background 属性设置单张幻灯片的背景而不更改幻灯片母版，则该幻灯片的 FollowMasterBackground 属性必须设置为 False。

**示例**

``` JavaScript
/*以下示例将当前演示文稿的幻灯片母版的背景设为预设的底纹*/
Application.ActivePresentation.SlideMaster.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset)


/*以下示例将当前演示文稿中第一张幻灯片的背景设为预设的底纹*/
function test(){
    let slides2 =Application.ActivePresentation.Slides.Item(1)
    slides2.FollowMasterBackground = false
    slides2.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset)
}
```

### SlideRange.ColorScheme

返回或设置 ColorScheme 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。

**语法**

**express.ColorScheme**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例将当前演示文稿第一张和第三张幻灯片的标题颜色设为绿色*/
function test(){
    let mySlides = Application.ActivePresentation.Slides.Range([1, 3])
    mySlides.ColorScheme.Colors(ppTitle).RGB = (0, 255, 0)
}
```

### SlideRange.Comments

返回一个 Comments 对象，该对象代表批注的集合。

**语法**

**express.Comments**

\*express是一个代表 SlideRange 对象的变量。

### SlideRange.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口*/
function test(){
    let windows = Application.Windows
    for(let i = 2; i <= windows2.Count; i++) {
        Application.windows.Item(2).Close()
    }
}
```

### SlideRange.Design

返回代表设计的 Design 对象。

**语法**

**express.Design**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例返回当前所在演示文档名*/
Application.ActivePresentation.Slides.Item(1).Design.Application.ActiveWindow
```

### SlideRange.HeadersFooters

返回一个 HeadersFooters 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。

**语法**

**express.HeadersFooters**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例在当前演示文稿的备注母版中设置页脚文本*/
function test(){
    let notesMaster = Application.ActivePresentation.NotesMaster.HeadersFooters
        notesMaster.Footer.Text = "Regional Sales"
}
```

### SlideRange.Hyperlinks

返回一个代表指定幻灯片中的所有超链接的 Hyperlinks 集合。只读。

**语法**

**express.Hyperlinks**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例允许用户更新当前演示文稿中所有超链接中已过时的 Internet 地址*/
function test(){
    let oldAddr =("Old Internet address")
    let newAddr = ("New Internet address")

    for(let s = 1; s <= Application.ActivePresentation.Slides.Count; s++) {
        for(let h = 1; h <= Application.ActivePresentation.Slides.Item(s).Hyperlinks.Count; h++) {
            if(Application.ActivePresentation.Slides.Item(s).Hyperlinks.Item(h).Address.toLowerCase() == oldAddr) {
                Application.ActivePresentation.Slides.Item(s).Hyperlinks.Item(h).Address = newAddr
            }
        }
    }
}
```

### SlideRange.Layout

返回或设置一个 **PpSlideLayout** 常量，该常量代表幻灯片版式。可读/写。

**语法**

**express.Layout**

\*express是一个代表 SlideRange 对象的变量。

**说明**

PpSlideLayout 可以是下列 PpSlideLayout 常量之一。 ppLayoutBlank ppLayoutChart ppLayoutChartAndText ppLayoutClipartAndText ppLayoutClipArtAndVerticalText ppLayoutFourObjects ppLayoutLargeObject ppLayoutMediaClipAndText ppLayoutMixed ppLayoutObject ppLayoutObjectAndText ppLayoutObjectOverText ppLayoutOrgchart ppLayoutTable ppLayoutText ppLayoutTextAndChart ppLayoutTextAndClipart ppLayoutTextAndMediaClip ppLayoutTextAndObject ppLayoutTextAndTwoObjects ppLayoutTextOverObject ppLayoutTitle ppLayoutTitleOnly ppLayoutTwoColumnText ppLayoutTwoObjectsAndText ppLayoutTwoObjectsOverText ppLayoutVerticalText ppLayoutVerticalTitleAndText ppLayoutVerticalTitleAndTextOverChart

### SlideRange.Master

返回一个 Master 对象，该对象代表幻灯片母版。只读。

**语法**

**express.Master**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例为当前演示文稿中第一张幻灯片的幻灯片母版设置背景填充*/
Application.ActivePresentation.Slides.Item(1).Master.Background.Fill.PresetGradient(msoGradientDiagonalUp, 1, msoGradientDaybreak)
```

### SlideRange.Name

向一个演示文稿插入幻灯片时，Powerpoint 自动为幻灯片分配 Sliden 形式的名称，其中 n 是一个整数，代表幻灯片在演示文稿中创建的顺序。例如，在某个演示文稿中插入的第一张幻灯片被自动命名为 Slide1。如果将幻灯片从一个演示文稿复制到另一个演示文稿，则该幻灯片会失去它在第一个演示文稿中的名称并在第二个演示文稿中被自动分配新名称。幻灯片范围必须至少包含一张幻灯片。 **String** 类型，可读/写。

**语法**

**express.Name**

\*express是一个代表 SlideRange 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则 .Shapes.Item("Rectangle 2") 将返回对该形状的引用。

### SlideRange.NotesPage

返回一个 SlideRange 对象，该对象代表指定幻灯片或幻灯片区域的备注页。只读。

**语法**

**express.NotesPage**

\*express是一个代表 SlideRange 对象的变量。

**说明**

NotesPage 属性返回一张或一组幻灯片的备注页，且只允许对这些备注页进行修改。如果要进行会影响所有备注页的更改，请使用 NotesMaster 返回代表备注母版的 Slide 对象。

**示例**

``` JavaScript
/*以下示例为当前演示文稿中第一张幻灯片的备注页设置背景填充*/
function test(){
    let myNotePage = Application.ActivePresentation.Slides.Item(1).NotesPage
    myNotePage.FollowMasterBackground = false
    myNotePage.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset)
}
```

### SlideRange.Shapes

返回一个 Shapes 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。

**语法**

**express.Shapes**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例添加一个宽 100 磅、高 50 磅的矩形，它的左上角距当前演示文稿第一张幻灯片的左边 5 磅、上边 25 磅*/
function test(){
    let firstSlide = Application.ActivePresentation.Slides.Item(1)
    firstSlide.Shapes.AddShape(msoShapeRectangle, 5, 25, 100, 50)
}

/*以下示例设置当前演示文稿第一张幻灯片的第三个形状的填充纹理*/
function test(){
    let newRect = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
    newRect.Fill.PresetTextured(msoTextureOak)
}

/*假设当前演示文稿第一张幻灯片包含一个标题，以下示例的第二行和第三行设置该演示文稿第一张幻灯片的标题文本*/
function test(){
    let newRect = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
    newRect.Fill.PresetTextured(msoTextureOak)
}

/*假设当前演示文稿第二张幻灯片中的第二个形状包含文本框架，以下示例向该幻灯片添加一系列的段落。请注意：Chr(13) 用于在该文本中插入段落标记*/
function test(){
    let tShape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(2)
    tShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item"
}

/*对于大多数幻灯片版式，第一个形状为文本占位符。以下示例与上例完成相同功能*/
function test(){
    let testShape = Application.ActivePresentation.Slides.Item(2).Shapes.Placeholders.Item(2)
    testShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item"
}
```

### SlideRange.SlideID

返回指定幻灯片的唯一 ID 号。只读。 **Long** 类型。

**语法**

**express.SlideID**

\*express是一个代表 SlideRange 对象的变量。

**说明**

与 SlideIndex 属性不同，在演示文稿中添加或重新排列幻灯片时，Slide 对象的 SlideID 属性不会改变。因此，与使用具有幻灯片索引号的 Item 方法相比，使用具有幻灯片 ID 号的 FindBySlideID 方法是从 Slides 集合返回特定 Slide 对象的更可靠方法。

**示例**

``` JavaScript
/*以下示例示范添加并 Slide 检索 Slide 对象的唯一 ID 号*/
function test(){
    let gslides = Application.ActivePresentation.Slides
    // Get slide ID
    let graphSlideID = gslides.Add(2, ppLayoutChart).SlideID
}
```

### SlideRange.SlideIndex

返回 **Slides** 集合内指定幻灯片的索引号。只读。 **Long** 类型。

**语法**

**express.SlideIndex**

\*express是一个代表 SlideRange 对象的变量。

**说明**

与 SlideID 属性不同，在演示文稿中添加或重新排列幻灯片时，Slide 对象的 SlideIndex 属性会改变。因此，与使用具有幻灯片索引号的 Item 方法相比，使用具有幻灯片 ID 号的 FindBySlideID 方法是从 Slides 集合返回特定 Slide 对象的更可靠方法。

**示例**

``` JavaScript
/*以下示例显示第一个幻灯片放映窗口中当前放映幻灯片的索引号*/
Application.ActivePresentation.SlideShowWindow.View.Slide.SlideIndex
```

### SlideRange.SlideNumber

返回幻灯片编号。只读 Integer 类型。

**语法**

**express.SlideNumber**

\*express是一个代表 SlideRange 对象的变量。

**说明**

Slide 对象的 SlideNumber 属性是显示幻灯片编号时在幻灯片右下角出现的实际编号。此编号由幻灯片在演示文稿中的编号（ SlideIndex 属性值）以及演示文稿的起始幻灯片编号（FirstSlideNumber 属性值）决定。幻灯片编号始终等于起始幻灯片编号 + 幻灯片索引号 - 1。

**示例**

``` JavaScript
/*本示例显示第一张幻灯片编号的变化对某特定幻灯片编号的影响*/
function test(){
    let mySlide = Application.ActivePresentation
    mySlide.PageSetup.FirstSlideNumber = 1     // starts slide numbering at 1
    MsgBox(mySlide.Slides.Item(2).SlideNumber) // returns 2

    mySlide.PageSetup.FirstSlideNumber = 10    // starts slide numbering at 10
    MsgBox(mySlide.Slides.Item(2).SlideNumber) // returns 11
}
```

### SlideRange.SlideShowTransition

返回一个 SlideShowTransition 对象，该对象代表指定幻灯片切换的特殊效果。只读。

**语法**

**express.SlideShowTransition**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例设置在幻灯片放映过程中，当前演示文稿的第二张幻灯片在 5 秒后自动换片，并在幻灯片切换时播放犬吠声*/
function test(){
    let transitionSet = Application.ActivePresentation.Slides.Item(2).SlideShowTransition
    transitionSet.AdvanceOnTime = true
    transitionSet.AdvanceTime = 5
    transitionSet.SoundEffect.ImportFromFile("C:\\windows\\media\\dogbark.wav")

    ActivePresentation.SlideShowSettings.AdvanceMode = ppSlideShowUseSlideTimings
}
```

### SlideRange.ThemeColorScheme

返回一个 ThemeColorScheme 对象，该对象代表与指定幻灯片区域相关联的配色方案。

**语法**

**express.ThemeColorScheme**

\*express是一个代表 SlideRange 对象的变量。

### SlideRange.TimeLine

返回 TimeLine 对象，该对象代表幻灯片的动画时间线。

**语法**

**express.TimeLine**

\*express是一个代表 SlideRange 对象的变量。

**示例**

``` JavaScript
/*以下示例将一个跳动的动画添加到第一张幻灯片的第一个形状中*/
function test(){
    let sldFirst = Application.ActivePresentation.Slides.Item(1)
    let shpFirst = sldFirst.Shapes.Item(1)

    sldFirst.TimeLine.MainSequence.AddEffect(shpFirst, msoAnimEffectBounce)
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/SlideRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SlideShowTransition 对象

## 简介

包含关于幻灯片放映中指定幻灯片的切换方式的信息。

## 属性

| 序号 | 名称                                                          | 说明                                                                                                                                                                                     |
|:----:|:--------------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AdvanceOnClick](#SlideShowTransition.AdvanceOnClick)         | 决定在幻灯片放映期间单击指定的幻灯片后是否会切换。可读/写 MsoTriState 类型。                                                                                                             |
|  2   | [AdvanceOnTime](#SlideShowTransition.AdvanceOnTime)           | 决定指定的幻灯片在经过指定时间后是否自动切换。可读/写 MsoTriState 类型。                                                                                                                 |
|  3   | [AdvanceTime](#SlideShowTransition.AdvanceTime)               | 返回或设置以秒为单位的时间长度，该段时间过后，指定的幻灯片将会切换。可读/写。                                                                                                            |
|  4   | [Application](#SlideShowTransition.Application)               | 返回一个 Application 对象，该对象表示指定对象的创建者                                                                                                                                    |
|  5   | [EntryEffect](#SlideShowTransition.EntryEffect)               | 对于 AnimationSettings 对象，此属性返回或设置应用于指定形状中动画的特殊效果。对于 SlideShowTransition 对象，此属性返回或设置应用于指定幻灯片切换的特殊效果。PpEntryEffect 类型，可读写。 |
|  6   | [Hidden](#SlideShowTransition.Hidden)                         | 决定幻灯片放映过程中指定的幻灯片是否隐藏。可读/写 MsoTriState 类型。                                                                                                                     |
|  7   | [LoopSoundUntilNext](#SlideShowTransition.LoopSoundUntilNext) | 指定是否循环播放为指定的幻灯片切换所设置的声音，直到下一个声音开始。可读/写 MsoTriState 类型。                                                                                           |
|  8   | [Parent](#SlideShowTransition.Parent)                         | 返回指定对象的父对象。                                                                                                                                                                   |
|  9   | [SoundEffect](#SlideShowTransition.SoundEffect)               | 返回一个 SoundEffect 对象，该对象代表切换到指定幻灯片时播放的声音。                                                                                                                      |
|  10  | [Speed](#SlideShowTransition.Speed)                           | 返回或设置一个 PpTransitionSpeed 常量，该常量代表切换到指定幻灯片的速度。可读/写。                                                                                                       |

## 成员属性

### SlideShowTransition.AdvanceOnClick

决定在幻灯片放映期间单击指定的幻灯片后是否会切换。可读/写 **MsoTriState** 类型。

**语法**

**express.AdvanceOnClick**

\*express是一个代表 SlideShowTransition 对象的变量。

**说明**

要设置幻灯片在某段时间后自动切换，请将 **AdvanceOnTime** 属性设置为 **True** ，并将 **AdvanceTime** 属性设置为希望幻灯片放映的时间。如果同时将 **AdvanceOnClick** 和 **AdvanceOnTime** 属性设置为 **True** ，则幻灯片将在被单击时或经过指定的时间后（无论哪个先发生）进行切换。

| MsoTriState 可以是下列 MsoTriState 常量之一。            |
|----------------------------------------------------------|
| msoCTrue                                                 |
| msoFalse                                                 |
| msoTriStateMixed                                         |
| msoTriStateToggle                                        |
| msoTrue 在幻灯片放映过程中，单击指定的幻灯片后进行切换。 |

**示例**

``` JavaScript
/*本示例设置活动演示文稿中的第一张幻灯片在五秒钟过去后前进，或在单击鼠标时前进（具体取决于哪个事件首先发生）。*/
function test(){
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnClick = msoTrue
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnTime = msoTrue
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceTime = 5
}
```

### SlideShowTransition.AdvanceOnTime

决定指定的幻灯片在经过指定时间后是否自动切换。可读/写 **MsoTriState** 类型。

**语法**

**express.AdvanceOnTime**

\*express是一个代表 SlideShowTransition 对象的变量。

**说明**

使用 **AdvanceTime** 属性可以指定幻灯片自动切换前需等待的秒数。将 **SlideShowSettings** 对象的 **AdvanceMode** 属性设置为 **ppSlideShowUseSlideTimings** 可使幻灯片的间隔设置对整个幻灯片放映有效。

| MsoTriState 可以是下列 MsoTriState 常量之一。  |
|------------------------------------------------|
| msoCTrue                                       |
| msoFalse                                       |
| msoTriStateMixed                               |
| msoTriStateToggle                              |
| msoTrue 指定的幻灯片在经过指定时间后自动切换。 |

**示例**

``` JavaScript
/*本示例设置活动演示文稿中的第一张幻灯片在五秒钟过去后前进，或在单击鼠标时前进（具体取决于哪个事件首先发生）。*/
function test(){
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnClick = msoTrue
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnTime = msoTrue
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceTime = 5
}
```

### SlideShowTransition.AdvanceTime

返回或设置以秒为单位的时间长度，该段时间过后，指定的幻灯片将会切换。可读/写。

**语法**

**express.AdvanceTime**

\*express是一个代表 SlideShowTransition 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动演示文稿中的第一张幻灯片在五秒钟过去后前进，或在单击鼠标时前进（具体取决于哪个事件首先发生）。*/
function test(){
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnClick = msoTrue
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnTime = msoTrue
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceTime = 5
}
```

### SlideShowTransition.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者

**语法**

**express.Application**

\*express是一个代表 SlideShowTransition 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test(){
　　　　let shpOle = ActivePresentation.Slides.Item(1).Shapes
       for(let i = 1; i <= shpOle.Count; i++) {
   　　　　 if(shpOle.Item(i).Type == msoLinkedOLEObject) {
   　　　　     MsgBox(shpOle.Item(i).OLEFormat.Application.Name)
 　　　　   }
　　　　}}
```

### SlideShowTransition.EntryEffect

对于 **AnimationSettings** 对象，此属性返回或设置应用于指定形状中动画的特殊效果。对于 **SlideShowTransition** 对象，此属性返回或设置应用于指定幻灯片切换的特殊效果。PpEntryEffect 类型，可读写。

**语法**

**express.EntryEffect**

\*express是一个代表 SlideShowTransition 对象的变量。

**说明**

如果将指定形状的 **TextLevelEffect** 属性设为 **ppAnimateLevelNone** （默认值）或将 **Animate** 属性设为 **False** ，您将看不到通过 **EntryEffect** 属性应用的特殊效果。

**示例**

``` JavaScript
/*本示例在当前演示文稿中添加一张标题幻灯片，并设置该标题在幻灯片放映中动画时从右侧飞入。*/
function test(){
　　　　let PreTit = ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes.Item(1)
   　　　　 PreTit.TextFrame.TextRange.Text = "Sample title"
   　　　　 PreTit.AnimationSettings.TextLevelEffect = ppAnimateByAllLevels
   　　　　 PreTit.AnimationSettings.EntryEffect = ppEffectFlyFromRight
   　　　　 PreTit.AnimationSettings.Animate = true
}
```

### SlideShowTransition.Hidden

决定幻灯片放映过程中指定的幻灯片是否隐藏。可读/写 **MsoTriState** 类型。

**语法**

**express.Hidden**

\*express是一个代表 SlideShowTransition 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue 幻灯片放映过程中隐藏指定的幻灯片。    |

**示例**

本示例将活动演示文稿中的第二张幻灯片设为隐藏。

``` JavaScript
ActivePresentation.Slides.Item(2).SlideShowTransition.Hidden = msoTrue
```

### SlideShowTransition.LoopSoundUntilNext

指定是否循环播放为指定的幻灯片切换所设置的声音，直到下一个声音开始。可读/写 **MsoTriState** 类型。

**语法**

**express.LoopSoundUntilNext**

\*express是一个代表 SlideShowTransition 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                        |
|----------------------------------------------------------------------|
| msoCTrue                                                             |
| msoFalse                                                             |
| msoTriStateMixed                                                     |
| msoTriStateToggle                                                    |
| msoTrue 循环播放为指定的幻灯片切换所设置的声音，直到下一个声音开始。 |

**示例**

``` JavaScript
/*本示例指定转换到当前演示文稿第二张幻灯片时开始播放文件 Dudududu.wav，并持续播放直到开始下一个声音文件。*/
function test(){
　　　　ActivePresentation.Slides.Item(2).SlideShowTransition.SoundEffect.ImportFromFile("C:\\sndsys\\dudududu.wav")
　　　　ActivePresentation.Slides.Item(2).SlideShowTransition.LoopSoundUntilNext = msoTrue
}
```

### SlideShowTransition.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 SlideShowTransition 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
　　　　myShapes.TextRange.Text = "Test text"
　　　　myShapes.Parent.Rotation = 45
}
```

### SlideShowTransition.SoundEffect

返回一个 **SoundEffect** 对象，该对象代表切换到指定幻灯片时播放的声音。

**语法**

**express.SoundEffect**

\*express是一个代表 SlideShowTransition 对象的变量。

**示例**

``` JavaScript
/*以下示例设置在当前演示文稿第一张幻灯片的第一个形状演示动画时播放文件 Bass.wav。*/
function test(){
　　　　let ShapeSet = ActivePresentation.Slides.Item(1).Shapes.Item(1).AnimationSettings
　　　　    ShapeSet.Animate = true
　　　　    ShapeSet.TextLevelEffect = ppAnimateByAllLevels
　　　　    ShapeSet.SoundEffect.ImportFromFile("C:\\bass.wav")
}
```

### SlideShowTransition.Speed

返回或设置一个 **PpTransitionSpeed** 常量，该常量代表切换到指定幻灯片的速度。可读/写。

**语法**

**express.Speed**

\*express是一个代表 SlideShowTransition 对象的变量。

**说明**

| PpTransitionSpeed 可以是下列 PpTransitionSpeed 常量之一。 |
|-----------------------------------------------------------|
| ppTransitionSpeedFast                                     |
| ppTransitionSpeedMedium                                   |
| ppTransitionSpeedMixed                                    |
| ppTransitionSpeedSlow                                     |

**示例**

``` JavaScript
/*本示例设置切换到当前演示文稿第一张幻灯片的特殊效果，并指定切换为快速。*/
function test(){
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.EntryEffect = ppEffectStripsDownLeft
　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.Speed = ppTransitionSpeedFast
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/SlideShowTransition](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CommandBarPopup 对象

## 简介

代表命令栏上的一个弹出式控件。

## 说明

每个弹出式控件都包含一个 CommandBar 对象。要从弹出式控件返回命令栏，请对 CommandBar 对象应用 CommandBarPopup 属性。

使用 Controls(index) 可返回一个 CommandBarPopup 对象，其中 *index* 是该控件的索引号。请注意，该控件的 Type 属性必须是 msoControlPopup 、 msoControlGraphicPopup 、 msoControlButtonPopup 、 msoControlSplitButtonPopup 或 msoControlSplitButtonMRUPopup 。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

也可以使用 FindControl 方法返回一个 CommandBarPopup 对象。下面的示例在所有命令栏中搜索一个标志为“Graphics”的 CommandBarPopup 对象。

``` JavaScript
let myControl = Application.CommandBars.FindControl(Application.Enum.msoControlPopup, null, "Graphics")
```

## 方法

| 序号 | 名称                                  | 说明                                                                                          |
|:----:|:--------------------------------------|:----------------------------------------------------------------------------------------------|
|  1   | [Copy](#CommandBarPopup.Copy)         | 将一个命令栏弹出式控件复制到已有的命令栏中。                                                  |
|  2   | [Delete](#CommandBarPopup.Delete)     | 将 CommandBarPopup 对象从其集合中删除。                                                       |
|  3   | [Execute](#CommandBarPopup.Execute)   | 运行分配给指定 CommandBarPopup 控件的过程或内置命令。                                         |
|  4   | [Move](#CommandBarPopup.Move)         | 将指定的 CommandBarPopup 控件移动到已有的命令栏。                                             |
|  5   | [Reset](#CommandBarPopup.Reset)       | 将内置 CommandBarPopup 控件重置为其初始功能和图符。                                           |
|  6   | [SetFocus](#CommandBarPopup.SetFocus) | 将键盘的焦点移到指定的 CommandBarPopup 控件。如果弹出式控件被禁用或不可见，那么此方法将失败。 |

## 属性

| 序号 | 名称                                                    | 说明                                                                                                                                                                                          |
|:----:|:--------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CommandBarPopup.Application)             | 获取一个 Application 对象，代表 CommandBarPopup 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。                                                        |
|  2   | [BeginGroup](#CommandBarPopup.BeginGroup)               | 如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 True 。可读写。                                                                                                                      |
|  3   | [BuiltIn](#CommandBarPopup.BuiltIn)                     | 如果指定的命令栏弹出式控件是容器应用程序的内置命令栏，则为 True 。如果是自定义命令栏，则返回 False 。只读。                                                                                   |
|  4   | [Caption](#CommandBarPopup.Caption)                     | 获取或设置命令栏控件的标题文字。可读写。                                                                                                                                                      |
|  5   | [CommandBar](#CommandBarPopup.CommandBar)               | 获取一个 CommandBar 对象，该对象代表由指定的弹出式控件显示的菜单。只读。                                                                                                                      |
|  6   | [Controls](#CommandBarPopup.Controls)                   | 获取一个代表弹出式控件上的所有控件的 CommandBarControls 对象。只读。                                                                                                                          |
|  7   | [Creator](#CommandBarPopup.Creator)                     | 获取一个 32 位整数，指示创建 CommandBarPopup 对象时所使用的应用程序。只读。                                                                                                                   |
|  8   | [DescriptionText](#CommandBarPopup.DescriptionText)     | 获取或设置命令栏弹出式控件的说明。可读写。                                                                                                                                                    |
|  9   | [Enabled](#CommandBarPopup.Enabled)                     | 如果启用了 CommandBarPopup ，则此属性为 True 。可读写。                                                                                                                                       |
|  10  | [Height](#CommandBarPopup.Height)                       | 获取或设置 CommandBarPopup 控件的高度。可读写。                                                                                                                                               |
|  11  | [HelpContextId](#CommandBarPopup.HelpContextId)         | 获取或设置附加于 CommandBarPopup 控件的“帮助”主题的帮助上下文 ID 号。可读写。                                                                                                                 |
|  12  | [HelpFile](#CommandBarPopup.HelpFile)                   | 获取或设置附加于 CommandBarPopup 控件的“帮助”主题的文件名。可读写。                                                                                                                           |
|  13  | [Id](#CommandBarPopup.Id)                               | 获取内置的 CommandBarPopup 控件的 ID。只读。                                                                                                                                                  |
|  14  | [Index](#CommandBarPopup.Index)                         | 获取一个 Long 类型的值，该值代表集合中 CommandBarPopup 对象的索引号。只读。                                                                                                                   |
|  15  | [IsPriorityDropped](#CommandBarPopup.IsPriorityDropped) | 如果当前已根据使用统计信息和布局空间将 CommandBarPopup 控件从菜单或工具栏中删除，则获取 True 。（请注意，这与通过 Visible 属性设置的控件可见性不同）。只读。                                  |
|  16  | [Left](#CommandBarPopup.Left)                           | 获取指定的 CommandBarPopup 相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。                                                                                     |
|  17  | [OLEMenuGroup](#CommandBarPopup.OLEMenuGroup)           | 获取或设置一个 MsoOLEMenuGroup 常量，代表在将 OLE 服务器菜单组与 OLE 客户端菜单组合并时（即将一个容器应用程序类型的对象嵌入到另一个应用程序中时）指定的命令栏弹出式控件所属的菜单组。可读写。 |
|  18  | [OLEUsage](#CommandBarPopup.OLEUsage)                   | 获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 CommandBarPopup 控件。可读写。                                                              |
|  19  | [OnAction](#CommandBarPopup.OnAction)                   | 获取或设置一个 Visual Basic 过程的名称，当用户单击或更改 “CommandBarPopup” 控件的值时，将运行该过程。可读写。                                                                                 |
|  20  | [Parameter](#CommandBarPopup.Parameter)                 | 获取或设置一个应用程序可用于执行 CommandBarPopup 控件中命令的字符串。可读写。                                                                                                                 |
|  21  | [Parent](#CommandBarPopup.Parent)                       | 获取 CommandBarPopup 对象的 Parent 对象。只读。                                                                                                                                               |
|  22  | [Priority](#CommandBarPopup.Priority)                   | 获取或设置 CommandBarPopup 控件的优先级。可读写。                                                                                                                                             |
|  23  | [Tag](#CommandBarPopup.Tag)                             | 获取或设置有关 CommandBarPopup 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。                                                                                         |
|  24  | [TooltipText](#CommandBarPopup.TooltipText)             | 获取或设置 CommandBarPopup 的 “屏幕提示” 中显示的文本。可读写。                                                                                                                               |
|  25  | [Top](#CommandBarPopup.Top)                             | 获取指定的 CommandBarPopup 控件顶边到屏幕顶边的距离（以像素为单位）。只读。                                                                                                                   |
|  26  | [Type](#CommandBarPopup.Type)                           | 获取 CommandBarPopup 控件的类型。只读。                                                                                                                                                       |
|  27  | [Visible](#CommandBarPopup.Visible)                     | 获取或设置 CommandBarPopup 控件的 Visible 属性。可读写。                                                                                                                                      |
|  28  | [Width](#CommandBarPopup.Width)                         | 获取或设置指定的 CommandBarPopup 控件的宽度（以像素为单位）。可读写。                                                                                                                         |

## 成员方法

### CommandBarPopup.Copy

将一个命令栏弹出式控件复制到已有的命令栏中。

**语法**

**express.Copy ( *Bar* , *Barfore* )**

\*express是一个代表 CommandBarPopup 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                   |
|-----------|-----------|----------|------------------------------------------------------------------------------------------------------------------------|
| *Bar*     | 可选      | Variant  | 一个代表目标命令栏的 CommandBar 对象。如果省略此参数，控件将被复制到它原来所在的命令栏中。                             |
| *Barfore* | 可选      | Variant  | 一个指示新控件在命令栏上的位置的数值。新控件将插入到位于此位置的控件之前。如果省略此参数，控件将被复制到命令栏的末尾。 |

### CommandBarPopup.Delete

将 **CommandBarPopup** 对象从其集合中删除。

**语法**

**express.Delete ( *Temporary* )**

\*express是一个代表 CommandBarPopup 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                          |
|-------------|-----------|----------|-------------------------------------------------------------------------------|
| *Temporary* | 可选      | Variant  | 如果为 True，则从当前会话中删除此控件。应用程序在下次会话中将再次显示该控件。 |

**说明**

| 注释                                                   |
|--------------------------------------------------------|
| WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。 |

### CommandBarPopup.Execute

运行分配给指定 **CommandBarPopup** 控件的过程或内置命令。

**语法**

**express.Execute ()**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本 ET 示例创建一个命令栏，然后向其中添加内置命令栏按钮控件。该按钮将执行 ET AutoSum 函数。本示例使用 Execute 方法在显示该命令栏时计算选定单元格区域的总计。
function test()
{
    let cbrCustBar = Application.CommandBars.Add("Custom")
    let ctlAutoSum = cbrCustBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("AutoSum").Id)
    cbrCustBar.Visible = true 
    ctlAutoSum.Execute()
}
```

### CommandBarPopup.Move

将指定的 **CommandBarPopup** 控件移动到已有的命令栏。

**语法**

**express.Move ( *Bar* , *Before* )**

\*express是一个代表 CommandBarPopup 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                          |
|----------|-----------|----------|-----------------------------------------------------------------------------------------------|
| *Bar*    | 可选      | Variant  | 一个代表控件的目标命令栏的 Command 对象。如果忽略该参数，则控件将移动到当前所在命令栏的末端。 |
| *Before* | 可选      | Variant  | 表示控件位置的数字。控件将插到该位置的控件之前。如果忽略该参数，控件插入到同一命令栏。        |

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例将名为 Custom 的命令栏上的第一个组合框控件移动到该命令栏上的第七个控件之前。本示例将 Tag 设置为“Selection box”并赋予控件较低的优先级，以便在一行容纳不下所有控件时将其隐藏。
function test()
{
    let allcontrols = Application.CommandBars.Item("Custom").Controls
    for(let i = 1; i <= allcontrols.Count; i++)
    {
        if(allcontrols.Item(i).Type == Application.Enum.msoControlComboBox)
        {
            let newallcontrols = allcontrols.Item(i)
            newallcontrols.Move(null,7)
            newallcontrols.Tag = "Selection box"
            newallcontrols.Priority = 5
            break;
        }
    }
}
```

### CommandBarPopup.Reset

将内置 **CommandBarPopup** 控件重置为其初始功能和图符。

**语法**

**express.Reset ()**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

重置一个内置控件将恢复该控件原来对应的动作，并将各属性恢复成初始状态。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在所有命令栏中搜索标签为“Graphics”的 CommandBarPopup 对象，然后将其重置为默认状态。
function test()
{
    let myControl = Application.CommandBars.FindControl(Application.Enum.msoControlPopup, null, "Graphics") 
    myControl.Reset()
}
```

### CommandBarPopup.SetFocus

将键盘的焦点移到指定的 **CommandBarPopup** 控件。如果弹出式控件被禁用或不可见，那么此方法将失败。

**语法**

**express.SetFocus ()**

\*express是一个代表 CommandBarPopup 对象的变量。

## 成员属性

### CommandBarPopup.Application

获取一个 **Application** 对象，代表 **CommandBarPopup** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.BeginGroup

如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 **True** 。可读写。

**语法**

**express.BeginGroup**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在活动菜单栏上将最后一个控件作为一个新组的开始。
function test()
{
    let myMenuBar = Application.CommandBars.ActiveMenuBar
    let lastMenu = myMenuBar.Controls.Item(myMenuBar.Controls.Count)
    lastMenu.BeginGroup = true
}
```

### CommandBarPopup.BuiltIn

如果指定的命令栏弹出式控件是容器应用程序的内置命令栏，则为 **True** 。如果是自定义命令栏，则返回 **False** 。只读。

**语法**

**express.BuiltIn**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Caption

获取或设置命令栏控件的标题文字。可读写。

**语法**

**express.Caption**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                     |
|------------------------------------------|
| 控件的标题也可作为默认的“屏幕提示”显示。 |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在自定义命令栏中添加一个带拼写检查按钮图标的命令栏控件，然后将其标题设置为“Spelling checker”。
function test()
{
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)
    myBar.Visible = true 
    let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, 2)
    myControl.DescriptionText = "Starts the spelling checker"
    myControl.Caption = "Spelling checker"
}
```

### CommandBarPopup.CommandBar

获取一个 **CommandBar** 对象，该对象代表由指定的弹出式控件显示的菜单。只读。

**语法**

**express.CommandBar**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例将变量 fourthLevel 设置为名为“Drawing”的命令栏上的第四个控件。
let fourthLevel = Application.CommandBars.Item("Drawing").Controls.Item(1).CommandBar.Controls.Item(4)
```

### CommandBarPopup.Controls

获取一个代表弹出式控件上的所有控件的 **CommandBarControls** 对象。只读。

**语法**

**express.Controls**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Creator

获取一个 32 位整数，指示创建 **CommandBarPopup** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.DescriptionText

获取或设置命令栏弹出式控件的说明。可读写。

**语法**

**express.DescriptionText**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

该说明不显示给用户，但有助于其他开发者将控件行为归档。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Enabled

如果启用了 **CommandBarPopup** ，则此属性为 **True** 。可读写。

**语法**

**express.Enabled**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Height

获取或设置 **CommandBarPopup** 控件的高度。可读写。

**语法**

**express.Height**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.HelpContextId

获取或设置附加于 **CommandBarPopup** 控件的“帮助”主题的帮助上下文 ID 号。可读写。

**语法**

**express.HelpContextId**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

若要使用此属性，必须同时设置 HelpFile 属性。帮助主题响应 Shift+F1 按键操作。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.HelpFile

获取或设置附加于 **CommandBarPopup** 控件的“帮助”主题的文件名。可读写。

**语法**

**express.HelpFile**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

若要使用该属性，必须同时设置 HelpContextID 属性。帮助主题响应 Shift+F1 用户按键操作。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Id

获取内置的 **CommandBarPopup** 控件的 ID。只读。

**语法**

**express.Id**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

弹出式控件的 ID 决定了该对象的内置操作。

### CommandBarPopup.Index

获取一个 **Long** 类型的值，该值代表集合中 **CommandBarPopup** 对象的索引号。只读。

**语法**

**express.Index**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

第一个命令栏控件的位置为 1。 **CommandBarControls** 集合中的分隔符不计算在内。

### CommandBarPopup.IsPriorityDropped

如果当前已根据使用统计信息和布局空间将 **CommandBarPopup** 控件从菜单或工具栏中删除，则获取 **True** 。（请注意，这与通过 **Visible** 属性设置的控件可见性不同）。只读。

**语法**

**express.IsPriorityDropped**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

如果 **IsPriorityDropped** 为 **True** ，则 **Visible** 属性设置为 **True** 的控件将不能在个性化菜单或工具栏中立即显示出来。

为了确定何时将某一特定菜单项的 **IsPriorityDropped** 属性设置为 **True** ，WPS Office会记录下该菜单项的使用次数的总和，并记录下用户使用了同一菜单中的其他菜单项而没有使用该特定菜单项的应用程序会话数。当该会话数达到一定的阈值之后，便会减少使用次数总和。当总次数减为零时，便将 **IsPriorityDropped** 设置为 **True** 。编程人员不能设置会话值和阈值，也不能设置 **IsPriorityDropped** 属性。但他们可以使用 **AdaptiveMenus** 属性来禁用应用程序中特定菜单的自适应菜单。

为了确定何时将某一特定工具栏控件的 **IsPriorityDropped** 属性设置为 **True** ，Office 维护了一个列表，记录下了该工具栏中各个控件的最后使用顺序。工具栏会在空间允许的前提下尽可能多地显示控件，控件按使用时间由近向远排列，最近使用过的控件显示在最前面。 **Priority** 设置为 1 的控件将始终显示，而且为了显示这样的控件，在必要时工具栏还将换行。编程人员可使用 **Priority** 属性来确保始终显示特定的工具栏控件，或重新定位工具栏以便使其有足够的空间显示所有的控件。

使用下表可以预测在菜单项的 **IsPriorityDropped** 属性被设置为 **True** 之前，它仍将出现在个性化菜单中的会话数。

| 命令栏控件的使用次数 | 应用程序的会话数 |
|----------------------|------------------|
| 0, 1                 | 3                |
| 2                    | 6                |
| 3                    | 9                |
| 4, 5                 | 12               |
| 6- 8                 | 17               |
| 9-13                 | 23               |
| 14-24                | 29               |
| 25 及 25 以上        | 31               |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Left

获取指定的 **CommandBarPopup** 相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。

**语法**

**express.Left**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.OLEMenuGroup

获取或设置一个 **MsoOLEMenuGroup** 常量，代表在将 OLE 服务器菜单组与 OLE 客户端菜单组合并时（即将一个容器应用程序类型的对象嵌入到另一个应用程序中时）指定的命令栏弹出式控件所属的菜单组。可读写。

**语法**

**express.OLEMenuGroup**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

通过此属性，加载项应用程序可以指定其命令栏控件在 Office 应用程序中的表示方式。如果容器或服务器中的任何一个不实现命令栏，则执行常规的 OLE 菜单合并：菜单栏以及服务器的所有工具栏将进行合并，但容器的工具栏不会被合并。本属性只与菜单栏上的弹出式控件有关，因为菜单是根据其菜单组的类别进行合并的。

如果要合并的两个应用程序均实现了命令栏，则命令栏控件根据 **OLEUsage** 属性进行合并。

| 注释                       |
|----------------------------|
| 该属性对内置控件是只读的。 |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例检查名为“Custom”的命令栏上一个新自定义弹出式控件的 OLEMenuGroup 属性，并将该属性设置为 msoOLEMenuGroupNone。 
function test()
{
    let myControl = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlPopup, null, null, null, false)
    myControl.OLEMenuGroup = Application.Enum.msoOLEMenuGroupNone
}
```

### CommandBarPopup.OLEUsage

获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 **CommandBarPopup** 控件。可读写。

**语法**

**express.OLEUsage**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

通过此属性，用户可以在某 Office 应用程序与另一个 Office 应用程序合并时，指定加载项应用程序的各个命令栏控件在该 Office 应用程序中的表示方式。如果客户端和服务器均实现命令栏，则命令栏控件将逐个嵌入客户端。系统将从服务器中略去标记为只用于客户端（或既不用于客户端也不用于服务器）的自定义控件，而从客户端中略去标记为只用于服务器（或既不用于客户端也不用于服务器）的控件。剩下的控件再进行合并。

如果要合并的应用程序之一不是 Office 应用程序，则使用常规 OLE 菜单合并，该过程将由 **OLEMenuGroup** 属性控制。

### CommandBarPopup.OnAction

获取或设置一个 Visual Basic 过程的名称，当用户单击或更改 **“CommandBarPopup”** 控件的值时，将运行该过程。可读写。

**语法**

**express.OnAction**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

容器应用程序将确定该值是否是一个合法的宏名。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Parameter

获取或设置一个应用程序可用于执行 **CommandBarPopup** 控件中命令的字符串。可读写。

**语法**

**express.Parameter**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

如果为内置控件设置指定的参数，那么应用程序将修改其默认的操作，前提是该应用程序能分析和使用这个新值。如果为自定义控件设置参数，则可以将其用于向 Visual Basic 过程发送信息，或保存有关该控件的信息（类似于另一个 Tag 属性值）。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例将一个新参数赋予控件并将焦点置于新按钮上。 
function test()
{
    let myControl = Application.CommandBars.Item("Custom").Controls.Item(4)
    myControl.Copy(null, 1)
    myControl.Parameter = "2"
    myControl.SetFocus()
}
```

### CommandBarPopup.Parent

获取 **CommandBarPopup** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Priority

获取或设置 **CommandBarPopup** 控件的优先级。可读写。

**语法**

**express.Priority**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。无法容纳在一行中的控件将在命令栏上从右向左进行删减。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例设置命令栏弹出式控件的说明性文字和优先级。 
function test()
{
    let popControl = Application.CommandBars.FindControl(Application.Enum.msoControlPopup, null, "Graphics")
    popControl.DescriptionText = "Graphics Selection dialog"
    popControl.Priority = 5
}
```

### CommandBarPopup.Tag

获取或设置有关 **CommandBarPopup** 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。

**语法**

**express.Tag**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.TooltipText

获取或设置 **CommandBarPopup** 的 **“屏幕提示”** 中显示的文本。可读写。

**语法**

**express.TooltipText**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

默认情况下， **Caption** 属性的值将用作 **“屏幕提示”** 。

### CommandBarPopup.Top

获取指定的 **CommandBarPopup** 控件顶边到屏幕顶边的距离（以像素为单位）。只读。

**语法**

**express.Top**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Type

获取 **CommandBarPopup** 控件的类型。只读。

**语法**

**express.Type**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Visible

获取或设置 **CommandBarPopup** 控件的 **Visible** 属性。可读写。

**语法**

**express.Visible**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarPopup.Width

获取或设置指定的 **CommandBarPopup** 控件的宽度（以像素为单位）。可读写。

**语法**

**express.Width**

\*express是一个代表 CommandBarPopup 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CommandBarPopup](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomXMLValidationError 对象

## 简介

代表 CustomXMLValidationErrors 集合中的单个验证错误。

## 说明

当根据架构验证操作时（如添加节点时），验证错误可能被触发，或者在不满足某个条件时（例如，如果开始日期晚于结束日期），可能由用户触发。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

## 方法

| 序号 | 名称                                       | 说明                                                   |
|:----:|:-------------------------------------------|:-------------------------------------------------------|
|  1   | [Delete](#CustomXMLValidationError.Delete) | 删除代表数据验证错误的 CustomXMLValidationError 对象。 |

## 属性

| 序号 | 名称                                                 | 说明                                                                                                        |
|:----:|:-----------------------------------------------------|:------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLValidationError.Application) | 获取一个代表 CustomXMLValidationError 对象的容器应用程序的 Application 对象。只读。                         |
|  2   | [Creator](#CustomXMLValidationError.Creator)         | 获取一个 32 位整数，指示创建 CustomXMLValidationError 对象时所使用的应用程序。只读。                        |
|  3   | [ErrorCode](#CustomXMLValidationError.ErrorCode)     | 获取一个代表 CustomXMLValidationError 对象中的验证错误的数字。只读。                                        |
|  4   | [Name](#CustomXMLValidationError.Name)               | 获取 CustomXMLValidationError 对象中的错误名称。如果不存在错误，则该属性将返回 Nothing 。只读。             |
|  5   | [Node](#CustomXMLValidationError.Node)               | 获取 CustomXMLValidationError 对象中的一个节点（如果存在）。如果不存在节点，则该属性将返回 Nothing 。只读。 |
|  6   | [Parent](#CustomXMLValidationError.Parent)           | 获取 CustomXMLValidationError 对象的 Parent 对象。只读。                                                    |
|  7   | [Text](#CustomXMLValidationError.Text)               | 获取 CustomXMLValidationError 对象中的文本。只读。                                                          |
|  8   | [Type](#CustomXMLValidationError.Type)               | 获取 CustomXMLValidationError 对象中所生成的错误的类型。只读。                                              |

## 成员方法

### CustomXMLValidationError.Delete

删除代表数据验证错误的 **CustomXMLValidationError** 对象。

**语法**

**express.Delete ()**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

|                                                                                                                                                                      |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例将删除验证错误。
Application.ActivePresentation.CustomXMLParts.Item(1).Errors.Item(1).Delete()
```

## 成员属性

### CustomXMLValidationError.Application

获取一个代表 **CustomXMLValidationError** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Creator

获取一个 32 位整数，指示创建 **CustomXMLValidationError** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.ErrorCode

获取一个代表 **CustomXMLValidationError** 对象中的验证错误的数字。只读。

**语法**

**express.ErrorCode**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Name

获取 **CustomXMLValidationError** 对象中的错误名称。如果不存在错误，则该属性将返回 **Nothing** 。只读。

**语法**

**express.Name**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Node

获取 **CustomXMLValidationError** 对象中的一个节点（如果存在）。如果不存在节点，则该属性将返回 **Nothing** 。只读。

**语法**

**express.Node**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Parent

获取 **CustomXMLValidationError** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Text

获取 **CustomXMLValidationError** 对象中的文本。只读。

**语法**

**express.Text**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLValidationError.Type

获取 **CustomXMLValidationError** 对象中所生成的错误的类型。只读。

**语法**

**express.Type**

\*express是一个代表 CustomXMLValidationError 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLValidationError](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DocumentProperty 对象

## 简介

代表容器文档的一个自定义或内置文档属性。 DocumentProperty 对象是 DocumentProperties 集合的成员。

## 说明

使用 BuiltinDocumentProperties(index) 可返回一个代表特定内置文档属性的 DocumentProperty 对象，其中 index 是该内置文档属性的名称或索引号。使用 CustomDocumentProperties(index) 可返回一个代表特定自定义文档属性的 DocumentProperty 对象，其中 *index* 是该自定义文档属性的名称或索引号。下面的列表包含所有可用内置文档属性的名称：

|              |              |
|--------------|--------------|
| 标题         | 字数         |
| 主题         | 字符数       |
| 作者         | 安全性       |
| 关键字       | 类别         |
| 批注         | 格式         |
| 批注         | 经理         |
| 上一个作者   | 单位         |
| 修订次数     | 字节数       |
| 应用程序名   | 行数         |
| 上次打印日期 | 段落数       |
| 创建日期     | 幻灯片数     |
| 上次保存时间 | 备注数       |
| 编辑时间总计 | 隐藏幻灯片数 |
| 页数         | 多媒体剪辑数 |

容器应用程序不一定为每个内置文档属性都定义一个属性值。如果所给的应用程序没有为某内置文档属性定义一个属性值，那么返回该文档属性的 Value 属性将产生错误。

## 方法

| 序号 | 名称                               | 说明                   |
|:----:|:-----------------------------------|:-----------------------|
|  1   | [Delete](#DocumentProperty.Delete) | 删除自定义的文档属性。 |

## 属性

| 序号 | 名称                                             | 说明                                                                                                                                    |
|:----:|:-------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#DocumentProperty.Application)     | 获取一个 Application 对象，代表 DocumentProperty 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Creator](#DocumentProperty.Creator)             | 获取一个 32 位整数，指示创建 DocumentProperty 对象时所使用的应用程序。只读。                                                            |
|  3   | [LinkSource](#DocumentProperty.LinkSource)       | 获取或设置所链接的自定义文档属性的来源。可读/写。                                                                                       |
|  4   | [LinkToContent](#DocumentProperty.LinkToContent) | 如果自定义文档属性的值链接到容器文档的内容，则为 True 。如果该值是静态的，则为 False 。可读/写。                                        |
|  5   | [Name](#DocumentProperty.Name)                   | 获取或设置文档属性的名称。可读写。                                                                                                      |
|  6   | [Parent](#DocumentProperty.Parent)               | 获取 DocumentProperty 对象的 Parent 对象。只读。                                                                                        |
|  7   | [Type](#DocumentProperty.Type)                   | 获取或设置文档属性类型。对于内置文档属性为只读；对于自定义文档属性为可读/写。                                                           |
|  8   | [Value](#DocumentProperty.Value)                 | 获取或设置文档属性的值。可读/写。                                                                                                       |

## 成员方法

### DocumentProperty.Delete

删除自定义的文档属性。

**语法**

**express.Delete ()**

\*express是一个代表 DocumentProperty 对象的变量。

**说明**

不能删除内置的文档属性。

**示例**

``` JavaScript
//本示例删除一个自定义文档属性。为了使该示例正确运行，必须有名为“CustomNumber”的自定义 DocumentProperty 对象。
Application.ActiveDocument.CustomDocumentProperties.Item("CustomNumber").Delete()
```

## 成员属性

### DocumentProperty.Application

获取一个 **Application** 对象，代表 **DocumentProperty** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 DocumentProperty 对象的变量。

### DocumentProperty.Creator

获取一个 32 位整数，指示创建 **DocumentProperty** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 DocumentProperty 对象的变量。

### DocumentProperty.LinkSource

获取或设置所链接的自定义文档属性的来源。可读/写。

**语法**

**express.LinkSource**

\*express是一个代表 DocumentProperty 对象的变量。

**说明**

本属性只用于自定义文档属性，不能用于内置文档属性。

指定链接的链接源由容器应用程序定义。

设置 **LinkSource** 属性会将 **LinkToContent** 属性设置为 **True** 。

**示例**

``` JavaScript
//本示例显示自定义文档属性的链接状态。要使本示例正常工作，dp 必须是一个有效的 DocumentProperty 对象。
function test(dp){
    let tf = dp.LinkToContent ? "" : "not "
    let stat = "This property is " + tf + "linked"

    if(dp.LinkToContent){
        stat = stat + "\n" + "The link source is " + dp.LinkSource
    }

    alert(stat)
}
```

### DocumentProperty.LinkToContent

如果自定义文档属性的值链接到容器文档的内容，则为 **True** 。如果该值是静态的，则为 **False** 。可读/写。

**语法**

**express.LinkToContent**

\*express是一个代表 DocumentProperty 对象的变量。

**说明**

该属性只用于自定义文档属性。对内置文档属性，该属性值为 **False** 。

使用 **LinkSource** 属性设置指定链接属性的来源。设置 **LinkSource** 属性会将 **LinkToContent** 属性设置为 **True** 。

**示例**

``` JavaScript
//本示例显示自定义文档属性的链接状态。要使本示例正常工作，dp 必须是一个有效的 DocumentProperty 对象。
function test(dp){
    let tf = dp.LinkToContent ? "" : "not "
    let stat = "This property is " + tf + "linked"

    if(dp.LinkToContent){
        stat = stat + "\n" + "The link source is " + dp.LinkSource
    }

    alert(stat)
}
```

### DocumentProperty.Name

获取或设置文档属性的名称。可读写。

**语法**

**express.Name**

\*express是一个代表 DocumentProperty 对象的变量。

**说明**

**DocumentProperty** 对象代表容器文档的自定义或内置文档属性。

**示例**

``` JavaScript
//本示例显示一个文档属性的名称、类型和值。必须向该过程传递一个有效的 DocumentProperty 对象。 
function test(dp){
    alert("value = " + dp.Value + "\n" +
        "type = " + dp.Type + "\n" +
        "name = " + dp.Name)
}
```

### DocumentProperty.Parent

获取 **DocumentProperty** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DocumentProperty 对象的变量。

**示例**

### DocumentProperty.Type

获取或设置文档属性类型。对于内置文档属性为只读；对于自定义文档属性为可读/写。

**语法**

**express.Type**

\*express是一个代表 DocumentProperty 对象的变量。

**说明**

返回值将是一个 **MsoDocProperties** 常量。

### DocumentProperty.Value

获取或设置文档属性的值。可读/写。

**语法**

**express.Value**

\*express是一个代表 DocumentProperty 对象的变量。

**说明**

如果容器应用程序没有定义某个内置文档属性的值，则读取该文档属性的 **Value** 属性将导致出错。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/DocumentProperty](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ODSOColumn 对象

## 简介

代表数据源中的一个字段。 ODSOColumn 对象是 ODSOColumns 集合的成员。

## 说明

ODSOColumns 集合包括邮件合并数据源中的所有数据字段（例如，姓名、地址和城市）。

不能向 ODSOColumns 集合添加字段。 ODSOColumns 集合自动包括数据源中的所有数据字段。

使用 Columns ( *index* ) 返回单个 ODSOColumn 对象，其中 *index* 是数据字段名或索引号。索引号表示数据字段在邮件合并数据源中的位置。

## 属性

| 序号 | 名称                                   | 说明                                                                                                                               |
|:----:|:---------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ODSOColumn.Application) | 获取一个 Application 对象，代表 ODSOColum n 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Creator](#ODSOColumn.Creator)         | 获取一个 32 位整数，指示创建 ODSOColumn 对象时所使用的应用程序。只读。                                                             |
|  3   | [Index](#ODSOColumn.Index)             | 获取一个 Long 类型的值，代表集合中 ODSOColumn 对象的索引号。只读。                                                                 |
|  4   | [Name](#ODSOColumn.Name)               | 获取邮件合并数据源中的数据字段的名称。只读。                                                                                       |
|  5   | [Parent](#ODSOColumn.Parent)           | 获取 ODSOColumn 对象的 Parent 对象。只读。                                                                                         |
|  6   | [Value](#ODSOColumn.Value)             | 获取邮件合并数据源中的数据字段的值。只读。                                                                                         |

## 成员属性

### ODSOColumn.Application

获取一个 **Application** 对象，代表 **ODSOColum** n 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 ODSOColumn 对象的变量。

### ODSOColumn.Creator

获取一个 32 位整数，指示创建 **ODSOColumn** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 ODSOColumn 对象的变量。

### ODSOColumn.Index

获取一个 **Long** 类型的值，代表集合中 **ODSOColumn** 对象的索引号。只读。

**语法**

**express.Index**

\*express是一个代表 ODSOColumn 对象的变量。

### ODSOColumn.Name

获取邮件合并数据源中的数据字段的名称。只读。

**语法**

**express.Name**

\*express是一个代表 ODSOColumn 对象的变量。

### ODSOColumn.Parent

获取 **ODSOColumn** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ODSOColumn 对象的变量。

### ODSOColumn.Value

获取邮件合并数据源中的数据字段的值。只读。

**语法**

**express.Value**

\*express是一个代表 ODSOColumn 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/ODSOColumn](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PickerFields 对象

## 简介

PickerField 对象的集合。每个 PickerField 对象都代表选取器对话框的一个列定义。

## 属性

| 序号 | 名称                                     | 说明                                                                         |
|:----:|:-----------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#PickerFields.Application) | 获取一个 Application 对象，该对象代表 PickerFields 对象的容器应用程序。只读  |
|  2   | [Count](#PickerFields.Count)             | 检索包含在 PickerFields 集合中的 PickerField 对象数的计数。只读              |
|  3   | [Creator](#PickerFields.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 PickerFields 对象的应用程序。只读 |
|  4   | [Item](#PickerFields.Item)               | 检索位于指定索引处的 PickerField 对象。只读                                  |

## 成员属性

### PickerFields.Application

获取一个 **Application** 对象，该对象代表 **PickerFields** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PickerFields 对象的变量。

### PickerFields.Count

检索包含在 **PickerFields** 集合中的 **PickerField** 对象数的计数。只读

**语法**

**express.Count**

\*express是一个代表 PickerFields 对象的变量。

### PickerFields.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PickerFields** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PickerFields 对象的变量。

### PickerFields.Item

检索位于指定索引处的 **PickerField** 对象。只读

**语法**

**express.Item**

\*express是一个代表 PickerFields 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PickerFields](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartArtColors 对象

## 简介

SmartArt 颜色样式的集合。

## 说明

模拟 WPS Office Fluent 功能区用户界面中“SmartArt 工具”上“设计”组内的“更改颜色”命令上的命令。

``` JavaScript
/*下面的代码设置 Smart Art 图表的配色方案。*/Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Color = Application.SmartArtColors.Item(1)
```

## 方法

| 序号 | 名称                         | 说明                                                          |
|:----:|:-----------------------------|:--------------------------------------------------------------|
|  1   | [Item](#SmartArtColors.Item) | 检索位于指定索引处或具有指定的唯一 Id 的 SmartArtColor 对象。 |

## 属性

| 序号 | 名称                                       | 说明                                                                             |
|:----:|:-------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtColors.Application) | 获取一个 Application 对象，该对象代表 SmartArtColors 对象的容器应用程序。只读。  |
|  2   | [Count](#SmartArtColors.Count)             | 检索包含在 SmartArtColors 集合中的 SmartArtColor 对象数的计数。只读。            |
|  3   | [Creator](#SmartArtColors.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtColors 对象的应用程序。只读。 |
|  4   | [Parent](#SmartArtColors.Parent)           | 返回调用对象。只读。                                                             |

## 成员方法

### SmartArtColors.Item

检索位于指定索引处或具有指定的唯一 Id 的 **SmartArtColor** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SmartArtColors 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                            |
|---------|-----------|----------|-----------------------------------------------------------------|
| *Index* | 必选      | Variant  | 指定一个代表索引的整数或一个代表 Smart Art 颜色的 Id 的字符串。 |

**返回值**

SmartArtColor

## 成员属性

### SmartArtColors.Application

获取一个 **Application** 对象，该对象代表 **SmartArtColors** 对象的容器应用程序。只读。

**语法**

**express.Application**

\*express是一个代表 SmartArtColors 对象的变量。

### SmartArtColors.Count

检索包含在 **SmartArtColors** 集合中的 **SmartArtColor** 对象数的计数。只读。

**语法**

**express.Count**

\*express是一个代表 SmartArtColors 对象的变量。

### SmartArtColors.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtColors** 对象的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SmartArtColors 对象的变量。

### SmartArtColors.Parent

返回调用对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SmartArtColors 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtColors](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartArtQuickStyle 对象

## 简介

代表 Smart Art 快速样式

## 说明

``` JavaScript
/*下面的代码更改 WPP 中的 Smart Art 图表的快速样式。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(1)
```

## 方法

| 序号 | 名称                                   | 说明 |
|:----:|:---------------------------------------|:-----|
|  1   | [Reserve](#SmartArtQuickStyle.Reserve) |      |

## 属性

| 序号 | 名称                                           | 说明                                                                               |
|:----:|:-----------------------------------------------|:-----------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtQuickStyle.Application) | 获取一个 Application 对象，该对象代表 SmartArtQuickStyle 对象的容器应用程序。只读  |
|  2   | [Category](#SmartArtQuickStyle.Category)       | 检索与 SmartArt 快速样式关联的主类别名称。只读                                     |
|  3   | [Creator](#SmartArtQuickStyle.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtQuickStyle 对象的应用程序。只读 |
|  4   | [Description](#SmartArtQuickStyle.Description) | 检索 SmartArt 快速样式的说明。只读。                                               |
|  5   | [Id](#SmartArtQuickStyle.Id)                   | 检索关联的 SmartArt 快速样式的唯一 Id。只读。                                      |
|  6   | [Name](#SmartArtQuickStyle.Name)               | 检索 SmartArt 快速样式的字符串名称。只读。                                         |
|  7   | [Parent](#SmartArtQuickStyle.Parent)           | 返回调用对象。只读                                                                 |

## 成员方法

### SmartArtQuickStyle.Reserve

**语法**

**express.Reserve ()**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

## 成员属性

### SmartArtQuickStyle.Application

获取一个 **Application** 对象，该对象代表 **SmartArtQuickStyle** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Category

检索与 SmartArt 快速样式关联的主类别名称。只读

**语法**

**express.Category**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtQuickStyle** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Description

检索 SmartArt 快速样式的说明。只读。

**语法**

**express.Description**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Id

检索关联的 SmartArt 快速样式的唯一 Id。只读。

**语法**

**express.Id**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

**说明**

与此属性关联的 ID 区分大小写。

### SmartArtQuickStyle.Name

检索 SmartArt 快速样式的字符串名称。只读。

**语法**

**express.Name**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

### SmartArtQuickStyle.Parent

返回调用对象。只读

**语法**

**express.Parent**

\*express是一个代表 SmartArtQuickStyle 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtQuickStyle](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartArtQuickStyles 对象

## 简介

代表 Smart Art 快速样式的集合。

## 说明

``` JavaScript
/*下面的代码更改 WPP 中的 Smart Art 图表的快速样式。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(1)
```

## 方法

| 序号 | 名称                              | 说明                                                               |
|:----:|:----------------------------------|:-------------------------------------------------------------------|
|  1   | [Item](#SmartArtQuickStyles.Item) | 在指定的索引处或通过指定的唯一 ID 来检索 SmartArtQuickStyle 对象。 |

## 属性

| 序号 | 名称                                            | 说明                                                                                |
|:----:|:------------------------------------------------|:------------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtQuickStyles.Application) | 获取一个 Application 对象，该对象代表 SmartArtQuickStyles 对象的容器应用程序。只读  |
|  2   | [Count](#SmartArtQuickStyles.Count)             | 检索包含在 SmartArtQuickStyles 集合中的 SmartArtQuickStyle 对象数的计数。只读       |
|  3   | [Creator](#SmartArtQuickStyles.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtQuickStyles 对象的应用程序。只读 |
|  4   | [Parent](#SmartArtQuickStyles.Parent)           | 返回调用对象。只读                                                                  |

## 成员方法

### SmartArtQuickStyles.Item

在指定的索引处或通过指定的唯一 ID 来检索 **SmartArtQuickStyle** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                   |
|---------|-----------|----------|------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 指定一个代表索引的整数或一个代表 SmartArtQuickStyle 对象位置的字符串。 |

**返回值**

SmartArtQuickStyle

## 成员属性

### SmartArtQuickStyles.Application

获取一个 **Application** 对象，该对象代表 **SmartArtQuickStyles** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

### SmartArtQuickStyles.Count

检索包含在 SmartArtQuickStyles 集合中的 SmartArtQuickStyle 对象数的计数。只读

**语法**

**express.Count**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

### SmartArtQuickStyles.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtQuickStyles** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

### SmartArtQuickStyles.Parent

返回调用对象。只读

**语法**

**express.Parent**

\*express是一个代表 SmartArtQuickStyles 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtQuickStyles](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
