# ToggleButton 对象

## 简介

创建一个切换开关的按钮。

## 说明

创建一个切换开关的按钮。

## 方法

| 序号 | 名称                       | 说明                       |
|:----:|:---------------------------|:---------------------------|
|  1   | [Move](#ToggleButton.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                         | 说明                                                      |
|:----:|:---------------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#ToggleButton.AutoSize)           | 指定该对象是否根据内容自动设置它的大小                    |
|  2   | [BackColor](#ToggleButton.BackColor)         | 获取或设置指定对象的内部颜色                              |
|  3   | [BackStyle](#ToggleButton.BackStyle)         | 您可以使用该属性指定控件是否透明                          |
|  4   | [Caption](#ToggleButton.Caption)             | 获取或设置出现在控件中的文本                              |
|  5   | [ControlSource](#ToggleButton.ControlSource) | 您可以使用该属性指定控件中显示的数据                      |
|  6   | [Enabled](#ToggleButton.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                  |
|  7   | [Font](#ToggleButton.Font)                   | 指定对象的字体，只读                                      |
|  8   | [ForeColor](#ToggleButton.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                      |
|  9   | [Height](#ToggleButton.Height)               | 获取或设置指定对象的高度                                  |
|  10  | [HeightPolicy](#ToggleButton.HeightPolicy)   | 获取或设置指定对像的高度策略                              |
|  11  | [Left](#ToggleButton.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  12  | [Name](#ToggleButton.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  13  | [TabIndex](#ToggleButton.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  14  | [TabStop](#ToggleButton.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  15  | [TextAlign](#ToggleButton.TextAlign)         | 该属性用于指定对象中文本的对齐方式                        |
|  16  | [Top](#ToggleButton.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置              |
|  17  | [Value](#ToggleButton.Value)                 | 指定该按钮是否被按下。true是按下，false是弹起。           |
|  18  | [Visible](#ToggleButton.Visible)             | 返回或设置该对象是否可见                                  |
|  19  | [Width](#ToggleButton.Width)                 | 获取或设置指定对象的宽度                                  |
|  20  | [WidthPolic](#ToggleButton.WidthPolic)       | 获取或设置指定对象的宽度策略                              |
|  21  | [WordWrap](#ToggleButton.WordWrap)           | 指定控件的内容是否在行尾自动换行                          |

## 成员方法

### ToggleButton.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 ToggleButton 对象的变量。

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
{#_examples}/*将ToggleButton1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.ToggleButton1.Move(100,100,100,100)
}
```

## 成员属性

### ToggleButton.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置ToggleButton1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.ToggleButton1.AutoSize = true
}
```

### ToggleButton.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将ToggleButton1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.ToggleButton1.BackColor = 0x665898
}
```

### ToggleButton.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

**示例**

``` JavaScript
/*将ToggleButton1设置为透明*/
function func1()
{
    UserForm1.ToggleButton1.BackStyle = 0
}
```

### ToggleButton.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将ToggleButton1中的文本设置为"ToggleButton1"*/
function func1()
{
    UserForm1.ToggleButton1.Caption = "ToggleButton1"
}
```

### ToggleButton.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### ToggleButton.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.ToggleButton1.Enabled = true
}
```

### ToggleButton.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

指定对象的字体，只读

### ToggleButton.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将ToggleButton1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.ToggleButton1.ForeColor = 0x658978
}
```

### ToggleButton.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将ToggleButton1控件的高度设置为100*/
function func1()
{
    UserForm1.ToggleButton1.Height = 100
}
```

### ToggleButton.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将ToggleButton1的高度设置为固定*/
function func1()
{
    UserForm1.ToggleButton1.HeightPolicy = 2
}
```

### ToggleButton.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ToggleButton1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.ToggleButton1.Left = 100
}
```

### ToggleButton.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### ToggleButton.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将ToggleButton1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.ToggleButton1.TabIndex = 2
}
```

### ToggleButton.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将ToggleButton1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.ToggleButton1.TabStop = true
}
```

### ToggleButton.TextAlign

该属性用于指定对象中文本的对齐方式

**语法**

**express.TextAlign**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

该属性用于指定对象中文本的对齐方式，包含以下值：

| 设置   | 值  | 描述                     |
|--------|-----|--------------------------|
| Left   | 1   | 文本数字和日期都左对齐   |
| Center | 2   | 文本数字和日期都居中对齐 |
| Right  | 3   | 文本数字和日期都右对齐   |

**示例**

``` JavaScript
/*将ToggleButton1中的文本设置为居中对齐*/
function func1()
{
    UserForm1.ToggleButton1.TextAlign = 2
}
```

### ToggleButton.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ToggleButton1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.ToggleButton1.Top = 100
}
```

### ToggleButton.Value

指定该按钮是否被按下。true是按下，false是弹起。

**语法**

**express.Value**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

指定该按钮是否被按下。true是按下，false是弹起。

**示例**

``` JavaScript
/*将ToggleButton1设置为按下*/
function func1()
{
    UserForm1.ToggleButton1.Value = true
}
```

### ToggleButton.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将ToggleButton1设置为可见*/
function func1()
{
    UserForm1.ToggleButton1.Visible = true
}
```

### ToggleButton.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将ToggleButton1的宽度设置为100*/
function func1()
{
    UserForm1.ToggleButton1.Width = 100
}
```

### ToggleButton.WidthPolic

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolic**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将ToggleButton1的宽度设置为固定*/
function func1()
{
    UserForm1.ToggleButton1.WidthPolicy = 2
}
```

### ToggleButton.WordWrap

指定控件的内容是否在行尾自动换行

**语法**

**express.WordWrap**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

指定控件的内容是否在行尾自动换行

**示例**

``` JavaScript
/*将ToggleButton1设置为控件的内容在行尾自动换行*/
function func1()
{
    UserForm1.ToggleButton1.WordWrap = true
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/ToggleButton](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
