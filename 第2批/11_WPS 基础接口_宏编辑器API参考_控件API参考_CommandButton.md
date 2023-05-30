# CommandButton 对象

## 简介

创建一个可以让用户选择，以完成一个命令的按钮

## 说明

创建一个可以让用户选择，以完成一个命令的按钮

## 方法

| 序号 | 名称                        | 说明                           |
|:----:|:----------------------------|:-------------------------------|
|  1   | [Move](#CommandButton.Move) | 将指定对象移动到参数指定的坐标 |

## 属性

| 序号 | 名称                                        | 说明                                                      |
|:----:|:--------------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#CommandButton.AutoSize)         | 指定该对象是否根据内容自动设置它的大小                    |
|  2   | [BackColor](#CommandButton.BackColor)       | 获取或设置指定对象的内部颜色                              |
|  3   | [BackStyle](#CommandButton.BackStyle)       | 您可以使用该属性指定控件是否透明                          |
|  4   | [Caption](#CommandButton.Caption)           | 获取或设置出现在控件中的文本                              |
|  5   | [Enabled](#CommandButton.Enabled)           | 您可以通过该属性设置或返回是否启用该控件                  |
|  6   | [Font](#CommandButton.Font)                 | 指定对象的字体，只读                                      |
|  7   | [ForeColor](#CommandButton.ForeColor)       | 您可以通过该属性指定控件中文本的颜色                      |
|  8   | [Height](#CommandButton.Height)             | 获取或设置指定对象的高度                                  |
|  9   | [HeightPolicy](#CommandButton.HeightPolicy) | 获取或设置指定对像的高度策略                              |
|  10  | [Left](#CommandButton.Left)                 | 您可以通过该属性指定控件在对象或报表中的位置              |
|  11  | [Name](#CommandButton.Name)                 | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  12  | [TabIndex](#CommandButton.TabIndex)         | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  13  | [TabStop](#CommandButton.TabStop)           | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  14  | [Top](#CommandButton.Top)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  15  | [Visible](#CommandButton.Visible)           | 返回或设置该对象是否可见                                  |
|  16  | [Width](#CommandButton.Width)               | 获取或设置指定对象的宽度                                  |
|  17  | [WidthPolicy](#CommandButton.WidthPolicy)   | 获取或设置指定对象的宽度策略                              |
|  18  | [WordWrap](#CommandButton.WordWrap)         | 指定控件的内容是否在行尾自动换行                          |

## 成员方法

### CommandButton.Move

将指定对象移动到参数指定的坐标

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 CommandButton 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                               |
|----------|-----------|----------|------------------------------------|
| *Left*   | 必选      | Number   | 对象移动后相对于窗口的左边缘的距离 |
| *Top*    | 可选      | Number   | 对象移动后相对于窗口的上边缘的距离 |
| *Width*  | 可选      | Number   | 对象移动后的预期宽度               |
| *Height* | 可选      | Number   | 对象移动后的预期长度               |

**说明**

将指定对象移动到参数指定的坐标

**示例**

``` JavaScript
/*将CommandButton移动到窗体的左上角，并将移动后CommandButton的长宽设置为100*/
function func1()
{
    UserForm1.CommandButton1.Move(100,100,100,100)
}
```

## 成员属性

### CommandButton.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 CommandButton 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*将该对象设置为根据内容自动设置大小*/
function func1()
{
    UserForm1.CommandButton1.AutoSize = true
}
```

### CommandButton.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 CommandButton 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将该CommandButton设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.CommandButton1.BackColor = 0x665898
}
```

### CommandButton.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 CommandButton 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

| 设置        | 值  | 描述                                                                              |
|-------------|-----|-----------------------------------------------------------------------------------|
| Normal      | 1   | （除option group以外的所有控件的默认值）控件内部的颜色是由BackColor属性设置的颜色 |
| Transparent | 0   | （option group的默认值）控件是透明的，控件后面的窗体和报表是可见的                |

**示例**

``` JavaScript
/*将该CommandButton设置为透明*/
function func1()
{
    UserForm1.CommandButton1.BackStyle = 0
}
```

### CommandButton.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 CommandButton 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将该CommandButton中的文本设置为"Button 1"*/
function func1()
{
    UserForm1.CommandButton1.Caption = "Button 1"
}
```

### CommandButton.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 CommandButton 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.CommandButton1.Enabled = true
}
```

### CommandButton.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 CommandButton 对象的变量。

**说明**

指定对象的字体，只读

### CommandButton.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 CommandButton 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将CommandButton中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.CommandButton1.ForeColor = 0x658978
}
```

### CommandButton.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 CommandButton 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将CommandButton控件的高度设置为100*/
function func1()
{
    UserForm1.CommandButton1.Height = 100
}
```

### CommandButton.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 CommandButton 对象的变量。

**说明**

获取或设置指定对像的高度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将CommandButton1的高度设置为固定*/
function func1()
{
    UserForm1.CommandButton1.HeightPolicy = 2
}
```

### CommandButton.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 CommandButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将CommandButton的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.CommandButton1.Left = 100
}
```

### CommandButton.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 CommandButton 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### CommandButton.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 CommandButton 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将CommandButton设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.CommandButton1.TabIndex = 2
}
```

### CommandButton.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 CommandButton 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

| 设置 | 值    | 描述                                            |
|------|-------|-------------------------------------------------|
| Yes  | True  | （默认）您可以通过按下Tab键将焦点移动到该控件上 |
| No   | False | 您无法通过按下Tab键将焦点移动到该控件上         |

**示例**

``` JavaScript
/*将CommandButton设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.CommandButton1.TabStop = true
}
```

### CommandButton.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 CommandButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将CommandButton的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.CommandButton1.Top = 100
}
```

### CommandButton.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 CommandButton 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将CommandButton设置为可见*/
function func1()
{
    UserForm1.CommandButton1.Visible = true
}
```

### CommandButton.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 CommandButton 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将CommandButton的宽度设置为100*/
function func1()
{
    UserForm1.CommandButton1.Width = 100
}
```

### CommandButton.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 CommandButton 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将CommandButton1的宽度设置为固定*/
function func1()
{
    UserForm1.CommandButton1.WidthPolicy = 2
}
```

### CommandButton.WordWrap

指定控件的内容是否在行尾自动换行

**语法**

**express.WordWrap**

\*express是一个代表 CommandButton 对象的变量。

**说明**

指定控件的内容是否在行尾自动换行

**示例**

``` JavaScript
/*将Command Button设置为控件的内容在行尾自动换行*/
function func1()
{
    UserForm1.CommandButton1.WordWrap = true
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/CommandButton](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
