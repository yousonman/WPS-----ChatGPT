# Label 对象

## 简介

用于显示用户不能编辑的文本或图像，例如图形下的标题文本。

## 说明

用于显示用户不能编辑的文本或图像，例如图形下的标题文本。

## 方法

| 序号 | 名称                | 说明                       |
|:----:|:--------------------|:---------------------------|
|  1   | [Move](#Label.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                | 说明                                                      |
|:----:|:------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#Label.AutoSize)         | 指定该对象是否根据内容自动设置它的大小                    |
|  2   | [BackColor](#Label.BackColor)       | 获取或设置指定对象的内部颜色                              |
|  3   | [BackStyle](#Label.BackStyle)       | 您可以使用该属性指定控件是否透明                          |
|  4   | [BorderColor](#Label.BorderColor)   | 您可以通过该属性指定控件边框的颜色                        |
|  5   | [BorderStyle](#Label.BorderStyle)   | 指定控件边框的显示方式                                    |
|  6   | [Caption](#Label.Caption)           | 获取或设置出现在控件中的文本                              |
|  7   | [Enabled](#Label.Enabled)           | 您可以通过该属性设置或返回是否启用该控件                  |
|  8   | [Font](#Label.Font)                 | 指定对象的字体，只读                                      |
|  9   | [ForeColor](#Label.ForeColor)       | 您可以通过该属性指定控件中文本的颜色                      |
|  10  | [Height](#Label.Height)             | 获取或设置指定对象的高度                                  |
|  11  | [HeightPolicy](#Label.HeightPolicy) | 获取或设置指定对像的高度策略                              |
|  12  | [Left](#Label.Left)                 | 您可以通过该属性指定控件在对象或报表中的位置              |
|  13  | [Name](#Label.Name)                 | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  14  | [TabIndex](#Label.TabIndex)         | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  15  | [TabStop](#Label.TabStop)           | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  16  | [TextAlign](#Label.TextAlign)       | 该属性用于指定对象中文本的对齐方式                        |
|  17  | [Top](#Label.Top)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  18  | [Visible](#Label.Visible)           | 返回或设置该对象是否可见                                  |
|  19  | [Width](#Label.Width)               | 获取或设置指定对象的宽度                                  |
|  20  | [WidthPolicy](#Label.WidthPolicy)   | 获取或设置指定对象的宽度策略                              |
|  21  | [WordWrap](#Label.WordWrap)         | 指定控件的内容是否在行尾自动换行                          |

## 成员方法

### Label.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Label 对象的变量。

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
/*将Label移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.Label1.Move(100,100,100,100)
}
```

## 成员属性

### Label.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 Label 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置Label控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.Label1.AutoSize = true
}
```

### Label.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将Label1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.Label1.BackColor = 0x665898
}
```

### Label.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 Label 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

**示例**

``` JavaScript
/*将Label1设置为透明*/
function func1()
{
    UserForm1.Label1.BackStyle = 0
}
```

### Label.BorderColor

您可以通过该属性指定控件边框的颜色

**语法**

**express.BorderColor**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件边框的颜色

**示例**

``` JavaScript
/*将Label1的边框设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.Label1.BorderColor = 0x665898
}
```

### Label.BorderStyle

指定控件边框的显示方式

**语法**

**express.BorderStyle**

\*express是一个代表 Label 对象的变量。

**说明**

指定控件边框的显示方式，可设置为以下值

| 设置   | 值  | 描述       |
|--------|-----|------------|
| None   | 0   | 无边框     |
| Single | 1   | 边框为实线 |

**示例**

``` JavaScript
/*将Label1的边框显示方式设置为实线*/
function func1()
{
    UserForm1.Label1.BorderStyle = 1
}
```

### Label.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将Label1中的文本设置为"Label 1"*/
function func1()
{
    UserForm1.Label1.Caption = "Label 1"
}
```

### Label.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.Label1.Enabled = true
}
```

### Label.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 Label 对象的变量。

**说明**

指定对象的字体，只读

### Label.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将Label1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.Label1.ForeColor = 0x658978
}
```

### Label.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将Label1控件的高度设置为100*/
function func1()
{
    UserForm1.Label1.Height = 100
}
```

### Label.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对像的高度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将Label1的高度设置为固定*/
function func1()
{
    UserForm1.Label1.HeightPolicy = 2
}
```

### Label.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将Label1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.Label1.Left = 100
}
```

### Label.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### Label.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将Label1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.Label1.TabIndex = 2
}
```

### Label.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将Label1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.Label1.TabStop = true
}
```

### Label.TextAlign

该属性用于指定对象中文本的对齐方式

**语法**

**express.TextAlign**

\*express是一个代表 Label 对象的变量。

**说明**

该属性用于指定对象中文本的对齐方式，包含以下值

| 设置   | 值  | 描述                     |
|--------|-----|--------------------------|
| Left   | 1   | 文本数字和日期都左对齐   |
| Center | 2   | 文本数字和日期都居中对齐 |
| Right  | 3   | 文本数字和日期都右对齐   |

**示例**

``` JavaScript
/*将Label1中的文本设置为居中对齐*/
function func1()
{
    UserForm1.Label1.TextAlign = 2
}
```

### Label.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将Label1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.Label1.Top = 100
}
```

### Label.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 Label 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将Label1设置为可见*/
function func1()
{
    UserForm1.Label1.Visible = true
}
```

### Label.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将Label1的宽度设置为100*/
function func1()
{
    UserForm1.Label1.Width = 100
}
```

### Label.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将Label1的宽度设置为固定*/
function func1()
{
    UserForm1.Label1.WidthPolicy = 2
}
```

### Label.WordWrap

指定控件的内容是否在行尾自动换行

**语法**

**express.WordWrap**

\*express是一个代表 Label 对象的变量。

**说明**

指定控件的内容是否在行尾自动换行

**示例**

``` JavaScript
/*将Label1设置为控件的内容在行尾自动换行*/
function func1()
{
    UserForm1.Label1.WordWrap = true
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/Label](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
