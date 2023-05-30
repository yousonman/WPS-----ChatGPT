# Frame 对象

## 简介

可以创建一个图形或控件的功能组。要为多个控件创建组，可先画框架，接着在框架中添加控件。

## 说明

可以创建一个图形或控件的功能组。要为多个控件创建组，可先画框架，接着在框架中添加控件。

## 方法

| 序号 | 名称                | 说明                       |
|:----:|:--------------------|:---------------------------|
|  1   | [Move](#Frame.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                | 说明                                                      |
|:----:|:------------------------------------|:----------------------------------------------------------|
|  1   | [BackColor](#Frame.BackColor)       | 获取或设置指定对象的内部颜色                              |
|  2   | [BorderColor](#Frame.BorderColor)   | 您可以通过该属性指定控件边框的颜色                        |
|  3   | [BorderStyle](#Frame.BorderStyle)   | 指定控件边框的显示方式                                    |
|  4   | [Caption](#Frame.Caption)           | 获取或设置出现在控件中的文本                              |
|  5   | [Enabled](#Frame.Enabled)           | 您可以通过该属性设置或返回是否启用该控件                  |
|  6   | [Font](#Frame.Font)                 | 指定对象的字体，只读                                      |
|  7   | [ForeColor](#Frame.ForeColor)       | 您可以通过该属性指定控件中文本的颜色                      |
|  8   | [Height](#Frame.Height)             | 获取或设置指定对象的高度                                  |
|  9   | [HeightPolicy](#Frame.HeightPolicy) | 获取或设置指定对像的高度策略                              |
|  10  | [Left](#Frame.Left)                 | 您可以通过该属性指定控件在对象或报表中的位置              |
|  11  | [Name](#Frame.Name)                 | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  12  | [TabIndex](#Frame.TabIndex)         | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  13  | [TabStop](#Frame.TabStop)           | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  14  | [Top](#Frame.Top)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  15  | [Visible](#Frame.Visible)           | 返回或设置该对象是否可见                                  |
|  16  | [Width](#Frame.Width)               | 获取或设置指定对象的宽度                                  |
|  17  | [WidthPolicy](#Frame.WidthPolicy)   | 获取或设置指定对象的宽度策略                              |

## 成员方法

### Frame.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Top* , *Left* , *Width* , *Height* )**

\*express是一个代表 Frame 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                               |
|----------|-----------|----------|------------------------------------|
| *Top*    | 必选      | Number   | 对象移动后相对于窗口的左边缘的距离 |
| *Left*   | 可选      | Number   | 对象移动后相对于窗口的上边缘的距离 |
| *Width*  | 可选      | Number   | 对象移动后的预期宽度               |
| *Height* | 可选      | Number   | 对象移动后的预期长度               |

**说明**

将对象移动到参数指定的位置

**示例**

``` JavaScript
/*将Frame1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.Frame1.Move(100,100,100,100)
}
```

## 成员属性

### Frame.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将Frame1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.Frame1.BackColor = 0x665898
}
```

### Frame.BorderColor

您可以通过该属性指定控件边框的颜色

**语法**

**express.BorderColor**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件边框的颜色

**示例**

``` JavaScript
/*将Frame1的边框设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.Frame1.BorderColor = 0x665898
}
```

### Frame.BorderStyle

指定控件边框的显示方式

**语法**

**express.BorderStyle**

\*express是一个代表 Frame 对象的变量。

**说明**

指定控件边框的显示方式，可设置为以下值：

| 设置   | 值  | 描述     |
|--------|-----|----------|
| None   | 0   | 无边框   |
| Single | 1   | 实现边框 |

**示例**

``` JavaScript
/*将Frame1的边框显示方式设置为实线*/
function func1()
{
    UserForm1.Frame1.BorderStyle = 1
}
```

### Frame.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将Frame1中的文本设置为"Frame1"*/
function func1()
{
    UserForm1.Frame1.Caption = "Frame1"
}
```

### Frame.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.Frame1.Enabled = true
}
```

### Frame.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 Frame 对象的变量。

**说明**

指定对象的字体，只读

### Frame.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将Frame1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.Frame1.ForeColor = 0x658978
}
```

### Frame.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将Frame1控件的高度设置为100*/
function func1()
{
    UserForm1.Frame1.Height = 100
}
```

### Frame.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将Frame1的高度设置为固定*/
function func1()
{
    UserForm1.Frame1.HeightPolicy = 2
}
```

### Frame.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将Frame1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.Frame1.Left = 100
}
```

### Frame.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### Frame.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将Frame1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.Frame1.TabIndex = 2
}
```

### Frame.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将Frame1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.Frame1.TabStop = true
}
```

### Frame.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 Frame 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将Frame1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.Frame1.Top = 100
}
```

### Frame.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 Frame 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将Frame1设置为可见*/
function func1()
{
    UserForm1.Frame1.Visible = true
}
```

### Frame.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将Frame1的宽度设置为100*/
function func1()
{
    UserForm1.Frame1.Width = 100
}
```

### Frame.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 Frame 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将Frame1的宽度设置为固定*/
function func1()
{
    UserForm1.Frame1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/Frame](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
