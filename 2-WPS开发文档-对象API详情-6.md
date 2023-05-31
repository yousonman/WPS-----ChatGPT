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
# SpinButton 对象

## 简介

一种与其它控件并用微调控制的控件，也可以用来对一个范围的值或工程列表做向前或向后的遍历。

## 说明

一种与其它控件并用微调控制的控件，也可以用来对一个范围的值或工程列表做向前或向后的遍历。

## 方法

| 序号 | 名称                     | 说明                       |
|:----:|:-------------------------|:---------------------------|
|  1   | [Move](#SpinButton.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                       | 说明                                                      |
|:----:|:-------------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#SpinButton.AutoSize)           | 指定该对象是否根据内容自动设置它的大小                    |
|  2   | [BackColor](#SpinButton.BackColor)         | 获取或设置指定对象的内部颜色                              |
|  3   | [ControlSource](#SpinButton.ControlSource) | 您可以使用该属性指定控件中显示的数据                      |
|  4   | [Enabled](#SpinButton.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                  |
|  5   | [Font](#SpinButton.Font)                   | 指定对象的字体，只读                                      |
|  6   | [ForeColor](#SpinButton.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                      |
|  7   | [Height](#SpinButton.Height)               | 获取或设置指定对象的高度                                  |
|  8   | [HeightPolicy](#SpinButton.HeightPolicy)   | 获取或设置指定对像的高度策略                              |
|  9   | [Left](#SpinButton.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  10  | [Max](#SpinButton.Max)                     | 为该控件的Value属性指定最大可接受值。                     |
|  11  | [Min](#SpinButton.Min)                     | 为该控件的Value属性指定最小可接受值。                     |
|  12  | [Name](#SpinButton.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  13  | [SmallChange](#SpinButton.SmallChange)     | 指定当用户单击 滚动箭头时发生的移动量。                   |
|  14  | [TabIndex](#SpinButton.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  15  | [TabStop](#SpinButton.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  16  | [Top](#SpinButton.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置              |
|  17  | [Value](#SpinButton.Value)                 | 您可以通过该属性指定一个在Max和Min之间的值                |
|  18  | [Visible](#SpinButton.Visible)             | 返回或设置该对象是否可见                                  |
|  19  | [Width](#SpinButton.Width)                 | 获取或设置指定对象的宽度                                  |
|  20  | [WidthPolicy](#SpinButton.WidthPolicy)     | 获取或设置指定对象的宽度策略                              |

## 成员方法

### SpinButton.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 SpinButton 对象的变量。

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
/*将SpinButton1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.SpinButton1.Move(100,100,100,100)
}
```

## 成员属性

### SpinButton.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 SpinButton 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置SpinButton1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.SpinButton1.AutoSize = true
}
```

### SpinButton.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 SpinButton 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将SpinButton1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.SpinButton1.BackColor = 0x665898
}
```

### SpinButton.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 SpinButton 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### SpinButton.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 SpinButton 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.SpinButton1.Enabled = true
}
```

### SpinButton.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 SpinButton 对象的变量。

**说明**

指定对象的字体，只读

### SpinButton.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 SpinButton 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将SpinButton1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.SpinButton1.ForeColor = 0x658978
}
```

### SpinButton.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 SpinButton 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将SpinButton1控件的高度设置为100*/
function func1()
{
    UserForm1.SpinButton1.Height = 100
}
```

### SpinButton.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 SpinButton 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将SpinButton1的高度设置为固定*/
function func1()
{
    UserForm1.SpinButton1.HeightPolicy = 2
}
```

### SpinButton.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 SpinButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将SpinButton1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.SpinButton1.Left = 100
}
```

### SpinButton.Max

为该控件的Value属性指定最大可接受值。

**语法**

**express.Max**

\*express是一个代表 SpinButton 对象的变量。

**说明**

为该控件的Value属性指定最大可接受值。

**示例**

``` JavaScript
/*将SpinButton1的Value属性的最大可接受值设为2000*/
function func1()
{
    UserForm1.SpinButton1.Max = 2000
}
```

### SpinButton.Min

为该控件的Value属性指定最小可接受值。

**语法**

**express.Min**

\*express是一个代表 SpinButton 对象的变量。

**说明**

为该控件的Value属性指定最小可接受值。

**示例**

``` JavaScript
/*将SpinButton1的Value属性的最小可接受值设为0*/
function func1()
{
    UserForm1.SpinButton1.Mix = 0
}
```

### SpinButton.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 SpinButton 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### SpinButton.SmallChange

指定当用户单击 滚动箭头时发生的移动量。

**语法**

**express.SmallChange**

\*express是一个代表 SpinButton 对象的变量。

**说明**

指定当用户单击滚动箭头时发生的移动量。

**示例**

``` JavaScript
/*将SpinButton1在用户点击滚动箭头时发生的移动量设为100*/
function func1()
{
    UserForm1.SpinButton1.LargeChange = 100
}
```

### SpinButton.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 SpinButton 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将SpinButton1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.SpinButton1.TabIndex = 2
}
```

### SpinButton.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 SpinButton 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将SpinButton1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.SpinButton1.TabStop = true
}
```

### SpinButton.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 SpinButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将SpinButton1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.SpinButton1.Top = 100
}
```

### SpinButton.Value

您可以通过该属性指定一个在Max和Min之间的值

**语法**

**express.Value**

\*express是一个代表 SpinButton 对象的变量。

**说明**

您可以通过该属性指定一个在Max和Min之间的值

**示例**

``` JavaScript
/*将SpinButton1的Valued的值为500*/
function func1()
{
    UserForm1.SpinButton1.Value = 500
}
```

### SpinButton.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 SpinButton 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将SpinButton1设置为可见*/
function func1()
{
    UserForm1.SpinButton1.Visible = true
}
```

### SpinButton.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 SpinButton 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将SpinButton1的宽度设置为100*/
function func1()
{
    UserForm1.SpinButton1.Width = 100
}
```

### SpinButton.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 SpinButton 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可以设置为以下值：

| 设置      | 值  | 描述       |
|-----------|-----|------------|
| Expanding | 1   | 柯柏年宽度 |
| Fixed     | 2   | 固定宽度   |

**示例**

``` JavaScript
/*将SpinButton1的宽度设置为固定*/
function func1()
{
    UserForm1.SpinButton1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/SpinButton](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Interior 对象

## 简介

{\_description}

## 方法

| 序号 | 名称                         | 说明 |
|:----:|:-----------------------------|:-----|
|  1   | [Reserve](#Interior.Reserve) |      |

## 属性

| 序号 | 名称                         | 说明 |
|:----:|:-----------------------------|:-----|
|  1   | [Reserve](#Interior.Reserve) |      |

## 成员方法

### Interior.Reserve

**语法**

**express.Reserve ()**

\*express是一个代表 Interior 对象的变量。

**示例**

``` JavaScript
{#_examples}
```

## 成员属性

### Interior.Reserve

**语法**

**express.Reserve**

\*express是一个代表 Interior 对象的变量。

**示例**

``` JavaScript
{#_examples}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/枚举/Interior](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AddIn 对象

## 简介

代表单个加载宏，不论该加载宏是否已加载。

## 说明

AddIn 对象是 <span style="color:#000000;text-decoration-line:none;">AddIns</span> 集合的成员。 AddIns 集合包含 ET 所有可用加载宏的列表（不论这些加载宏是否已安装）。此列表与 “加载宏” 对话框中显示的加载宏列表对应。

示例 使用 AddIns ( *index* )（其中 *index* 是加载宏标题或索引号）可返回单个 AddIn 对象。下例安装“分析工具库”加载宏。

``` JavaScript
Application.AddIns.Item("analysis toolpak").Installed = true
```

请勿混淆加载宏标题（出现在 “加载宏” 对话框中）与加载宏名称（加载宏的文件名）。必须严格按照 “加载宏” 对话框中的标题书写加载宏标题，但不必匹配大小写。

加载宏索引号代表加载宏在 “加载宏” 对话框内 “可用加载宏” 框中的位置。下例创建一个列表，包含可用加载宏的指定属性。

``` JavaScript
function test()
{
    let sheet = Application.Worksheets.Item("Sheet1")
    sheet.Rows.Item(1).Font.Bold = true
    sheet.Range("a1:d1").Value2 = ["Name", "Full Name", "Title", "Installed"]
    for(let i = 1; i <= Application.AddIns.Count; i++) {
        sheet.Cells(i + 1, 1).Value2 = Application.AddIns.Item(i).Name
        sheet.Cells(i + 1, 2).Value2 = Application.AddIns.Item(i).FullName
        sheet.Cells(i + 1, 3).Value2 = Application.AddIns.Item(i).Title
        sheet.Cells(i + 1, 4).Value2 = Application.AddIns.Item(i).Installed
    }
    sheet.Range("a1").CurrentRegion.Columns.AutoFit()
}
```

Add 方法将加载宏添加到可用加载宏列表中，但不安装加载宏。将加载宏的 Installed 属性设为 True 可安装加载宏。要安装可用加载宏列表中没有的加载宏，必须先使用 Add 方法，然后设置 Installed 属性。此操作一步即可完成，如下例中所示（注意， Add 方法中应使用加载宏的名称，而不使用标题）。

``` JavaScript
Application.AddIns.Add("generic.xll").Installed = true
```

使用 Workbooks ( *index* )（其中 *index* 是加载宏文件名而非标题）可返回对与某一加载宏相对应的工作簿的引用。因为加载宏通常不出现在 Workbooks 集合中，所以必须使用其文件名来指定。此示例将变量 *wb* 设置为 Myaddin.xla 的工作簿。

``` JavaScript
let wb = Application.Workbooks.Item("myaddin.xla")
```

下例将变量 *wb* 设置为“分析工具库”加载宏的工作簿。

``` JavaScript
let wb = Application.Workbooks.Item(Application.AddIns.Item("analysis toolpak").Name)
```

如果 Installed 属性为 True ，但调用加载宏中的函数仍旧失败，那么可能并未真正地加载了该加载宏。这是因为 Addin 对象代表了加载宏的存在及安装状态，但并不代表加载宏工作簿中的实际内容。为保证已安装的加载宏被加载，应当打开该加载宏工作簿。下例中，如果加载宏“My Addin”未出现在 Workbooks 集合中，就打开该加载宏。

``` JavaScript
function test()
{
    try {    
        // turn off error checking
        let wbMyAddin = Application.Workbooks.Item(Application.AddIns.Item("My Addin").Name)
        let lastError = Application.Err
    }
    catch(exception) {
        if(lastError != 0) {
            // the add-in workbook isn't currently open. Manually open it.
            let wbMyAddin = Application.Workbooks.Open(Application.AddIns.Item("My Addin").FullName)
        }
    }
}
```

## 属性

| 序号 | 名称                              | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#AddIn.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [CLSID](#AddIn.CLSID)             | 返回一个只读的唯一标识符，或识别对象的 CLSID。 String 类型。                                                                                                                                                                    |
|  3   | [Creator](#AddIn.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [FullName](#AddIn.FullName)       | 返回对象的名称（以字符串表示），包括其磁盘路径。 String 型，只读。                                                                                                                                                              |
|  5   | [Installed](#AddIn.Installed)     | 如果安装了此加载宏，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                 |
|  6   | [IsOpen](#AddIn.IsOpen)           | 如果加载项当前打开，则返回 True 。只读 Boolean 类型                                                                                                                                                                             |
|  7   | [Name](#AddIn.Name)               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  8   | [Parent](#AddIn.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  9   | [Path](#AddIn.Path)               | 返回一个 String 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。                                                                                                                                                |
|  10  | [progID](#AddIn.progID)           | 返回对象的程序标识符。 String 型，只读。                                                                                                                                                                                        |

## 成员属性

### AddIn.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 AddIn 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
  let myObject = Application.ActiveWorkbook
  if(myObject.Application.Value == "ET") {
      alert("This is an ET Application object.")
  }
  else {
      alert("This is not an ET Application object.")
  }
}
```

### AddIn.CLSID

返回一个只读的唯一标识符，或识别对象的 CLSID。 **String** 类型。

**语法**

**express.CLSID**

\*express是一个代表 AddIn 对象的变量。

### AddIn.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 AddIn 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### AddIn.FullName

返回对象的名称（以字符串表示），包括其磁盘路径。 **String** 型，只读。

**语法**

**express.FullName**

\*express是一个代表 AddIn 对象的变量。

**示例**

``` JavaScript
/*此示例显示每一个可用加载宏的路径及文件名。*/
function test()
{
  for(let i = 1; i <= Application.AddIns.Count; i++) {
      alert(Application.AddIns.Item(i).FullName)
  }
}
```

### AddIn.Installed

如果安装了此加载宏，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Installed**

\*express是一个代表 AddIn 对象的变量。

**说明**

将此属性设为 **True** 将安装指定加载宏，并调用 Auto_Add 函数。将此属性设为 **False** 则移去指定加载宏，并调用 Auto_Remove 函数。

**示例**

``` JavaScript
/*本示例用一个消息框显示 Solver 加载宏的安装状态。*/
function test()
{
  let a = Application.AddIns.Item("Solver Add-In")
  if(a.Installed == true) {
      alert("The Solver add-in is installed")
  }
  else {
      alert("The Solver add-in is not installed")
  }
}
```

### AddIn.IsOpen

如果加载项当前打开，则返回 **True** 。只读 **Boolean** 类型

**语法**

**express.IsOpen**

\*express是一个代表 AddIn 对象的变量。

### AddIn.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 AddIn 对象的变量。

### AddIn.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 AddIn 对象的变量。

### AddIn.Path

返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。

**语法**

**express.Path**

\*express是一个代表 AddIn 对象的变量。

### AddIn.progID

返回对象的程序标识符。 **String** 型，只读。

**语法**

**express.progID**

\*express是一个代表 AddIn 对象的变量。

**示例**

``` JavaScript
/*此示例为第一张工作表中所有 OLE 对象创建程序标识符列表。*/
function test()
{
  let rw = 0
  let o = Application.Worksheets.Item(1).OLEObjects()                               
  for(let i = 1; i <= o.Count; i++) {
      rw = rw + 1
      Application.Worksheets.Item(2).Cells.Item(rw, 1).Value2 = o.Item(i).ProgId
  }
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AddIn](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# ColorStop 对象

## 简介

表示某一区域或选定内容中渐变填充的光圈点。

## 说明

ColorStop 对象可用来设置单元格填充属性，包括 Color 、 ThemeColor 和 TintAndShade 属性。

``` JavaScript
/*以下示例演示如何对 ColorStop 对象应用属性。*/
function test(){
let myInterior = Application.Selection.Interior
    myInterior.Pattern = xlPatternLinearGradient
    myInterior.Gradient.Degree = 135
    myInterior.Gradient.ColorStops.Clear()

let myColorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)
    myColorStops1.ThemeColor = xlThemeColorDark1
    myColorStops1.TintAndShade = 0

let myColorStops2 = Application.Selection.Interior.Gradient.ColorStops.Add(0.5)
    myColorStops2.ThemeColor = xlThemeColorAccent1
    myColorStops2.TintAndShade = 0

let myColorStops3 = Application.Selection.Interior.Gradient.ColorStops.Add(1)
    myColorStops3.ThemeColor = xlThemeColorDark1
    myColorStops3.TintAndShade = 0
}
```

## 方法

| 序号 | 名称                        | 说明           |
|:----:|:----------------------------|:---------------|
|  1   | [Delete](#ColorStop.Delete) | 删除指定对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                        |
|:----:|:----------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ColorStop.Application)   | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 Application 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 Application 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
|  2   | [Color](#ColorStop.Color)               | 返回或设置指定对象的颜色。可读/写。                                                                                                                                                                                         |
|  3   | [Creator](#ColorStop.Creator)           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                  |
|  4   | [Parent](#ColorStop.Parent)             | 返回指定对象的父对象。只读。                                                                                                                                                                                                |
|  5   | [Position](#ColorStop.Position)         | 返回或设置 ColorStop 对象的位置。可读/写。                                                                                                                                                                                  |
|  6   | [ThemeColor](#ColorStop.ThemeColor)     | 返回或设置指定对象的主颜色。可读/写。                                                                                                                                                                                       |
|  7   | [TintAndShade](#ColorStop.TintAndShade) | 返回或设置指定对象的淡色和阴影。可读/写。                                                                                                                                                                                   |

## 成员方法

### ColorStop.Delete

删除指定对象。

**语法**

**express.Delete ()**

\*express是一个代表 ColorStop 对象的变量。

## 成员属性

### ColorStop.Application

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 ColorStop 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
}
```

### ColorStop.Color

返回或设置指定对象的颜色。可读/写。

**语法**

**express.Color**

\*express是一个代表 ColorStop 对象的变量。

### ColorStop.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ColorStop 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。Creator 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ColorStop.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ColorStop 对象的变量。

### ColorStop.Position

返回或设置 **ColorStop** 对象的位置。可读/写。

**语法**

**express.Position**

\*express是一个代表 ColorStop 对象的变量。

### ColorStop.ThemeColor

返回或设置指定对象的主颜色。可读/写。

**语法**

**express.ThemeColor**

\*express是一个代表 ColorStop 对象的变量。

**示例**

``` JavaScript
/*对当前选定内容应用主题颜色。*/

function test(){
Range("A1:A10").Select()
let myColorStop = Application.Selection.Interior.Gradient.ColorStops.Add(1)
    myColorStop.ThemeColor = xlThemeColorAccent1
    myColorStop.TintAndShade = 0
}
```

### ColorStop.TintAndShade

返回或设置指定对象的淡色和阴影。可读/写。

**语法**

**express.TintAndShade**

\*express是一个代表 ColorStop 对象的变量。

**示例**

``` JavaScript
/*对当前选定内容应用淡色和阴影。*/
function test(){
Range("A1:A10").Select()
let myColorStop = Application.Selection.Interior.Gradient.ColorStops.Add(1)
    myColorStop.ThemeColor = xlThemeColorAccent1
    myColorStop.TintAndShade = 0
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ColorStop](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomViews 对象

## 简介

由自定义工作簿视图组成的集合。

## 说明

每个视图由一个 CustomView 对象代表。

使用 CustomViews 属性可返回 CustomViews 集合。使用 Add 方法可新建自定义视图，并将它添加到 CustomViews 集合中。

``` JavaScript
/*下列新建一个名为“Summary”的自定义视图。*/
Application.ActiveWorkbook.CustomViews.Add("Summary", true, true)
```

## 方法

| 序号 | 名称                      | 说明                   |
|:----:|:--------------------------|:-----------------------|
|  1   | [Add](#CustomViews.Add)   | 新建自定义视图。       |
|  2   | [Item](#CustomViews.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CustomViews.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#CustomViews.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#CustomViews.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#CustomViews.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### CustomViews.Add

新建自定义视图。

**语法**

**express.Add ( *ViewName* , *PrintSettings* , *RowColSettings* )**

\*express是一个代表 CustomViews 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                    |
|------------------|-----------|----------|-------------------------------------------------------------------------|
| *ViewName*       | 必选      | String   | 新视图的名称。                                                          |
| *PrintSettings*  | 可选      | Variant  | 如果为 True，则在自定义视图中包括打印设置。                             |
| *RowColSettings* | 可选      | Variant  | 如果为 True，则在自定义视图中包括隐藏行和隐藏列（包括筛选信息）的设置。 |

**返回值**

一个 CustomView 对象，它代表新的自定义视图。

**示例**

``` JavaScript
/*此示例在活动工作簿中新建一个自定义视图，并命名为“Summary”。*/
Application.ActiveWorkbook.CustomViews.Add("Summary", true, true)
```

### CustomViews.Item

从集合中返回一个对象。

**语法**

**express.Item ( *ViewName* )**

\*express是一个代表 CustomViews 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                 |
|------------|-----------|----------|----------------------|
| *ViewName* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 CustomView 对象。

**示例**

``` JavaScript
/*本示例在判断名为“Current Inventory”的自定义视图中是否包含打印设置。*/
alert(Application.ActiveWorkbook.CustomViews.Item("Current Inventory").PrintSettings == true)
```

## 成员属性

### CustomViews.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomViews 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
}
```

### CustomViews.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 CustomViews 对象的变量。

### CustomViews.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CustomViews 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### CustomViews.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomViews 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/CustomViews](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DataTable 对象

## 简介

代表一张图表模拟运算表。

## 说明

使用 DataTable 属性可返回 DataTable 对象。

``` JavaScript
/*下例向嵌入式图表中添加带有外边框的模拟运算表。*/
function test(){
let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    myChart.HasDataTable = true
    myChart.DataTable.HasBorderOutline = true
}
```

## 方法

| 序号 | 名称                        | 说明       |
|:----:|:----------------------------|:-----------|
|  1   | [Delete](#DataTable.Delete) | 删除对象。 |
|  2   | [Select](#DataTable.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                                  | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#DataTable.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Border](#DataTable.Border)                           | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  3   | [Creator](#DataTable.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Font](#DataTable.Font)                               | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                      |
|  5   | [Format](#DataTable.Format)                           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  6   | [HasBorderHorizontal](#DataTable.HasBorderHorizontal) | 如果图表模拟运算表具有水平单元格边框，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                               |
|  7   | [HasBorderOutline](#DataTable.HasBorderOutline)       | 如果图表模拟运算表具有外部边框，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                     |
|  8   | [HasBorderVertical](#DataTable.HasBorderVertical)     | 如果图表模拟运算表具有垂直单元格边框，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                               |
|  9   | [Parent](#DataTable.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  10  | [ShowLegendKey](#DataTable.ShowLegendKey)             | 如果数据标签图例项标示可见，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                 |

## 成员方法

### DataTable.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 DataTable 对象的变量。

### DataTable.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 DataTable 对象的变量。

## 成员属性

### DataTable.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 DataTable 对象的变量。

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

### DataTable.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 DataTable 对象的变量。

### DataTable.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DataTable 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### DataTable.Font

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 DataTable 对象的变量。

### DataTable.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 DataTable 对象的变量。

### DataTable.HasBorderHorizontal

如果图表模拟运算表具有水平单元格边框，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasBorderHorizontal**

\*express是一个代表 DataTable 对象的变量。

**示例**

``` JavaScript
/*本示例使嵌入图表模拟运算表显示时带有外边框，但无单元格边框。*/
function test(){
let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    myChart.HasDataTable = true
    let myDataTable = myChart.DataTable
        myDataTable.HasBorderHorizontal = false
        myDataTable.HasBorderVertical = false
        myDataTable.HasBorderOutline = true
}
```

### DataTable.HasBorderOutline

如果图表模拟运算表具有外部边框，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasBorderOutline**

\*express是一个代表 DataTable 对象的变量。

**示例**

``` JavaScript
/*本示例使嵌入图表模拟运算表显示时带有外边框，但无单元格边框。*/
function test(){
let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    myChart.HasDataTable = true
    let myDataTable = myChart.DataTable
        myDataTable.HasBorderHorizontal = false
        myDataTable.HasBorderVertical = false
        myDataTable.HasBorderOutline = true
}
```

### DataTable.HasBorderVertical

如果图表模拟运算表具有垂直单元格边框，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasBorderVertical**

\*express是一个代表 DataTable 对象的变量。

**示例**

``` JavaScript
/*本示例使嵌入图表模拟运算表显示时带有外边框，但无单元格边框。*/
function test(){
let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    myChart.HasDataTable = true
    let myDataTable = myChart.DataTable
        myDataTable.HasBorderHorizontal = false
        myDataTable.HasBorderVertical = false
        myDataTable.HasBorderOutline = true
}
```

### DataTable.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DataTable 对象的变量。

### DataTable.ShowLegendKey

如果数据标签图例项标示可见，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowLegendKey**

\*express是一个代表 DataTable 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/DataTable](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# HiLoLines 对象

## 简介

代表图表组中的高低点连线。

## 说明

高低点连线连接图表组内每一分类中的最高数据点和最底数据点。只有二维折线图组可以有高低点连线。此对象不是集合。没有代表单个高低点连线的对象；或者打开图表组中所有数据点的高低点连线，或者将其全部关闭。

如果 HasHiLoLines 属性是 False ， HiLoLines 对象的大多数属性都会被禁用。

使用 HiLoLines 属性可返回 HiLoLines 对象。

``` JavaScript
/*下例使用 HasHiLowLines 属性将 HiLowLines 添加到工作表一上的嵌入图表一（该图表必须是折线图）。之后，该示例会将高低点连线设为蓝色。*/
function test(){
Application.Worksheets.Item(1).ChartObjects(1).Activate
ActiveChart.ChartGroups(1).HasHiLoLines = true
ActiveChart.ChartGroups(1).HiLoLines.Border.Color = (0, 0, 255)
}
```

## 方法

| 序号 | 名称                        | 说明       |
|:----:|:----------------------------|:-----------|
|  1   | [Delete](#HiLoLines.Delete) | 删除对象。 |
|  2   | [Select](#HiLoLines.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#HiLoLines.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Border](#HiLoLines.Border)           | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  3   | [Creator](#HiLoLines.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Format](#HiLoLines.Format)           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  5   | [Name](#HiLoLines.Name)               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  6   | [Parent](#HiLoLines.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### HiLoLines.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 HiLoLines 对象的变量。

**返回值**

Variant

### HiLoLines.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 HiLoLines 对象的变量。

**返回值**

Variant

## 成员属性

### HiLoLines.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 HiLoLines 对象的变量。

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

### HiLoLines.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 HiLoLines 对象的变量。

### HiLoLines.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 HiLoLines 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### HiLoLines.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 HiLoLines 对象的变量。

### HiLoLines.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 HiLoLines 对象的变量。

### HiLoLines.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 HiLoLines 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/HiLoLines](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LegendKey 对象

## 简介

代表图表图例中的图例标示。

## 说明

每个图例项标示都是一个图形，它将图例项标示以及与之相关联的图表中的系列或趋势线链接起来。图例项标示链接到与之相关联的系列或趋势线的方式为：更改一个的格式会同时更改另一个的格式。

使用 LegendKey 属性可返回 LegendKey 对象。下例更改名为“Sheet1”的工作表上嵌入式图表一的图例顶部的图例项标示的标志背景色。这样会同时更改与该图例项相关的系列中每一个数据点的格式。相关的系列必须支持数据标志。

``` JavaScript
Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).LegendKey.MarkerBackgroundColorIndex = 5
```

## 方法

| 序号 | 名称                                    | 说明                 |
|:----:|:----------------------------------------|:---------------------|
|  1   | [ClearFormats](#LegendKey.ClearFormats) | 清除对象的格式设置。 |
|  2   | [Delete](#LegendKey.Delete)             | 删除对象。           |

## 属性

| 序号 | 名称                                                                | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#LegendKey.Application)                               | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#LegendKey.Creator)                                       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Format](#LegendKey.Format)                                         | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  4   | [Height](#LegendKey.Height)                                         | 返回一个 Double 值，它代表对象的高度（以磅为单位）。                                                                                                                                                                            |
|  5   | [InvertIfNegative](#LegendKey.InvertIfNegative)                     | 如果指定项与一个负数相对应时 ET 就将其反色，则为 True 。 Boolean 类型，可读写。                                                                                                                                                 |
|  6   | [Left](#LegendKey.Left)                                             | 返回一个 Double 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。                                                                                                                                                      |
|  7   | [MarkerBackgroundColor](#LegendKey.MarkerBackgroundColor)           | 将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。 Long 型，可读写。                                                                                                                       |
|  8   | [MarkerBackgroundColorIndex](#LegendKey.MarkerBackgroundColorIndex) | 返回或设置数据标志的背景色，表示为当前调色板中的索引或下列 XlColorIndex 常量之一： xlColorIndexAutomatic 或 xlColorIndexNone 。仅适用于折线图、散点图和雷达图。 Long 类型，可读写。                                             |
|  9   | [MarkerForegroundColor](#LegendKey.MarkerForegroundColor)           | 将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。 Long 型，可读写。                                                                                                                       |
|  10  | [MarkerForegroundColorIndex](#LegendKey.MarkerForegroundColorIndex) | 返回或设置数据标志的前景色，表示为当前调色板中的索引或下列 XlColorIndex 常量之一： xlColorIndexAutomatic 或 xlColorIndexNone 。仅适用于折线图、散点图和雷达图。 Long 类型，可读写。                                             |
|  11  | [MarkerSize](#LegendKey.MarkerSize)                                 | 返回或设置数据标志的大小，以 磅?（磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。） 为单位。可以是 2 到 72 之间的一个值。 Long 型，可读写。                                                    |
|  12  | [MarkerStyle](#LegendKey.MarkerStyle)                               | 返回或设置折线图、散点图或雷达图中数据点或数据系列的数据标志样式。 XlMarkerStyle 类型，可读写。                                                                                                                                 |
|  13  | [Parent](#LegendKey.Parent)                                         | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  14  | [PictureType](#LegendKey.PictureType)                               | 返回或设置一个 XlChartPictureType 值，它代表图片在图例项标示上的显示方式。                                                                                                                                                      |
|  15  | [PictureUnit2](#LegendKey.PictureUnit2)                             | 如果 PictureType 属性设置为 xlStackScale ，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 Double 类型。                                                                                                            |
|  16  | [Shadow](#LegendKey.Shadow)                                         | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  17  | [Smooth](#LegendKey.Smooth)                                         | 如果图例项标示有平滑线，则为 True 。可读写。                                                                                                                                                                                    |
|  18  | [Top](#LegendKey.Top)                                               | 返回一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                            |
|  19  | [Width](#LegendKey.Width)                                           | 返回一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                            |

## 成员方法

### LegendKey.ClearFormats

清除对象的格式设置。

**语法**

**express.ClearFormats ()**

\*express是一个代表 LegendKey 对象的变量。

**返回值**

Variant

### LegendKey.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 LegendKey 对象的变量。

**返回值**

Variant

**说明**

删除 **LegendKey** 对象将删除整个数据系列。

## 成员属性

### LegendKey.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 LegendKey 对象的变量。

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

### LegendKey.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 LegendKey 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### LegendKey.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.Height

返回一个 **Double** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.InvertIfNegative

如果指定项与一个负数相对应时 ET 就将其反色，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.InvertIfNegative**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.Left

返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.MarkerBackgroundColor

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。 **Long** 型，可读写。

**语法**

**express.MarkerBackgroundColor**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.MarkerBackgroundColorIndex

返回或设置数据标志的背景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一： **xlColorIndexAutomatic** 或 **xlColorIndexNone** 。仅适用于折线图、散点图和雷达图。 **Long** 类型，可读写。

**语法**

**express.MarkerBackgroundColorIndex**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.MarkerForegroundColor

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。 **Long** 型，可读写。

**语法**

**express.MarkerForegroundColor**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.MarkerForegroundColorIndex

返回或设置数据标志的前景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一： **xlColorIndexAutomatic** 或 **xlColorIndexNone** 。仅适用于折线图、散点图和雷达图。 **Long** 类型，可读写。

**语法**

**express.MarkerForegroundColorIndex**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.MarkerSize

返回或设置数据标志的大小，以 磅?（磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。） 为单位。可以是 2 到 72 之间的一个值。 **Long** 型，可读写。

**语法**

**express.MarkerSize**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.MarkerStyle

返回或设置折线图、散点图或雷达图中数据点或数据系列的数据标志样式。 **XlMarkerStyle** 类型，可读写。

**语法**

**express.MarkerStyle**

\*express是一个代表 LegendKey 对象的变量。

**说明**

|                                                         |
|---------------------------------------------------------|
| **XlMarkerStyle** 可为以下 **XlMarkerStyle** 常量之一。 |
| **xlMarkerStyleAutomatic** 。自动设置标记               |
| **xlMarkerStyleCircle** 。圆形标记                      |
| **xlMarkerStyleDash** 。长条形标记                      |
| **xlMarkerStyleDiamond** 。菱形标记                     |
| **xlMarkerStyleDot** 。短条形标记                       |
| **xlMarkerStyleNone** 。无标记                          |
| **xlMarkerStylePicture** 。图片标记                     |
| **xlMarkerStylePlus** 。带加号的方形标记                |
| **xlMarkerStyleSquare** 。方形标记                      |
| **xlMarkerStyleStar** 。带星号的方形标记                |
| **xlMarkerStyleTriangle** 。三角形标记                  |
| **xlMarkerStyleX** 。带 X 记号的方形标记                |

### LegendKey.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.PictureType

返回或设置一个 **XlChartPictureType** 值，它代表图片在图例项标示上的显示方式。

**语法**

**express.PictureType**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.PictureUnit2

如果 **PictureType** 属性设置为 **xlStackScale** ，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 **Double** 类型。

**语法**

**express.PictureUnit2**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.Smooth

如果图例项标示有平滑线，则为 **True** 。可读写。

**语法**

**express.Smooth**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.Top

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 LegendKey 对象的变量。

### LegendKey.Width

返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 LegendKey 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/LegendKey](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotCaches 对象

## 简介

代表工作簿中数据透视表的内存缓存的集合。

## 说明

每一个内存缓存都由一个 PivotCache 对象代表。

使用 PivotCaches 方法可返回 PivotCaches 集合。下例为活动工作簿中的所有内存缓存设置 RefreshOnFileOpen 属性。

``` JavaScript
function test(){for (let i=1;i<=ActiveWorkbook.PivotCaches.Count;i++) {
    ActiveWorkbook.PivotCaches().Item(i).RefreshOnFileOpen = true
}
}
```

## 方法

| 序号 | 名称                          | 说明                   |
|:----:|:------------------------------|:-----------------------|
|  1   | [Create](#PivotCaches.Create) | 创建新的 PivotCache。  |
|  2   | [Item](#PivotCaches.Item)     | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#PivotCaches.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#PivotCaches.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#PivotCaches.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#PivotCaches.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### PivotCaches.Create

创建新的 PivotCache。

**语法**

**express.Create ( *SourceType* , *SourceData* , *Version* )**

\*express是一个代表 PivotCaches 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型               | 说明                                                                                               |
|--------------|-----------|------------------------|----------------------------------------------------------------------------------------------------|
| *SourceType* | 必选      | XlPivotTableSourceType | SourceType 可以为以下 xlPivotTableSourceType 常量之一：xlConsolidation、xlDatabase 或 xlExternal。 |
| *SourceData* | 可选      | Variant                | 新数据透视表缓存的数据。                                                                           |
| *Version*    | 可选      | Variant                | 数据透视表的版本。可以为 xlPivotTableVersionList 常量之一。                                        |

**返回值**

PivotCache

**说明**

使用此方法创建 PivotCache 时，不支持以下两个 **xlPivotTableSourceType** 常量： **xlPivotTable** 和 **xlScenario** 。如果提供这两个常量之一，将返回运行时错误。

如果 *SourceType* 不为 **xlExternal** ，则 *SourceData* 参数为必需。该参数可以为 **Range** 对象（当 *SourceType* 为 **xlConsolidation** 或 **xlDatabase** 时），或为 ET 工作簿连接对象（当 *SourceType* 为 **xlExternal** 时）。

如果不提供数据透视表的版本，则版本默认为 **xlPivotTableVersion12** 。不允许使用 **xlPivotTableVersionCurrent** 常量，如果提供该常量，将返回运行时错误。

### PivotCaches.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 PivotCaches 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 PivotCache 对象。

**示例**

``` JavaScript
/*本示例刷新第一个缓存。*/

ActiveWorkbook.PivotCaches().Item(1).Refresh()
```

## 成员属性

### PivotCaches.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotCaches 对象的变量。

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

### PivotCaches.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 PivotCaches 对象的变量。

### PivotCaches.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotCaches 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotCaches.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotCaches 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotCaches](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotFormulas 对象

## 简介

代表数据透视表的公式的集合。每个公式都由一个 PivotFormula 对象代表。

## 说明

本对象及其相关属性和方法对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效，这是因为它不支持计算字段和计算项。

使用 PivotFormulas 属性可返回 PivotFormulas 集合。下例为活动工作表上的第一张数据透视表创建一个数据透视表公式列表。

``` JavaScript
let r = 1
for(let i=1;i <= ActiveSheet.PivotTables(1).PivotFormulas.Count;i++) {
    Cells.Item(r, 1).Value2 = ActiveSheet.PivotTables(1).PivotFormulas.Item(i).Formula
    r++
}
```

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormulas.Add">Add</a></td>
<td style="text-align: left;" data-valian="middle"><p>新建数据透视表公式。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormulas.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormulas.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>从集合中返回一个对象。</p></td>
</tr>
</tbody>
</table>

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormulas.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormulas.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 Long 值，它代表集合中对象的数量。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormulas.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormulas.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定对象的父对象。只读。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotFormulas.Add

新建数据透视表公式。

**语法**

**express.Add ( *Formula* , *UseStandardFormula* )**

\*express是一个代表 PivotFormulas 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                 |
|----------------------|-----------|----------|----------------------|
| *Formula*            | 必选      | String   | 新的数据透视表公式。 |
| *UseStandardFormula* | 可选      | Variant  | 标准数据透视表公式。 |

**返回值**

一个 PivotFormula 对象，它代表新的数据透视表公式。

**示例**

此示例在第一张工作表上为第一个数据透视表创建一个新的数据透视表公式。

``` JavaScript
Worksheets.Item(1).PivotTables(1).PivotFormulas.Add("Year['1998'] Apples = (Year['1997'] Apples) * 2")
```

### PivotFormulas.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator ()**

\*express是一个代表 PivotFormulas 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotFormulas.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 PivotFormulas 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 PivotFormula 对象。

**示例**

本示例显示第一张工作表中第一个数据透视表的第一个公式。

``` JavaScript
MsgBox(Worksheets.Item(1).PivotTables(1).PivotFormulas.Item(1).Formula)
```

## 成员属性

### PivotFormulas.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotFormulas 对象的变量。

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

### PivotFormulas.Count

table { word-break:break-all; }

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 PivotFormulas 对象的变量。

### PivotFormulas.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotFormulas 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotFormulas.Parent

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotFormulas 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotFormulas](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotLayout 对象

## 简介

代表数据透视图报表中字段的位置。

## 说明

使用 PivotLayout 属性可返回一个 PivotLayout 对象。下例创建在第一个数据透视图报表中所使用的数据透视表字段名称的列表。

``` JavaScript
function ListFieldNames() {
   let objNewSheet = Worksheets.Add()
   let intRow = 1
   for(let i=1;i <= Charts.Item("Chart1").PivotLayout.PivotFields().Count;i++) {         
        objNewSheet.Cells.Item(intRow, 1).Value2 = Charts.Item("Chart1").PivotLayout.PivotFields(i).Caption   
        intRow++
   } 
}
```

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLayout.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLayout.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLayout.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLayout.PivotTable">PivotTable</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>PivotTable</span> 对象，它代表与数据透视图相关联的数据透视表。</p></td>
</tr>
</tbody>
</table>

## 成员属性

### PivotLayout.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotLayout 对象的变量。

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

### PivotLayout.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotLayout 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotLayout.Parent

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotLayout 对象的变量。

### PivotLayout.PivotTable

返回一个 **PivotTable** 对象，它代表与数据透视图相关联的数据透视表。

**语法**

**express.PivotTable**

\*express是一个代表 PivotLayout 对象的变量。

**示例**

此示例将 Sheet1 上的数据透视表的当前页设置为名为“Canada”的页。

``` JavaScript
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
pvtTable.PivotFields("Country").CurrentPage = "Canada"
```

此示例确定与“Sales”图表相关联的数据透视表，然后将名为“Oregon”的页设置为数据透视表的当前页。

``` JavaScript
let objPT = Charts.Item("Sales").PivotLayout.PivotTable
objPT.PivotFields("State").CurrentPageName = "Oregon"
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotLayout](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotTable 对象

## 简介

代表工作表上的数据透视表。

## 说明

PivotTable 对象是 PivotTables 集合的成员。 PivotTables 集合包含某一张工作表上的所有 PivotTable 对象。

因为对数据透视表进行编程可能会很复杂，所以，最方便的做法是将数据透视表操作录制到宏中，然后再修订所录制的宏代码。

使用 PivotTables ( *index* )（其中 *index* 是数据透视表索引号或名称）可返回一个 PivotTable 对象。下例使 Sheet3 上第一张数据透视表中的字段“Year”成为行字段。

``` JavaScript
Worksheets.Item("Sheet3").PivotTables(1).PivotFields("Year").Orientation = xlRowField
```

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.AddDataField">AddDataField</a></td>
<td style="text-align: left;" data-valian="middle"><p>将数据字段添加到数据透视表中。返回一个 <span>PivotField</span> 对象，该对象表示新的数据字段。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.AddFields">AddFields</a></td>
<td style="text-align: left;" data-valian="middle"><p>向数据透视表或数据透视图中添加行字段、列字段和页字段。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.AllocateChanges">AllocateChanges</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>对基于 OLAP 数据源的数据透视表中所有编辑过的单元格执行回写操作。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CalculatedFields">CalculatedFields</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>CalculatedFields</span> 集合，该集合表示指定数据透视表中的所有计算字段。只读。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ChangeConnection">ChangeConnection</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>更改指定 <span>PivotTable</span> 的连接。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ChangePivotCache">ChangePivotCache</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>更改指定 <span>PivotTable</span> 的 <span>PivotCache</span> 。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ClearAllFilters">ClearAllFilters</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>ClearAllFilters 方法将删除当前应用于数据透视表的所有筛选。包括删除 PivotTable 对象的 PivotFilters 集合中的所有筛选，删除应用的所有手动筛选以及将报表筛选区域中的所有筛选字段设置为默认项。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ClearTable">ClearTable</a></td>
<td style="text-align: left;" data-valian="middle"><p>ClearTable 方法用于清除数据透视表。清除数据透视表包括删除所有字段以及删除应用于数据透视表的所有筛选和排序。此方法将数据透视表重置为该表刚创建而尚未添加任何字段时的状态。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CommitChanges">CommitChanges</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>对基于 OLAP 数据源的数据透视表的数据源执行提交操作。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ConvertToFormulas">ConvertToFormulas</a></td>
<td style="text-align: left;" data-valian="middle"><p>ConvertToFormulas 方法是 ET 中的新方法，可用于将数据透视表转换为多维数据集公式。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CreateCubeFile">CreateCubeFile</a></td>
<td style="text-align: left;" data-valian="middle"><p>创建数据透视表的多维数据集文件，该数据透视表与 <span>联机分析处理 (OLAP) ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。）</span> 数据源相连接。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DiscardChanges">DiscardChanges</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>放弃基于 OLAP 数据源的数据透视表中经过编辑的单元格中的所有更改。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.GetData">GetData</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回数据透视表中数据字段的值。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.GetPivotData">GetPivotData</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Range</span> 对象，并返回有关数据透视表中数据项的信息。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ListFormulas">ListFormulas</a></td>
<td style="text-align: left;" data-valian="middle"><p>在分离工作表上创建数据透视表的计算项和计算字段的列表。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PivotCache">PivotCache</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>PivotCache</span> 对象，该对象表示指定数据透视表的缓存。只读。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PivotFields">PivotFields</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，该对象表示数据透视表中单个字段（ <span>PivotField</span> 对象）或可见及隐藏字段集合（ <span>PivotFields</span> 对象）。只读。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PivotSelect">PivotSelect</a></td>
<td style="text-align: left;" data-valian="middle"><p>选定数据透视表的一部分。</p></td>
</tr>
</tbody>
</table>

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CacheIndex">CacheIndex</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置数据透视表缓存的索引号。 Long 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CalculatedMembers">CalculatedMembers</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>CalculatedMembers</span> 集合，该集合表示 OLAP 数据透视表的所有计算成员和计算方法。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CalculatedMembersInFilter">CalculatedMembersInFilter</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置是否估算筛选器中的 OLAP 服务器计算成员。可读/写</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ChangeList">ChangeList</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 <span>PivotTableChangeList</span> 集合，该集合表示已对基于 OLAP 数据源的指定数据透视表所做更改的列表。只读</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ColumnFields">ColumnFields</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，该对象表示单个的数据透视表字段（ <span>PivotField</span> 对象）或当前显示为列字段的所有字段的集合（ <span>PivotFields</span> 对象）。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ColumnGrand">ColumnGrand</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果数据透视表显示列总计，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ColumnRange">ColumnRange</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Range</span> 对象，该对象表示数据透视表中包含列区域的区域。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CompactLayoutColumnHeader">CompactLayoutColumnHeader</a></td>
<td style="text-align: left;" data-valian="middle"><p>指定透视数据表在压缩行布局表单中时标题中显示的标题。可读写 String 。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CompactLayoutRowHeader">CompactLayoutRowHeader</a></td>
<td style="text-align: left;" data-valian="middle"><p>指定透视数据表在压缩行布局表单中时行标题中显示的标题。可读写 String 。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CompactRowIndent">CompactRowIndent</a></td>
<td style="text-align: left;" data-valian="middle"><p>当启用压缩行布局表单时，返回或设置透视项目的缩进增量。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.CubeFields">CubeFields</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 <span>CubeFields</span> 集合。每个 <span>CubeField</span> 对象都包含 <span>多维数据集 ?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。）</span> 字段元素的属性。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DataBodyRange">DataBodyRange</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Range</span> 对象，它代表数据透视表中数值区域。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DataFields">DataFields</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，该对象表示单个数据透视表字段（ <span>PivotField</span> 对象）或当前显示为数据字段的所有字段集合（ <span>PivotFields</span> 对象）。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DataLabelRange">DataLabelRange</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Range</span> 对象，该对象表示在数据透视表中包含数据字段的标签的区域。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DataPivotField">DataPivotField</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>PivotField</span> 对象，该对象表示数据透视表中所有数据字段。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DisplayContextTooltips">DisplayContextTooltips</a></td>
<td style="text-align: left;" data-valian="middle"><p>控制是否显示数据透视表单元格的工具提示。可读写 Boolean 。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DisplayEmptyColumn">DisplayEmptyColumn</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果对数值轴的 OLAP 提供程序的查询中包括非空 MDX 关键字，则返回 True 。在结果集中，OLAP 提供程序不返回空列。如果省略非空关键字，则返回 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DisplayEmptyRow">DisplayEmptyRow</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果对分类轴的 OLAP 提供程序的查询中包括非空 MDX 关键字，则返回 True 。在结果集中，OLAP 提供程序不返回空行。如果省略非空关键字，则返回 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">21</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DisplayErrorString">DisplayErrorString</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果数据透视表在有错误的单元格中显示用户自定义的错误字符串，则该值为 True 。默认值为 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">22</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DisplayFieldCaptions">DisplayFieldCaptions</a></td>
<td style="text-align: left;" data-valian="middle"><p>控制是否在网格中显示筛选器按钮以及行和列的透视字段标题。可读/写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">23</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DisplayImmediateItems">DisplayImmediateItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean ，用于指明当数据透视表的数据区域为空时，行和列区域中的项是否可见。如果该属性为 False ，则当数据透视表的数据区域为空时，将隐藏行和列区域中的项。默认值为 True 。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">24</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DisplayMemberPropertyTooltips">DisplayMemberPropertyTooltips</a></td>
<td style="text-align: left;" data-valian="middle"><p>控制是否在工具提示中显示成员属性。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">25</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.DisplayNullString">DisplayNullString</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果数据透视表在包含空值的单元格中显示用户自定义的字符串，则该值为 True 。默认值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">26</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.EnableDataValueEditing">EnableDataValueEditing</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则当用户覆盖数据透视表数据区域中的值时禁用警告。设置为 True 也可以使用户更改先前无法更改的数据值。默认值为 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">27</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.EnableDrilldown">EnableDrilldown</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果启用“显示明细数据”，则该值为 True 。默认值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">28</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.EnableFieldDialog">EnableFieldDialog</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当用户双击数据透视表字段时， “数据透视表字段” 对话框可用，则该属性值为 True 。默认值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">29</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.EnableFieldList">EnableFieldList</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 False ，则禁止显示数据透视表字段列表的功能。如果已经显示字段列表，则该列表将消失。默认值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">30</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.EnableWizard">EnableWizard</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 “数据透视表向导” 可用，则该属性值为 True 。默认值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">31</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.EnableWriteback">EnableWriteback</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置是否对指定的数据透视表启用回写到数据源。默认值为 False 。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">32</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ErrorString">ErrorString</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 值，它代表 <span>DisplayErrorString</span> 属性为 True 时如果单元格中有错误而显示的字符串。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">33</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.FieldListSortAscending">FieldListSortAscending</a></td>
<td style="text-align: left;" data-valian="middle"><p>控制数据透视表字段列表中字段的排序顺序。当此属性设置为 True 时，字段按升序顺序排序。当它设置为 False 时，字段按数据源顺序排序。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">34</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.GrandTotalName">GrandTotalName</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置显示在指定数据透视表的总计列或行标题中的文本串标志。默认值为字符串“Grand Total”。 String 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">35</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.HiddenFields">HiddenFields</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，该对象表示单个数据透视表字段（ <span>PivotField</span> 对象），或表示当前未显示为行、列、页或数据字段的所有字段的集合（ <span>PivotFields</span> 对象）。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">36</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.InGridDropZones">InGridDropZones</a></td>
<td style="text-align: left;" data-valian="middle"><p>此属性用于为 PivotTable 对象切换网格中的拖放区域。在一些情况下，它还会影响数据透视表的布局。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">37</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.InnerDetail">InnerDetail</a></td>
<td style="text-align: left;" data-valian="middle"><p>当最内部行或列字段的 ShowDetail 属性设为 True 时，返回或设置这些详细数据的字段名称。 String 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">38</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.LayoutRowDefault">LayoutRowDefault</a></td>
<td style="text-align: left;" data-valian="middle"><p>此属性指定初次将透视字段添加到数据透视表中时它们的布局设置。可读/写 xlLayoutRowType 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">39</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.Location">Location</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取或设置一个 String 类型的值，该值代表指定 <span>PivotTable</span> 主体中左上角的单元格。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">40</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.ManualUpdate">ManualUpdate</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果数据透视表仅在用户请求时重新计算，则该值为 True 。默认值为 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">41</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.MDX">MDX</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 String 类型的值，该值表示将发送给提供程序以填充当前数据透视表视图的多维表达式 (MDX)。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">42</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.MergeLabels">MergeLabels</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定的数据透视表的外部行项、列项、分类汇总和总计标志使用合并单元格，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">43</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.Name">Name</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 值，它代表对象的名称。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">44</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.NullString">NullString</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置当 <span>DisplayNullString</span> 为 True 时，在包含 null 值的单元格中显示的字符串。默认值为空字符串 ("")。 String 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">45</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PageFieldOrder">PageFieldOrder</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置页字段在数据透视表布局上的顺序。可为以下 <span>XlOrder</span> 常量之一： xlDownThenOver 或 xlOverThenDown 。默认常量是 xlDownThenOver 。 Long 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">46</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PageFields">PageFields</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，该对象表示单个数据透视表字段（ <span>PivotField</span> 对象），或者当前显示为页字段的所有字段的一个集合（ <span>PivotFields</span> 对象）。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">47</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PageFieldStyle">PageFieldStyle</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置用于边界页字段区域中的样式。默认值为 null 字符串（默认时无样式）。 String 类型，可读写。</p>
<p>语法</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">48</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PageFieldWrapCount">PageFieldWrapCount</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置数据透视表中每行或每列的页字段数目。 Long 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">49</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PageRange">PageRange</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Range</span> 对象，该对象表示包含数据透视表中页区域的区域。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">50</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PageRangeCells">PageRangeCells</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Range</span> 对象，该对象仅表示指定数据透视表中（包含页字段和项下拉列表）的单元格。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">51</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">52</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PivotColumnAxis">PivotColumnAxis</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 PivotAxis 对象，该对象表示整个列轴。只读 PivotAxis 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">53</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PivotFormulas">PivotFormulas</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>PivotFormulas</span> 对象，该对象表示指定数据透视表的公式集合。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">54</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PivotRowAxis">PivotRowAxis</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 PivotAxis 对象，该对象表示整个行轴。只读 PivotAxis 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">55</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PivotSelection">PivotSelection</a></td>
<td style="text-align: left;" data-valian="middle"><p>以标准数据透视表的选定区域格式返回或设置数据透视表的选定区域。 String 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">56</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PivotSelectionStandard">PivotSelectionStandard</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置用英语（美国）设置的标准 <span>数据透视表 ?（数据透视表：一种交互的、交叉制表的 ET 报表，用于对多种来源（包括 ET 的外部数据）的数据（如数据库记录）进行汇总和分析。）</span> 格式的数据透视表选定内容的 String 类型的数值。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">57</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PreserveFormatting">PreserveFormatting</a></td>
<td style="text-align: left;" data-valian="middle"><p>当使用数据透视、筛选或更改页字段项等操作刷新或重新计算报告时，如果格式处于被保护状态，则为 True 。对于查询表，如果数据前五行的任何常规格式设置将应用于查询表数据的新行，则此属性为 True 。对未使用的单元格不进行设置。如果将查询表应用的最近一次自动套用格式应用于新数据行，则此属性为 False 。默认值是 True 。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">58</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PrintDrillIndicators">PrintDrillIndicators</a></td>
<td style="text-align: left;" data-valian="middle"><p>指定是否随数据透视表打印深化指示符。可读写 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">59</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.PrintTitles">PrintTitles</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果基于数据透视表设置工作表的打印标题，则该属性值为 True 。如果使用工作表的打印标题，则该属性值为 False 。默认值为 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">60</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.RefreshDate">RefreshDate</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回数据透视表或缓存最近一次刷新的日期。 Date 型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">61</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.RefreshName">RefreshName</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回最近一次刷新数据透视表的人员的名字。 String 型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">62</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.RepeatItemsOnEachPrintedPage">RepeatItemsOnEachPrintedPage</a></td>
<td style="text-align: left;" data-valian="middle"><p>当打印指定的数据透视表时，如果每页第一行上都显示行、列和项标志，则该值为 True 。如果仅在第一页上打印这些标志，则该值为 False 。默认值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">63</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.RowFields">RowFields</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，该对象表示数据透视表中的单个字段（ <span>PivotField</span> 对象），或者当前未显示为行字段的所有字段的一个集合（ <span>PivotFields</span> 对象）。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">64</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotTable.RowGrand">RowGrand</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果数据透视表显示行的总数，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotTable.AddDataField

将数据字段添加到数据透视表中。返回一个 **PivotField** 对象，该对象表示新的数据字段。

**语法**

**express.AddDataField ( *Field* , *Caption* , *Function* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
|------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Field*    | 必选      | Object   | 服务器上的唯一字段。如果源数据是联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。），则唯一字段是多维数据集字段。如果源数据不是 OLAP（非 OLAP 源数据?（非 OLAP 源数据：非 OLAP 源数据是数据透视表或数据透视图使用的基本数据，该数据来自 OLAP 数据库之外的源。这些数据源包括关系数据库、 ET 工作表中的清单以及文本文件数据库。）），则唯一字段是数据透视表字段。 |
| *Caption*  | 可选      | Variant  | 数据透视表中使用的标签，用于识别该数据字段。                                                                                                                                                                                                                                                                                                                                                                                                                  |
| *Function* | 可选      | Variant  | 在已添加数据字段中执行的函数。                                                                                                                                                                                                                                                                                                                                                                                                                                |

**返回值**

PivotField

**示例**

本示例将标题为“Total Score”的数据字段添加到名为“PivotTable1”的数据透视表中。

| 注释                                                          |
|---------------------------------------------------------------|
| 本示例假定存在这样的表格，该表格中包含一个标题为“Score”的列。 |

``` JavaScript
function AddMoreFields() {
    let pTab = ActiveSheet.PivotTables("PivotTable1")
    pTab.AddDataField(ActiveSheet.PivotTables("PivotTable1").PivotFields("Score"), "Total Score")
}
```

### PivotTable.AddFields

向数据透视表或数据透视图中添加行字段、列字段和页字段。

**语法**

**express.AddFields ( *RowFields* , *ColumnFields* , *PageFields* , *AddToTable* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                                                                      |
|----------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------|
| *RowFields*    | 可选      | Variant  | 指定要作为行添加或要添加到类别坐标轴中的字段名（或字段名数组）。                                                                          |
| *ColumnFields* | 可选      | Variant  | 指定要作为列添加或要添加到序列坐标轴中的字段名（或字段名数组）。                                                                          |
| *PageFields*   | 可选      | Variant  | 指定要作为页添加或要添加到页区域中的字段名（或字段名数组）。                                                                              |
| *AddToTable*   | 可选      | Variant  | 仅适用于数据透视表。如果为 True，则将指定的字段添加到报表中（不替换现有字段）。如果为 False，则用新的字段替换现有的字段。默认值为 False。 |

**返回值**

Variant

**说明**

必须指定其中某个字段参数。

字段名指定由 **PivotField** 对象的 **SourceName** 属性返回的唯一名称。

这个方法对 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源不可用。

**示例**

本示例以“Status”字段和“Closed_By”字段替换 Sheet1 的第一个数据透视表的现有列字段。

``` JavaScript
Worksheets.Item("Sheet1").PivotTables(1).AddFields(null,["Status", "Closed_By"])
```

### PivotTable.AllocateChanges

table { word-break:break-all; }

对基于 OLAP 数据源的数据透视表中所有编辑过的单元格执行回写操作。

**语法**

**express.AllocateChanges ()**

\*express是一个代表 PivotTable 对象的变量。

**说明**

table { word-break:break-all; }

**AllocateChanges** 方法将执行 **UPDATE CUBE** 语句，以应用自上次提交应用更改操作以来，或创建数据透视表以来（如果从未执行过提交应用更改），数据透视表的值区域中发生的所有更改。如果对基于非 OLAP 数据源的数据透视表执行此方法，那么将生成运行时错误。

### PivotTable.CalculatedFields

返回一个 **CalculatedFields** 集合，该集合表示指定数据透视表中的所有计算字段。只读。

**语法**

**express.CalculatedFields ()**

\*express是一个代表 PivotTable 对象的变量。

**返回值**

CalculatedFields

**示例**

本示例使计算结果字段不能被拖至行。

``` JavaScript
for(let i=1;i <= Worksheets.Item(1).PivotTables("Pivot1").CalculatedFields().Count;i++) {
    Worksheets.Item(1).PivotTables("Pivot1").CalculatedFields().Item(i).DragToRow = false
}
```

### PivotTable.ChangeConnection

table { word-break:break-all; }

更改指定 **PivotTable** 的连接。

**语法**

**express.ChangeConnection ( *conn* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型           | 说明                                                   |
|--------|-----------|--------------------|--------------------------------------------------------|
| *conn* | 必选      | WorkbookConnection | 一个代表数据透视表的新连接的 WorkbookConnection 对象。 |

**说明**

table { word-break:break-all; }

**ChangeConnection** 方法只能用于连接到外部数据源的 **PivotTable** 。如果将 **ChangeConnection** 方法与使用工作表中存储的数据作为数据源的 **PivotTable** 一起使用，将发生运行时错误。

### PivotTable.ChangePivotCache

table { word-break:break-all; }

更改指定 **PivotTable** 的 **PivotCache** 。

**语法**

**express.ChangePivotCache ( *bstr* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                            |
|--------|-----------|----------|---------------------------------------------------------------------------------|
| *bstr* | 必选      | String   | 一个 PivotTable 或 PivotCache 对象，该对象代表指定 PivotTable 的新 PivotCache。 |

**说明**

table { word-break:break-all; }

**ChangePivotCache** 方法只能用于使用工作表中存储的数据作为数据源的 **PivotTable** 。如果将 **ChangePivotCache** 方法与连接到外部数据源的 **PivotTable** 一起使用，将发生运行时错误。

### PivotTable.ClearAllFilters

table { word-break:break-all; }

**ClearAllFilters** 方法将删除当前应用于数据透视表的所有筛选。包括删除 **PivotTable** 对象的 **PivotFilters** 集合中的所有筛选，删除应用的所有手动筛选以及将报表筛选区域中的所有筛选字段设置为默认项。

**语法**

**express.ClearAllFilters ()**

\*express是一个代表 PivotTable 对象的变量。

**说明**

table { word-break:break-all; }

调用 **ClearAllFilters** 方法后， **PivotFilters** 集合将为空。

### PivotTable.ClearTable

**ClearTable** 方法用于清除数据透视表。清除数据透视表包括删除所有字段以及删除应用于数据透视表的所有筛选和排序。此方法将数据透视表重置为该表刚创建而尚未添加任何字段时的状态。

**语法**

**express.ClearTable ()**

\*express是一个代表 PivotTable 对象的变量。

**说明**

**ClearTable** 函数不使用任何参数，可用于关系数据透视表和 OLAP 数据透视表。

**示例**

以下示例清除活动工作表中的数据透视表。

``` JavaScript
ActiveSheet.PivotTables(1).ClearTable()
```

### PivotTable.CommitChanges

table { word-break:break-all; }

对基于 OLAP 数据源的数据透视表的数据源执行提交操作。

**语法**

**express.CommitChanges ()**

\*express是一个代表 PivotTable 对象的变量。

**说明**

table { word-break:break-all; }

**CommitChanges** 方法将 **COMMIT TRANSACTION** 语句发送给 OLAP 服务器，并且清除通过输入值而编辑的所有单元格，但是不清除值单元格中的公式。如果对基于非 OLAP 数据源的数据透视表执行此方法，那么将生成运行时错误。

### PivotTable.ConvertToFormulas

**ConvertToFormulas** 方法是 ET 中的新方法，可用于将数据透视表转换为多维数据集公式。可读/写 **Boolean** 类型。

**语法**

**express.ConvertToFormulas ( *ConvertFilters* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                 |
|------------------|-----------|----------|------------------------------------------------------|
| *ConvertFilters* | 必选      | Boolean  | 包含 True 或 False，以指明 ReportFilter 区域的状态。 |

**说明**

此参数控制是否转换数据透视表的 **ReportFilter** 区域。

**示例**

在以下示例中，不转换 **ReportFilter** 区域。

``` JavaScript
function ConvertToCubeFormulas() {
    ActiveSheet.PivotTables("PivotTable1").ConvertToFormulas(false)
}
```

### PivotTable.CreateCubeFile

创建数据透视表的多维数据集文件，该数据透视表与 联机分析处理 (OLAP) ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源相连接。

**语法**

**express.CreateCubeFile ( *File* , *Measures* , *Levels* , *Members* , *Properties* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                                                   |
|--------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------|
| *File*       | 必选      | String   | 要创建的多维数据集文件的名称。如果该文件已存在，则将被覆盖。                                                                                           |
| *Measures*   | 可选      | Variant  | 度量的唯一名称的数组，该度量是扇面的一部分。                                                                                                           |
| *Levels*     | 可选      | Variant  | 字符串数组。每个数组项都是一个唯一的级别名称。该参数表示扇面中的层次结构的最低级别。                                                                   |
| *Members*    | 可选      | Variant  | 字符串数组的数组。元素按顺序对应于由 Levels 数组表示的分层结构。每个元素都是字符串数组的一个数组，由维（包括在扇面中）中的最高级别成员的唯一名称组成。 |
| *Properties* | 可选      | Variant  | 如果该值为 False，则扇区中不包含成员属性。默认值为 True                                                                                                |

**返回值**

String

**示例**

本示例在驱动器 C:\\ 下创建一个名为“CustomCubeFile”的多维数据集文件，扇面中不包括任何成员属性。由于本示例省略了 *Measures* 、 *Levels* 和 *Members* 参数，因此该多维数据集文件将停止匹配数据透视表的视图。本示例假定与 OLAP 数据源相连接的数据透视表位于活动工作表上。

``` JavaScript
function UseCreateCubeFile() {
    ActiveSheet.PivotTables(1).CreateCubeFile("C:\\CustomCubeFile",null,null,null,false)
}
```

### PivotTable.DiscardChanges

table { word-break:break-all; }

放弃基于 OLAP 数据源的数据透视表中经过编辑的单元格中的所有更改。

**语法**

**express.DiscardChanges ()**

\*express是一个代表 PivotTable 对象的变量。

**说明**

table { word-break:break-all; }

对于基于 OLAP 数据源的数据透视表，此方法删除在值单元格中输入的所有值和公式，然后运行数据透视表更新操作从数据源中检索最新值。它将所有编辑过的值单元格的对应数据源值设置为 **NULL** ，并且还对 OLAP 服务器执行 **ROLLBACK TRANSACTION** 语句。

如果尝试对基于非 OLAP 数据源的数据透视表执行此方法，则会生成运行时错误。

### PivotTable.GetData

返回数据透视表中数据字段的值。

**语法**

**express.GetData ( *Name* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                            |
|--------|-----------|----------|-------------------------------------------------------------------------------------------------|
| *Name* | 必选      | String   | 描述数据透视表中的单个单元格，使用的语法类似于 PivotSelect 方法或者计算项公式中的数据透表引用。 |

**返回值**

双精度

**示例**

本示例显示一月的苹果收入之和 (Data field = Revenue、Product = Apples、Month = January)。

``` JavaScript
MsgBox(ActiveSheet.PivotTables(1).GetData("'Sum of Revenue' Apples January"))
```

### PivotTable.GetPivotData

返回一个 **Range** 对象，并返回有关数据透视表中数据项的信息。

**语法**

**express.GetPivotData ( *DataField* , *Field1~14* , *Item1~14* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                             |
|-------------|-----------|----------|----------------------------------|
| *DataField* | 可选      | Variant  | 包含数据透视表数据的字段的名称。 |
| *Field1~14* | 可选      | Variant  | 数据透视表中的列或行字段的名称。 |
| *Item1~14*  | 可选      | Variant  | Field中的项目的名称。            |

**返回值**

Range

**示例**

在本例中，ET 向用户返回仓库中椅子的数量。本示例假定数据透视表位于活动工作表上。此外，本示例还假设：在数据透视表中，数据字段的标题为“Quantity”，同时还存在标题为“Warehouse”的字段，并且在“Warehouse”字段中有标题为“Chairs”的数据项。

``` JavaScript
function UseGetPivotData() {
    // Get PivotData for the quantity of chairs in the warehouse.
    let rngTableItem = ActiveCell.PivotTable.GetPivotData("Quantity", "Warehouse", "Chairs")
    MsgBox("The quantity of chairs in the warehouse is: " + rngTableItem.Value())
}
```

### PivotTable.ListFormulas

在分离工作表上创建数据透视表的计算项和计算字段的列表。

**语法**

**express.ListFormulas ()**

\*express是一个代表 PivotTable 对象的变量。

**说明**

本方法对 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

本示例为第一个工作表上的第一个数据透视表创建一个计算项和计算字段的列表。

``` JavaScript
Worksheets.Item(1).PivotTables(1).ListFormulas()
```

### PivotTable.PivotCache

返回一个 **PivotCache** 对象，该对象表示指定数据透视表的缓存。只读。

**语法**

**express.PivotCache ()**

\*express是一个代表 PivotTable 对象的变量。

**返回值**

PivotCache

**示例**

本示例使第一张工作表上的第一张数据透视表的缓存在构造时进行优化。

``` JavaScript
Worksheets.Item(1).PivotTables("Pivot1").PivotCache.OptimizeCache = true
```

### PivotTable.PivotFields

返回一个对象，该对象表示数据透视表中单个字段（ **PivotField** 对象）或可见及隐藏字段集合（ **PivotFields** 对象）。只读。

**语法**

**express.PivotFields ( *Index* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                       |
|---------|-----------|----------|----------------------------|
| *Index* | 可选      | Variant  | 要返回的字段的名称或编号。 |

**返回值**

Object

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，不存在隐藏字段，返回的对象或集合反映了当前可见的内容。

**示例**

本示例向新工作表列表添加数据透视表的字段名称。

``` JavaScript
let nwSheet = Worksheets.Add()
nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i=1;i <= pvtTable.PivotFields().Count;i++) {
    rw++
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields(i).Name
}
```

### PivotTable.PivotSelect

选定数据透视表的一部分。

**语法**

**express.PivotSelect ( *Name* , *Mode* , *UseStandardName* )**

\*express是一个代表 PivotTable 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型          | 说明                                          |
|-------------------|-----------|-------------------|-----------------------------------------------|
| *Name*            | 必选      | String            | 要选择的数据透视表的一部分。                  |
| *Mode*            | 可选      | XlPTSelectionMode | 指定结构化选择模式。                          |
| *UseStandardName* | 可选      | Variant           | 如果为 True，则表示将在其他位置运行录制的宏。 |

**说明**

只能使用指定模式来选择数据透视表中的相应项。例如，无法通过 **xlButton** 模式来选择数据和标志；同样，也无法通过 **xlDataOnly** 模式来选择按钮。

**示例**

本示例选定第一张工作表上的第一张数据透视表中的所有日期标志。

``` JavaScript
Worksheets.Item(1).PivotTables(1).PivotSelect("date[All]", xlLabelOnly)
```

## 成员属性

### PivotTable.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
}
```

### PivotTable.CacheIndex

返回或设置数据透视表缓存的索引号。 **Long** 类型，可读写。

**语法**

**express.CacheIndex**

\*express是一个代表 PivotTable 对象的变量。

**说明**

如果设置 **CacheIndex** 属性，以便于数据透视表使用第二张数据透视表的高速缓存，则第一个报表的字段必须是第二个报表中字段的有效子集。

**示例**

``` JavaScript
/*本示例将数据透视表“Pivot1”的缓存设置成数据透视表“Pivot2”的缓存。*/
Application.Worksheets.Item(1).PivotTables("Pivot1").CacheIndex = Worksheets.Item(1).PivotTables("Pivot2").CacheIndex
 
```

### PivotTable.CalculatedMembers

返回一个 **CalculatedMembers** 集合，该集合表示 OLAP 数据透视表的所有计算成员和计算方法。

**语法**

**express.CalculatedMembers**

\*express是一个代表 PivotTable 对象的变量。

**说明**

本属性用于 联机分析处理 (OLAP) ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 源，非 OLAP 源将返回一个运行时错误。

**示例**

本示例向数据透视表添加一个设置。假定数据透视表位于一个与 OLAP 数据源相连的活动工作表上，此 OLAP 数据源中包含一个标题为“\[Product\].\[All Products\]”的字段。

``` JavaScript
function test(){
  let pvtTable = Application.ActiveSheet.PivotTables(1)
    // Add the calculated member.
    pvtTable.CalculatedMembers.Add("[Beef]", "'{[Product].[All Products].Children}'", null, xlCalculatedSet)
}
```

### PivotTable.CalculatedMembersInFilter

返回或设置是否估算筛选器中的 OLAP 服务器计算成员。可读/写

**语法**

**express.CalculatedMembersInFilter**

\*express是一个代表 PivotTable 对象的变量。

**说明**

如果在筛选器中估算计算成员，则为 **True** ；否则为 **False** 。

此属性的值对应于基于 OLAP 数据源的数据透视表的 **“数据透视表选项”** 对话框 **“汇总和筛选”** 选项卡上的 **“估算筛选器中的 OLAP 服务器计算成员”** 复选框的设置。

### PivotTable.ChangeList

返回 **PivotTableChangeList** 集合，该集合表示已对基于 OLAP 数据源的指定数据透视表所做更改的列表。只读

**语法**

**express.ChangeList**

\*express是一个代表 PivotTable 对象的变量。

### PivotTable.ColumnFields

返回一个对象，该对象表示单个的数据透视表字段（ **PivotField** 对象）或当前显示为列字段的所有字段的集合（ **PivotFields** 对象）。只读。

**语法**

**express.ColumnFields**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例将数据透视表的列字段名称添加到一张新工作表上的数据清单中。*/
function test(){
let nwSheet = Worksheets.Add()
nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i=1;i <= pvtTable.ColumnFields().Count;i++) {
    rw++
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.ColumnFields(i).Name
}
}
```

### PivotTable.ColumnGrand

如果数据透视表显示列总计，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ColumnGrand**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例设置数据透视表显示列总计。*/
function test(){
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
pvtTable.ColumnGrand = true
}
```

### PivotTable.ColumnRange

返回一个 **Range** 对象，该对象表示数据透视表中包含列区域的区域。只读。

**语法**

**express.ColumnRange**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例为数据透视表选择列标题。*/
function test(){
Worksheets.Item("Sheet1").Activate()
Range("A3").Select()
ActiveCell.PivotTable.ColumnRange.Select()
}
```

### PivotTable.CompactLayoutColumnHeader

指定透视数据表在压缩行布局表单中时标题中显示的标题。可读写 **String** 。

**语法**

**express.CompactLayoutColumnHeader**

\*express是一个代表 PivotTable 对象的变量。

**说明**

用户在网格中自定义这些标题时，这些标题也会反映到这些属性中。

### PivotTable.CompactLayoutRowHeader

指定透视数据表在压缩行布局表单中时行标题中显示的标题。可读写 **String** 。

**语法**

**express.CompactLayoutRowHeader**

\*express是一个代表 PivotTable 对象的变量。

**说明**

用户在网格中自定义这些标题时，这些标题也会反映到这些属性中。

### PivotTable.CompactRowIndent

当启用压缩行布局表单时，返回或设置透视项目的缩进增量。可读写。

**语法**

**express.CompactRowIndent**

\*express是一个代表 PivotTable 对象的变量。

**说明**

默认值为 1。此设置的有效值为 0 到 ET 中指定的最大缩进。

### PivotTable.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotTable 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotTable.CubeFields

返回 **CubeFields** 集合。每个 **CubeField** 对象都包含 多维数据集 ?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。） 字段元素的属性。只读。

**语法**

**express.CubeFields**

\*express是一个代表 PivotTable 对象的变量。

**示例**

本示例为工作表 Sheet1 上第一个基于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 的数据透视表中的数据字段，创建一个 多维数据集 ?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。） 字段名称列表。

``` JavaScript
function test(){
let objNewSheet = Worksheets.Add()
objNewSheet.Activate()
let intRow = 1
for(let i=1;i <= Worksheets.Item("Sheet1").PivotTables(1).CubeFields.Count;i++) {
    if(Worksheets.Item("Sheet1").PivotTables(1).CubeFields.Item(i).Orientation == xlDataField) {
        objNewSheet.Cells.Item(intRow, 1).Value2 = Worksheets.Item("Sheet1").PivotTables(1).CubeFields.Item(i).Name
        intRow++
    }
}
}
```

### PivotTable.DataBodyRange

返回一个 **Range** 对象，它代表数据透视表中数值区域。只读。

**语法**

**express.DataBodyRange**

\*express是一个代表 PivotTable 对象的变量。

### PivotTable.DataFields

返回一个对象，该对象表示单个数据透视表字段（ **PivotField** 对象）或当前显示为数据字段的所有字段集合（ **PivotFields** 对象）。只读。

**语法**

**express.DataFields**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例向一张新工作表中添加数据透视表数据字段的名称列表。*/
function test(){
let nwSheet = Worksheets.Add()
nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i=1;i <= pvtTable.DataFields().Count;i++) {
    rw++
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.DataFields(i).Name
}
}
```

### PivotTable.DataLabelRange

返回一个 **Range** 对象，该对象表示在数据透视表中包含数据字段的标签的区域。只读。

**语法**

**express.DataLabelRange**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例选定数据透视表中的数据字段标签。*/
function test(){
Worksheets.Item("Sheet1").Activate()
Range("A3").Select()
ActiveCell.PivotTable.DataLabelRange.Select()
}
```

### PivotTable.DataPivotField

返回一个 **PivotField** 对象，该对象表示数据透视表中所有数据字段。只读。

**语法**

**express.DataPivotField**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例将第二个 PivotItem 对象移到第一个位置。假定数据透视表位于活动的工作表上，并且数据透视表包含数据字段。*/
function test(){
 let pvtTable = ActiveSheet.PivotTables(1)
    // Move second PivotItem to the first position in PivotTable.
    pvtTable.DataPivotField.PivotItems(2).Position = 1
}
```

### PivotTable.DisplayContextTooltips

控制是否显示数据透视表单元格的工具提示。可读写 **Boolean** 。

**语法**

**express.DisplayContextTooltips**

\*express是一个代表 PivotTable 对象的变量。

**说明**

如果 **DisplayContextTooltips** 设置为 **False** ，将不显示数据透视表单元格的上下文工具提示。如果此属性设置为 **True** ，将显示行、列和数据区中数据透视表单元格的工具提示。

### PivotTable.DisplayEmptyColumn

如果对数值轴的 OLAP 提供程序的查询中包括非空 MDX 关键字，则返回 **True** 。在结果集中，OLAP 提供程序不返回空列。如果省略非空关键字，则返回 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayEmptyColumn**

\*express是一个代表 PivotTable 对象的变量。

**说明**

数据透视表必须与一个 联机分析处理 (OLAP) ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源相连，以避免运行时错误。

**示例**

``` JavaScript
/*本示例确定数据透视表中空列的显示设置。本示例假定与 OLAP 数据源相连的数据透视表位于活动工作表上。*/
function test(){
 let pvtTable = ActiveSheet.PivotTables(1)
    // Determine display setting for empty columns.
    if(pvtTable.DisplayEmptyColumn == false) {
        alert("Empty columns will not be displayed.")
    }
    else {
        alert("Empty columns will be displayed.")
    }
}
```

### PivotTable.DisplayEmptyRow

如果对分类轴的 OLAP 提供程序的查询中包括非空 MDX 关键字，则返回 **True** 。在结果集中，OLAP 提供程序不返回空行。如果省略非空关键字，则返回 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayEmptyRow**

\*express是一个代表 PivotTable 对象的变量。

**说明**

数据透视表必须与一个 联机分析处理 (OLAP) ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源相连，以避免运行时错误。

**示例**

``` JavaScript
/*本示例确定用于数据透视表中的空行的显示设置。本示例假定与 OLAP 数据源相连的数据透视表位于活动工作表上。*/
function test(){
  let pvtTable = ActiveSheet.PivotTables(1)
    // Determine display setting for empty rows.
    if(pvtTable.DisplayEmptyRow == false) {
        alert("Empty rows will not be displayed.")
    }
    else {
        alert("Empty rows will be displayed.")
    }
}
```

### PivotTable.DisplayErrorString

如果数据透视表在有错误的单元格中显示用户自定义的错误字符串，则该值为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayErrorString**

\*express是一个代表 PivotTable 对象的变量。

**说明**

使用 **ErrorString** 属性可设置用户自定义的错误字符串。

当透视处理计算字段时，该属性对于消除被零除错误尤其有用。

**示例**

``` JavaScript
/*本示例使数据透视表在有错误的单元格中显示连字符。*/
function test(){
Worksheets.Item(1).PivotTables("Pivot1").ErrorString = "-"
Worksheets.Item(1).PivotTables("Pivot1").DisplayErrorString = true
}
```

### PivotTable.DisplayFieldCaptions

控制是否在网格中显示筛选器按钮以及行和列的透视字段标题。可读/写。

**语法**

**express.DisplayFieldCaptions**

\*express是一个代表 PivotTable 对象的变量。

**说明**

默认值为 **True** 。

### PivotTable.DisplayImmediateItems

返回或设置 **Boolean** ，用于指明当数据透视表的数据区域为空时，行和列区域中的项是否可见。如果该属性为 **False** ，则当数据透视表的数据区域为空时，将隐藏行和列区域中的项。默认值为 **True** 。

**语法**

**express.DisplayImmediateItems**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例确定如何创建数据透视表并通知用户。本示例假定数据透视表位于活动工作表上。*/
function test(){
let pvtTable = ActiveSheet.PivotTables(1)
// Determine how the PivotTable was created.
if(pvtTable.DisplayImmediateItems == true) {
alert("Fields have been added to the row or column areas for the PivotTable report.")
}
else {
alert("The PivotTable was created by using object-model calls.")
}
}
```

### PivotTable.DisplayMemberPropertyTooltips

控制是否在工具提示中显示成员属性。可读/写 **Boolean** 类型。

**语法**

**express.DisplayMemberPropertyTooltips**

\*express是一个代表 PivotTable 对象的变量。

**说明**

如果 **DisplayMemberPropertyTooltips** 设置为 **False** ，则将不会在工具提示中显示任何成员属性。如果它设置为 **True** ，则将在工具提示中显示成员属性。

### PivotTable.DisplayNullString

如果数据透视表在包含空值的单元格中显示用户自定义的字符串，则该值为 **True** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayNullString**

\*express是一个代表 PivotTable 对象的变量。

**说明**

使用 **NullString** 属性可设置自定义空字符串。

**示例**

``` JavaScript
/*本示例使数据透视表在包含空值的单元格中显示“NA”。*/
function test(){
Application.Worksheets.Item(1).PivotTables("Pivot1").NullString = "NA"
Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayNullString = true
}
```

``` JavaScript
/*本示例使数据透视表在包含空值的单元格中显示零 (0)。*/
Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayNullString = false
```

### PivotTable.EnableDataValueEditing

如果为 **True** ，则当用户覆盖数据透视表数据区域中的值时禁用警告。设置为 **True** 也可以使用户更改先前无法更改的数据值。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableDataValueEditing**

\*express是一个代表 PivotTable 对象的变量。

**说明**

在数据值上进行的任何编辑，在刷新时都将丢失。

**示例**

``` JavaScript
/*本示例确定对覆盖数据区域中的值时的警告设置，并通知用户。本示例假定数据透视表位于活动工作表上。*/
function test(){
 let pvtTable = ActiveSheet.PivotTables(1)
    // Determine alert setting.
    if(pvtTable.EnableDataValueEditing == false) {
        alert("Alert is enabled.")
    }
    else {
        alert("Alert is disabled.")
    }
}
```

### PivotTable.EnableDrilldown

如果启用“显示明细数据”，则该值为 **True** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableDrilldown**

\*express是一个代表 PivotTable 对象的变量。

**说明**

为数据透视表设置该属性将对该表中的所有字段有效。

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性值总是 **True** 。

**示例**

``` JavaScript
/*本示例对第一张工作表上第一个数据透视表中的所有字段禁用明细。*/
Application.Worksheets.Item(1).PivotTables("Pivot1").EnableDrilldown = false
```

### PivotTable.EnableFieldDialog

如果当用户双击数据透视表字段时， **“数据透视表字段”** 对话框可用，则该属性值为 **True** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableFieldDialog**

\*express是一个代表 PivotTable 对象的变量。

**说明**

为数据透视表设置该属性将对该表中的所有字段有效。

**示例**

``` JavaScript
/*本示例对“Year”字段禁用“数据透视表字段”对话框。*/
Application.Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").EnableFieldDialog = false
 
```

### PivotTable.EnableFieldList

如果为 **False** ，则禁止显示数据透视表字段列表的功能。如果已经显示字段列表，则该列表将消失。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableFieldList**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例确定是否能显示数据透视表的字段列表并通知用户。本示例假定数据透视表位于活动工作表上。*/
function test(){
 let pvtTable = ActiveSheet.PivotTables(1)
    // Determine if field list can be displayed.
    if(pvtTable.EnableFieldList == true) {
        alert("The field list for the PivotTable can be displayed.")
    }
    else {
        alert("The field list for the PivotTable cannot be displayed.")
    }

}
```

### PivotTable.EnableWizard

如果 **“数据透视表向导”** 可用，则该属性值为 **True** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableWizard**

\*express是一个代表 PivotTable 对象的变量。

**说明**

在设置该属性时，字段井并不会显示在工作表上。

**示例**

``` JavaScript
/*本示例对第一张工作表上第一个数据透视表禁用“数据透视表向导”.*/
Application.Worksheets.Item(1).PivotTables("Pivot1").EnableWizard = false
```

### PivotTable.EnableWriteback

返回或设置是否对指定的数据透视表启用回写到数据源。默认值为 **False** 。可读/写。

**语法**

**express.EnableWriteback**

\*express是一个代表 PivotTable 对象的变量。

**说明**

对于 OLAP 数据源，如果将 **EnableWriteback** 属性设置为 **True** ，那么将启用回写，并禁用用户在覆盖数据透视表数据区中的值时出现的警告。对于非 OLAP 数据源，如果将 **EnableWriteback** 属性设置为 **True** ，则在代码中启用回写，还将允许用户更改以前无法更改的数据值。

不能将 **PivotTable** 对象的 **EnableWriteback** 和 **EnableDataValueEditing** 属性同时设置为 **True** 。

如果将 **EnableDataValueEditing** 属性设置为 **True** ，然后再将 **EnableWriteback** 属性设置为 **True** ， **EnableDataValueEditing** 属性就会自动设置为 **False** ，数据透视表将被刷新并且对数据值进行的所有编辑操作都将丢失。

如果将 **EnableWriteback** 属性设置为 **True** ，然后再将 **EnableDataValueEditing** 属性设置为 **True** ， **EnableWriteback** 属性就会自动设置为 **False** ，不刷新数据透视表，但会还原数据源值。

对于非 OLAP 数据源，设置此属性将会生成运行时错误。

### PivotTable.ErrorString

返回或设置一个 **String** 值，它代表 **DisplayErrorString** 属性为 **True** 时如果单元格中有错误而显示的字符串。

**语法**

**express.ErrorString**

\*express是一个代表 PivotTable 对象的变量。

**说明**

该属性的默认值是个空字符串 ("")。

**示例**

``` JavaScript
/*对于指定数据透视表中有错误的单元格，本示例将在单元格中显示连字符。*/
function test(){
Application.Worksheets.Item(1).PivotTables("Pivot1").ErrorString = "-"
Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayErrorString = true
}
```

### PivotTable.FieldListSortAscending

控制数据透视表字段列表中字段的排序顺序。当此属性设置为 **True** 时，字段按升序顺序排序。当它设置为 **False** 时，字段按数据源顺序排序。可读/写。

**语法**

**express.FieldListSortAscending**

\*express是一个代表 PivotTable 对象的变量。

**说明**

不适用于与 OLAP 数据源相连的数据透视表。

### PivotTable.GrandTotalName

返回或设置显示在指定数据透视表的总计列或行标题中的文本串标志。默认值为字符串“Grand Total”。 **String** 类型，可读写。

**语法**

**express.GrandTotalName**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例将活动工作表的第二个数据透视表中总计标题的标志设置为“Regional Total”。*/
Application.ActiveSheet.PivotTables("PivotTable2").GrandTotalName = "Regional Total"
 
```

### PivotTable.HiddenFields

返回一个对象，该对象表示单个数据透视表字段（ **PivotField** 对象），或表示当前未显示为行、列、页或数据字段的所有字段的集合（ **PivotFields** 对象）。只读。

**语法**

**express.HiddenFields**

\*express是一个代表 PivotTable 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性总返回一个空集合。

**示例**

``` JavaScript
/*本示例将指定隐藏字段的名称添加进一张新工作表中的列表中。*/
function test(){
let nwSheet = Worksheets.Add()
nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i=1;i <= pvtTable.HiddenFields.Count;i++) {
    rw++
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.HiddenFields.Item(i).Name
}
}
```

### PivotTable.InGridDropZones

此属性用于为 **PivotTable** 对象切换网格中的拖放区域。在一些情况下，它还会影响数据透视表的布局。可读/写 **Boolean** 类型。

**语法**

**express.InGridDropZones**

\*express是一个代表 PivotTable 对象的变量。

**说明**

当 **InGridDropZones** 属性设置为 **True** 时，存在网格中的拖放区域。当此属性设置为 **False** 时，不存在网格中的拖放区域。数据透视表的布局也会随此属性一起改变。

### PivotTable.InnerDetail

当最内部行或列字段的 **ShowDetail** 属性设为 **True** 时，返回或设置这些详细数据的字段名称。 **String** 类型，可读写。

**语法**

**express.InnerDetail**

\*express是一个代表 PivotTable 对象的变量。

**说明**

该属性不可用于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

**示例**

``` JavaScript
/*本示例当最内部行或列字段的 ShowDetail 属性设为 True 时，显示这些详细数据的字段名称。*/
function test(){
let pvtTable = Application.Worksheets.Item("Sheet1").Range("A3").PivotTable
alert(pvtTable.InnerDetail)
}
```

### PivotTable.LayoutRowDefault

此属性指定初次将透视字段添加到数据透视表中时它们的布局设置。可读/写 **xlLayoutRowType** 类型。

**语法**

**express.LayoutRowDefault**

\*express是一个代表 PivotTable 对象的变量。

**说明**

**xlLayoutRowType** 可以是 **xlCompactRow** 、 **xlTabularRow** 或 **xlOutlineRow** 。

### PivotTable.Location

获取或设置一个 **String** 类型的值，该值代表指定 **PivotTable** 主体中左上角的单元格。可读/写。

**语法**

**express.Location**

\*express是一个代表 PivotTable 对象的变量。

### PivotTable.ManualUpdate

如果数据透视表仅在用户请求时重新计算，则该值为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.ManualUpdate**

\*express是一个代表 PivotTable 对象的变量。

**说明**

程序终止之后，以及在 Microsoft Visual Basic 编辑器的立即窗口执行了该语句之后，本属性的值立即设为 **False** 。

**示例**

``` JavaScript
/*本示例使数据透视表仅在用户请求时进行重新计算。*/
function test(){
Application.Worksheets.Item(1).PivotTables("Pivot1").ManualUpdate = true
}
```

### PivotTable.MDX

返回一个 **String** 类型的值，该值表示将发送给提供程序以填充当前数据透视表视图的多维表达式 (MDX)。只读。

**语法**

**express.MDX**

\*express是一个代表 PivotTable 对象的变量。

**说明**

查询非联机分析处理 (OLAP) 数据透视表的此值，或者当不存在数据透视表视图（无数据项）时，将返回一个运行错误。

**示例**

``` JavaScript
/*本示例返回数据透视表的 MDX 字符串。本示例假定数据透视表位于活动工作表上。*/
function test(){
 let pvtTable = ActiveSheet.PivotTables(1)
    alert("The MDX string for the PivotTable is: " + pvtTable.MDX)
}
```

### PivotTable.MergeLabels

如果指定的数据透视表的外部行项、列项、分类汇总和总计标志使用合并单元格，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MergeLabels**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例使第一张工作表上的第一个数据透视表使用合并单元格外部行项、列项、分类汇总和总计的标志。*/
function test(){
Application.Worksheets.Item(1).PivotTables(1).MergeLabels = true
}
```

### PivotTable.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 PivotTable 对象的变量。

### PivotTable.NullString

返回或设置当 **DisplayNullString** 为 **True** 时，在包含 null 值的单元格中显示的字符串。默认值为空字符串 ("")。 **String** 类型，可读写。

**语法**

**express.NullString**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例使数据透视表在包含空值的单元格中显示“NA”。*/
function test(){
Application.Worksheets.Item(1).PivotTables("Pivot1").NullString = "NA"
Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayNullString = true
}
```

### PivotTable.PageFieldOrder

返回或设置页字段在数据透视表布局上的顺序。可为以下 **XlOrder** 常量之一： **xlDownThenOver** 或 **xlOverThenDown** 。默认常量是 **xlDownThenOver** 。 **Long** 类型，可读写。

**语法**

**express.PageFieldOrder**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例使数据透视表先在一行中显示三个页字段，然后再显示下一个新行。*/
function test(){
Application.Worksheets.Item(1).PivotTables("Pivot1").PageFieldOrder = xlOverThenDown
Application.Worksheets.Item(1).PivotTables("Pivot1").PageFieldWrapCount = 3
}
```

### PivotTable.PageFields

返回一个对象，该对象表示单个数据透视表字段（ **PivotField** 对象），或者当前显示为页字段的所有字段的一个集合（ **PivotFields** 对象）。只读。

**语法**

**express.PageFields**

\*express是一个代表 PivotTable 对象的变量。

**说明**

一个层次只能包含一个页字段。

对基于数据透视表高速缓存的数据透视表而言，返回的数据透视表字段的集合反应了当前高速缓存中的内容。

**示例**

``` JavaScript
/*本示例向一张新工作表上的列表中添加页字段名称。*/
function test(){
let nwSheet = Application.Worksheets.Add()
nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i=1;i <= pvtTable.PageFields().Count;i++) {
    rw++
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PageFields(i).Name
}
}
```

### PivotTable.PageFieldStyle

返回或设置用于边界页字段区域中的样式。默认值为 null 字符串（默认时无样式）。 **String** 类型，可读写。

**语法**

**语法**

**express.PageFieldStyle**

\*express是一个代表 PivotTable 对象的变量。

**说明**

此样式用作背景区域的默认样式，并且在用户指定格式前起作用。当对某字段进行数据透视时，从页字段区域移到其他位置的单元格保留此样式。

**示例**

``` JavaScript
/*本示例将第一张工作表上的第一个数据透视表中的页字段区域设置为 PurpleAndGold 样式。*/
Application.Worksheets.Item(1).PivotTables("Pivot1").PageFieldStyle = "PurpleAndGold"
```

### PivotTable.PageFieldWrapCount

返回或设置数据透视表中每行或每列的页字段数目。 **Long** 类型，可读写。

**语法**

**express.PageFieldWrapCount**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例使数据透视表先在一行中显示三个页字段，然后再显示下一个新行。*/
function test(){
let pTab = Application.Worksheets.Item(1).PivotTables("Pivot1")
pTab.PageFieldOrder = xlOverThenDown
pTab.PageFieldWrapCount = 3
}
```

### PivotTable.PageRange

返回一个 **Range** 对象，该对象表示包含数据透视表中页区域的区域。只读。

**语法**

**express.PageRange**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例选择数据透视表中的页面页眉。*/
function test(){
Application.Worksheets.Item("Sheet1").Activate()
Range("A3").Select()
Application.ActiveCell.PivotTable.PageRange.Select()
}
```

### PivotTable.PageRangeCells

返回一个 **Range** 对象，该对象仅表示指定数据透视表中（包含页字段和项下拉列表）的单元格。

**语法**

**express.PageRangeCells**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例仅选择包含页字段和数据项下拉列表的数据透视表中的单元格。*/
Application.Worksheets.Item(1).PivotTables(1).PageRangeCells.Select()
```

### PivotTable.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotTable 对象的变量。

### PivotTable.PivotColumnAxis

返回 **PivotAxis** 对象，该对象表示整个列轴。只读 **PivotAxis** 类型。

**语法**

**express.PivotColumnAxis**

\*express是一个代表 PivotTable 对象的变量。

### PivotTable.PivotFormulas

返回一个 **PivotFormulas** 对象，该对象表示指定数据透视表的公式集合。只读。

**语法**

**express.PivotFormulas**

\*express是一个代表 PivotTable 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性将返回一个空集合。

**示例**

``` JavaScript
function test(){
let r = 0
for(let i=1;i <= Application.ActiveSheet.PivotTables(1).PivotFormulas.Count;i++) {
    r++
    Cells.Item(r, 1).Value2 = Application.ActiveSheet.PivotTables(1).PivotFormulas.Item(i).Formula
}
}
```

### PivotTable.PivotRowAxis

返回 **PivotAxis** 对象，该对象表示整个行轴。只读 **PivotAxis** 类型。

**语法**

**express.PivotRowAxis**

\*express是一个代表 PivotTable 对象的变量。

### PivotTable.PivotSelection

以标准数据透视表的选定区域格式返回或设置数据透视表的选定区域。 **String** 类型，可读写。

**语法**

**express.PivotSelection**

\*express是一个代表 PivotTable 对象的变量。

**说明**

设置该属性等同于在 *Mode* 参数设置为 **xlDataAndLabel** 时调用 **PivotSelect** 方法。

**示例**

``` JavaScript
/*本示例在第一张工作表上的第一张数据透视表中选择名为“Bob”的销售员的数据和标签。*/
Application.Worksheets.Item(1).PivotTables(1).PivotSelection = "Salesman[Bob]"
```

### PivotTable.PivotSelectionStandard

返回或设置用英语（美国）设置的标准 数据透视表 ?（数据透视表：一种交互的、交叉制表的 ET 报表，用于对多种来源（包括 ET 的外部数据）的数据（如数据库记录）进行汇总和分析。） 格式的数据透视表选定内容的 **String** 类型的数值。可读写。

**语法**

**express.PivotSelectionStandard**

\*express是一个代表 PivotTable 对象的变量。

**说明**

**PivotSelection** 属性是国际通用的，而 **PivotSelectionStandard** 方法则不是。

**示例**

``` JavaScript
/*本示例选择数据透视表中标题为“1.57”的字段，并在其前面插入一个空白列字段。本示例假定数据透视表位于包含标题为“1.57”的列字段的活动工作表上。*/
function test(){
 let pvtTable = Application.ActiveSheet.PivotTables(1)
    pvtTable.PivotSelectionStandard = "1.57"
    Selection.Insert()
}
```

### PivotTable.PreserveFormatting

当使用数据透视、筛选或更改页字段项等操作刷新或重新计算报告时，如果格式处于被保护状态，则为 **True** 。对于查询表，如果数据前五行的任何常规格式设置将应用于查询表数据的新行，则此属性为 **True** 。对未使用的单元格不进行设置。如果将查询表应用的最近一次自动套用格式应用于新数据行，则此属性为 **False** 。默认值是 **True** 。

**语法**

**express.PreserveFormatting**

\*express是一个代表 PivotTable 对象的变量。

**说明**

对于数据库查询表，默认的格式设置为 **xlSimple** 。

刷新查询表时，将对查询表应用新的自动套用格式样式。只要 **PreserveFormatting** 的值为 **False** ，则 AutoFormat（自动套用格式）就会被设置为 **None** 。这样，任何在 **PreserveFormatting** 被设置为 **False** 或在查询表刷新之前设置的自动套用格式都不会起作用，且相应产生的查询表也不会被应用任何格式。

**示例**

``` JavaScript
/*此示例保留第一张工作表上的第一个数据透视表的格式。*/
Application.Worksheets.Item(1).PivotTables("Pivot1").PreserveFormatting = true
```

此示例演示了将 **PreserveFormatting** 设置为 **False** 后，如何操作才能将“自动套用格式”设置为 **xlRangeAutoFormatNone** ，而并不是指定的 **xlRangeAutoFormatColor1** 格式。

``` JavaScript
function test(){
let QTab = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)
QTab.Range.AutoFormat = xlRangeAutoFormatColor1
QTab.PreserveFormatting = false
QTab.Refresh()
}
```

### PivotTable.PrintDrillIndicators

指定是否随数据透视表打印深化指示符。可读写 **Boolean** 类型。

**语法**

**express.PrintDrillIndicators**

\*express是一个代表 PivotTable 对象的变量。

**说明**

如果随数据透视表打印深化指示符，则返回 **True** ；如果不随数据透视表打印深化指示符，则返回 **False** 。

当 **ShowDrillIndicators** 属性设置为 **False** 时， **PrintDrillIndicators** 属性没有效果。如果在数据透视表中未显示深化指示符，则不会打印这些指示符。

### PivotTable.PrintTitles

如果基于数据透视表设置工作表的打印标题，则该属性值为 **True** 。如果使用工作表的打印标题，则该属性值为 **False** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintTitles**

\*express是一个代表 PivotTable 对象的变量。

**说明**

行打印标题设置为包含数据透视表的列字段项的行。列打印标题设置为包含行项目的列。

**示例**

``` JavaScript
/*本示例指定当打印活动工作表上的第四个数据透视表时，同时也打印该工作表的打印标题集。*/
Application.ActiveSheet.PivotTables("PivotTable4").PrintTitles = true
```

### PivotTable.RefreshDate

返回数据透视表或缓存最近一次刷新的日期。 **Date** 型，只读。

**语法**

**express.RefreshDate**

\*express是一个代表 PivotTable 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，在每次查询后都会更新此属性。

**示例**

``` JavaScript
/*此示例显示数据透视表最近一次刷新的日期。*/
function test(){
Set pvtTable = Application.Worksheets("Sheet1").Range("A3").PivotTable
dateString = Format(pvtTable.RefreshDate, "Long Date")
MsgBox "The data was last refreshed on " & dateString
}
```

### PivotTable.RefreshName

返回最近一次刷新数据透视表的人员的名字。 **String** 型，只读。

**语法**

**express.RefreshName**

\*express是一个代表 PivotTable 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，在每次查询后都会更新此属性。

**示例**

``` JavaScript
/*此示例显示最近一次更新数据透视表的用户的名字。*/
function test(){
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
alert("The data was last refreshed by " + pvtTable.RefreshName)
}
```

### PivotTable.RepeatItemsOnEachPrintedPage

当打印指定的数据透视表时，如果每页第一行上都显示行、列和项标志，则该值为 **True** 。如果仅在第一页上打印这些标志，则该值为 **False** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.RepeatItemsOnEachPrintedPage**

\*express是一个代表 PivotTable 对象的变量。

**说明**

对于工作表，ET 打印行和列标签，而不是打印标题。使用 **PrintTitles** 属性可判断是否为数据透视表设置了打印标题。

**示例**

``` JavaScript
/*在打印活动工作表的第四个数据透视表时，本示例使 ET 在每页上重复打印标签。*/
Application.ActiveSheet.PivotTables("PivotTable4").RepeatItemsOnEachPrintedPage = true
```

### PivotTable.RowFields

返回一个对象，该对象表示数据透视表中的单个字段（ **PivotField** 对象），或者当前未显示为行字段的所有字段的一个集合（ **PivotFields** 对象）。只读。

**语法**

**express.RowFields**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例在新工作表上的清单中添加数据透视表的行字段名称。*/
function test(){
let nwSheet = Application.Worksheets.Add()
nwSheet.Activate()
let pvtTable = 
Worksheets("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i = 
 1;  i < = pvtTable.RowFields().Count; i++) {
    rw = rw + 1
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.RowFields(i).Name
}
}
```

### PivotTable.RowGrand

如果数据透视表显示行的总数，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.RowGrand**

\*express是一个代表 PivotTable 对象的变量。

**示例**

``` JavaScript
/*本示例设置数据透视表以显示行的总计。*/
function test(){
let pvtTable = Application.Worksheets.Item("Sheet1").Range("A3").PivotTable
pvtTable.RowGrand = true
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotTable](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CoAuthor 对象

## 简介

代表文档中的单个共同作者。 CoAuthor 对象是 CoAuthors 集合的一个成员。 CoAuthors 集合包含文档中的所有共同作者（当前正在编辑文档的作者）。

## 说明

使用 CoAuthors ( *Index* ) 可以返回单个 CoAuthor 对象，其中 *Index* 是索引号。

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
<td>当一个新的共同作者开始编辑文档时，该作者可能需要一分钟或更长时间才会显示在文档中。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

``` JavaScript
/*下面的代码示例返回活动文档中第一个共同作者的名称。*/
function test() {
    let author = Application.ActiveDocument.CoAuthoring.Authors.Item(1)
    alert("The name of the first co-author in this document is " + author.Name)
}
```

## 属性

| 序号 | 名称                                   | 说明                                                                        |
|:----:|:---------------------------------------|:----------------------------------------------------------------------------|
|  1   | [Application](#CoAuthor.Application)   | 返回一个代表 WPS 应用程序的 Application 对象。只读。                        |
|  2   | [Creator](#CoAuthor.Creator)           | 返回表示用于创建指定对象的应用程序的 32 位整数。只读 Long 类型。            |
|  3   | [EmailAddress](#CoAuthor.EmailAddress) | 返回一个字符串，指明指定合著者的电子邮件地址。只读。                        |
|  4   | [ID](#CoAuthor.ID)                     | 返回一个 String 类型的值，该值指定某个指定作者的唯一标识符。只读。          |
|  5   | [IsMe](#CoAuthor.IsMe)                 | 如果此作者代表当前用户，则返回 true。只读。                                 |
|  6   | [Locks](#CoAuthor.Locks)               | 返回一个 CoAuthLocks 集合，该集合代表文档中与指定共同作者相关联的锁。只读。 |
|  7   | [Name](#CoAuthor.Name)                 | 返回一个 String 类型的值，该值包含指定的共同作者的显示名称。只读。          |
|  8   | [Parent](#CoAuthor.Parent)             | 返回一个 Object 类型值，该值代表指定 CoAuthor 对象的父对象。                |

## 成员属性

### CoAuthor.Application

返回一个代表 WPS 应用程序的 Application 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CoAuthor 对象的变量。

### CoAuthor.Creator

返回表示用于创建指定对象的应用程序的 32 位整数。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CoAuthor 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表 **string** “WPS”。该属性主要设计用于 Apple Macintosh 平台，在该平台上，每个应用程序都有一个由四个字符组成的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的详细信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

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
<td>该值也可用常量 <strong>wdCreatorCode</strong> 表示。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

### CoAuthor.EmailAddress

返回一个字符串，指明指定合著者的电子邮件地址。只读。

**语法**

**express.EmailAddress**

\*express是一个代表 CoAuthor 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示活动文档中第一个合著者的电子邮件地址。*/
function test() {
    if (!Application.ActiveDocument.CoAuthoring.Authors.Count) {
        alert(Application.ActiveDocument.CoAuthoring.Authors.Item(1).EmailAddress)
    }
    alert("There are no co-authors in this document.")
}
```

### CoAuthor.ID

返回一个 **String** 类型的值，该值指定某个指定作者的唯一标识符。只读。

**语法**

**express.ID**

\*express是一个代表 CoAuthor 对象的变量。

**说明**

不应假定 **ID** 属性返回的唯一标识符具有特定的长度或格式。

**示例**

``` JavaScript
/*下面的代码示例显示活动文档中的每个共同作者的唯一标识符。*/
function test() {
    let allAuthors = Application.ActiveDocument.CoAuthoring.Authors
    for (let i = 1; i <= allAuthors.Count; i++) {
        alert("The ID for  " + allAuthors.Item(i).Name + " is " + allAuthors.Item(i).ID + ".")
    }
}
```

### CoAuthor.IsMe

如果此作者代表当前用户，则返回 true。只读。

**语法**

**express.IsMe**

\*express是一个代表 CoAuthor 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例检查活动文档，以查看 CoAuthors 集合中的第一个共同作者是否为当前用户。*/
function test() {
    if (Application.ActiveDocument.CoAuthoring.Authors.Item(1).IsMe) {
        alert("The current user is the first coauthor.")
    }
}
```

### CoAuthor.Locks

返回一个 CoAuthLocks 集合，该集合代表文档中与指定共同作者相关联的锁。只读。

**语法**

**express.Locks**

\*express是一个代表 CoAuthor 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示活动文档中与第一个共同作者相关联的锁的数量。*/
function test() {
    let lockCount
    let coAuth = Application.ActiveDocument.CoAuthoring.Authors.Item(1)
    lockCount = coAuth.Locks.Count
    alert("There are " + lockCount + " locks in the active document for " + coAuth.Name + ".")
}
```

### CoAuthor.Name

返回一个 **String** 类型的值，该值包含指定的共同作者的显示名称。只读。

**语法**

**express.Name**

\*express是一个代表 CoAuthor 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示活动文档中第一个共同作者的名称。*/
function test() {
    let author = Application.ActiveDocument.CoAuthoring.Authors.Item(1)
    alert("The name of the first co-author in this document is " + author.Name)
}
```

### CoAuthor.Parent

返回一个 **Object** 类型值，该值代表指定 **CoAuthor** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 CoAuthor 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/CoAuthor](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Dialog 对象

## 简介

代表一个内置对话框。 Dialog 对象是 Dialogs 集合的成员。 Dialogs 集合包含 WPS 中的所有内置对话框。不能在 Dialogs 集合中创建新的内置对话框，也不能向其中添加内置对话框。

## 说明

使用 Dialogs ( *Index* ) 返回单个 Dialog 对象（其中 *Index* 是标识对话框的 WdWordDialog 常量）。以下示例显示和执行在内置的 “打开” 对话框中采取的操作。

``` JavaScript
Application.Dialogs.Item(wdDialogFileOpen).Show()
```

WdWordDialog 常量由“wdDialog”前缀加上菜单和对话框的名称组成。例如， “页面设置” 对话框的该常量为 wdDialogFilePageSetup ，而 “新建” 对话框的该常量为 wdDialogFileNew 。

## 方法

| 序号 | 名称                       | 说明                                                                                                                  |
|:----:|:---------------------------|:----------------------------------------------------------------------------------------------------------------------|
|  1   | [Display](#Dialog.Display) | 显示指定的 WPS 内置对话框，直到用户关闭该对话框或已超过指定时限。返回一个 Number 类型值，指明关闭对话框时单击的按钮。 |
|  2   | [Execute](#Dialog.Execute) | 应用 WPS 对话框的当前设置。                                                                                           |
|  3   | [Show](#Dialog.Show)       | 显示和执行在指定的 WPS 内置对话框中的初始操作。返回一个 Number 类型值，指明关闭对话框时单击的按钮。                   |
|  4   | [Update](#Dialog.Update)   | 更新内置 WPS 对话框中显示的值。                                                                                       |

## 属性

| 序号 | 名称                                 | 说明                                                                            |
|:----:|:-------------------------------------|:--------------------------------------------------------------------------------|
|  1   | [Application](#Dialog.Application)   | 返回一个代表 WPS 应用程序的 Application 对象。                                  |
|  2   | [CommandBarId](#Dialog.CommandBarId) | 返回一个 Number 类型的值，该值代表内置 WPS 对话框的工具栏控件 ID。只读。        |
|  3   | [CommandName](#Dialog.CommandName)   | 返回用来显示指定内置对话框的过程名。 String 类型，只读。                        |
|  4   | [Creator](#Dialog.Creator)           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。    |
|  5   | [DefaultTab](#Dialog.DefaultTab)     | 显示指定的对话框时，该属性返回或设置活动选项卡。 WdWordDialogTab 类型，可读写。 |
|  6   | [Parent](#Dialog.Parent)             | 返回一个 Object 类型值，该值代表指定 Dialog 对象的父对象。                      |
|  7   | [Type](#Dialog.Type)                 | 返回内置 WPS 对话框的类型。只读 WdWordDialog 类型。                             |

## 成员方法

### Dialog.Display

显示指定的 WPS 内置对话框，直到用户关闭该对话框或已超过指定时限。返回一个 **Number** 类型值，指明关闭对话框时单击的按钮。

**语法**

**express.Display ( *TimeOut* )**

\*express是一个代表 Dialog 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                                           |
|-----------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------|
| *TimeOut* | 可选      | Variant  | 自动关闭对话框前 WPS 等待的时间。每一单位大约为 0.001 秒。并发系统活动可能会增加有效时间值。如果省略这一参数，对话框只有在用户关闭它时才关闭。 |

**说明**

**Display** 方法返回下列可能值。

| 返回值     | 说明                                                     |
|------------|----------------------------------------------------------|
| -2         | **“关闭”** 按钮。                                        |
| -1         | **“确定”** 按钮。                                        |
| 0（零）    | **“取消”** 按钮。                                        |
| \> 0（零） | 命令按钮：1 代表第一个按钮，2 代表第二个按钮，依此类推。 |

**示例**

``` JavaScript
/* 本示例显示“关于”对话框 */
Application.Dialogs.Item(wdDialogHelpAbout).Display()
```

### Dialog.Execute

应用 WPS 对话框的当前设置。

**语法**

**express.Execute ()**

\*express是一个代表 Dialog 对象的变量。

### Dialog.Show

显示和执行在指定的 WPS 内置对话框中的初始操作。返回一个 **Number** 类型值，指明关闭对话框时单击的按钮。

**语法**

**express.Show ( *TimeOut* )**

\*express是一个代表 Dialog 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                                           |
|-----------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------|
| *TimeOut* | 可选      | Variant  | 自动关闭对话框前 WPS 等待的时间。每一单位大约为 0.001 秒。并发系统活动可能会增加有效时间值。如果省略这一参数，对话框只有在用户关闭它时才关闭。 |

**说明**

下表显示 **Show** 方法所返回的值的含义。

| 返回值     | 说明                                                     |
|------------|----------------------------------------------------------|
| -2         | “关闭” 按钮。                                            |
| -1         | “确定” 按钮。                                            |
| 0（零）    | “取消” 按钮。                                            |
| \> 0（零） | 命令按钮：1 代表第一个按钮，2 代表第二个按钮，依此类推。 |

**示例**

``` JavaScript
/* 本示例将显示并执行在“打开”对话框中启动的任意操作。文件名设置为“*.*”可显示所有文件名*/
function test() {
    let fpx = Application.Dialogs.Item(wdDialogFileOpen)
    fpx.Name = "*.*"
    fpx.Show()
}
```

### Dialog.Update

更新内置 WPS 对话框中显示的值。

**语法**

**express.Update ()**

\*express是一个代表 Dialog 对象的变量。

**示例**

``` JavaScript
/* 本示例返回引用“字体”对话框的 Dialog 对象，将应用于 Selection 对象的字体更改为 Arial，更新对话框的值，并显示“字体”对话框 */
function test() {
    let myDialog = Application.Dialogs.Item(wdDialogFormatFont)
    Application.Selection.Font.Name = "Arial"
    myDialog.Update()
    myDialog.Show()
}
```

## 成员属性

### Dialog.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Dialog 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Dialog.CommandBarId

返回一个 **Number** 类型的值，该值代表内置 WPS 对话框的工具栏控件 ID。只读。

**语法**

**express.CommandBarId**

\*express是一个代表 Dialog 对象的变量。

### Dialog.CommandName

返回用来显示指定内置对话框的过程名。 **String** 类型，只读。

**语法**

**express.CommandName**

\*express是一个代表 Dialog 对象的变量。

**说明**

有关使用 WPS 内置对话框的详细信息，请参阅 显示内置的 WPS 对话框。

### Dialog.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Dialog 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Dialog.DefaultTab

显示指定的对话框时，该属性返回或设置活动选项卡。 **WdWordDialogTab** 类型，可读写。

**语法**

**express.DefaultTab**

\*express是一个代表 Dialog 对象的变量。

**示例**

``` JavaScript
/* 本示例显示所选的“页面设置”对话框中的“纸张来源”选项卡 */
function test() {
    let rng = Application.Dialogs.Item(wdDialogFilePageSetup)
    rng.DefaultTab = wdDialogFilePageSetupTabPaperSource
    rng.Show()
}
```

### Dialog.Parent

返回一个 **Object** 类型值，该值代表指定 **Dialog** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Dialog 对象的变量。

### Dialog.Type

返回内置 WPS 对话框的类型。只读 WdWordDialog 类型。

**语法**

**express.Type**

\*express是一个代表 Dialog 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Dialog](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DropDown 对象

## 简介

代表包含一个窗体选项列表的下拉型窗体域。

## 说明

使用 FormFields (index) 返回单个 FormField 对象（其中 index 是与下拉型窗体域相关的索引编号或书签名称）。可将 DropDown 属性与 FormField 对象配合使用，返回 DropDown 的对象。以下示例选择活动文档中名为“DropDown”的下拉型窗体域中的第一项。

``` JavaScript
ActiveDocument.FormFields.Item("DropDown1").DropDown.Value = 1
```

索引编号代表窗体域在 FormFields 集合中的位置。以下示例检查活动文档中第一个窗体域的类型。如果是下拉型窗体域，则选中第二项。

``` JavaScript
function test()
{
if(ActiveDocument.FormFields.Item(1).Type == wdFieldFormDropDown) {
    ActiveDocument.FormFields.Item(1).DropDown.Value = 2
}
}
```

以下示例在向 *ffield* 窗体域中添加选项以前，判断该下拉型窗体域是否有效。

``` JavaScript
function test()
{
let ffield = ActiveDocument.FormFields.Item(1).DropDown
if(ffield.Valid == true) {
    ffield.ListEntries.Add("Hello")
}
else {
    MsgBox("First field is not a drop down")
}
}
```

可将 Add 方法用于 FormFields 集合来添加下拉型窗体域。以下示例在活动文档的开头添加一个下拉型窗体域，然后在该窗体域中添加项目。

``` JavaScript
function test()
{
let ffield = ActiveDocument.FormFields.Add( ActiveDocument.Range(0, 0), wdFieldFormDropDown)
let ff = ActiveDocument.FormFields.Item(1)
ff.Name = "Colors"
ff.DropDown.ListEntries
ff.DropDown.ListEntries.Add("Blue")
ff.DropDown.ListEntries.Add("Green")
ff.DropDown.ListEntries.Add("Red")
}
```

## 属性

| 序号 | 名称                                 | 说明                                                                            |
|:----:|:-------------------------------------|:--------------------------------------------------------------------------------|
|  1   | [Application](#DropDown.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                  |
|  2   | [Creator](#DropDown.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。    |
|  3   | [Default](#DropDown.Default)         | 返回或设置一个 Long 类型值，该值代表默认的下拉型项目。可读写。                  |
|  4   | [ListEntries](#DropDown.ListEntries) | 返回一个 ListEntries 集合，该集合代表 DropDown 对象中所有项。                   |
|  5   | [Parent](#DropDown.Parent)           | 返回一个 Object 类型值，该值代表指定 DropDown 对象的父对象。                    |
|  6   | [Valid](#DropDown.Valid)             | 如果指定的窗体域对象为有效的下拉窗体域，则该属性为 True 。 Boolean 类型，只读。 |
|  7   | [Value](#DropDown.Value)             | 返回或设置下拉型窗体域中所选项目的编号。可读写 Long 类型。                      |

## 成员属性

### DropDown.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 DropDown 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### DropDown.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DropDown 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### DropDown.Default

返回或设置一个 **Long** 类型值，该值代表默认的下拉型项目。可读写。

**语法**

**express.Default**

\*express是一个代表 DropDown 对象的变量。

**说明**

下拉型窗体域的第一项为 1，第二项为 2，以此类推。

**示例**

``` JavaScript
/*本示例为 Sales.doc 中名为“Colors”的下拉型窗体域设置默认项。*/
Documents.Item("Sales.doc").FormFields.Item("Colors").DropDown.Default = 2
```

### DropDown.ListEntries

返回一个 **ListEntries** 集合，该集合代表 **DropDown** 对象中所有项。

**语法**

**express.ListEntries**

\*express是一个代表 DropDown 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例在名为“DropDown1”的下拉型窗体域中检索活动项文字。*/
function test()
{
let myField = ActiveDocument.FormFields.Item("DropDown1").DropDown
let num = myField.Value
let myName = myField.ListEntries.Item(Number(num)).Name
}
```

``` JavaScript
/*本示例在活动下拉型窗体域中检索项的总数（应该先保护文档的窗体）。如果有两个或更多的项，本示例将第二项设置为活动项。*/
function test()
{
let myField = Selection.FormFields.Item(1)
if(myField.Type == wdFieldFormDropDown) {
    num = myField.DropDown.ListEntries.Count
    if(num > = 2) {
        myField.DropDown.Value = 2
}
}
```

### DropDown.Parent

返回一个 **Object** 类型值，该值代表指定 **DropDown** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 DropDown 对象的变量。

### DropDown.Valid

如果指定的窗体域对象为有效的下拉窗体域，则该属性为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.Valid**

\*express是一个代表 DropDown 对象的变量。

**说明**

在应用 **DropDown** 属性之前，请使用 **FormField** 对象的 **Type** 属性来确定窗体域的类型 ( **wdFieldFormDropDown** )。该措施可确保 **FormField** 对象为要求的类型。

### DropDown.Value

返回或设置下拉型窗体域中所选项目的编号。可读写 **Long** 类型。

**语法**

**express.Value**

\*express是一个代表 DropDown 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/DropDown](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Endnote 对象

## 简介

代表一个尾注。 Endnote 对象是 Endnotes 集合的一个成员，该集合代表所选内容、范围或文档的尾注。

## 说明

使用 Endnotes ( *Index* ) 可返回一个 Endnote 对象，其中 *Index* 是索引号。索引号代表尾注在所选内容、范围或文档中的位置。以下示例将所选内容中的第一个尾注设置成红色。

``` JavaScript
function test() {
    if(Application.Selection.Endnotes.Count >= 1) {
        Application.Selection.Endnotes.Item(1).Reference.Font.ColorIndex = wdRed
    }
}
```

使用 Add 方法可向 Endnotes 集合添加尾注。以下示例紧接着所选内容添加一个尾注。

``` JavaScript
function test() {
    Application.Selection.Collapse(wdCollapseEnd)
    Application.ActiveDocument.Endnotes.Add(Selection.Range , null, "The Willow Tree, (Lone Creek Press, 1996).")
}
```

## 方法

| 序号 | 名称                      | 说明             |
|:----:|:--------------------------|:-----------------|
|  1   | [Delete](#Endnote.Delete) | 删除指定的尾注。 |

## 属性

| 序号 | 名称                                | 说明                                                                         |
|:----:|:------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Endnote.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Creator](#Endnote.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  3   | [Index](#Endnote.Index)             | 返回一个 Long 类型的值，该值代表项目在集合中的位置。只读。                   |
|  4   | [Parent](#Endnote.Parent)           | 返回一个 Object 类型值，该值代表指定 Endnote 对象的父对象。                  |
|  5   | [Range](#Endnote.Range)             | 返回 Range 对象，该对象表示文档包含在指定对象中的一部分。                    |
|  6   | [Reference](#Endnote.Reference)     | 返回 Range 对象，该对象表示尾注引用标记。                                    |

## 成员方法

### Endnote.Delete

删除指定的尾注。

**语法**

**express.Delete ()**

\*express是一个代表 Endnote 对象的变量。

## 成员属性

### Endnote.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 Endnote 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Endnote.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Endnote 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Endnote.Index

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Endnote 对象的变量。

### Endnote.Parent

返回一个 **Object** 类型值，该值代表指定 **Endnote** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Endnote 对象的变量。

### Endnote.Range

返回 **Range** 对象，该对象表示文档包含在指定对象中的一部分。

**语法**

**express.Range**

\*express是一个代表 Endnote 对象的变量。

**说明**

有关从文档返回一个区域或从形状集合返回形状区域的信息，请参阅 **Range** 方法。

**示例**

``` JavaScript
/*以下示例更改活动文档中第一个尾注的文本。*/
function test() {
  let range = Application.ActiveDocument.Endnotes.Item(1).Range
  range.Delete()
  range.Text = "new endnote text"
}
```

### Endnote.Reference

返回 **Range** 对象，该对象表示尾注引用标记。

**语法**

**express.Reference**

\*express是一个代表 Endnote 对象的变量。

**说明**

此示例将 *myRange* 设置为活动文档中的第一个尾注引用标记，然后复制该引用标记。

**示例**

``` JavaScript
function test() {
  let myRange = Application.ActiveDocument.Endnotes.Item(1).Reference
  myRange.Copy()
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Endnote](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# EndnoteOptions 对象

## 简介

代表指定给文档中尾注的某范围或所选内容的属性。

## 说明

使用 Range 或 Selection 对象的 EndnoteOptions 属性可返回 EndnoteOptions 对象。

使用 EndnoteOptions 对象可为文档的不同区域指定不同的尾注属性。例如，您可能希望将长文档说明部分的尾注用小写罗马数字显示，而文档中其余部分的尾注用阿拉伯数字显示。以下示例使用 NumberingRule 、 NumberStyle 和 StartingNumber 属性设置活动文档中第一节的尾注格式。

``` JavaScript
function BookIntro() {
    //Sets the range as section one of the active document
    let rngIntro = Application.ActiveDocument.Sections.Item(1).Range

    //Formats the EndnoteOptions properties
    let re = rngIntro.EndnoteOptions
        re.NumberingRule = wdRestartSection
        re.NumberStyle = 2
        re.StartingNumber = 1
}
```

## 属性

| 序号 | 名称                                             | 说明                                                                          |
|:----:|:-------------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#EndnoteOptions.Application)       |                                                                               |
|  2   | [Creator](#EndnoteOptions.Creator)               | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。  |
|  3   | [Location](#EndnoteOptions.Location)             | 返回或设置所有尾注的位置。 WdEndnoteLocation 类型，可读写。                   |
|  4   | [NumberStyle](#EndnoteOptions.NumberStyle)       | 返回或设置尾注的编号样式。 WdNoteNumberStyle 类型，可读写。                   |
|  5   | [NumberingRule](#EndnoteOptions.NumberingRule)   | 返回或设置在分页符或分节符之后脚注或尾注的编号方式。可读写 WdNumberingRule 。 |
|  6   | [Parent](#EndnoteOptions.Parent)                 | 返回一个 Object 类型值，该值代表指定 EndnoteOptions 对象的父对象。            |
|  7   | [StartingNumber](#EndnoteOptions.StartingNumber) | 返回或设置尾注的起始编号。 Long 类型，可读写。                                |

## 成员属性

### EndnoteOptions.Application

**语法**

**express.Application**

\*express是一个代表 EndnoteOptions 对象的变量。

### EndnoteOptions.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 EndnoteOptions 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### EndnoteOptions.Location

返回或设置所有尾注的位置。 **WdEndnoteLocation** 类型，可读写。

**语法**

**express.Location**

\*express是一个代表 EndnoteOptions 对象的变量。

**示例**

``` JavaScript
/*本示例将所有的尾注置于节的结尾。*/
Application.ActiveDocument.Endnotes.Location = wdEndOfSection
```

### EndnoteOptions.NumberStyle

返回或设置尾注的编号样式。 **WdNoteNumberStyle** 类型，可读写。

**语法**

**express.NumberStyle**

\*express是一个代表 EndnoteOptions 对象的变量。

**说明**

某些 **WdNoteNumberStyle** 常量可能不可用，具体取决于所选择或安装的语言支持（例如，美国英语）。

### EndnoteOptions.NumberingRule

返回或设置在分页符或分节符之后脚注或尾注的编号方式。可读写 WdNumberingRule 。

**语法**

**express.NumberingRule**

\*express是一个代表 EndnoteOptions 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档的每一分节符后重新开始尾注的编号。*/
Application.ActiveDocument.Endnotes.NumberingRule = wdRestartSection
```

### EndnoteOptions.Parent

返回一个 **Object** 类型值，该值代表指定 **EndnoteOptions** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 EndnoteOptions 对象的变量。

### EndnoteOptions.StartingNumber

返回或设置尾注的起始编号。 **Long** 类型，可读写。

**语法**

**express.StartingNumber**

\*express是一个代表 EndnoteOptions 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/EndnoteOptions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Footnote 对象

## 简介

代表位于页面底端或文字下方的脚注。 Footnote 对象是 Footnotes 集合的一个成员。 Footnotes 集合代表所选内容、范围或文档中的脚注。

## 说明

使用 Footnotes ( *Index* ) 可返回一个 Footnote 对象，其中 *Index* 是索引号。索引号代表脚注在所选内容、范围或文档中的位置。以下示例对所选内容中的第一个脚注应用红色格式。

``` JavaScript
if(Application.Selection.Footnotes.Count >= 1) {
    Application.Selection.Footnotes.Item(1).Reference.Font.ColorIndex = wdRed
}
```

使用 Add 方法可将脚注添加到 Footnotes 集合。以下示例紧接所选内容之后插入一条自动编号的脚注。

``` JavaScript
function test() {
    Application.Selection.Collapse(wdCollapseEnd)
    Application.ActiveDocument.Footnotes.Add(Selection.Range ,"The Willow Tree, (Lone Creek Press, 1996).")
}
```

## 方法

| 序号 | 名称                       | 说明             |
|:----:|:---------------------------|:-----------------|
|  1   | [Delete](#Footnote.Delete) | 删除指定的脚注。 |

## 属性

| 序号 | 名称                                 | 说明                                                                         |
|:----:|:-------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Footnote.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Creator](#Footnote.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  3   | [Index](#Footnote.Index)             | 返回一个 Long 类型的值，该值代表项目在集合中的位置。只读。                   |
|  4   | [Parent](#Footnote.Parent)           | 返回一个 Object 类型值，该值代表指定 Footnote 对象的父对象。                 |
|  5   | [Range](#Footnote.Range)             | 返回一个 Range 对象，该对象代表脚注中包含的文档部分。                        |
|  6   | [Reference](#Footnote.Reference)     | 返回一个 Range 对象，该对象代表脚注引用标记。                                |

## 成员方法

### Footnote.Delete

删除指定的脚注。

**语法**

**express.Delete ()**

\*express是一个代表 Footnote 对象的变量。

## 成员属性

### Footnote.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Footnote 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Footnote.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Footnote 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Footnote.Index

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Footnote 对象的变量。

### Footnote.Parent

返回一个 **Object** 类型值，该值代表指定 **Footnote** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Footnote 对象的变量。

### Footnote.Range

返回一个 **Range** 对象，该对象代表脚注中包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 Footnote 对象的变量。

### Footnote.Reference

返回一个 **Range** 对象，该对象代表脚注引用标记。

**语法**

**express.Reference**

\*express是一个代表 Footnote 对象的变量。

**示例**

``` JavaScript
/*以下示例将 myRange 设置为活动文档中的第一个脚注引用标记，然后复制该引用标记。*/
function test()
{
let myRange = ActiveDocument.Footnotes.Item(1).Reference
myRange.Copy()
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Footnote](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FormFields 对象

## 简介

FormField 对象的集合，这些对象代表所选内容、范围或文档中的所有窗体域。

## 方法

| 序号 | 名称                     | 说明                                                          |
|:----:|:-------------------------|:--------------------------------------------------------------|
|  1   | [Add](#FormFields.Add)   | 返回一个 FormField 对象，该对象代表添加到区域中的新的窗体域。 |
|  2   | [Item](#FormFields.Item) | 返回集合中的单个 FormField 对象。                             |

## 属性

| 序号 | 名称                                   | 说明                                                                         |
|:----:|:---------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#FormFields.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#FormFields.Count)             | 返回一个 Long 类型的值，该值代表集合中域的数量。只读。                       |
|  3   | [Creator](#FormFields.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#FormFields.Parent)           | 返回一个 Object 类型值，该值代表指定 FormFields 对象的父对象。               |
|  5   | [Shaded](#FormFields.Shaded)           | 如果将底纹应用于窗体域，则该属性值为 True 。 Boolean 类型，可读写。          |

## 成员方法

### FormFields.Add

返回一个 **FormField** 对象，该对象代表添加到区域中的新的窗体域。

**语法**

**express.Add ( *Range* , *Type* )**

\*express是一个代表 FormFields 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型    | 说明                                                           |
|---------|-----------|-------------|----------------------------------------------------------------|
| *Range* | 必选      | Range 对象  | 要添加窗体域的区域。如果该区域未折叠，那么窗体域将替换此区域。 |
| *Type*  | 必选      | WdFieldType | 要添加的窗体域的类型。                                         |

**说明**

**安全性:** 动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

**示例**

``` JavaScript
/*本示例在选定内容的末尾处添加一个复选框，然后命名并选定该复选框。*/
function test() {
  Application.Selection.Collapse(wdCollapseEnd)
  
  let ffield = Application.ActiveDocument.FormFields.Add(Selection.Range, wdFieldFormCheckBox)
  ffield.Name = "Check_Box_1"
  ffield.CheckBox.Value = true
}
```

### FormFields.Item

返回集合中的单个 **FormField** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 FormFields 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

## 成员属性

### FormFields.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 FormFields 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### FormFields.Count

返回一个 **Long** 类型的值，该值代表集合中域的数量。只读。

**语法**

**express.Count**

\*express是一个代表 FormFields 对象的变量。

### FormFields.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FormFields 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### FormFields.Parent

返回一个 **Object** 类型值，该值代表指定 **FormFields** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 FormFields 对象的变量。

### FormFields.Shaded

如果将底纹应用于窗体域，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Shaded**

\*express是一个代表 FormFields 对象的变量。

**说明**

使用底纹便于在文档中查找窗体域，且不影响打印输出。

**示例**

``` JavaScript
/*本示例从 Employment Form.doc 的窗体域中清除底纹。*/
Application.Documents.Item("Employment Form.doc").FormFields.Shaded = false

/*本示例向活动文档的窗体域添加底纹，并且保护文档中的窗体。*/
function test() {
  Application.ActiveDocument.FormFields.Shaded = true
  Application.ActiveDocument.Protect(wdAllowOnlyFormFields, true)
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/FormFields](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PageNumber 对象

## 简介

代表页眉或页脚中的页码。 PageNumber 对象是 PageNumbers 集合的一个成员。 PageNumbers 集合包含单个页眉或页脚中的所有页码。

## 说明

使用 PageNumbers ( *Index* ) 可返回一个 PageNumber 对象，其中 *Index* 为索引号。在大多数情况下，一个页眉或页脚只包含一个页码（索引号为 1）。以下示例将活动文档第一节的主页眉中的起始页码居中。

``` JavaScript
Application.ActiveDocument.Sections.Item(1).Headers(wdHeaderFooterPrimary).PageNumbers.Item(1).Alignment = wdAlignPageNumberCenter
```

使用 Add 方法可在页眉或页脚中添加页码（PAGE 域）。以下示例向第一节和任何后续节中的主页脚添加一个页码。

``` JavaScript
function test()
{
    let sect = Application.ActiveDocument.Sections.Item(1)
    sect.Footers(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberLeft, true)
}
```

## 属性

| 序号 | 名称                               | 说明                                                                          |
|:----:|:-----------------------------------|:------------------------------------------------------------------------------|
|  1   | [Alignment](#PageNumber.Alignment) | 返回或设置一个 WdPageNumberAlignment 常量，该常量代表页码的对齐方式，可读写。 |

## 成员属性

### PageNumber.Alignment

返回或设置一个 **WdPageNumberAlignment** 常量，该常量代表页码的对齐方式，可读写。

**语法**

**express.Alignment**

\*express是一个代表 PageNumber 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/PageNumber](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XMLNode 对象

## 简介

表示应用于文档的单个 XML 元素。

## 说明

每个已应用于文档的 XML 元素都显示为 “XML 结构” 任务窗格的树视图控件中的一个节点。树视图中的每个节点都是 XMLNode 对象的一个实例。树视图的层次结构表明节点是否包含子节点。

使用 XMLNodes 集合的 Item 方法可返回一个 XMLNode 对象。使用 Validate 方法可根据所应用的架构验证 XML 元素是否有效、必需的子元素是否存在并按要求的顺序显示。运行 Validate 方法后，使用 ValidationStatus 属性可验证元素是否有效，使用 ValidationErrorText 属性可显示关于如下内容的信息：用户应采用哪些措施来确保文档符合该 XML 架构的规则。

以下示例验证活动文档中的每个 XML 元素。如果根据架构发现元素无效，则该示例向用户返回一条消息解释问题所在。

``` JavaScript
function test() {
    let objNode = Application.ActiveDocument.XMLNodes
    for(let i = 1; i <= objNode.Count; i++) {
        objNode.Item(i).Validate()
        if(objNode.ValidationStatus != wdXMLValidationStatusOK) {
            alert(objNode.ValidationErrorText(true))
        }
    }
}
```

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#XMLNode.Copy">Copy</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定的 XML 元素（不包括 XML 标记）复制到剪贴板。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#XMLNode.Cut">Cut</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定的 XML 元素从文档中移到剪贴板上。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#XMLNode.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>删除 XML 文档中指定的 XML 元素。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#XMLNode.RemoveChild">RemoveChild</a></td>
<td style="text-align: left;" data-valian="middle"><p>从指定的元素中删除子元素。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#XMLNode.SelectNodes">SelectNodes</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 XMLNodes 集合，该集合代表与 XPath 参数匹配的所有子元素（它们的顺序与它们在指定的 XML 元素中出现的顺序相同）。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#XMLNode.SelectSingleNode">SelectSingleNode</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 XMLNode 对象，该对象代表指定的 XML 元素中第一个与 XPath 参数匹配的子元素。</p>
<p>返回 XMLNode值</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#XMLNode.SetValidationError">SetValidationError</a></td>
<td style="text-align: left;" data-valian="middle"><p>更改向用户显示的指定节点的验证错误文本，并强制 WPS 将节点报告为无效。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#XMLNode.Validate">Validate</a></td>
<td style="text-align: left;" data-valian="middle"><p>针对附加到文档的 XML 架构验证单个 XML 元素。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                                                  | 说明                                                                                                                  |
|:----:|:------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#XMLNode.Application)                   | 返回一个代表 WPS 应用程序的 Application 对象。                                                                        |
|  2   | [Attributes](#XMLNode.Attributes)                     | 返回一个 XMLNodes 集合，该集合代表指定元素的属性。                                                                    |
|  3   | [BaseName](#XMLNode.BaseName)                         | 返回一个 String 类型的值，该值代表不带任何前缀的元素名称。                                                            |
|  4   | [ChildNodeSuggestions](#XMLNode.ChildNodeSuggestions) | 返回一个 XMLChildNodeSuggestions 集合，代表指定元素的子元素列表。                                                     |
|  5   | [ChildNodes](#XMLNode.ChildNodes)                     | 返回一个 XMLNodes 集合，该集合代表指定元素的子元素。                                                                  |
|  6   | [Creator](#XMLNode.Creator)                           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                          |
|  7   | [FirstChild](#XMLNode.FirstChild)                     | 返回一个 DiagramNode 对象，该对象代表父节点的第一个子节点。只读。                                                     |
|  8   | [HasChildNodes](#XMLNode.HasChildNodes)               | 返回 Boolean 值，该值表示 XML 节点是否有子节点。只读。                                                                |
|  9   | [LastChild](#XMLNode.LastChild)                       | 返回一个 XMLNode 对象，该对象代表 XML 元素的最后一个子节点。                                                          |
|  10  | [Level](#XMLNode.Level)                               | 返回一个 WdXMLNodeLevel 常量，该常量代表 XML 元素是段落的一部分、是整个段落、包含在表格单元格中还是包含表格行。只读。 |
|  11  | [NamespaceURI](#XMLNode.NamespaceURI)                 | 返回一个 String 类型的值，该值代表指定对象的架构命名空间的统一资源标识符 (URI)。只读。                                |
|  12  | [NextSibling](#XMLNode.NextSibling)                   | 返回一个 XMLNode 对象，该对象代表文档中与指定元素同级的下一个元素。                                                   |
|  13  | [NodeType](#XMLNode.NodeType)                         | 返回一个 WdXMLNodeType 常量，该常量代表节点的类型。                                                                   |
|  14  | [NodeValue](#XMLNode.NodeValue)                       | 返回或设置一个 String 类型的值，该值代表 XML 元素的值。可读写。                                                       |
|  15  | [OwnerDocument](#XMLNode.OwnerDocument)               | 返回一个代表指定 XML 元素的父文档的 Document 对象。                                                                   |
|  16  | [Parent](#XMLNode.Parent)                             | 返回一个 Object 类型值，该值代表指定 XMLNode 对象的父对象。                                                           |
|  17  | [ParentNode](#XMLNode.ParentNode)                     | 返回一个代表指定元素的父元素的 XMLNode 对象。                                                                         |
|  18  | [PlaceholderText](#XMLNode.PlaceholderText)           | 设置或返回一个 String 类型的值，该值代表为不包含文本的元素显示的文本。                                                |
|  19  | [PreviousSibling](#XMLNode.PreviousSibling)           | 返回一个 XMLNode 对象，该对象代表文档中与指定元素同级的前一个元素。                                                   |
|  20  | [Range](#XMLNode.Range)                               | 返回一个 Range 对象，该对象代表包含在指定对象中的文档部分。只读。                                                     |
|  21  | [Text](#XMLNode.Text)                                 | 返回或设置 XML 元素中包含的文本。 String 类型，可读写。                                                               |
|  22  | [ValidationErrorText](#XMLNode.ValidationErrorText)   | 返回一个 String 类型的值，代表关于 XMLNode 对象验证错误的说明。                                                       |
|  23  | [WordOpenXML](#XMLNode.WordOpenXML)                   | 返回一个 String 类型的值，该值以 WPS Open XML 格式表示节点的 XML。只读。                                              |
|  24  | [XML](#XMLNode.XML)                                   | 返回 String 值，该值表示包含在 XML 节点中的文本（有或没有 XML 标记）。只读。                                          |

## 成员方法

### XMLNode.Copy

将指定的 XML 元素（不包括 XML 标记）复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.Cut

将指定的 XML 元素从文档中移到剪贴板上。

**语法**

**express.Cut ()**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.Delete

删除 XML 文档中指定的 XML 元素。

**语法**

**express.Delete ()**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.RemoveChild

从指定的元素中删除子元素。

**语法**

**express.RemoveChild ( *ChildElement* )**

\*express是一个代表 XMLNode 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明             |
|----------------|-----------|----------|------------------|
| *ChildElement* | 必选      | XMLNode  | 要删除的子元素。 |

**示例**

``` JavaScript
/* 以下示例删除活动文档中第一个元素的第一个子元素。 */
Application.ActiveDocument.XMLNodes.Item(1).RemoveChild(Application.ActiveDocument.XMLNodes.Item(1).ChildNodes.Item(1))
```

### XMLNode.SelectNodes

返回一个 **XMLNodes** 集合，该集合代表与 XPath 参数匹配的所有子元素（它们的顺序与它们在指定的 XML 元素中出现的顺序相同）。

**语法**

**express.SelectNodes ( *XPath* , *PrefixMapping* , *FastSearchSkippingTextNodes* )**

\*express是一个代表 XMLNode 对象的变量。

**参数**

| 名称                          | 必选/可选 | 数据类型 | 说明                                                                                                                |
|-------------------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------|
| *XPath*                       | 必选      | String   | 指定有效的 XPath 字符串。                                                                                           |
| *PrefixMapping*               | 可选      | String   | 提供用于搜索的架构中的前缀。如果 XPath 参数使用名称来搜索元素，则可使用 PrefixMapping 参数。                        |
| *FastSearchSkippingTextNodes* | 可选      |          | 如果该参数值为 True，则搜索指定节点时跳过所有文本节点。如果该参数值为 False，则搜索时包含文本节点。默认值为 False。 |

### XMLNode.SelectSingleNode

返回一个 **XMLNode** 对象，该对象代表指定的 XML 元素中第一个与 XPath 参数匹配的子元素。

返回 XMLNode值

**语法**

**express.SelectSingleNode ( *XPath* , *PrefixMapping* , *FastSearchSkippingTextNodes* )**

\*express是一个代表 XMLNode 对象的变量。

**参数**

| 名称                          | 必选/可选 | 数据类型 | 说明                                                                                                                |
|-------------------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------|
| *XPath*                       | 必选      | String   | 指定有效的 XPath 字符串。                                                                                           |
| *PrefixMapping*               | 可选      | String   | 提供用于搜索的架构中的前缀。如果 XPath 参数使用名称来搜索元素，则可使用 PrefixMapping 参数。                        |
| *FastSearchSkippingTextNodes* | 可选      | Boolean  | 如果该参数值为 True，则搜索指定节点时跳过所有文本节点。如果该参数值为 False，则搜索时包含文本节点。默认值为 False。 |

### XMLNode.SetValidationError

更改向用户显示的指定节点的验证错误文本，并强制 WPS 将节点报告为无效。

**语法**

**express.SetValidationError ( *Status* , *ErrorText* , *ClearedAutomatically* )**

\*express是一个代表 XMLNode 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型              | 说明                                                                                                                                                                                 |
|------------------------|-----------|-----------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Status*               | 必选      | WdXMLValidationStatus | 指定设置验证状态错误文本 (wdXMLValidationStatusCustom)，还是清除验证状态错误文本 (wdXMLValidationStatusOK)。                                                                         |
| *ErrorText*            | 可选      | Variant               | 显示给用户的文本。如果 Status 参数设置为 wdXMLValidationStatusOK，则将该参数保留为空。                                                                                               |
| *ClearedAutomatically* | 可选      | Boolean               | 如果值为 True，则指定节点上出现下一验证事件时，将自动清除错误消息。如果值为 False，则需要运行 SetValidationError 方法（Status 参数为 wdXMLValidationStatusOK）来清除自定义错误文本。 |

**说明**

若要设置自定义错误文本，请使用 **wdXMLValidationStatusCustom** 常量。

**示例**

``` JavaScript
/*以下示例指定自定义验证错误文本。*/
function test() {
    let objNode = Application.ActiveDocument.XMLNodes.Item(1)
    objNode.SetValidationError(wdXMLValidationStatusCustom,"Error Text",true)
}
```

### XMLNode.Validate

针对附加到文档的 XML 架构验证单个 XML 元素。

**语法**

**express.Validate ()**

\*express是一个代表 XMLNode 对象的变量。

**说明**

使用具有 **ValidationStatus** 和 **ValidationErrorText** 属性的 **Validate** 方法，根据所应用的架构确定 XML 元素是否有效，并且确定向用户显示的错误文本内容。使用 **SetValidationError** 方法可以用自定义验证错误来替代架构冲突。

当您运行 **Validate** 方法时，WPS 会用具有验证错误的 XML 节点的集合填充 **Document** 对象的 **XMLSchemaViolations** 属性。

**示例**

``` JavaScript
/* 以下示例检查活动文档中的每个元素和属性，并显示一条消息，其中包含根据架构没有通过验证的元素和属性以及原因说明*/
function test() {
    let objNode = Application.ActiveDocument.XMLNodes
    let strValid

    for(let i=1; i <= objNode.Count; i++) {
        objNode.Item(i).Validate()
        if(objNode.Item(i).ValidationStatus != wdXMLValidationStatusOK) {
            strValid = strValid + objNode.Item(i).BaseName + '\t'  + objNode.Item(i).ValidationErrorText + '\r\n'
        }
    }    
    alert("The following elements do not validate against " + "the schema." + '\r\n' + '\r\n' + strValid + '\r\n' + "You should fix these elements before continuing.")
}
```

## 成员属性

### XMLNode.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 XMLNode 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XMLNode.Attributes

返回一个 **XMLNodes** 集合，该集合代表指定元素的属性。

**语法**

**express.Attributes**

\*express是一个代表 XMLNode 对象的变量。

**说明**

使用 **Attributes** 属性返回的 **XMLNodes** 集合中的所有 **XMLNode** 对象的 **NodeType** 属性值均为 **wdXMLNodeAttribute** 。

### XMLNode.BaseName

返回一个 **String** 类型的值，该值代表不带任何前缀的元素名称。

**语法**

**express.BaseName**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.ChildNodeSuggestions

返回一个 **XMLChildNodeSuggestions** 集合，代表指定元素的子元素列表。

**语法**

**express.ChildNodeSuggestions**

\*express是一个代表 XMLNode 对象的变量。

**说明**

**XMLChildNodeSuggestions** 集合中的每个 **XMLChildNodeSuggestion** 对象都是允许的、可能做为子元素的 XML 元素列表中的一项，该列表位于 **“XML 结构”** 任务窗格底部。

**示例**

``` JavaScript
/*下列示例循环遍历活动文档中选定的第一个元素的建议子元素，并在插入点插入所有允许的元素。*/
function test() {
    let objSuggestion
    let objNode = Application.Selection.XMLParentNode
    
    for(objSuggestion = 1; objSuggestion <= objNode.ChildNodeSuggestions.Count; objSuggestion++) {
        XMLChildNodeSuggestions.Item(objSuggestion).Insert()
        Application.Selection.MoveRight()
    }
}
```

### XMLNode.ChildNodes

返回一个 **XMLNodes** 集合，该集合代表指定元素的子元素。

**语法**

**express.ChildNodes**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XMLNode 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XMLNode.FirstChild

返回一个 **DiagramNode** 对象，该对象代表父节点的第一个子节点。只读。

**语法**

**express.FirstChild**

\*express是一个代表 XMLNode 对象的变量。

**说明**

使用 **LastChild** 属性访问最后一个子节点。使用 **Root** 属性访问图表中的父节点。

**示例**

``` JavaScript
/*本示例将一个组织图表添加至活动文档，添加三个节点，并将第一个和最后一个子节点赋给变量。*/
function test() {
    let shpDiagram
    let dgnRoot
    let dgnFirstChild
    let dgnLastChild
    let intCount

    //Add organizational chart diagram to the current document
    shpDiagram = Application.ActiveDocument.Shapes.AddDiagram(msoDiagramOrgChart, 10, 15, 400, 475)

    //Add the first node to the diagram
    dgnRoot = shpDiagram.DiagramNode.Children.AddNode()

    //Add three child nodes
    for(intCount = 1; intCount <= 3; intCount++) {
        dgnRoot.Children.AddNode()
    }

    //Assign the first and last child nodes to variables
    dgnFirstChild = dgnRoot.Children.FirstChild
    dgnLastChild = dgnRoot.Children.LastChild
}
```

### XMLNode.HasChildNodes

返回 **Boolean** 值，该值表示 XML 节点是否有子节点。只读。

**语法**

**express.HasChildNodes**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.LastChild

返回一个 **XMLNode** 对象，该对象代表 XML 元素的最后一个子节点。

**语法**

**express.LastChild**

\*express是一个代表 XMLNode 对象的变量。

**示例**

``` JavaScript
/*以下示例访问活动文档中第二个元素的最后一个子元素。*/
let objNode = Application.ActiveDocument.XMLNodes.Item(2).LastChild
```

### XMLNode.Level

返回一个 WdXMLNodeLevel 常量，该常量代表 XML 元素是段落的一部分、是整个段落、包含在表格单元格中还是包含表格行。只读。

**语法**

**express.Level**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.NamespaceURI

返回一个 **String** 类型的值，该值代表指定对象的架构命名空间的统一资源标识符 (URI)。只读。

**语法**

**express.NamespaceURI**

\*express是一个代表 XMLNode 对象的变量。

**说明**

| 注释                                                                            |
|---------------------------------------------------------------------------------|
| 如果创建的是用于 WPS 的 XML 架构，则强烈建议在架构中指定 targetNamespace 设置。 |

### XMLNode.NextSibling

返回一个 **XMLNode** 对象，该对象代表文档中与指定元素同级的下一个元素。

**语法**

**express.NextSibling**

\*express是一个代表 XMLNode 对象的变量。

**说明**

如果指定元素是 **XMLNodes** 集合中的最后一个元素，则此属性返回 **Nothing** 。

**示例**

``` JavaScript
/*以下示例返回活动文档中第二个元素的下一个同级元素。*/
let objNode = Application.ActiveDocument.XMLNodes.Item(2).NextSibling
```

### XMLNode.NodeType

返回一个 **WdXMLNodeType** 常量，该常量代表节点的类型。

**语法**

**express.NodeType**

\*express是一个代表 XMLNode 对象的变量。

**说明**

**XMLNode** 对象可以是 XML 元素或元素的属性。请使用 **NodeType** 属性确定要使用的节点的类型，以便您不会试图对该节点执行无效操作。例如，虽然 **Attributes** 属性出现在 **XMLNode** 对象的可用属性列表中，但它仅适用于元素节点。

### XMLNode.NodeValue

返回或设置一个 **String** 类型的值，该值代表 XML 元素的值。可读写。

**语法**

**express.NodeValue**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.OwnerDocument

返回一个代表指定 XML 元素的父文档的 **Document** 对象。

**语法**

**express.OwnerDocument**

\*express是一个代表 XMLNode 对象的变量。

**示例**

``` JavaScript
/*以下示例访问活动选定内容中第一个 XML 元素的父文档。*/
let objDoc = Application.Selection.XMLNodes.Item(1).OwnerDocument
```

### XMLNode.Parent

返回一个 **Object** 类型值，该值代表指定 **XMLNode** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.ParentNode

返回一个代表指定元素的父元素的 **XMLNode** 对象。

**语法**

**express.ParentNode**

\*express是一个代表 XMLNode 对象的变量。

**示例**

``` JavaScript
/*以下示例访问活动文档中选定文本的 XML 父节点。*/
let objNode = Application.Selection.XMLParentNode.ParentNode
```

### XMLNode.PlaceholderText

设置或返回一个 **String** 类型的值，该值代表为不包含文本的元素显示的文本。

**语法**

**express.PlaceholderText**

\*express是一个代表 XMLNode 对象的变量。

**说明**

仅当清除了 **“XML 结构”** 任务窗格中的 **“在文档中显示 XML 标记”** 复选框时，才在 WPS 中显示占位符文本。 **“在文档中显示 XML 标记”** 复选框对应于 **ShowXMLMarkup** 属性。

### XMLNode.PreviousSibling

返回一个 **XMLNode** 对象，该对象代表文档中与指定元素同级的前一个元素。

**语法**

**express.PreviousSibling**

\*express是一个代表 XMLNode 对象的变量。

**说明**

如果指定元素是 **XMLNodes** 集合中的第一个元素，则此属性返回 **Nothing** 。

**示例**

``` JavaScript
/*以下示例返回活动文档中第三个元素的前一个同级元素。*/
let objNode = Application.ActiveDocument.XMLNodes.Item(3).PreviousSibling
```

### XMLNode.Range

返回一个 **Range** 对象，该对象代表包含在指定对象中的文档部分。只读。

**语法**

**express.Range**

\*express是一个代表 XMLNode 对象的变量。

### XMLNode.Text

返回或设置 XML 元素中包含的文本。 **String** 类型，可读写。

**语法**

**express.Text**

\*express是一个代表 XMLNode 对象的变量。

**说明**

**Text** 属性返回 XML 元素中的无格式纯文本。设置该属性，可替换现有文本。

### XMLNode.ValidationErrorText

返回一个 **String** 类型的值，代表关于 **XMLNode** 对象验证错误的说明。

**语法**

**express.ValidationErrorText**

\*express是一个代表 XMLNode 对象的变量。

**说明**

**ValidationErrorText(Boolean *Advanced* ), *Advanced* 代表显示的错误文字是验证错误说明的高级版本，该版本来自 WPS 所包含的 MSXML 5.0 组件。**

**示例**

``` JavaScript
/*下列示例检查活动文档中的每个元素，并显示一条消息，其中包含未通过架构验证的元素和属性，并说明相关原因。*/
function test() {
    let objNode = Application.ActiveDocument.XMLNodes
    let strValid

    for(i = 1; i <= objNode.Count; i++) { 
        objNode.Item(i).Validate()
        if(objNode.Item(i).ValidationStatus != wdXMLValidationStatusOK) {
            strValid = strValid + objNode.Item(i).BaseName + "\t" + objNode.Item(i).ValidationErrorText + "\r\n"
        }
    }

    alert("The following elements do not validate against " + "the schema." + "\r\n" + "\r\n" + strValid + "\r\n" + "You should fix these elements before continuing.")
}
```

### XMLNode.WordOpenXML

返回一个 **String** 类型的值，该值以 WPS Open XML 格式表示节点的 XML。只读。

**语法**

**express.WordOpenXML**

\*express是一个代表 XMLNode 对象的变量。

**说明**

此属性只返回文档中为表示指定的 XML 节点所需的 XML。

### XMLNode.XML

返回 **String** 值，该值表示包含在 XML 节点中的文本（有或没有 XML 标记）。只读。

**语法**

**express.XML**

\*express是一个代表 XMLNode 对象的变量。

**说明**

参数

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                                                                                                |
|------------|---------------|--------------|-----------------------------------------------------------------------------------------------------------------------------------------|
| *DataOnly* | 可选          | **Boolean**  | 指定返回的 XML 是否带标记。如果为 **True** ，将返回 XML 节点中包含的文本，但不带 XML 标记。如果为 **False** ，将返回带 XML 标记的文本。 |

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XMLNode](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Design 对象

## 简介

代表单张幻灯片设计模板。 Design 对象是 Designs 和 SlideRange 集合以及 Master 和 Slide 对象的成员。

## 说明

使用 Master 、 Slide 或 SlideRange 对象的 Design 属性访问 Design 对象，例如：

``` JavaScript
Application.ActivePresentation.SlideMaster.Design
```

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Design
```

``` JavaScript
ActivePresentation.Slides.Range().Design
```

分别使用 Designs 集合的 Add 、 Item 、 Clone 或 Load 方法添加、引用、克隆或加载一个 Design 对象。例如，若要添加一个设计模板，请使用 ` ActivePresentation.Designs.Add designName:="MyDesign" `

## 方法

| 序号 | 名称                     | 说明                                                                             |
|:----:|:-------------------------|:---------------------------------------------------------------------------------|
|  1   | [Delete](#Design.Delete) | 删除指定的对象。                                                                 |
|  2   | [MoveTo](#Design.MoveTo) | 将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。 |

## 属性

| 序号 | 名称                               | 说明                                                                                     |
|:----:|:-----------------------------------|:-----------------------------------------------------------------------------------------|
|  1   | [Application](#Design.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                  |
|  2   | [Index](#Design.Index)             | 返回 Long 类型值，该值代表动画效果或设计的索引号。只读。                                 |
|  3   | [Name](#Design.Name)               | 返回或设置指定对象的名称。 String 类型，可读/写。                                        |
|  4   | [Parent](#Design.Parent)           | 返回指定对象的父对象。                                                                   |
|  5   | [Preserved](#Design.Preserved)     | 设置或返回一个 MsoTriState 类型的常量，该常量代表是否保护设计母版以避免被更改。可读/写。 |
|  6   | [SlideMaster](#Design.SlideMaster) | 返回一个 Master 对象，该对象代表幻灯片母版。                                             |

## 成员方法

### Design.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 Design 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

### Design.MoveTo

将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。

**语法**

**express.MoveTo ( *toPos* )**

\*express是一个代表 Design 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                     |
|---------|-----------|----------|--------------------------|
| *toPos* | 必选      | Long     | 动画效果要移动到的索引。 |

**示例**

``` JavaScript
/*本示例在指定形状的动画效果集合中将一个动画效果移动到第二个动画效果的位置上。*/
function MoveEffect() {
    let sldFirst = Application.ActivePresentation.Slides.Item(1)
    let shpFirst = sldFirst.Shapes.Item(1)
    let effAdd = sldFirst.TimeLine.MainSequence.AddEffect(shpFirst, msoAnimEffectBlinds)
    effAdd.MoveTo(2)
}
```

``` JavaScript
/*本示例将当前演示文稿中的第二张幻灯片移动到第一张幻灯片的位置上。*/
 Application.ActivePresentation.Slides.Item(2).MoveTo(1)
```

## 成员属性

### Design.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Design 对象的变量。

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
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### Design.Index

返回 **Long** 类型值，该值代表动画效果或设计的索引号。只读。

**语法**

**express.Index**

\*express是一个代表 Design 对象的变量。

**示例**

以下示例显示第一张幻灯片的主动画序列中所有效果的名称和索引号。

``` JavaScript
 alert("Effect Name: " + Application.ActivePresentation.Slides.Item(1).DisplayName + "\n" + "Effect Index: " + Application.ActivePresentation.Slides.Item(1).Index)
```

### Design.Name

返回或设置指定对象的名称。 **String** 类型，可读/写。

**语法**

**express.Name**

\*express是一个代表 Design 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则 ` .Shapes("Rectangle 2") ` 将返回对该形状的引用。

### Design.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Design 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

### Design.Preserved

设置或返回一个 **MsoTriState** 类型的常量，该常量代表是否保护设计母版以避免被更改。可读/写。

**语法**

**express.Preserved**

\*express是一个代表 Design 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue 不适用于此属性。                     |
| msoFalse 设计母版尚未保护并且可进行编辑。     |
| msoTriStateMixed 不适用于此属性。             |
| msoTriStateToggle 不适用于此属性。            |
| msoTrue 设计母版已保护且无法进行编辑。        |

**示例**

下面的代码行锁定并保护第一个设计母版。

``` JavaScript
 Application.ActivePresentation.Designs.Item(1).Preserved = msoTrue
```

### Design.SlideMaster

返回一个 **Master** 对象，该对象代表幻灯片母版。

**语法**

**express.SlideMaster**

\*express是一个代表 Design 对象的变量。

**示例**

``` JavaScript
/*以下示例设置当前演示文稿的幻灯片母版的背景图案。*/
function test(){
　　　　Application.ActivePresentation.SlideMaster.Background.Fill.PresetTextured(msoTextureGreenMarble)
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Design](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Designs 对象

## 简介

代表幻灯片设计模板的集合。

## 说明

使用 Presentation 对象的 Designs 属性引用一个设计模板。

若要添加或克隆单个的设计模板，请分别使用 Designs 集合的 Add 或 Clone 方法。若要引用单个的设计模板，请使用 Item 方法。

若要加载一个设计模板，请使用 Load 方法。

## 方法

| 序号 | 名称                    | 说明                                                                         |
|:----:|:------------------------|:-----------------------------------------------------------------------------|
|  1   | [Add](#Designs.Add)     | 返回一个代表新幻灯片设计的 Design 对象。                                     |
|  2   | [Clone](#Designs.Clone) | 创建 Design 对象的一个副本。                                                 |
|  3   | [Item](#Designs.Item)   | 从指定集合中返回单个对象。                                                   |
|  4   | [Load](#Designs.Load)   | 返回一个 Design 对象，该对象代表一个已加载到指定演示文稿的母版列表中的设计。 |

## 属性

| 序号 | 名称                                | 说明                                                    |
|:----:|:------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#Designs.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#Designs.Count)             | 返回指定集合中的对象数目。                              |
|  3   | [Parent](#Designs.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### Designs.Add

返回一个代表新幻灯片设计的 **Design** 对象。

**语法**

**express.Add ( *designName* , *Index* )**

\*express是一个代表 Designs 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                 |
|--------------|-----------|----------|------------------------------------------------------------------------------------------------------|
| *designName* | 必选      | String   | 设计的名称。                                                                                         |
| *Index*      | 必选      | Integer  | 设计的索引号。其默认值为 -1，表示如果省略 Index 参数，则新的幻灯片设计将添加到现有幻灯片设计的末尾。 |

**返回值**

Design

### Designs.Clone

创建 **Design** 对象的一个副本。

**语法**

**express.Clone ( *pOriginal* , *Index* )**

\*express是一个代表 Designs 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                               |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------|
| *pOriginal* | 必选      | Design   | Design 对象。原始设计。                                                                            |
| *Index*     | 必选      | Long     | 设计将要复制到的 Designs 集合中的索引位置。如果忽略 Index，复制的设计则添加到 Designs 集合的最后。 |

**返回值**

Design

**示例**

### Designs.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Designs 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

Design

### Designs.Load

返回一个 **Design** 对象，该对象代表一个已加载到指定演示文稿的母版列表中的设计。

**语法**

**express.Load ( *TemplateName* , *Index* )**

\*express是一个代表 Designs 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                              |
|----------------|-----------|----------|---------------------------------------------------------------------------------------------------|
| *TemplateName* | 必选      | String   | 设计模板的路径。                                                                                  |
| *Index*        | 必选      | Long     | 设计模板在设计模板集合中的索引号。默认值为 -1，表示将设计模板添加到演示文稿中设计模板列表的末尾。 |

## 成员属性

### Designs.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Designs 对象的变量。

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
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### Designs.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 Designs 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭第二个窗口。*/
function test(){
　　　　let dCount = Application.Windows
    for (let i = 2; i <= dCount.Count; i++) {
　　　　    dCount.Item(2).Close()
　　　　}
}
```

### Designs.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Designs 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Designs](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# GroupShapes 对象

## 简介

代表组合形状中的各个形状。每个形状都由一个 Shape 对象代表。对本对象使用 Item 方法，可以在不取消分组的情况下处理组合中的各个形状。

## 方法

| 序号 | 名称                        | 说明                       |
|:----:|:----------------------------|:---------------------------|
|  1   | [Item](#GroupShapes.Item)   | 从指定集合中返回单个对象。 |
|  2   | [Range](#GroupShapes.Range) | 返回一个 ShapeRange 对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                          |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#GroupShapes.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                       |
|  2   | [Count](#GroupShapes.Count)             | 返回指定集合中的对象数目。                                                                                                                                    |
|  3   | [Creator](#GroupShapes.Creator)         | 返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
|  4   | [Parent](#GroupShapes.Parent)           | 返回指定对象的父对象。                                                                                                                                        |

## 成员方法

### GroupShapes.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 GroupShapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

Shape

### GroupShapes.Range

返回一个 **ShapeRange** 对象。

**语法**

**express.Range ( *Index* )**

\*express是一个代表 GroupShapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                       |
|---------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要包含在范围中的各个形状。可以是指定形状的索引号的 Integer、指定形状名称的 String，或者是包含整数或字符串的数组。如果省略此参数，则 Range 方法将返回指定集合中的所有对象。 |

**返回值**

ShapeRange

**说明**

虽然可以使用 **Range** 方法返回任意数目的形状或幻灯片，但如果只需要返回该集合的一个成员，则使用 **Item** 方法会更为简单。例如， ` Shapes(1) ` 比 ` Shapes.Range(1) ` 简单， ` Slides(2) ` 比 ` Slides.Range(2) ` 简单。

若要为 **Index** 指定一个整数或字符串数组，可以使用 **Array** 函数。例如，以下指令返回用名称指定的两个形状。

` let myArray = ["Oval 4", "Rectangle 5"] let myRange = ActivePresentation.Slides.Item(1).Shapes.Range(myArray) `

## 成员属性

### GroupShapes.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 GroupShapes 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPre) {
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

### GroupShapes.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 GroupShapes 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除窗口1外的所有窗口。*/
function test(){
　　　　for(let i=2; i<=Application.Windows.Count; i++) {
 　　　　   Application.Windows.Item(i).Close()
　　　　}
}
```

### GroupShapes.Creator

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 GroupShapes 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

``` JavaScript
/*以下示例显示关于 myObject 创建者的消息。*/
function test(){
　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
　　　　if (myObject.Creator == "0x50575054") {
　　　　    alert("This is aWPP object")
　　　　}
　　　　else {
　　　　    alert("This is not aWPP object")
　　　　}
}
```

### GroupShapes.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 GroupShapes 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/GroupShapes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Options 对象

## 简介

代表 WPP 中的应用程序选项。

## 属性

| 序号 | 名称                                                | 说明                                                                                                      |
|:----:|:----------------------------------------------------|:----------------------------------------------------------------------------------------------------------|
|  1   | [DisplayPasteOptions](#Options.DisplayPasteOptions) | WPP 的 MsoTrue 属性值用于显示 “粘贴选项” 按钮，该按钮显示在新粘贴文本的正下方。可读/写 MsoTriState 类型。 |

## 成员属性

### Options.DisplayPasteOptions

WPP 的 **MsoTrue** 属性值用于显示 **“粘贴选项”** 按钮，该按钮显示在新粘贴文本的正下方。可读/写 **MsoTriState** 类型。

**语法**

**express.DisplayPasteOptions**

\*express是一个代表 Options 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue                                       |

**示例**

如果 **“粘贴选项”** 按钮为禁用状态，本示例将启用该选项。

``` JavaScript
function ShowPasteOptionsButton(){
let opt = Application.Options
    if(opt.DisplayPasteOptions == false){
        opt.DisplayPasteOptions = true
    }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Options](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Presentations 对象

## 简介

WPP中所有 Presentation 对象的集合。每个 Presentation 对象代表WPP 中当前打开的一个演示文稿。

## 说明

Presentations 集合不包含开放式加载宏，它是一种特殊类型的隐藏演示文稿。但是如果知道某单个开放式加载宏的文件名，就可以返回它。例如， Presentations("oscar.ppa") 将以 Presentation 对象的形式返回名为“Oscar.ppa”的开放式加载宏。但是，建议使用 AddIns 集合返回开放式加载宏。

示例：

使用 Presentations 属性返回 Presentation 集合。使用 Add 方法创建一个新演示文稿并将其添加到集合中。以下示例创建一个新演示文稿，在其中添加一张幻灯片，然后保存该演示文稿。

``` JavaScript
function test()
{
    let newPres = Application.Presentations.Add(true)
    newPres.Slides.Add(1, 1)
    newPres.SaveAs("Sample")
}
```

使用 Presentations ( *index* ) 返回单个 Presentation 对象，其中 *index* 是演示文稿的名称或索引号。以下示例打印第一个演示文稿。

``` JavaScript
Application.Presentations.Item(1).PrintOut()
```

使用 [Open](#Presentations.Open) 方法打开演示文稿并将其添加到 Presentations 集合中。以下示例以只读方式打开文件“sales.ppt”。

``` JavaScript
Application.Presentations.Open("sales.ppt",true)
```

## 方法

| 序号 | 名称                                      | 说明                                                                                                       |
|:----:|:------------------------------------------|:-----------------------------------------------------------------------------------------------------------|
|  1   | [Add](#Presentations.Add)                 | 创建一个演示文稿。返回一个代表新演示文稿的 Presentation 对象。                                             |
|  2   | [CanCheckOut](#Presentations.CanCheckOut) | 值为 True 时，WPP 可以将指定的演示文稿从服务器签出。 Boolean 类型，可读/写。                               |
|  3   | [CheckOut](#Presentations.CheckOut)       | 从服务器向本地计算机复制指定演示文稿以进行编辑。返回表示签出演示文稿的本地路径和文件名的 String 类型的值。 |
|  4   | [Item](#Presentations.Item)               | 从指定集合中返回单个对象。                                                                                 |
|  5   | [Open](#Presentations.Open)               | 打开指定的演示文稿。返回一个 Presentation 对象，该对象代表已打开的演示文稿。                               |

## 属性

| 序号 | 名称                                      | 说明                                                    |
|:----:|:------------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#Presentations.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#Presentations.Count)             | 返回指定集合中的对象数目。                              |
|  3   | [Parent](#Presentations.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### Presentations.Add

创建一个演示文稿。返回一个代表新演示文稿的 Presentation 对象。

**语法**

**express.Add ( *WithWindow* )**

\*express是一个代表 Presentations 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型           | 说明                           |
|--------------|-----------|--------------------|--------------------------------|
| *WithWindow* | 可选      | MsoTriState 类型。 | 是否在可视窗口中创建演示文稿。 |

**示例**

``` JavaScript
/*本示例创建一个演示文稿，并在其中添加一张幻灯片，然后保存该演示文稿。*/
function test()
{
    let add = Application.Presentations.Add()
    add.Slides.Add(1, ppLayoutTitle)
    add.SaveAs("Sample")
}
```

### Presentations.CanCheckOut

值为 **True** 时，WPP 可以将指定的演示文稿从服务器签出。 **Boolean** 类型，可读/写。

**语法**

**express.CanCheckOut ( *FileName* )**

\*express是一个代表 Presentations 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                         |
|------------|-----------|----------|------------------------------|
| *FileName* | 必选      | String   | 演示文稿的服务器路径和名称。 |

**说明**

若要利用WPP 中的协作功能，演示文稿必须保存在 SharePoint Portal Server 上。

**示例**

``` JavaScript
//本示例验证演示文稿是否可以签出且未被其他用户签出。如果演示文稿可以签出，则将演示文稿复制到本地计算机以进行编辑。
function CheckOutPresentation(strPresentation) {
    if (Application.Presentations.CanCheckOut(strPresentation) == true) {
        Application.Presentations.CheckOut(strPresentation)
    }
    else {
        alert("You are unable to check out this " + "presentation at this time.")
    }
}
```

### Presentations.CheckOut

从服务器向本地计算机复制指定演示文稿以进行编辑。返回表示签出演示文稿的本地路径和文件名的 **String** 类型的值。

**语法**

**express.CheckOut ( *FileName* )**

\*express是一个代表 Presentations 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                         |
|------------|-----------|----------|------------------------------|
| *FileName* | 必选      | String   | 演示文稿的服务器路径和名称。 |

**说明**

若要利用 WPP 中的协作功能，演示文稿必须保存在 SharePoint Portal Server 上。

**示例**

``` JavaScript
//本示例验证演示文稿是否可以签出且未被其他用户签出。如果演示文稿可以签出，则将演示文稿复制到本地计算机以进行编辑。
function CheckOutPresentation(strPresentation) {
    if (Application.Presentations.CanCheckOut(strPresentation) == true) {
        Application.Presentations.CheckOut(strPresentation)
    }
    else {
        alert("You are unable to check out this " + "presentation at this time.")
    }
}
```

### Presentations.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Presentations 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

### Presentations.Open

打开指定的演示文稿。返回一个 Presentation 对象，该对象代表已打开的演示文稿。

**语法**

**express.Open ( *FileName* , *ReadOnly* , *Untitled* , *WithWindow* , *OpenConflictDocument* )**

\*express是一个代表 Presentations 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型    | 说明                                           |
|------------------------|-----------|-------------|------------------------------------------------|
| *FileName*             | 必选      | String      | 要打开的文件名。                               |
| *ReadOnly*             | 可选      | MsoTriState | 指定是以可读/写状态还是以只读状态打开文件。    |
| *Untitled*             | 可选      | MsoTriState | 指定文件是否有标题。                           |
| *WithWindow*           | 可选      | MsoTriState | 指定文件是否可见。                             |
| *OpenConflictDocument* | 可选      | MsoTriState | 指定是否为带有脱机冲突的演示文稿打开冲突文件。 |

**说明**

安装适当的文件转换程序后，WPP 可打开具有下列 MS-DOS 文件扩展名的文件：.ch3、.cht、.doc、.htm、.html、.mcw、.pot、.ppa、.pps、.ppt、.pre、.rtf、.sh3、.shw、.txt、.wk1、.wk3、.wk4、.wpd、.wpf、.wps 和 .xls。

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse** 默认值。以可读/写状态打开文件。  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 以只读状态打开文件。              |

|                                                               |
|---------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。                 |
| **msoCTrue**                                                  |
| **msoFalse** 默认值。自动使用文件名作为被打开演示文稿的标题。 |
| **msoTriStateMixed**                                          |
| **msoTriStateToggle**                                         |
| **msoTrue** 打开一个没有标题的文件。等于创建一个文件副本。    |

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse** 隐藏被打开的演示文稿。           |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 默认值。在可视窗口中打开文件。    |

|                                                     |
|-----------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。       |
| **msoCTrue**                                        |
| **msoFalse** 默认值。打开服务器文件并忽略冲突文档。 |
| **msoTriStateMixed**                                |
| **msoTriStateToggle**                               |
| **msoTrue** 打开冲突文件并覆盖服务器文件。          |

**示例**

``` JavaScript
/*本示例以只读状态打开一个演示文稿。*/
Application.Presentations.Open("C:\\My Documents\\pres1.ppt", true)
```

## 成员属性

### Presentations.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Presentations 对象的变量。

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

### Presentations.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 Presentations 对象的变量。

**示例**

``` JavaScript
/*以下示例显示所有打开的演示文稿的名字。*/
function test()
{
    let pres = Application.Presentations
    for(let i=1; i<=pres.Count; i++) 
    {
        alert(pres.Item(i).Name)
    }
}
```

### Presentations.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Presentations 对象的变量。

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

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Presentations](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Shape 对象

## 简介

代表绘图层中的对象，例如自选图形、任意多边形、OLE 对象或图片。

## 说明

| 注释                                                                                                                                                                                                                                                                                                                                                                                   |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 共有三个代表形状的对象： Shapes 集合，代表文档中的所有形状； ShapeRange 集合，代表文档中指定的形状子集（例如， ShapeRange 对象可以代表文档中的第一个和第四个形状，或代表文档中的所有选定形状）； Shape 对象，代表文档中的单个形状。若要同时使用多个形状或使用选定范围中的多个形状，请使用 ShapeRange 集合。有关如何使用一个形状或同时使用多个形状的概述，请参阅 使用形状（图形对象）。 |

以下示例说明如何执行下列操作：

- 按名称或编号索引，返回幻灯片上现有的形状。
- 返回幻灯片上新建的形状。
- 返回选定范围中的形状。
- 返回幻灯片上的幻灯片标题和其他占位符。
- 返回与连接符的端点相连的形状。
- 返回演示文稿的默认形状。
- 返回新建的任意多边形。
- 返回组合内的单个形状。
- 返回新组合的一组形状。

## 示例

使用 Shapes ( *index* ) 返回一个代表幻灯片上形状的 Shape 对象，其中 *index* 是形状名称或索引号。

向 Shapes 集合添加新的图形时，将对该新添加的图形赋以默认的名称。若要为图形指定更有意义的名称，可使用 Name 属性。下例向 myDocument 添加矩形，将其命名为“Red Square”，然后设置该矩形的前景色和线型。

``` JavaScript
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let addshape = myDocument.Shapes.AddShape(msoShapeRectangle, 144, 144, 72, 72)
    addshape.Name = "Red Square"
    addshape.Fill.ForeColor.RGB = (255, 0, 0)
    addshape.Line.DashStyle = msoLineDashDot    
} 
```

若要在幻灯片中添加形状并返回一个代表新建形状的 Shape 对象，请使用 Shapes 集合的下列方法之一： AddConnector 、 AddLine 、 AddMediaObject 、 AddOLEObject 、 AddPicture 、 AddPlaceholder 、 AddShape 、 AddTable 、 AddTextbox 、 AddTextEffect 。

使用 Selection.ShapeRange ( *index* ) 返回一个代表选定范围中形状的 Shape 对象，其中 *index* 是形状名称或索引号。以下示例设置活动窗口选定范围中第一个形状的填充（假定选定范围中至少有一个形状）。

``` JavaScript
Application.ActiveWindow.Selection.ShapeRange.Item(1).Fill.ForeColor.RGB = (255, 0, 0)
```

使用 Shapes.Title 返回代表幻灯片标题的 Shape 对象。使用 Shapes.AddTitle 在无标题的幻灯片中添加标题并返回代表新建标题的 Shape 对象。使用 Shapes.Placeholders ( *index* ) 返回一个代表占位符的 Shape 对象，其中 *index* 是占位符的索引号。如果没有改变过幻灯片中形状的排列顺序，则以下三个语句是等价的（假设第一张幻灯片有标题）。

``` JavaScript
function test() 
{
    Application.ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.Font.Italic = true
    Application.ActivePresentation.Slides.Item(1).Shapes.Placeholders.Item(1).TextFrame.TextRange.Font.Italic = true
    Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Font.Italic = true
}
```

若要返回代表演示文稿默认形状的 Shape 对象，请使用 DefaultShape 属性。 若要返回一个 Shape 对象，该对象代表连接符所连接的形状之一，请使用 BeginConnectedShape 或 EndConnectedShape 属性。

使用 BuildFreeform 和 AddNodes 方法定义新任意多边形的几何形状，使用 ConvertToShape 方法创建任意多边形并返回代表该形状的 Shape 对象。

使用 GroupItems ( *index* ) 返回代表组合形状中的单个形状的 Shape 对象，其中 *index* 是形状的名称或组合中的索引号。

使用 Group 或 Regroup 方法可对某一范围的形状进行组合，并返回代表新组合的单个 Shape 对象。形成一个组合后，可以像使用任何其他形状一样使用该组合。

## 方法

| 序号 | 名称                                          | 说明                                                                                                                                                             |
|:----:|:----------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Apply](#Shape.Apply)                         | 适用于指定的形状格式，该格式已使用 PickUp 方法复制。                                                                                                             |
|  2   | [Copy](#Shape.Copy)                           | 将指定对象复制到剪贴板。                                                                                                                                         |
|  3   | [Cut](#Shape.Cut)                             | 删除指定对象并将其放到剪贴板中。                                                                                                                                 |
|  4   | [Delete](#Shape.Delete)                       | 删除指定的对象。                                                                                                                                                 |
|  5   | [Flip](#Shape.Flip)                           | 绕水平或垂直轴翻转指定形状。                                                                                                                                     |
|  6   | [IncrementLeft](#Shape.IncrementLeft)         | 以指定点数水平移动指定形状。                                                                                                                                     |
|  7   | [IncrementRotation](#Shape.IncrementRotation) | 按指定度数改变指定形状绕 z 轴的旋转量。使用 Rotation 属性设置形状的绝对旋转量。                                                                                  |
|  8   | [IncrementTop](#Shape.IncrementTop)           | 以指定点数垂直移动指定形状。                                                                                                                                     |
|  9   | [PickUp](#Shape.PickUp)                       | 复制指定形状的格式。用 Apply 方法可将复制的格式应用于其他形状。                                                                                                  |
|  10  | [ScaleHeight](#Shape.ScaleHeight)             | 以指定的比例缩放图形高度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于除图片和 OLE 对象以外的其他图形来说，缩放总是相对于当前高度而言。 |
|  11  | [ScaleWidth](#Shape.ScaleWidth)               | 根据指定的系数缩放图形宽度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于图片和 OLE 对象以外的其他图形来说，总是相对于当前宽度进行缩放。 |
|  12  | [Select](#Shape.Select)                       | 选择指定的对象。                                                                                                                                                 |
|  13  | [ZOrder](#Shape.ZOrder)                       | 将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-顺序中的位置）。                                                                                    |

## 属性

| 序号 | 名称                                              | 说明                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
|:----:|:--------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AnimationSettings](#Shape.AnimationSettings)     | 返回一个 AnimationSettings 对象，该对象代表所有可应用于指定形状的动画的特殊效果。                                                                                                                                                                                                                                                                                                                                                                      |
|  2   | [Application](#Shape.Application)                 | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                                                                                                                                                                                                |
|  3   | [AutoShapeType](#Shape.AutoShapeType)             | 返回或设置指定的 Shape 或 ShapeRange 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。 MsoAutoShapeType 类型，可读写。                                                                                                                                                                                                                                                                                                  |
|  4   | [BackgroundStyle](#Shape.BackgroundStyle)         | 设置或返回指定对象的背景样式。可读/写。                                                                                                                                                                                                                                                                                                                                                                                                                |
|  5   | [Chart](#Shape.Chart)                             | 返回当前 Shape 对象的 Chart 对象。只读。                                                                                                                                                                                                                                                                                                                                                                                                               |
|  6   | [Child](#Shape.Child)                             | 如果该形状是子形状，或者如果形状区域内的所有形状都是同一个父形状的子形状，则属性值为 MsoTrue 。只读。 MsoTriState 类型。                                                                                                                                                                                                                                                                                                                               |
|  7   | [ConnectionSiteCount](#Shape.ConnectionSiteCount) | 返回指定形状中的连结点的数量。只读。 Number 类型。                                                                                                                                                                                                                                                                                                                                                                                                     |
|  8   | [Fill](#Shape.Fill)                               | 返回一个 FillFormat 对象，该对象包含指定形状的填充格式属性。只读。                                                                                                                                                                                                                                                                                                                                                                                     |
|  9   | [Glow](#Shape.Glow)                               | 返回指定形状的发光格式。只读 msoGlowType 类型。                                                                                                                                                                                                                                                                                                                                                                                                        |
|  10  | [GroupItems](#Shape.GroupItems)                   | 返回一个 GroupShapes 对象，该对象代表指定形状组中的单个形状。使用 GroupShapes 对象的 Item 方法可返回形状组中的单个形状。适用于代表组合形状的 Shape 或 ShapeRange 对象。只读。                                                                                                                                                                                                                                                                          |
|  11  | [HasChart](#Shape.HasChart)                       | 返回指定对象代表的形状是否包含图表。只读。                                                                                                                                                                                                                                                                                                                                                                                                             |
|  12  | [HasTable](#Shape.HasTable)                       | 返回指定的形状是否为表。只读。 MsoTriState 类型。                                                                                                                                                                                                                                                                                                                                                                                                      |
|  13  | [HasTextFrame](#Shape.HasTextFrame)               | 返回指定形状是否有文本框。只读。 MsoTriState 类型。                                                                                                                                                                                                                                                                                                                                                                                                    |
|  14  | [Height](#Shape.Height)                           | 以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读， Number 类型；用于所有其他对象时可读/写， Number 类型。                                                                                                                                                                                                                                                                                                                                    |
|  15  | [Id](#Shape.Id)                                   | 返回一个 Number 类型值，该值标识形状或形状范围。只读。                                                                                                                                                                                                                                                                                                                                                                                                 |
|  16  | [Left](#Shape.Left)                               | 返回或设置一个 Single 类型值，该值代表从形状边界框的左边缘到幻灯片左边缘的距离（以磅为单位）。可读/写。                                                                                                                                                                                                                                                                                                                                                |
|  17  | [MediaType](#Shape.MediaType)                     | 返回 OLE 媒体类型。只读。 PpMediaType 类型。                                                                                                                                                                                                                                                                                                                                                                                                           |
|  18  | [Name](#Shape.Name)                               | 创建形状时，WPP 自动以 ShapeType Number 的形式为其分配一个名称，其中 ShapeType 指明形状或自选图形的类型，Number 是在幻灯片上形状的集合中具有唯一性的整数。例如，为幻灯片上的形状自动生成的名称可以为 Placeholder 1、Oval 2 和 Rectangle 3。为了避免与自动分配的名称发生冲突，对用户定义的名称请不要使用 ShapeType Number 形式，其中 ShapeType 为一个用于自动生成名称的值，Number 为任意的正整数。形状范围必须至少包含一个形状。 String 类型，可读/写。 |
|  19  | [PictureFormat](#Shape.PictureFormat)             | 返回一个 PictureFormat 对象，该对象包含指定形状的图片格式属性。适用于代表图片或 OLE 对象的 Shape 或 ShapeRange 对象。只读。                                                                                                                                                                                                                                                                                                                            |
|  20  | [PlaceholderFormat](#Shape.PlaceholderFormat)     | 返回一个 PlaceholderFormat 对象，该对象包含占位符特有的属性。只读。                                                                                                                                                                                                                                                                                                                                                                                    |
|  21  | [Shadow](#Shape.Shadow)                           | 返回一个只读的 ShadowFormat 对象，该对象包含指定形状的阴影格式属性。                                                                                                                                                                                                                                                                                                                                                                                   |
|  22  | [ShapeStyle](#Shape.ShapeStyle)                   | 设置或返回指定对象的形状样式索引。可读/写。                                                                                                                                                                                                                                                                                                                                                                                                            |
|  23  | [Table](#Shape.Table)                             | 返回一个 Table 对象，该对象代表某个形状或形状区域内的一个表格。只读。                                                                                                                                                                                                                                                                                                                                                                                  |
|  24  | [TextEffect](#Shape.TextEffect)                   | 返回一个 TextEffectFormat 对象，该对象包含指定形状的文本效果格式属性。适用于代表艺术字的 Shape 或 ShapeRange 对象。                                                                                                                                                                                                                                                                                                                                    |
|  25  | [TextFrame](#Shape.TextFrame)                     | 返回一个 TextFrame 对象，该对象包含指定形状或母版文本样式的对齐方式和定位属性。                                                                                                                                                                                                                                                                                                                                                                        |
|  26  | [TextFrame2](#Shape.TextFrame2)                   | 返回与指定的 Shape 对象（包含指定形状的对齐方式和定位属性）相关联的 TextFrame2 对象。只读。                                                                                                                                                                                                                                                                                                                                                            |
|  27  | [ThreeD](#Shape.ThreeD)                           | 返回一个 ThreeDFormat 对象，该对象包含指定形状的三维效果格式属性。只读。                                                                                                                                                                                                                                                                                                                                                                               |
|  28  | [Top](#Shape.Top)                                 | 返回或设置一个 Number 类型值，该值代表形状边界框的上边缘到文档上边缘的距离。可读/写。                                                                                                                                                                                                                                                                                                                                                                  |
|  29  | [Type](#Shape.Type)                               | 返回一个 MsoShapeType 常量，该常量代表单个形状或形状范围内多个形状的类型。只读。                                                                                                                                                                                                                                                                                                                                                                       |
|  30  | [Visible](#Shape.Visible)                         | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 MsoTriState 类型。                                                                                                                                                                                                                                                                                                                                                                         |
|  31  | [Width](#Shape.Width)                             | 以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读， Number 类型；用于所有其他对象时可读/写， Number 类型。                                                                                                                                                                                                                                                                                                                                    |
|  32  | [ZOrderPosition](#Shape.ZOrderPosition)           | 返回指定的形状在 z-顺序中的位置。只读。 Number 类型。                                                                                                                                                                                                                                                                                                                                                                                                  |

## 成员方法

### Shape.Apply

适用于指定的形状格式，该格式已使用 **PickUp** 方法复制。

**语法**

**express.Apply ()**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例复制 myDocument 上第一个形状的格式，然后将该格式应用于第二个形状。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.Item(1).PickUp()
    myDocument.Shapes.Item(2).Apply()
}
```

### Shape.Copy

将指定对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 Shape 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

``` JavaScript
/*本示例在第一个窗口中拷贝选定幻灯片到剪贴板然后将其粘贴到第二窗口的视图中。*/
/*如果剪贴板的内容不能粘贴到第二个窗口的视图中（例如，如果试图将一个形状粘贴到幻灯片浏览视图中），则本示例失败。*/
function test() 
{
    Application.Windows.Item(1).Selection.Copy()
    Application.Windows.Item(2).View.Paste()
}

/*本示例将当前演示文稿中第一张幻灯片的第一个和第二个形状复制到剪贴板，然后将该副本粘贴到第二张灯片。*/
function test() 
{
    let active = Application.ActivePresentation
    active.Slides.Item(1).Shapes.Range([1, 2]).Copy()
    active.Slides.Item(2).Shapes.Paste()
}

/*本示例将当前演示文稿的第一张幻灯片复制到剪贴板。*/
Application.ActivePresentation.Slides.Item(1).Copy()

/*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()
```

### Shape.Cut

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut ()**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中。*/
Application.Windows.Item(1).Selection.Cut()

/*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中。*/
Application.ActivePresentation.Slides.Item(1).Cut()

/*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut()
```

### Shape.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 Shape 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**示例**

``` JavaScript
/*本示例从活动演示文稿的第一张幻灯片中删除所有任意多边形。*/
function test() 
{
    let shapes = Application.ActivePresentation.Slides.Item(1).Shapes
    for(let i = shapes.Count; i >= 1; i--)
    {
        if(shapes.Item(i).Type == msoFreeform)
        {
            shapes.Item(i).Delete()
        }
    }
}
```

### Shape.Flip

绕水平或垂直轴翻转指定形状。

**语法**

**express.Flip ( *FlipCmd* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型   | 说明                             |
|-----------|-----------|------------|----------------------------------|
| *FlipCmd* | 必选      | MsoFlipCmd | 指定形状是水平翻转还是垂直翻转。 |

**说明**

|                                             |
|---------------------------------------------|
| MsoFlipCmd 可以是下列 MsoFlipCmd 常量之一。 |
| **msoFlipHorizontal**                       |
| **msoFlipVertical**                         |

### Shape.IncrementLeft

以指定点数水平移动指定形状。

**语法**

**express.IncrementLeft ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                               |
|-------------|-----------|----------|------------------------------------------------------------------------------------|
| *Increment* | 必选      | Number   | 以点为单位指定形状水平移动的距离。正值表示将形状向右移动，负值表示将形状向左移动。 |

**示例**

``` JavaScript
/*本示例将 myDocument 上的第一个形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let duplicate = myDocument.Shapes(1)
    duplicate.Fill.PresetTextured(msoTextureGranite)
    duplicate.IncrementLeft(70)
    duplicate.IncrementTop(-50)
    duplicate.IncrementRotation(30)
}
```

### Shape.IncrementRotation

按指定度数改变指定形状绕 z 轴的旋转量。使用 **Rotation** 属性设置形状的绝对旋转量。

**语法**

**express.IncrementRotation ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                             |
|-------------|-----------|----------|--------------------------------------------------------------------------------------------------|
| *Increment* | 必选      | Number   | 以度为单位指定形状的水平旋转量。正值表示使形状按顺时针方向旋转；负值表示使形状按逆时针方向旋转。 |

**说明**

要将三维形状绕 X 轴或 Y 轴旋转，请分别使用 **IncrementRotationX** 方法或 **IncrementRotationY** 方法。

**示例**

``` JavaScript
/*本示例将 myDocument 上的第一个形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let duplicate = myDocument.Shapes(1)
    duplicate.Fill.PresetTextured(msoTextureGranite)
    duplicate.IncrementLeft(70)
    duplicate.IncrementTop(-50)
    duplicate.IncrementRotation(30)
}
```

### Shape.IncrementTop

以指定点数垂直移动指定形状。

**语法**

**express.IncrementTop ( *Increment* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                 |
|-------------|-----------|----------|--------------------------------------------------------------------------------------|
| *Increment* | 必选      | Number   | 以点为单位指定形状对象的垂直移动距离。正值表示将形状向下移动，负值表示将形状向上移动 |

**示例**

``` JavaScript
/*本示例将 myDocument 上的第一个形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let duplicate = myDocument.Shapes(1)
    duplicate.Fill.PresetTextured(msoTextureGranite)
    duplicate.IncrementLeft(70)
    duplicate.IncrementTop(-50)
    duplicate.IncrementRotation(30)
}
```

### Shape.PickUp

复制指定形状的格式。用 **Apply** 方法可将复制的格式应用于其他形状。

**语法**

**express.PickUp ()**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例复制 myDocument 上第一个形状的格式，然后将该格式应用于第二个形状。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.Item(1).PickUp()
    myDocument.Shapes.Item(2).Apply()
}
```

### Shape.ScaleHeight

以指定的比例缩放图形高度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于除图片和 OLE 对象以外的其他图形来说，缩放总是相对于当前高度而言。

**语法**

**express.ScaleHeight ( *Factor* , *RelativeToOriginalSize* , *fScale* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型     | 说明                                                                                                   |
|--------------------------|-----------|--------------|--------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Number       | 指定形状调整后的高度与当前或原始高度的比例。例如，若要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选      | MsoTriState  | 指定是否相对于形状的当前或原始尺寸来缩放形状。                                                         |
| *fScale*                 | 可选      | MsoScaleFrom | 在缩放形状时，形状中位置不变的部分。                                                                   |

**说明**

|                                                                                                      |
|------------------------------------------------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。                                                        |
| **msoCTrue**                                                                                         |
| **msoFalse** ：相对于形状的当前尺寸缩放。                                                            |
| **msoTriStateMixed**                                                                                 |
| **msoTriStateToggle**                                                                                |
| **msoTrue** 相对于原始大小缩放该形状。仅当指定形状是图片或 OLE 对象时，才为此参数指定 **msoTrue** 。 |

|                                                 |
|-------------------------------------------------|
| MsoScaleFrom 可以是下列 MsoScaleFrom 常量之一。 |
| **msoScaleFromBottomRight**                     |
| **msoScaleFromMiddle**                          |
| **msoScaleFromTopLeft** ：默认值。              |

**示例**

``` JavaScript
/*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let s = myDocument.Shapes
    for (let i = 1; i <= s.Count; i++)
    {
        switch(s.Item(i).Type)
        {
            case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture:
                s.Item(i).ScaleHeight (1.75, msoTrue)
                s.Item(i).ScaleWidth (1.75, msoTrue)
                break
            default:
                s.Item(i).ScaleHeight (1.75, msoFalse)
                s.Item(i).ScaleWidth (1.75, msoFalse)
        }
    }
}
```

### Shape.ScaleWidth

根据指定的系数缩放图形宽度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于图片和 OLE 对象以外的其他图形来说，总是相对于当前宽度进行缩放。

**语法**

**express.ScaleWidth ( *Factor* , *RelativeToOriginalSize* , *fScale* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型     | 说明                                                                                                   |
|--------------------------|-----------|--------------|--------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Number       | 指定形状调整后的宽度与当前或原始宽度的比例。例如，若要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选      | MsoTriState  | 指定是否相对于形状的当前或原始尺寸来缩放形状。                                                         |
| *fScale*                 | 可选      | MsoScaleFrom | 在缩放形状时，形状中位置不变的部分。                                                                   |

**说明**

|                                                                                                      |
|------------------------------------------------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。                                                        |
| **msoCTrue**                                                                                         |
| **msoFalse** ：相对于形状的当前尺寸缩放。                                                            |
| **msoTriStateMixed**                                                                                 |
| **msoTriStateToggle**                                                                                |
| **msoTrue** 相对于原始大小缩放该形状。仅当指定形状是图片或 OLE 对象时，才为此参数指定 **msoTrue** 。 |

|                                                 |
|-------------------------------------------------|
| MsoScaleFrom 可以是下列 MsoScaleFrom 常量之一。 |
| **msoScaleFromBottomRight**                     |
| **msoScaleFromMiddle**                          |
| **msoScaleFromTopLeft** ：默认值。              |

**示例**

``` JavaScript
/*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let s = myDocument.Shapes
    for (let i = 1; i <= s.Count; i++)
    {
        switch(s.Item(i).Type)
        {
            case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture:
                s.Item(i).ScaleHeight (1.75, msoTrue)
                s.Item(i).ScaleWidth (1.75, msoTrue)
                break
            default:
                s.Item(i).ScaleHeight (1.75, msoFalse)
                s.Item(i).ScaleWidth (1.75, msoFalse)
        }
    }
}
```

### Shape.Select

选择指定的对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型    | 说明                                       |
|-----------|-----------|-------------|--------------------------------------------|
| *Replace* | 可选      | MsoTriState | 指定该选择是否替换之前所做的任何一项选择。 |

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片。

|                                                        |
|--------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。          |
| **msoCTrue**                                           |
| **msoFalse** 将该选择添加到前一项选择中。              |
| **msoTriStateMixed**                                   |
| **msoTriStateToggle**                                  |
| **msoTrue** 默认值。该选择替换之前所做的任何一项选择。 |

**示例**

``` JavaScript
/*本示例选择活动演示文稿第一张幻灯片上的第一个和第三个形状。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Range([1, 3]).Select()
/*本示例将活动演示文稿第一张幻灯片上的第二个和第四个形状添加到前一项选择中。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Range([2, 4]).Select(false)
```

### Shape.ZOrder

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-顺序中的位置）。

**语法**

**express.ZOrder ( *ZOrderCmd* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型     | 说明                                     |
|-------------|-----------|--------------|------------------------------------------|
| *ZOrderCmd* | 必选      | MsoZOrderCmd | 指定针对其他形状将指定形状移动到的位置。 |

**说明**

|                                                 |
|-------------------------------------------------|
| MsoZOrderCmd 可以是下列 MsoZOrderCmd 常量之一。 |
| **msoBringForward**                             |
| **msoBringInFrontOfText** 仅用于 WPS。          |
| **msoBringToFront**                             |
| **msoSendBackward**                             |
| **msoSendBehindText** 仅用于 WPS。              |
| **msoSendToBack**                               |

可用 **ZOrderPosition** 属性确定形状在 z-顺序中的当前位置。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加椭圆，如果幻灯片中有至少一个其他形状，则将该椭圆放置在 z-顺序中倒数第二的位置上。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myaddshape = myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)
    while(myaddshape.ZOrderPosition > 2)
    {
        myaddshape.ZOrder(msoSendBackward)
    }
}
```

## 成员属性

### Shape.AnimationSettings

返回一个 **AnimationSettings** 对象，该对象代表所有可应用于指定形状的动画的特殊效果。

**语法**

**express.AnimationSettings**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例设置在创建幻灯片时，当前演示文稿第二张幻灯片的第一个形状从左侧飞入。*/
function test() 
{
    let asets = Application.ActivePresentation.Slides.Item(2).Shapes.Item(1).AnimationSettings
    asets.EntryEffect = ppEffectFlyFromLeft
    asets.TextLevelEffect = ppAnimateByAllLevels
}
```

### Shape.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres)
{
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test() 
{
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for(let i = 1; i <= shpOle.Count; i++)
    {
        if(shpOle.Item(i).Type == msoLinkedOLEObject)
        {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### Shape.AutoShapeType

返回或设置指定的 **Shape** 或 **ShapeRange** 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。 **MsoAutoShapeType** 类型，可读写。

**语法**

**express.AutoShapeType**

\*express是一个代表 Shape 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

使用 **ConnectorFormat** 对象的 **Type** 属性可设置或返回连接符类型。

**示例**

``` JavaScript
/*本示例在 myDocument 中用 32 角星替换所有 16 角星。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let s = myDocument.Shapes
    for(let i = 1; i <= s.Count; i++)
    {
        if(s.Item(i).AutoShapeType == msoShape16pointStar)
        {
            s.Item(i).AutoShapeType = msoShape32pointStar
        }
    }
}
```

### Shape.BackgroundStyle

设置或返回指定对象的背景样式。可读/写。

**语法**

**express.BackgroundStyle**

\*express是一个代表 Shape 对象的变量。

### Shape.Chart

返回当前 Shape 对象的 Chart 对象。只读。

**语法**

**express.Chart**

\*express是一个代表 Shape 对象的变量。

### Shape.Child

如果该形状是子形状，或者如果形状区域内的所有形状都是同一个父形状的子形状，则属性值为 **MsoTrue** 。只读。 **MsoTriState** 类型。

**语法**

**express.Child**

\*express是一个代表 Shape 对象的变量。

**说明**

|                                                                                     |
|-------------------------------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。                                       |
| **msoCTrue** 不应用于此属性。                                                       |
| **msoFalse** ：该形状不是子形状，或者形状区域内并非所有的子形状都属于同一个父形状。 |
| **msoTriStateMixed** 不应用于该属性。                                               |
| **msoTriStateToggle** 不应用于此属性。                                              |
| **msoTrue** ：该形状是子形状，或者形状区域内所有的子形状都属于同一个父形状。        |

**示例**

``` JavaScript
/*以下示例选择了画布中的第一个形状，如果选定的形状是子形状，则用指定的颜色填充该形状。以下示例假设当前演示文稿中的第一个形状是包含多个形状的绘图画布。*/
function test()
{
    //Select the first shape in the drawing canvas
    Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).CanvasItems.Item(1).Select()
    //Fill selected shape if it is a child shape
    let actsel = Application.ActiveWindow.Selection
    if(actsel.ShapeRange.Child == true)
    {
        actsel.ShapeRange.Fill.ForeColor.RGB = 100, 0, 200
    }
    else
    {
        alert("This shape is not a child shape.")
    }
}
```

### Shape.ConnectionSiteCount

返回指定形状中的连结点的数量。只读。 **Number** 类型。

**语法**

**express.ConnectionSiteCount**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加两个矩形，并且用两个连接符连接这两个矩形。两个连接符的起点都连接到第一个矩形的第一个连接位置，两个连接符的终点分别连接到第二个矩形的第一个和最后一个连接位置。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let s = myDocument.Shapes
    let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
    let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
    let lastsite = secondRect.ConnectionSiteCount
    let myaddcor = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat
    myaddcor.BeginConnect(firstRect, 1)
    myaddcor.EndConnect(secondRect, 1)
    myaddcor.BeginConnect(firstRect, 1)
    myaddcor.EndConnect(secondRect, lastsite)
}
```

### Shape.Fill

返回一个 **FillFormat** 对象，该对象包含指定形状的填充格式属性。只读。

**语法**

**express.Fill**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myshapes = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
    myshapes.ForeColor.RGB = 128, 0, 0
    myshapes.BackColor.RGB = 170, 170, 170
    myshapes.TwoColorGradient (msoGradientHorizontal, 1)
}
```

### Shape.Glow

返回指定形状的发光格式。只读 **msoGlowType** 类型。

**语法**

**express.Glow**

\*express是一个代表 Shape 对象的变量。

### Shape.GroupItems

返回一个 **GroupShapes** 对象，该对象代表指定形状组中的单个形状。使用 **GroupShapes** 对象的 **Item** 方法可返回形状组中的单个形状。适用于代表组合形状的 **Shape** 或 **ShapeRange** 对象。只读。

**语法**

**express.GroupItems**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加三个三角形，将它们组合，再设置整个组的颜色，然后只更改第二个三角形的颜色。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myshapes = myDocument.Shapes
    myshapes.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"
    myshapes.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"
    myshapes.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"
    let mygroup = myshapes.Range(["shpOne", "shpTwo", "shpThree"]).Group()
    mygroup.Fill.PresetTextured (msoTextureBlueTissuePaper)
    mygroup.GroupItems.Item(2).Fill.PresetTextured (msoTextureGreenMarble)
}
```

### Shape.HasChart

返回指定对象代表的形状是否包含图表。只读。

**语法**

**express.HasChart**

\*express是一个代表 Shape 对象的变量。

### Shape.HasTable

返回指定的形状是否为表。只读。 **MsoTriState** 类型。

**语法**

**express.HasTable**

\*express是一个代表 Shape 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** ：指定的形状是表。                |

**示例**

``` JavaScript
/*以下示例检查当前选定的形状是否为表格。如果是表格，此代码将第一列的宽度设置为 1 英寸（72 磅）。*/
function test() 
{
    let srange = Application.ActiveWindow.Selection.ShapeRange
    if(srange.HasTable == msoTrue)
    {
        srange.Table.Columns.Item(1).Width = 72
    }
}
```

### Shape.HasTextFrame

返回指定形状是否有文本框。只读。 **MsoTriState** 类型。

**语法**

**express.HasTextFrame**

\*express是一个代表 Shape 对象的变量。

**说明**

|                                                  |
|--------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。    |
| **msoCTrue**                                     |
| **msoFalse**                                     |
| **msoTriStateMixed**                             |
| **msoTriStateToggle**                            |
| **msoTrue** ：指定形状有文本框，因此可包含文本。 |

### Shape.Height

以磅为单位返回或设置指定对象的高度。用于 **Master** 对象时只读， **Number** 类型；用于所有其他对象时可读/写， **Number** 类型。

**语法**

**express.Height**

\*express是一个代表 Shape 对象的变量。

**说明**

一个 **Height** Shape 对象的 **Shape** Height 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

### Shape.Id

返回一个 **Number** 类型值，该值标识形状或形状范围。只读。

**语法**

**express.Id**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例向当前演示文稿添加一个新形状，再根据 ID 属性的值填充该形状。*/
function test()
{
    let myaddshape = Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShape5pointStar, 100, 100, 100, 100)
    switch (myaddshape.Id)
    {
        case myaddshape.Id >0 && myaddshape.Id <=500:
            myaddshape.Fill.ForeColor.RGB = 255, 0, 0
            break
        case myaddshape.Id >500 && myaddshape.Id <=1000:
            myaddshape.Fill.ForeColor.RGB = 255, 255, 0
            break
        case myaddshape.Id >1000 && myaddshape.Id <=1500:
            myaddshape.Fill.ForeColor.RGB = 255, 0, 255
            break
        case myaddshape.Id >1500 && myaddshape.Id <=2000:
            myaddshape.Fill.ForeColor.RGB = 0, 255, 0
            break
        case myaddshape.Id >2000 && myaddshape.Id <=2500:
            myaddshape.Fill.ForeColor.RGB = 0, 255, 255
            break
        default:
            myaddshape.Fill.ForeColor.RGB = 0, 0, 255
     }
}
```

### Shape.Left

返回或设置一个 **Single** 类型值，该值代表从形状边界框的左边缘到幻灯片左边缘的距离（以磅为单位）。可读/写。

**语法**

**express.Left**

\*express是一个代表 Shape 对象的变量。

### Shape.MediaType

返回 OLE 媒体类型。只读。 **PpMediaType** 类型。

**语法**

**express.MediaType**

\*express是一个代表 Shape 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| PpMediaType 可以是下列 PpMediaType 常量之一。 |
| **ppMediaTypeMixed**                          |
| **ppMediaTypeMovie**                          |
| **ppMediaTypeOther**                          |
| **ppMediaTypeSound**                          |

**示例**

``` JavaScript
/*以下示例设置在放映幻灯片时，当前演示文稿内第一张幻灯片上的所有原始声音对象循环播放直到被手动停止。*/
function test() 
{
    let so = Application.ActivePresentation.Slides.Item(1).Shapes
    for(let i = 1; i <= so.Count; i++)
    {
        if(so.Item(i).Type == msoMedia)
        {
            if(so.Item(i).MediaType == ppMediaTypeSound)
            {
                so.Item(i).AnimationSettings.PlaySettings.LoopUntilStopped = true
            }
        }
    }  
}
```

### Shape.Name

创建形状时，WPP 自动以 ShapeType Number 的形式为其分配一个名称，其中 ShapeType 指明形状或自选图形的类型，Number 是在幻灯片上形状的集合中具有唯一性的整数。例如，为幻灯片上的形状自动生成的名称可以为 Placeholder 1、Oval 2 和 Rectangle 3。为了避免与自动分配的名称发生冲突，对用户定义的名称请不要使用 ShapeType Number 形式，其中 ShapeType 为一个用于自动生成名称的值，Number 为任意的正整数。形状范围必须至少包含一个形状。 **String** 类型，可读/写。

**语法**

**express.Name**

\*express是一个代表 Shape 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。

**示例**

``` JavaScript
/*本示例将当前演示文稿第一张幻灯片上的第二个对象的名称设为“big triangle”。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).Name = "big triangle"

/*本示例为当前演示文稿第一张幻灯片中名为“big triangle”的形状设置填充颜色。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item("big triangle").Fill.ForeColor.RGB = 0, 0, 255
```

### Shape.PictureFormat

返回一个 **PictureFormat** 对象，该对象包含指定形状的图片格式属性。适用于代表图片或 OLE 对象的 **Shape** 或 **ShapeRange** 对象。只读。

**语法**

**express.PictureFormat**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例设置了 myDocument 中第一个形状的亮度和对比度。第一个形状必须是图片或 OLE 对象。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let mypfmat = myDocument.Shapes.Item(1).PictureFormat
    mypfmat.Brightness = 0.3
    mypfmat.Contrast = 0.75 
}
```

### Shape.PlaceholderFormat

返回一个 PlaceholderFormat 对象，该对象包含占位符特有的属性。只读。

**语法**

**express.PlaceholderFormat**

\*express是一个代表 Shape 对象的变量。

### Shape.Shadow

返回一个只读的 **ShadowFormat** 对象，该对象包含指定形状的阴影格式属性。

**语法**

**express.Shadow**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例在当前演示文稿第一张幻灯片中添加一个有阴影的矩形。蓝色浮凸效果的阴影距该矩形的右边 3 磅，下边 2 磅。*/
function test() 
{
    let myShap = Application.ActivePresentation.Slides.Item(1).Shapes
    let myshadow = myShap.AddShape(msoShapeRectangle, 10, 10, 150, 90).Shadow
    myshadow.Type = msoShadow17
    myshadow.ForeColor.RGB = (0, 0, 128)
    myshadow.OffsetX = 3
    myshadow.OffsetY = 2
}
```

### Shape.ShapeStyle

设置或返回指定对象的形状样式索引。可读/写。

**语法**

**express.ShapeStyle**

\*express是一个代表 Shape 对象的变量。

### Shape.Table

返回一个 **Table** 对象，该对象代表某个形状或形状区域内的一个表格。只读。

**语法**

**express.Table**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例将第二张幻灯片第五个形状中表格第一列的宽度设置为 80 磅。*/
Application.ActivePresentation.Slides.Item(2).Shapes.Item(5).Table.Columns.Item(1).Width = 80
```

### Shape.TextEffect

返回一个 **TextEffectFormat** 对象，该对象包含指定形状的文本效果格式属性。适用于代表艺术字的 **Shape** 或 **ShapeRange** 对象。

**语法**

**express.TextEffect**

\*express是一个代表 Shape 对象的变量。

### Shape.TextFrame

返回一个 **TextFrame** 对象，该对象包含指定形状或母版文本样式的对齐方式和定位属性。

**语法**

**express.TextFrame**

\*express是一个代表 Shape 对象的变量。

**说明**

使用 **TextFrame** 对象的 **TextRange** 属性可返回文本框中的文本。

在应用 **TextFrame** 属性之前，可使用 **HasTextFrame** 属性来确定形状中是否包含文本框。

**示例**

``` JavaScript
/*以下示例在 myDocument 中添加一个矩形，向该矩形中添加文本，然后设置文本框的上边距。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myshapes = myDocument.Shapes.AddShape(msoShapeRectangle, 180, 175, 350, 140).TextFrame
    myshapes.TextRange.Text = "Here is some test text"
    myshapes.MarginTop = 10
}
```

### Shape.TextFrame2

返回与指定的 Shape 对象（包含指定形状的对齐方式和定位属性）相关联的 TextFrame2 对象。只读。

**语法**

**express.TextFrame2**

\*express是一个代表 Shape 对象的变量。

### Shape.ThreeD

返回一个 **ThreeDFormat** 对象，该对象包含指定形状的三维效果格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*以下示例为应用于 myDocument 中第一个形状的三维效果设置深度、延伸色、延伸方向以及照明方向。​*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let mythreed = myDocument.Shapes.Item(1).ThreeD
    mythreed.Visible = true
    mythreed.Depth = 50
    //RGB value for purple
    mythreed.ExtrusionColor.RGB = (255, 100, 255)
    mythreed.SetExtrusionDirection (msoExtrusionTop)
    mythreed.PresetLightingDirection = msoLightingLeft
}
```

### Shape.Top

返回或设置一个 **Number** 类型值，该值代表形状边界框的上边缘到文档上边缘的距离。可读/写。

**语法**

**express.Top**

\*express是一个代表 Shape 对象的变量。

### Shape.Type

返回一个 **MsoShapeType** 常量，该常量代表单个形状或形状范围内多个形状的类型。只读。

**语法**

**express.Type**

\*express是一个代表 Shape 对象的变量。

**说明**

|                                                 |
|-------------------------------------------------|
| MsoShapeType 可以是下列 MsoShapeType 常量之一。 |
| **msoAutoShape**                                |
| **msoCallout**                                  |
| **msoCanvas**                                   |
| **msoChart**                                    |
| **msoComment**                                  |
| **msoDiagram**                                  |
| **msoEmbeddedOLEObject**                        |
| **msoFormControl**                              |
| **msoFreeform**                                 |
| **msoGroup**                                    |
| **msoLine**                                     |
| **msoLinkedOLEObject**                          |
| **msoLinkedPicture**                            |
| **msoMedia**                                    |
| **msoOLEControlObject**                         |
| **msoPicture**                                  |
| **msoPlaceholder**                              |
| **msoScriptAnchor**                             |
| **msoShapeTypeMixed**                           |
| **msoTable**                                    |
| **msoTextBox**                                  |
| **msoTextEffect**                               |

### Shape.Visible

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 **MsoTriState** 类型。

**语法**

**express.Visible**

\*express是一个代表 Shape 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。 msoCTrue msoFalse msoTriStateMixed msoTriStateToggle msoTrue：指定对象或对象格式是可见的。

### Shape.Width

以磅为单位返回或设置指定对象的宽度。用于 **Master** 对象时只读， **Number** 类型；用于所有其他对象时可读/写， **Number** 类型。

**语法**

**express.Width**

\*express是一个代表 Shape 对象的变量。

### Shape.ZOrderPosition

返回指定的形状在 z-顺序中的位置。只读。 **Number** 类型。

**语法**

**express.ZOrderPosition**

\*express是一个代表 Shape 对象的变量。

**说明**

Shapes(1) 返回 z-顺序中的最后一个形状，Shapes(Shapes.Count) 返回 z-顺序中的第一个形状。

若要设置形状在 z-顺序中的位置，请用 **ZOrder** 方法。

形状在 z-顺序中的位置与 **Shapes** 集合中形状的索引号相对应。例如，如果幻灯片上有四个形状，表达式 myDocument.Shapes(1) 在 z-顺序的后面返回该形状，表达式 myDocument.Shapes(4) 在 z-顺序的前面返回该形状。

如向一个集合添加新图形，其默认的添加位置是 z-顺序中的前面位置。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加椭圆，如果幻灯片中有至少一个其他形状，则将该椭圆放置在 z-顺序中倒数第二的位置上。*/
function test() 
{
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myaddshape = myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)
    while (myaddshape.ZOrderPosition > 2)
    {
        myaddshape.ZOrder (msoSendBackward)
    }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Shape](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SlideShowWindow 对象

## 简介

代表运行幻灯片放映的窗口。

## 方法

| 序号 | 名称                                  | 说明           |
|:----:|:--------------------------------------|:---------------|
|  1   | [Activate](#SlideShowWindow.Activate) | 激活指定的对象 |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                                                            |
|:----:|:----------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Active](#SlideShowWindow.Active)             | 返回指定的窗格或窗口是否处于激活的状态。只读。MsoTriState 类型。                                                                                                                                |
|  2   | [Application](#SlideShowWindow.Application)   | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                         |
|  3   | [Height](#SlideShowWindow.Height)             | 以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。                                                                               |
|  4   | [IsFullScreen](#SlideShowWindow.IsFullScreen) | 返回指定幻灯片放映窗口是否占用全屏。只读 MsoTriState 类型。                                                                                                                                     |
|  5   | [Left](#SlideShowWindow.Left)                 | 返回或设置一个 Single 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。 |
|  6   | [Presentation](#SlideShowWindow.Presentation) | 返回一个 Presentation 对象，该对象代表创建指定文档窗口或幻灯片放映窗口的演示文稿。只读。                                                                                                        |
|  7   | [Top](#SlideShowWindow.Top)                   | 返回或设置一个 Single 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。   |
|  8   | [View](#SlideShowWindow.View)                 | 返回一个 SlideShowView 对象。只读。                                                                                                                                                             |
|  9   | [Width](#SlideShowWindow.Width)               | 以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。                                                                               |

## 成员方法

### SlideShowWindow.Activate

激活指定的对象

**语法**

**express.Activate ()**

\*express是一个代表 SlideShowWindow 对象的变量。

**示例**

``` JavaScript
/*本示例按顺序激活紧挨在活动窗口之后的文档窗口*/
 Application.Windows.Item(1).Activate()
```

## 成员属性

### SlideShowWindow.Active

返回指定的窗格或窗口是否处于激活的状态。只读。MsoTriState 类型。

**语法**

**express.Active**

\*express是一个代表 SlideShowWindow 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。 msoCTrue msoFalse msoTriStateMixed msoTriStateToggle msoTrue：指定的窗格或窗口处于活动状态。

**示例**

``` JavaScript
/*以下示例检查演示文稿文件 "test.ppt" 是否位于当前窗口。如果不是，则将当前处于激活状态的演示文稿的名称保存在变量 oldWin 中，并激活 "test.ppt" 演示文稿*/
function test(){
    if(Application.Presentations.Item("test.ppt").Windows.Item(1).Active == false) {
        let oldWin = Application.ActiveWindow
        Application.Presentations.Item("test.ppt").Windows.Item(1).Activate()
    }
}
```

### SlideShowWindow.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 SlideShowWindow 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/
function test() {
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")  
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/
function test(){
   let shpOle = ActivePresentation.Slides.Item(1).Shapes
    for(let i = 1; i <= shpOle.Count; i++) {
        if(shpOle.Item(i).Type == msoLinkedOLEObject) {
            MsgBox(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### SlideShowWindow.Height

以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Height**

\*express是一个代表 SlideShowWindow 对象的变量。

**说明**

一个 HeightShape 对象的 ShapeHeight 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

**示例**

``` JavaScript
/*以下示例将第二个文档窗口的高度设置为应用程序窗口高度的一半*/
Application.Windows.Item(2).Height = Application.Windows.Item(1).Height/2
```

### SlideShowWindow.IsFullScreen

返回指定幻灯片放映窗口是否占用全屏。只读 MsoTriState 类型。

**语法**

**express.IsFullScreen**

\*express是一个代表 SlideShowWindow 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。 msoCTrue msoFalse msoTriStateMixed msoTriStateToggle msoTrue 指定的幻灯片放映窗口占用全屏。

**示例**

``` JavaScript
/*本示例将全屏的幻灯片放映窗口的高度适当缩小以便可看到任务栏*/
function test(){
    let SliWin = Application.ActivePresentation
    if(SliWin.IsFullScreen) {
        SliWin.Height = SliWin.Height - 20
    }
}
```

### SlideShowWindow.Left

返回或设置一个 Single 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Left**

\*express是一个代表 SlideShowWindow 对象的变量。

### SlideShowWindow.Presentation

返回一个 Presentation 对象，该对象代表创建指定文档窗口或幻灯片放映窗口的演示文稿。只读。

**语法**

**express.Presentation**

\*express是一个代表 SlideShowWindow 对象的变量。

**说明**

如果在第一个文档窗口中当前放映的幻灯片来自嵌入的演示文稿，则 Windows(1).View.Slide.Parent 返回该嵌入演示文稿，并且 Windows(1).Presentation 返回在其中创建第一个文档窗口的演示文稿。

如果在第一个幻灯片放映窗口中当前放映的幻灯片来自嵌入的演示文稿，则 SlideShowWindows(1).View.Slide.Parent 返回该嵌入演示文稿，并且 SlideShowWindows(1).Presentation 返回在其中创建第一个幻灯片放映窗口的演示文稿。

**示例**

``` JavaScript
/*以下示例使第一个窗口中的演示文稿幻灯片编号延续到第二个窗口中*/
function test(){
    let firstPresSlides =  Application.Windows.Item(1).Presentation.Slides.Count
     Application.Windows.Item(2).Presentation.PageSetup.FirstSlideNumber = firstPresSlides + 1
}
```

### SlideShowWindow.Top

返回或设置一个 Single 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Top**

\*express是一个代表 SlideShowWindow 对象的变量。

**示例**

``` JavaScript
/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的文档窗口*/
function test(){
    let sngHeight = Application.Windows.Item(1).Height                     // available height
    let sngWidth = Application.Windows.Item(1).Width + Windows.Item(2).Width    // available width

    let dWindows1 = Application.Windows.Item(1)
        dWindows1.Width = sngWidth
        dWindows1.Height = sngHeight / 2
        dWindows1.Left = 0

    let dWindows2 = Application.Windows.Item(2)
        dWindows2.Width = sngWidth
        dWindows2.Height = sngHeight / 2
        dWindows2.Top = sngHeight / 2
        dWindows2.Left = 0
}
```

### SlideShowWindow.View

返回一个 SlideShowView 对象。只读。

**语法**

**express.View**

\*express是一个代表 SlideShowWindow 对象的变量。

**示例**

``` JavaScript
/*本示例将当前窗口中的视图设为幻灯片视图，然后显示第三张幻灯片*/
function test(){
    Application.ActiveWindow.ViewType = ppViewSlide
    Application.ActiveWindow.View.GotoSlide(3)
}
```

### SlideShowWindow.Width

以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Width**

\*express是一个代表 SlideShowWindow 对象的变量。

**示例**

``` JavaScript
/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的文档窗口*/
function test(){
    let sngHeight = Application.Windows.Item(1).Height                     // available height
    let sngWidth = Application.Windows.Item(1).Width + Windows.Item(2).Width    // available width

    let dWindows1 = Application.Windows.Item(1)
        dWindows1.Width = sngWidth
        dWindows1.Height = sngHeight / 2
        dWindows1.Left = 0

    let dWindows2 = Windows.Item(2)
        dWindows2.Width = sngWidth
        dWindows2.Height = sngHeight / 2
        dWindows2.Top = sngHeight / 2
        dWindows2.Left = 0
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/SlideShowWindow](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomXMLNodes 对象

## 简介

包含 CustomXMLNodes 对象的集合，这些对象代表文档中的 XML 节点。

## 说明

Attributes 和 ChildNodes 属性返回此类型节点的集合。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

## 属性

| 序号 | 名称                                       | 说明                                                                       |
|:----:|:-------------------------------------------|:---------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLNodes.Application) | 获取一个代表 CustomXMLNodes 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [Count](#CustomXMLNodes.Count)             | 获取 CustomXMLNodes 集合中 CustomXMLNode 对象的数目。只读。                |
|  3   | [Creator](#CustomXMLNodes.Creator)         | 获取一个 32 位整数，指示创建 CustomXMLNodes 对象时所使用的应用程序。只读。 |
|  4   | [Item](#CustomXMLNodes.Item)               | 获取 CustomXMLNodes 集合中的一个 CustomXMLNode 对象。只读。                |
|  5   | [Parent](#CustomXMLNodes.Parent)           | 获取 CustomXMLNodes 对象的 Parent 对象。只读。                             |

## 成员属性

### CustomXMLNodes.Application

获取一个代表 **CustomXMLNodes** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLNodes 对象的变量。

### CustomXMLNodes.Count

获取 **CustomXMLNodes** 集合中 **CustomXMLNode** 对象的数目。只读。

**语法**

**express.Count**

\*express是一个代表 CustomXMLNodes 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNodes.Creator

获取一个 32 位整数，指示创建 **CustomXMLNodes** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLNodes 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNodes.Item

获取 **CustomXMLNodes** 集合中的一个 **CustomXMLNode** 对象。只读。

**语法**

**express.Item**

\*express是一个代表 CustomXMLNodes 对象的变量。

### CustomXMLNodes.Parent

获取 **CustomXMLNodes** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLNodes 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLNodes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomXMLPart 对象

## 简介

代表 CustomXMLParts 集合中的单个 CustomXMLPart 。

## 说明

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

以下示例将部件添加到 CustomXMLPart 对象。

``` JavaScript
function AddPartToCollection() {
let myPart = Application.ActiveDocument.CustomXMLParts.Add("Mark Twain")                 
}
```

## 方法

| 序号 | 名称                                                | 说明                                                                             |
|:----:|:----------------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [AddNode](#CustomXMLPart.AddNode)                   | 将节点添加到 XML 树。                                                            |
|  2   | [Delete](#CustomXMLPart.Delete)                     | 从数据存储（ IXMLDataStore 接口）中删除当前 CustomXMLPart 。                     |
|  3   | [Load](#CustomXMLPart.Load)                         | 使模板作者可以通过现有文件填充 CustomXMLPart 。如果加载成功，则返回 True 。      |
|  4   | [LoadXML](#CustomXMLPart.LoadXML)                   | 允许模板作者依据 XML 字符串填充 CustomXMLPart 对象。如果加载成功，则返回 True 。 |
|  5   | [SelectNodes](#CustomXMLPart.SelectNodes)           | 从自定义 XML 部件中选择节点的集合。                                              |
|  6   | [SelectSingleNode](#CustomXMLPart.SelectSingleNode) | 选择自定义 XML 部件内与 XPath 表达式匹配的单个节点。                             |

## 属性

| 序号 | 名称                                                | 说明                                                                                                                                  |
|:----:|:----------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLPart.Application)           | 获取一个代表 CustomXMLPart 对象的容器应用程序的 Application 对象。只读。                                                              |
|  2   | [BuiltIn](#CustomXMLPart.BuiltIn)                   | 获取一个指示 CustomXMLPart 是否为内置的值。只读                                                                                       |
|  3   | [Creator](#CustomXMLPart.Creator)                   | 获取一个 32 位整数，指示创建 CustomXMLPart 对象时所使用的应用程序。只读。                                                             |
|  4   | [DocumentElement](#CustomXMLPart.DocumentElement)   | 获取文档中数据绑定区域的根元素。如果该区域为空，则该属性将返回 Nothing 。只读。                                                       |
|  5   | [Errors](#CustomXMLPart.Errors)                     | 获取一个提供对任何 XML 验证错误（如果有）的访问的 CustomXMLValidationErrors 对象。如果不存在验证错误，则此属性将返回 Nothing 。只读。 |
|  6   | [Id](#CustomXMLPart.Id)                             | 获取一个 String 类型的值，其中包含分配给当前 CustomXMLPart 对象的 GUID。只读。                                                        |
|  7   | [NamespaceManager](#CustomXMLPart.NamespaceManager) | 获取对当前 CustomXMLPart 对象使用的命名空间前缀映射集。只读。                                                                         |
|  8   | [NamespaceURI](#CustomXMLPart.NamespaceURI)         | 获取 CustomXMLPart 对象的命名空间的唯一地址标识符。只读。                                                                             |
|  9   | [Parent](#CustomXMLPart.Parent)                     | 获取 CustomXMLPart 对象的 Parent 对象。只读。                                                                                         |
|  10  | [SchemaCollection](#CustomXMLPart.SchemaCollection) | 获取或设置一个代表附加到文档中数据绑定区域的架构集的 CustomXMLSchemaCollection 对象。可读/写。                                        |
|  11  | [XML](#CustomXMLPart.XML)                           | 获取当前 CustomXMLPart 对象的 XML 表示形式。只读。                                                                                    |

## 成员方法

### CustomXMLPart.AddNode

将节点添加到 XML 树。

**语法**

**express.AddNode ( *Parent* , *Name* , *NamespaceURI* , *NextSibling* , *NodeType* , *NodeValue* )**

\*express是一个代表 CustomXMLPart 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型             | 说明                                                                                                                                                                                                  |
|----------------|-----------|----------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Parent*       | 必选      | CustomXMLNode        | 代表应该在其下添加此节点的节点。如果添加属性，则该参数指示应该向其添加属性的元素。                                                                                                                    |
| *Name*         | 可选      | String               | 代表要添加的节点的基本名称。                                                                                                                                                                          |
| *NamespaceURI* | 可选      | String               | 代表要追加的元素的命名空间。如果要追加 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 类型的节点，则此参数是必需的，否则它将被忽略。                                                            |
| *NextSibling*  | 可选      | CustomXMLNode        | 代表应该成为新节点的下一个同级项的节点。如果未指定此参数，则将该节点添加到父节点的子节点的末尾。如果添加 msoXMLNodeAttribute 类型的节点，则忽略此参数。如果该节点不是父节点的子节点，则显示一个错误。 |
| *NodeType*     | 可选      | MsoCustomXMLNodeType | 指定要追加的节点的类型。如果未指定该参数，则假定节点类型为 msoCustomXMLNodeElement。                                                                                                                  |
| *NodeValue*    | 可选      | String               | 用于为允许文本的那些节点设置所追加节点的值。如果节点不允许文本，则忽略该参数。                                                                                                                        |

**说明**

如果 **AddNode** 操作将导致树结构无效，则不执行追加操作，并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例说明了如何将节点添加到 CustomXMLPart 对象。 
function test()
{
    let cxp1
    let cxn
    let ad = Application.ActiveDocument
    // Add and populate a custom xml part
    cxp1 = ad.CustomXMLParts.Add("<invoice />")
        
    // Set the parent node 
    cxn = cxp1.SelectSingleNode("/invoice")
        
    // Add a node under the parent node
    cxp1.AddNode(cxn, "upccode", "urn:invoice:namespace")
}
```

### CustomXMLPart.Delete

从数据存储（ **IXMLDataStore** 接口）中删除当前 **CustomXMLPart** 。

**语法**

**express.Delete ()**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

如果尝试删除包含核心属性的部件，则不会执行操作并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例将添加一个自定义 XML 部件、使用标准选择节点，并删除部件和节点。
function test()
{
    let ad = Application.ActiveDocument
    // Example written for Word.
    // Add a custom XML part and then load the XML from a file.
    let cxp1 = ad.CustomXMLParts.Add()
                  
    // Delete custom XML part.
    cxp1.Delete()
}
```

### CustomXMLPart.Load

使模板作者可以通过现有文件填充 **CustomXMLPart** 。如果加载成功，则返回 **True** 。

**语法**

**express.Load ( *FilePath* )**

\*express是一个代表 CustomXMLPart 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                |
|------------|-----------|----------|-----------------------------------------------------|
| *FilePath* | 必选      | String   | 指向用户计算机上或包含要加载的 XML 的网络上的文件。 |

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例将添加一个自定义 XML 部件、使用标准选择节点，并删除部件和节点。
function test()
{
    try 
    {
        let ad = Application.ActiveDocument
        // Example written for Word.
        // Add a custom XML part and then load the XML from a file.
        let cxp1 = ad.CustomXMLParts.Add()
        cxp1.Load("c:\\invoice.xml")
        
        let cxn = cxp1.SelectSingleNode("//*[@quantity < 4]") 
        // Insert a subtree before the single node selected previously.
        cxn.InsertSubTreeBefore("<discounts><discount>0.10</discount></discounts>")  
                  
        // Delete custom XML part.
        cxp1.Delete()
        cxn.Delete()
    }
                
    // Exception handling. Show the message and resume.
    catch(Err) 
    {
        alert(Err.Description)
    }
}
```

### CustomXMLPart.LoadXML

允许模板作者依据 XML 字符串填充 **CustomXMLPart** 对象。如果加载成功，则返回 **True** 。

**语法**

**express.LoadXML ( *XML* )**

\*express是一个代表 CustomXMLPart 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明               |
|-------|-----------|----------|--------------------|
| *XML* | 必选      | String   | 包含要加载的 XML。 |

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例通过字符串将 XML 加载到自定义 XML 部件中。
function test()
{
    try 
    {
        // Add a custom XML part and then load the XML.
        let cxp1 = Application.ActiveDocument.CustomXMLParts.Add()
        cxp1.LoadXML("<discounts><discount>0.10</discount></discounts>")
    }
                
    // Exception handling. Show the message and resume.
    catch(Err)
    {
        alert(Err.Description)
    }
}
```

### CustomXMLPart.SelectNodes

从自定义 XML 部件中选择节点的集合。

**语法**

**express.SelectNodes ( *XPath* )**

\*express是一个代表 CustomXMLPart 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                |
|---------|-----------|----------|---------------------|
| *XPath* | 必选      | String   | 包含 XPath 表达式。 |

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例说明了添加自定义 XML 部件，选择与命名空间 URI 匹配的部件，然后在该部件内选择与 XPath 表达式匹配的节点。
function test()
{
    let cxp1
    let cxn
    
    // Add a custom xml part.
    Application.ActiveDocument.CustomXMLParts.Add()
    
    // Return the first custom xml part with the given index.
    cxp1 = Application.ActiveDocument.CustomXMLParts.Item(1) 
    
    // Get all of the nodes matching an XPath expression.
    cxns = cxp1.SelectNodes("//*[@unitPrice > 20]")

}
```

### CustomXMLPart.SelectSingleNode

选择自定义 XML 部件内与 XPath 表达式匹配的单个节点。

**语法**

**express.SelectSingleNode ( *XPath* )**

\*express是一个代表 CustomXMLPart 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                    |
|---------|-----------|----------|-------------------------|
| *XPath* | 必选      | String   | 包含一个 XPath 表达式。 |

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例说明了添加自定义 XML 部件，选择与命名空间 URI 匹配的部件，然后在该部件内选择与 XPath 表达式匹配的节点。
function test()
{
    let cxp1
    let cxn
    // Add a custom xml part.
    Application.ActiveDocument.CustomXMLParts.Add()
    
    // Return the first custom xml part with the given index.
    cxp1 = Application.ActiveDocument.CustomXMLParts.Item(1)        
    
    // Get a node using XPath.                             
    cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")
}
```

## 成员属性

### CustomXMLPart.Application

获取一个代表 **CustomXMLPart** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPart.BuiltIn

获取一个指示 **CustomXMLPart** 是否为内置的值。只读

**语法**

**express.BuiltIn**

\*express是一个代表 CustomXMLPart 对象的变量。

### CustomXMLPart.Creator

获取一个 32 位整数，指示创建 **CustomXMLPart** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPart.DocumentElement

获取文档中数据绑定区域的根元素。如果该区域为空，则该属性将返回 **Nothing** 。只读。

**语法**

**express.DocumentElement**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPart.Errors

获取一个提供对任何 XML 验证错误（如果有）的访问的 **CustomXMLValidationErrors** 对象。如果不存在验证错误，则此属性将返回 **Nothing** 。只读。

**语法**

**express.Errors**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPart.Id

获取一个 **String** 类型的值，其中包含分配给当前 **CustomXMLPart** 对象的 GUID。只读。

**语法**

**express.Id**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPart.NamespaceManager

获取对当前 **CustomXMLPart** 对象使用的命名空间前缀映射集。只读。

**语法**

**express.NamespaceManager**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPart.NamespaceURI

获取 **CustomXMLPart** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPart.Parent

获取 **CustomXMLPart** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPart.SchemaCollection

获取或设置一个代表附加到文档中数据绑定区域的架构集的 **CustomXMLSchemaCollection** 对象。可读/写。

**语法**

**express.SchemaCollection**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPart.XML

获取当前 **CustomXMLPart** 对象的 XML 表示形式。只读。

**语法**

**express.XML**

\*express是一个代表 CustomXMLPart 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLPart](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DocumentProperties 对象

## 简介

DocumentProperty 对象的集合。每个 DocumentProperty 对象均代表容器文档的一个内置属性或自定义属性。

## 说明

使用 Add 方法可创建一个新的自定义属性，并将其添加到 DocumentProperties 集合中。不能使用 Add 方法创建内置文档属性。

使用 BuiltinDocumentProperties(index) 可返回一个代表特定内置文档属性的 DocumentProperty 对象；其中 *index* 是该内置文档属性的索引号。使用 CustomDocumentProperties(index) 可返回一个代表特定自定义文档属性的 DocumentProperty 对象；其中 *index* 是该自定义文档属性的编号。

## 方法

| 序号 | 名称                           | 说明                                                                           |
|:----:|:-------------------------------|:-------------------------------------------------------------------------------|
|  1   | [Add](#DocumentProperties.Add) | 新建一个自定义的文档属性。只能在自定义 DocumentProperties 集合中新添文档属性。 |

## 属性

| 序号 | 名称                                           | 说明                                                                                                                                      |
|:----:|:-----------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#DocumentProperties.Application) | 获取一个 Application 对象，代表 DocumentProperties 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Count](#DocumentProperties.Count)             | 获取一个 Long 类型的值，指示 DocumentProperties 集合中的项数。只读。                                                                      |
|  3   | [Creator](#DocumentProperties.Creator)         | 获取一个 32 位整数，指示创建 DocumentProperties 对象时所使用的应用程序。只读。                                                            |
|  4   | [Item](#DocumentProperties.Item)               | 获取 DocumentProperties 集合中的一个 DocumentProperty 对象。只读。                                                                        |
|  5   | [Parent](#DocumentProperties.Parent)           | 获取 DocumentProperties 对象的 Parent 对象。只读。                                                                                        |

## 成员方法

### DocumentProperties.Add

新建一个自定义的文档属性。只能在自定义 **DocumentProperties** 集合中新添文档属性。

**语法**

**express.Add ( *Name* , *LinkToContent* , *Type* , *Value* , *LinkSource* )**

\*express是一个代表 DocumentProperties 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                 |
|-----------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Name*          | 必选      | String   | 属性的名称。                                                                                                                                                                                                                                         |
| *LinkToContent* | 必选      | Boolean  | 指定属性是否链接到容器文档中的内容。如果参数为 True，则必需设置 LinkSource 参数，如果为 False，则需要数值参数。                                                                                                                                      |
| *Type*          | 可选      | Variant  | 该属性的数据类型。可以是以下 MsoDocProperties 常量之一：msoPropertyBoolean、msoPropertyDate、msoPropertyFloat、msoPropertyNumber 或 msoPropertyString。                                                                                              |
| *Value*         | 可选      | Variant  | 如果未链接到容器文档中的内容，则为属性的值。该数值将进行转换以匹配类型参数指定的数据类型，如果不能转换，则将发生错误。如果 LinkToContent 为 True，将忽略该参数，新文档属性将被赋予默认值，直到容器应用程序更新链接的属性值（通常在保存文档时进行）。 |
| *LinkSource*    | 可选      | Variant  | 如果 LinkToContent 为 False，则忽略此参数。链接属性的来源。可使用的源链接的类型由容器应用程序决定。                                                                                                                                                  |

**说明**

如果在 **DocumentProperties** 集合中添加了自定义文档属性，该属性链接到 WPS Office文档中指定的值，则必须保存文档以查看 **DocumentProperty** 对象的更改。

**示例**

``` JavaScript
//本示例设计为在 WPS 中运行，它将向 DocumentProperties 集合中添加三个自定义文档属性。
function ValidateSchemas(objSourceCustomXMLSchemaCollection) {
    let customProps = Application.ActiveDocument.CustomDocumentProperties
    customProps.Add("CustomNumber", false, Application.Enum.msoPropertyTypeNumber, 1000)
    customProps.Add("CustomString", false, Application.Enum.msoPropertyTypeString, "This is a custom property.")
    customProps.Add("CustomDate", false, Application.Enum.msoPropertyTypeDate, new Date())
}
```

## 成员属性

### DocumentProperties.Application

获取一个 **Application** 对象，代表 **DocumentProperties** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 DocumentProperties 对象的变量。

### DocumentProperties.Count

获取一个 **Long** 类型的值，指示 **DocumentProperties** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 DocumentProperties 对象的变量。

**示例**

``` JavaScript
//本示例显示活动文档中自定义文档属性的数量。
alert("There are " + Application.ActiveDocument.CustomDocumentProperties.Count +
    " custom document properties in the active document.")
```

### DocumentProperties.Creator

获取一个 32 位整数，指示创建 **DocumentProperties** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 DocumentProperties 对象的变量。

### DocumentProperties.Item

获取 **DocumentProperties** 集合中的一个 **DocumentProperty** 对象。只读。

**语法**

**express.Item**

\*express是一个代表 DocumentProperties 对象的变量。

### DocumentProperties.Parent

获取 **DocumentProperties** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DocumentProperties 对象的变量。

**示例**

> WPS 开发文档： [WPS 基础接口/通用 API 参考/DocumentProperties](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FileDialogFilters 对象

## 简介

表示文件对话框中可选文件类型的 FileDialogFilter 对象的集合，该对话框由 FileDialog 对象显示。

## 说明

使用 FileDialog 对象的 Filters 属性可返回一个 FileDialogFilters 集合。

使用 Add 方法可向 FileDialogFilters 集合中添加 FileDialogFilter 对象。下面的示例使用 Clear 方法清除集合，然后向集合中添加筛选器。Clear 方法将集合完全清空；但是，如果在清除后没有向集合中添加任何筛选器，将自动添加“所有文件(\*.\*)”筛选器。

``` JavaScript
function Main() {

    //Declare a variable as a FileDialog object.

    //Create a FileDialog object as a File Picker dialog box.
    let fd = Application.FileDialog(Application.Enum.msoFileDialogFilePicker)

    //Declare a variable to contain the path
    //of each selected item. Even though the path is aString,
    //the variable must be a Variant because For Each...Next
    //routines only work with Variants and Objects.

    //Use a With...End With block to reference the FileDialog object.

        //Change the contents of the Files of Type list.
        //Empty the list by clearing the FileDialogFilters collection.
        fd.Filters.Clear()

        //Add a filter that includes all files.
        fd.Filters.Add("All files", "*.*")

        //Add a filter that includes GIF and JPEG images and make it the first item in the list.
        fd.Filters.Add("Images", "*.gif; *.jpg; *.jpeg", 1)

        //Use the Show method to display the File Picker dialog box and return the user's action.
        //The user pressed the button.
        if(fd.Show() == -1) {

            //Step through eachString in the FileDialogSelectedItems collection.
            for(let i = 1;i <= fd.SelectedItems.Count;i++){
                let vrtSelectedItem = fd.SelectedItems.Item(i)
                //vrtSelectedItem is aString that contains the path of each selected item.
                //You can use any file I/O functions that you want to work with this path.
                //This example displays the path in a message box.
                MsgBox("Path name: " + vrtSelectedItem)

            }
        }
        //The user pressed Cancel.
        else { }

    //Set the object variable to Nothing.
    fd = null

}
```

在更改 FileDialogFilters 集合时，请记住每个应用程序只能创建单个 FileDialog 对象的实例。这意味着在调用一个新对话框类型的 FileDialog 方法时， FileDialogFilters 集合将重置为其默认筛选器。

下面的示例循环访问 “SaveAs” 对话框的默认筛选器，并显示每个包括 ET 文件的筛选器的说明。

``` JavaScript
function test() {

    //Declare a variable as a FileDialogFilters collection.

    //Declare a variable as a FileDialogFilter object.

    //Set the FileDialogFilters collection variable to
    //the FileDialogFilters collection of the SaveAs dialog box.
    let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters

    //Iterate through the description and extensions of each
    //default filter in the SaveAs dialog box.
    for(let i = 1;i <= fdfs.Count;i++) {
        let fdf = fdfs.Item(i)
        //Display the description of filters that include
        //ET files
        if(fdf.Extensions.indexOf("xls") > -1) {
            alert("Description of filter: " + fdf.Description)
        }
    }

}
```

``` JavaScript
Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters.Clear()
```

## 方法

| 序号 | 名称                                | 说明                                                                                                                                 |
|:----:|:------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Add](#FileDialogFilters.Add)       | 在 “文件” 对话框的 “文件类型” 下拉列表框的筛选器列表中添加一个新的文件筛选器。返回一个代表新添的文件筛选器的 FileDialogFilter 对象。 |
|  2   | [Clear](#FileDialogFilters.Clear)   | 删除当前在文件对话框中应用的所有筛选器。                                                                                             |
|  3   | [Delete](#FileDialogFilters.Delete) | 删除文件对话框筛选。                                                                                                                 |
|  4   | [Item](#FileDialogFilters.Item)     | 获取属于指定的 FileDialogFilters 集合的成员的 FileDialogFilter 对象。                                                                |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                     |
|:----:|:----------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FileDialogFilters.Application) | 获取一个 Application 对象，代表 FileDialogFilters 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Count](#FileDialogFilters.Count)             | 获取一个 Long 类型的值，指示 FileDialogFilters 集合中的项数。只读。                                                                      |
|  3   | [Creator](#FileDialogFilters.Creator)         | 获取一个 32 位整数，指示创建 FileDialogFilters 对象时所使用的应用程序。只读。                                                            |
|  4   | [Parent](#FileDialogFilters.Parent)           | 获取 FileDialogFilters 对象的 Parent 对象。只读。                                                                                        |

## 成员方法

### FileDialogFilters.Add

在 **“文件”** 对话框的 **“文件类型”** 下拉列表框的筛选器列表中添加一个新的文件筛选器。返回一个代表新添的文件筛选器的 **FileDialogFilter** 对象。

**语法**

**express.Add ( *Description* , *Extensions* , *Position* )**

\*express是一个代表 FileDialogFilters 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                       |
|---------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Description* | 必选      | String   | 该文本表示要添加到筛选器列表中的文件扩展名的说明。                                                                                                                                                                                                                         |
| *Extensions*  | 必选      | String   | 该文本代表要添加到筛选器列表中的文件扩展名。可以指定多个扩展名，每个扩展名必须以分号分隔。例如，可以向以下字符串分配参数：“\*.txt; \*.htm”。 注释 : 不需要在扩展名两侧添加括号。在说明和扩展名字符串连接成一个文件筛选器项时，WPS Office将自动在扩展名字符串两侧添加括号。 |
| *Position*    | 必选      | Variant  | 表示新控件在筛选列表中位置的数字。新筛选将插入到该位置的筛选之前。如果忽略该参数，筛选将添加到指定列表的末端。                                                                                                                                                             |

**说明**

列表中的每个筛选器由两部分组成：文件扩展名（例如 .txt）和文件扩展名的文本说明（例如“文本文件”）。二者相结合，文件筛选器将在 **“文件类型”** 下拉列表框中显示为：文本文件(\*.txt)。请注意，向列表中添加筛选器时，不会删除默认的筛选器。仅当选中 **“窗口”** 选项时才显示筛选器。如果 *Position* 无效，将显示超出范围运行时错误。如果 *Description* 和 *Extensions* 值无效，将显示运行时错误（分析）。

### FileDialogFilters.Clear

删除当前在文件对话框中应用的所有筛选器。

**语法**

**express.Clear ()**

\*express是一个代表 FileDialogFilters 对象的变量。

**示例**

下面的示例循环访问 **“SaveAs”** 对话框的默认筛选器，并显示每个包括 ET 文件的筛选器的说明。

``` JavaScript
function test() {

    //Declare a variable as a FileDialogFilters collection.

    //Declare a variable as a FileDialogFilter object.

    //Set the FileDialogFilters collection variable to
    //the FileDialogFilters collection of the SaveAs dialog box.
    let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters

    //Iterate through the description and extensions of each
    //default filter in the SaveAs dialog box.
    for(let i = 1;i <= fdfs.Count;i++) {
        let fdf = fdfs.Item(i)
        //Display the description of filters that include
        //ET files
        if(fdf.Extensions.indexOf("xls") > -1) {
            MsgBox("Description of filter: " + fdf.Description)
        }
    }

}
```

### FileDialogFilters.Delete

删除文件对话框筛选。

**语法**

**express.Delete ( *filter* )**

\*express是一个代表 FileDialogFilters 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明             |
|----------|-----------|----------|------------------|
| *filter* | 可选      | Variant  | 要删除的筛选器。 |

### FileDialogFilters.Item

获取属于指定的 **FileDialogFilters** 集合的成员的 **FileDialogFilter** 对象。

**语法**

**express.Item ()**

\*express是一个代表 FileDialogFilters 对象的变量。

## 成员属性

### FileDialogFilters.Application

获取一个 **Application** 对象，代表 **FileDialogFilters** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 FileDialogFilters 对象的变量。

### FileDialogFilters.Count

获取一个 **Long** 类型的值，指示 **FileDialogFilters** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 FileDialogFilters 对象的变量。

### FileDialogFilters.Creator

获取一个 32 位整数，指示创建 **FileDialogFilters** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 FileDialogFilters 对象的变量。

### FileDialogFilters.Parent

获取 **FileDialogFilters** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FileDialogFilters 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/FileDialogFilters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ParagraphFormat2 对象

## 简介

代表文本范围的段落格式。

## 说明

以下示例将活动 PowerPoint 演示文稿第一张幻灯片第二个形状中的段落左对齐。

``` JavaScript
Application.ActiveWorkbook.ActiveChart.Shapes.Item(1).SmartArt.AllNodes.Item(1).TextFrame2.TextRange.Text="Node 1"
```

## 属性

| 序号 | 名称                                                             | 说明                                                                 |
|:----:|:-----------------------------------------------------------------|:---------------------------------------------------------------------|
|  1   | [Alignment](#ParagraphFormat2.Alignment)                         | 获取或设置一个值，该值指定段落的对齐方式。可读写。                   |
|  2   | [Application](#ParagraphFormat2.Application)                     | 获取一个代表包含该对象的应用程序的对象。只读。                       |
|  3   | [BaseLineAlignment](#ParagraphFormat2.BaseLineAlignment)         | 获取或设置一个常量，该常量代表段落中字体的垂直位置。可读写。         |
|  4   | [Bullet](#ParagraphFormat2.Bullet)                               | 获取段落的 BulletFormat2 对象。只读。                                |
|  5   | [Creator](#ParagraphFormat2.Creator)                             | 获取一个值，该值代表创建 ParagraphFormat2 对象的应用程序。只读。     |
|  6   | [FarEastLineBreakLevel](#ParagraphFormat2.FarEastLineBreakLevel) | 获取或设置指定段落的东亚换行符控制级别。可读写。                     |
|  7   | [FirstLineIndent](#ParagraphFormat2.FirstLineIndent)             | 获取或设置首行缩进或悬挂缩进的值（以磅值表示）。可读写。             |
|  8   | [HangingPunctuation](#ParagraphFormat2.HangingPunctuation)       | 确定指定段落中的标点是否可以溢出边界。可读写。                       |
|  9   | [IndentLevel](#ParagraphFormat2.IndentLevel)                     | 获取或设置一个值，该值代表分配给选定段落中的文本的缩进级别。可读写。 |
|  10  | [LeftIndent](#ParagraphFormat2.LeftIndent)                       | 返回或设置一个值，该值代表指定段落的左缩进值（以磅为单位）。可读写。 |
|  11  | [LineRuleAfter](#ParagraphFormat2.LineRuleAfter)                 | 确定是否将每段最后一行后面的行距设为特定的磅数或行数。可读写。       |
|  12  | [LineRuleBefore](#ParagraphFormat2.LineRuleBefore)               | 确定是否将每段首行前面的行距设为特定的磅数或行数。可读写。           |
|  13  | [LineRuleWithin](#ParagraphFormat2.LineRuleWithin)               | 确定是否将基线间的行距设为特定的磅数或行数。可读写。                 |
|  14  | [Parent](#ParagraphFormat2.Parent)                               | 获取 ParagraphFormat2 对象的父对象。只读。                           |
|  15  | [RightIndent](#ParagraphFormat2.RightIndent)                     | 获取或设置指定段落的右缩进量（以磅为单位）。可读写。                 |
|  16  | [SpaceAfter](#ParagraphFormat2.SpaceAfter)                       | 获取或设置指定段落的段后间距（以磅为单位）。可读写。                 |
|  17  | [SpaceBefore](#ParagraphFormat2.SpaceBefore)                     | 获取或设置指定段落的段前间距（以磅为单位）。可读写。                 |
|  18  | [SpaceWithin](#ParagraphFormat2.SpaceWithin)                     | 获取或设置指定段落中基准行之间的距离（以磅或行为单位）。可读写。     |
|  19  | [TabStops](#ParagraphFormat2.TabStops)                           | 获取一个代表指定段落的所有自定义制表位的 TabStops2 集合。只读。      |
|  20  | [TextDirection](#ParagraphFormat2.TextDirection)                 | 获取或设置指定段落的文本方向。可读写。                               |
|  21  | [WordWrap](#ParagraphFormat2.WordWrap)                           | 确定应用程序是否在指定段落的语句中对拉丁语文本换行。可读写。         |

## 成员属性

### ParagraphFormat2.Alignment

获取或设置一个值，该值指定段落的对齐方式。可读写。

**语法**

**express.Alignment**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.Application

获取一个代表包含该对象的应用程序的对象。只读。

**语法**

**express.Application**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.BaseLineAlignment

获取或设置一个常量，该常量代表段落中字体的垂直位置。可读写。

**语法**

**express.BaseLineAlignment**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.Bullet

获取段落的 **BulletFormat2** 对象。只读。

**语法**

**express.Bullet**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.Creator

获取一个值，该值代表创建 **ParagraphFormat2** 对象的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.FarEastLineBreakLevel

获取或设置指定段落的东亚换行符控制级别。可读写。

**语法**

**express.FarEastLineBreakLevel**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.FirstLineIndent

获取或设置首行缩进或悬挂缩进的值（以磅值表示）。可读写。

**语法**

**express.FirstLineIndent**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.HangingPunctuation

确定指定段落中的标点是否可以溢出边界。可读写。

**语法**

**express.HangingPunctuation**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.IndentLevel

获取或设置一个值，该值代表分配给选定段落中的文本的缩进级别。可读写。

**语法**

**express.IndentLevel**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.LeftIndent

返回或设置一个值，该值代表指定段落的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.LineRuleAfter

确定是否将每段最后一行后面的行距设为特定的磅数或行数。可读写。

**语法**

**express.LineRuleAfter**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.LineRuleBefore

确定是否将每段首行前面的行距设为特定的磅数或行数。可读写。

**语法**

**express.LineRuleBefore**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.LineRuleWithin

确定是否将基线间的行距设为特定的磅数或行数。可读写。

**语法**

**express.LineRuleWithin**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.Parent

获取 **ParagraphFormat2** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.RightIndent

获取或设置指定段落的右缩进量（以磅为单位）。可读写。

**语法**

**express.RightIndent**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.SpaceAfter

获取或设置指定段落的段后间距（以磅为单位）。可读写。

**语法**

**express.SpaceAfter**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.SpaceBefore

获取或设置指定段落的段前间距（以磅为单位）。可读写。

**语法**

**express.SpaceBefore**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.SpaceWithin

获取或设置指定段落中基准行之间的距离（以磅或行为单位）。可读写。

**语法**

**express.SpaceWithin**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.TabStops

获取一个代表指定段落的所有自定义制表位的 **TabStops2** 集合。只读。

**语法**

**express.TabStops**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.TextDirection

获取或设置指定段落的文本方向。可读写。

**语法**

**express.TextDirection**

\*express是一个代表 ParagraphFormat2 对象的变量。

### ParagraphFormat2.WordWrap

确定应用程序是否在指定段落的语句中对拉丁语文本换行。可读写。

**语法**

**express.WordWrap**

\*express是一个代表 ParagraphFormat2 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/ParagraphFormat2](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PickerResult 对象

## 简介

代表解析的或选定的项目数据。

## 说明

``` JavaScript
//下面的代码设置选取器对话框的属性，然后显示选取器对话框。
function test()
{
    // Configure the Picker Dialog properties.
    let objPickerDialog = Application.PickerDialog
    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
    objPickerDialog.Title = "Sample Picker Dialog"
    let objPickerProperties = objPickerDialog.Properties
    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)
    let objPickerExistingResults = objPickerDialog.CreatePickerResults()
    let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")
     
    // Show the Picker Dialog and get the results.
    let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)
}
```

## 属性

| 序号 | 名称                                               | 说明                                                                                            |
|:----:|:---------------------------------------------------|:------------------------------------------------------------------------------------------------|
|  1   | [Application](#PickerResult.Application)           | 获取一个 Application 对象，该对象代表 PickerResult 对象的容器应用程序。只读                     |
|  2   | [Creator](#PickerResult.Creator)                   | 获取一个 32 位整数，该整数指示在其中创建了 PickerResult 对象的应用程序。只读                    |
|  3   | [DisplayName](#PickerResult.DisplayName)           | 代表 PickerResult 的显示名称。可读/写                                                           |
|  4   | [DuplicateResults](#PickerResult.DuplicateResults) | 如果对结果进行解析的结果有多个候选项，则获取 PickerResult 集合。只读                            |
|  5   | [Fields](#PickerResult.Fields)                     | 代表 PickerFields 集合内子项的字段定义。只读                                                    |
|  6   | [Id](#PickerResult.Id)                             | 检索关联的 PickerResult 对象的唯一 Id。只读                                                     |
|  7   | [ItemData](#PickerResult.ItemData)                 | 获取或设置绑定到数据的不用于显示的项。可读/写                                                   |
|  8   | [SIPId](#PickerResult.SIPId)                       | 是 Office Communication Server 的标识符。它仅用于人员选取方案。可读/写                          |
|  9   | [SubItems](#PickerResult.SubItems)                 | 显示用于显示或不用于显示的 PickerResult 对象的字段数据。它用于在选取器对话框中传递列值。可读/写 |
|  10  | [Type](#PickerResult.Type)                         | 代表 PickerResult 对象的类型。可读/写                                                           |

## 成员属性

### PickerResult.Application

获取一个 **Application** 对象，该对象代表 **PickerResult** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PickerResult 对象的变量。

### PickerResult.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PickerResult** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PickerResult 对象的变量。

### PickerResult.DisplayName

代表 PickerResult 的显示名称。可读/写

**语法**

**express.DisplayName**

\*express是一个代表 PickerResult 对象的变量。

### PickerResult.DuplicateResults

如果对结果进行解析的结果有多个候选项，则获取 PickerResult 集合。只读

**语法**

**express.DuplicateResults**

\*express是一个代表 PickerResult 对象的变量。

### PickerResult.Fields

代表 PickerFields 集合内子项的字段定义。只读

**语法**

**express.Fields**

\*express是一个代表 PickerResult 对象的变量。

### PickerResult.Id

检索关联的 **PickerResult** 对象的唯一 Id。只读

**语法**

**express.Id**

\*express是一个代表 PickerResult 对象的变量。

### PickerResult.ItemData

获取或设置绑定到数据的不用于显示的项。可读/写

**语法**

**express.ItemData**

\*express是一个代表 PickerResult 对象的变量。

### PickerResult.SIPId

是 Office Communication Server 的标识符。它仅用于人员选取方案。可读/写

**语法**

**express.SIPId**

\*express是一个代表 PickerResult 对象的变量。

### PickerResult.SubItems

显示用于显示或不用于显示的 PickerResult 对象的字段数据。它用于在选取器对话框中传递列值。可读/写

**语法**

**express.SubItems**

\*express是一个代表 PickerResult 对象的变量。

### PickerResult.Type

代表 **PickerResult** 对象的类型。可读/写

**语法**

**express.Type**

\*express是一个代表 PickerResult 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PickerResult](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SignatureSetup 对象

## 简介

代表用于设置签名数据包的信息。

## 说明

``` JavaScript
/*以下示例将设置签名数据包的 SignatureSetup 对象的各种属性。*/
function test(){
    var objSigSetup = Application.ActiveWorkbook.Signatures
    objSigSetup.AllowComments = true 
    objSigSetup.ShowSignDate = true 
    objSigSetup.SigningInstructions = "Please sign this document." 
    objSigSetup.SuggestedSignerEmail = "jdow@example.com"
}
```

## 属性

| 序号 | 名称                                                         | 说明                                                                                     |
|:----:|:-------------------------------------------------------------|:-----------------------------------------------------------------------------------------|
|  1   | [AdditionalXml](#SignatureSetup.AdditionalXml)               | 获取或设置在设置过程中添加到签名的任何附加 XML 信息。可读/写。                           |
|  2   | [AllowComments](#SignatureSetup.AllowComments)               | 获取或设置一个 Boolean 类型的值，指定签署人是否可以在 “签名” 对话框中输入注释。可读/写。 |
|  3   | [Application](#SignatureSetup.Application)                   | 获取一个代表 SignatureSetup 对象的容器应用程序的 Application 对象。只读。                |
|  4   | [Creator](#SignatureSetup.Creator)                           | 获取一个 32 位整数，指示创建 SignatureSetup 对象时所使用的应用程序。只读。               |
|  5   | [Id](#SignatureSetup.Id)                                     | 获取文档的签名提供程序的 ID。只读。                                                      |
|  6   | [ReadOnly](#SignatureSetup.ReadOnly)                         | 获取一个 Boolean 类型的值，指示 SignatureSetup 对象是否为只读。只读。                    |
|  7   | [ShowSignDate](#SignatureSetup.ShowSignDate)                 | 获取或设置一个 Boolean 类型的值，指示是否应显示文档的签署日期。可读/写。                 |
|  8   | [SignatureProvider](#SignatureSetup.SignatureProvider)       | 获取一个标识所安装的签名提供程序加载项的值。只读。                                       |
|  9   | [SigningInstructions](#SignatureSetup.SigningInstructions)   | 获取或设置用于签署文档的指令。可读/写。                                                  |
|  10  | [SuggestedSigner](#SignatureSetup.SuggestedSigner)           | 获取或设置文档主要签署人的名称。可读/写。                                                |
|  11  | [SuggestedSignerEmail](#SignatureSetup.SuggestedSignerEmail) | 获取或设置文档签署人的电子邮件地址。可读/写。                                            |
|  12  | [SuggestedSignerLine2](#SignatureSetup.SuggestedSignerLine2) | 获取或设置建议签署人信息的第二行（例如，职务）。可读/写。                                |

## 成员属性

### SignatureSetup.AdditionalXml

获取或设置在设置过程中添加到签名的任何附加 XML 信息。可读/写。

**语法**

**express.AdditionalXml**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.AllowComments

获取或设置一个 **Boolean** 类型的值，指定签署人是否可以在 **“签名”** 对话框中输入注释。可读/写。

**语法**

**express.AllowComments**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.Application

获取一个代表 **SignatureSetup** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.Creator

获取一个 32 位整数，指示创建 **SignatureSetup** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.Id

获取文档的签名提供程序的 ID。只读。

**语法**

**express.Id**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.ReadOnly

获取一个 **Boolean** 类型的值，指示 **SignatureSetup** 对象是否为只读。只读。

**语法**

**express.ReadOnly**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.ShowSignDate

获取或设置一个 **Boolean** 类型的值，指示是否应显示文档的签署日期。可读/写。

**语法**

**express.ShowSignDate**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SignatureProvider

获取一个标识所安装的签名提供程序加载项的值。只读。

**语法**

**express.SignatureProvider**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SigningInstructions

获取或设置用于签署文档的指令。可读/写。

**语法**

**express.SigningInstructions**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SuggestedSigner

获取或设置文档主要签署人的名称。可读/写。

**语法**

**express.SuggestedSigner**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SuggestedSignerEmail

获取或设置文档签署人的电子邮件地址。可读/写。

**语法**

**express.SuggestedSignerEmail**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SuggestedSignerLine2

获取或设置建议签署人信息的第二行（例如，职务）。可读/写。

**语法**

**express.SuggestedSignerLine2**

\*express是一个代表 SignatureSetup 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SignatureSetup](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartArtLayouts 对象

## 简介

代表 Smart Art 布局图表的集合。

## 说明

选项包括基本列表、图片题注列表、垂直项目符号列表等。

``` JavaScript
/*下面的代码更改 WPP 中的 Smart Art 图表的图表样式。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Layout = Application.SmartArtLayouts.Item(1)
```

## 方法

| 序号 | 名称                          | 说明                                                           |
|:----:|:------------------------------|:---------------------------------------------------------------|
|  1   | [Item](#SmartArtLayouts.Item) | 检索位于指定索引处或具有指定的唯一 Id 的 SmartArtLayout 对象。 |

## 属性

| 序号 | 名称                                        | 说明                                                                              |
|:----:|:--------------------------------------------|:----------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtLayouts.Application) | 获取一个 Application 对象，该对象代表 SmartArtLayouts 对象的容器应用程序。只读。  |
|  2   | [Count](#SmartArtLayouts.Count)             | 检索包含在 SmartArtLayouts 集合中的 SmartArtLayout 对象数的计数。只读。           |
|  3   | [Creator](#SmartArtLayouts.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtLayouts 对象的应用程序。只读。 |
|  4   | [Parent](#SmartArtLayouts.Parent)           | 返回调用对象。只读。                                                              |

## 成员方法

### SmartArtLayouts.Item

检索位于指定索引处或具有指定的唯一 Id 的 **SmartArtLayout** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SmartArtLayouts 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                 |
|---------|-----------|----------|----------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 指定一个代表索引的整数或一个代表 SmartArtLayout 对象的位置的字符串。 |

**返回值**

SmartArtLayout

## 成员属性

### SmartArtLayouts.Application

获取一个 **Application** 对象，该对象代表 **SmartArtLayouts** 对象的容器应用程序。只读。

**语法**

**express.Application**

\*express是一个代表 SmartArtLayouts 对象的变量。

### SmartArtLayouts.Count

检索包含在 **SmartArtLayouts** 集合中的 **SmartArtLayout** 对象数的计数。只读。

**语法**

**express.Count**

\*express是一个代表 SmartArtLayouts 对象的变量。

### SmartArtLayouts.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtLayouts** 对象的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SmartArtLayouts 对象的变量。

### SmartArtLayouts.Parent

返回调用对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SmartArtLayouts 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtLayouts](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DataRange 对象

## 简介

代表容器应用程序的一个数据区域对象。

## 说明

DataRange为类似表格Range概念，对应于表格中的一个区域，用于存储缓存数据，如果注册进DataSource，表格中的对应区域数据修改，将会触发事件。

## 属性

| 序号 | 名称                                  | 说明                                                                     |
|:----:|:--------------------------------------|:-------------------------------------------------------------------------|
|  1   | [Address](#DataRange.Address)         | 代表DataRange对象的地址。形如“sheet1!\$a\$1”表示sheet1中A1单元格。只读。 |
|  2   | [TextInJSON](#DataRange.TextInJSON)   | json格式表示区域的格式化后的值。只读。                                   |
|  3   | [ValueInJSON](#DataRange.ValueInJSON) | json格式表示区域的值。只读。                                             |

## 成员属性

### DataRange.Address

代表DataRange对象的地址。形如“sheet1!\$a\$1”表示sheet1中A1单元格。只读。

**语法**

**express.Address**

\*express是一个代表 DataRange 对象的变量。

**类型**

String

### DataRange.TextInJSON

json格式表示区域的格式化后的值。只读。

**语法**

**express.TextInJSON**

\*express是一个代表 DataRange 对象的变量。

**类型**

String

### DataRange.ValueInJSON

json格式表示区域的值。只读。

**语法**

**express.ValueInJSON**

\*express是一个代表 DataRange 对象的变量。

> WPS 开发文档： [WPS 基础接口/WPS 扩展 API/公共/DataRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
