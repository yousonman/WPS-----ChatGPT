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
