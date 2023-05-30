# ScrollBar 对象

## 简介

提供在长列表工程或大量信息中快速浏览的图形工具，以比例方式指示出当前位置，或是做为一个输入设备，或速度及数量的指示器。

## 说明

提供在长列表工程或大量信息中快速浏览的图形工具，以比例方式指示出当前位置，或是做为一个输入设备，或速度及数量的指示器。

## 方法

| 序号 | 名称                    | 说明                       |
|:----:|:------------------------|:---------------------------|
|  1   | [Move](#ScrollBar.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                      | 说明                                                      |
|:----:|:------------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#ScrollBar.AutoSize)           | 定该对象是否根据内容自动设置它的大小                      |
|  2   | [BackColor](#ScrollBar.BackColor)         | 获取或设置指定对象的内部颜色                              |
|  3   | [ControlSource](#ScrollBar.ControlSource) | 您可以使用该属性指定控件中显示的数据                      |
|  4   | [Enabled](#ScrollBar.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                  |
|  5   | [Font](#ScrollBar.Font)                   | 指定对象的字体，只读                                      |
|  6   | [ForeColor](#ScrollBar.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                      |
|  7   | [Height](#ScrollBar.Height)               | 获取或设置指定对象的高度                                  |
|  8   | [HeightPolicy](#ScrollBar.HeightPolicy)   | 获取或设置指定对像的高度策略                              |
|  9   | [LargeChange](#ScrollBar.LargeChange)     | 指定当用户在滚动框和滚动箭头之间单击时发生的移动量。      |
|  10  | [Left](#ScrollBar.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  11  | [Max](#ScrollBar.Max)                     | 为该控件的Value属性指定最大可接受值。                     |
|  12  | [Min](#ScrollBar.Min)                     | 为该控件的Value属性指定最小可接受值。                     |
|  13  | [Name](#ScrollBar.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  14  | [Orientation](#ScrollBar.Orientation)     | 指定该控件是垂直定向还是水平定向                          |
|  15  | [SmallChange](#ScrollBar.SmallChange)     | 指定当用户单击滚动箭头时发生的移动量。                    |
|  16  | [TabIndex](#ScrollBar.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  17  | [TabStop](#ScrollBar.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  18  | [Top](#ScrollBar.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置              |
|  19  | [Value](#ScrollBar.Value)                 | 您可以通过该属性指定一个在Max和Min之间的值                |
|  20  | [Visible](#ScrollBar.Visible)             | 返回或设置该对象是否可见                                  |
|  21  | [Width](#ScrollBar.Width)                 | 获取或设置指定对象的宽度                                  |
|  22  | [WidthPolicy](#ScrollBar.WidthPolicy)     | 获取或设置指定对象的宽度策略                              |

## 成员方法

### ScrollBar.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 ScrollBar 对象的变量。

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
/*将ScrollBar1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.ScrollBar1.Move(100,100,100,100)
}
```

## 成员属性

### ScrollBar.AutoSize

定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置ScrollBar1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.ScrollBar1.AutoSize = true
}
```

### ScrollBar.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将ScrollBar1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.ScrollBar1.BackColor = 0x665898
}
```

### ScrollBar.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### ScrollBar.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.ScrollBar1.Enabled = true
}
```

### ScrollBar.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

指定对象的字体，只读

### ScrollBar.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将ScrollBar1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.ScrollBar1.ForeColor = 0x658978
}
```

### ScrollBar.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将ScrollBar1控件的高度设置为100*/
function func1()
{
    UserForm1.ScrollBar1.Height = 100
}
```

### ScrollBar.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将ScrollBar1的高度设置为固定*/
function func1()
{
    UserForm1.ScrollBar1.HeightPolicy = 2
}
```

### ScrollBar.LargeChange

指定当用户在滚动框和滚动箭头之间单击时发生的移动量。

**语法**

**express.LargeChange**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

指定当用户在滚动框和滚动箭头之间单击时发生的移动量。

**示例**

``` JavaScript
/*将ScrollBar1在用户点击滚动框和滚动箭头时发生的移动量设为200*/
function func1()
{
    UserForm1.ScrollBar1.LargeChange = 200
}
```

### ScrollBar.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ScrollBar1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.ScrollBar1.Left = 100
}
```

### ScrollBar.Max

为该控件的Value属性指定最大可接受值。

**语法**

**express.Max**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

为该控件的Value属性指定最大可接受值。

**示例**

``` JavaScript
/*将ScrollBar1的Value属性的最大可接受值设为2000*/
function func1()
{
    UserForm1.ScrollBar1.Max = 2000
}
```

### ScrollBar.Min

为该控件的Value属性指定最小可接受值。

**语法**

**express.Min**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

为该控件的Value属性指定最小可接受值。

**示例**

``` JavaScript
/*将ScrollBar1的Value属性的最小可接受值设为0*/
function func1()
{
    UserForm1.ScrollBar1.Mix = 0
}
```

### ScrollBar.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### ScrollBar.Orientation

指定该控件是垂直定向还是水平定向

**语法**

**express.Orientation**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

指定该控件是垂直定向还是水平定向，可以设置为以下值：

| 设置       | 值  | 描述                                 |
|------------|-----|--------------------------------------|
| Auto       | -1  | 根据控件的尺寸自动确定方向（默认）。 |
| Vertical   | 0   | 控件垂直呈现。                       |
| Horizontal | 1   | 控件水平呈现。                       |

**示例**

``` JavaScript
/*将ScrollBar1垂直呈现*/
function func1()
{
    UserForm1.ScrollBar1.Orientation = 0
}
```

### ScrollBar.SmallChange

指定当用户单击滚动箭头时发生的移动量。

**语法**

**express.SmallChange**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

指定当用户单击滚动箭头时发生的移动量。

**示例**

``` JavaScript
/*将ScrollBar1在用户点击滚动箭头时发生的移动量设为100*/
function func1()
{
    UserForm1.ScrollBar1.LargeChange = 100
}
```

### ScrollBar.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将ScrollBar1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.ScrollBar1.TabIndex = 2
}
```

### ScrollBar.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将ScrollBar1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.ScrollBar1.TabStop = true
}
```

### ScrollBar.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ScrollBar1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.ScrollBar1.Top = 100
}
```

### ScrollBar.Value

您可以通过该属性指定一个在Max和Min之间的值

**语法**

**express.Value**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定一个在Max和Min之间的值

**示例**

``` JavaScript
/*将ScrollBar1的Valued的值为500*/
function func1()
{
    UserForm1.ScrollBar1.Value = 500
}
```

### ScrollBar.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将ScrollBar1设置为可见*/
function func1()
{
    UserForm1.ScrollBar1.Visible = true
}
```

### ScrollBar.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将ScrollBar1的宽度设置为100*/
function func1()
{
    UserForm1.ScrollBar1.Width = 100
}
```

### ScrollBar.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将ScrollBar1的宽度设置为固定*/
function func1()
{
    UserForm1.ScrollBar1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/ScrollBar](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
