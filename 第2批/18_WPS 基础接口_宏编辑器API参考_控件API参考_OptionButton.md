# OptionButton 对象

## 简介

可以显示出多重选择，用户只能选择一个。

## 说明

可以显示出多重选择，用户只能选择一个。

## 方法

| 序号 | 名称                       | 说明                       |
|:----:|:---------------------------|:---------------------------|
|  1   | [Move](#OptionButton.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                         | 说明                                                      |
|:----:|:---------------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#OptionButton.AutoSize)           | 指定该对象是否根据内容自动设置它的大小                    |
|  2   | [BackColor](#OptionButton.BackColor)         | 获取或设置指定对象的内部颜色                              |
|  3   | [BackStyle](#OptionButton.BackStyle)         | 您可以使用该属性指定控件是否透明                          |
|  4   | [Caption](#OptionButton.Caption)             | 获取或设置出现在控件中的文本                              |
|  5   | [ControlSource](#OptionButton.ControlSource) | 您可以使用该属性指定控件中显示的数据                      |
|  6   | [Enabled](#OptionButton.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                  |
|  7   | [Font](#OptionButton.Font)                   | 指定对象的字体，只读                                      |
|  8   | [ForeColor](#OptionButton.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                      |
|  9   | [GroupName](#OptionButton.GroupName)         | 您可以通过设置组名，创建一组互斥的OptionButton            |
|  10  | [Height](#OptionButton.Height)               | 获取或设置指定对象的高度                                  |
|  11  | [HeightPolicy](#OptionButton.HeightPolicy)   | 获取或设置指定对像的高度策略                              |
|  12  | [Left](#OptionButton.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  13  | [Name](#OptionButton.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  14  | [TabIndex](#OptionButton.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  15  | [TabStop](#OptionButton.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  16  | [Top](#OptionButton.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置              |
|  17  | [Visible](#OptionButton.Visible)             | 返回或设置该对象是否可见                                  |
|  18  | [Width](#OptionButton.Width)                 | 获取或设置指定对象的宽度                                  |
|  19  | [WidthPolicy](#OptionButton.WidthPolicy)     | 获取或设置指定对象的宽度策略                              |
|  20  | [WordWrap](#OptionButton.WordWrap)           | 指定控件的内容是否在行尾自动换行                          |

## 成员方法

### OptionButton.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 OptionButton 对象的变量。

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
/*将OptionButton1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.OptionButton1.Move(100,100,100,100)
}
```

## 成员属性

### OptionButton.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 OptionButton 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置OptionButton1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.OptionButton1.AutoSize = true
}
```

### OptionButton.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将OptionButton1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.OptionButton1.BackColor = 0x665898
}
```

### OptionButton.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

**示例**

``` JavaScript
/*将OptionButton1设置为透明*/
function func1()
{
    UserForm1.OptionButton1.BackStyle = 0
}
```

### OptionButton.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将OptionButton1中的文本设置为"OptionButton1"*/
function func1()
{
    UserForm1.OptionButton1.Caption = "OptionButton1"
}
```

### OptionButton.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### OptionButton.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.OptionButton1.Enabled = true
}
```

### OptionButton.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 OptionButton 对象的变量。

**说明**

指定对象的字体，只读

### OptionButton.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将OptionButton1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.OptionButton1.ForeColor = 0x658978
}
```

### OptionButton.GroupName

您可以通过设置组名，创建一组互斥的OptionButton

**语法**

**express.GroupName**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过设置组名，创建一组互斥的OptionButton。同一组里的OptionButton将采用相同的设置，点击组内的任何一个选项都会将其他选项置为false。不同容器下相同的组名会被视为不同组（每个容器一个组）。

**示例**

``` JavaScript
/*将OptionButton1的组名设置为group one*/
function func1()
{
    UserForm1.OptionButton1.GroupName = "group one"
}
```

### OptionButton.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将OptionButton1控件的高度设置为100*/
function func1()
{
    UserForm1.OptionButton1.Height = 100
}
```

### OptionButton.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将OptionButton1的高度设置为固定*/
function func1()
{
    UserForm1.OptionButton1.HeightPolicy = 2
}
```

### OptionButton.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将OptionButton1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.OptionButton1.Left = 100
}
```

### OptionButton.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### OptionButton.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将OptionButton1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.OptionButton1.TabIndex = 2
}
```

### OptionButton.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将OptionButton1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.OptionButton1.TabStop = true
}
```

### OptionButton.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将OptionButton1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.OptionButton1.Top = 100
}
```

### OptionButton.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 OptionButton 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将OptionButton1设置为可见*/
function func1()
{
    UserForm1.OptionButton1.Visible = true
}
```

### OptionButton.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将OptionButton1的宽度设置为100*/
function func1()
{
    UserForm1.OptionButton1.Width = 100
}
```

### OptionButton.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将OptionButton1的宽度设置为固定*/
function func1()
{
    UserForm1.OptionButton1.WidthPolicy = 2
}
```

### OptionButton.WordWrap

指定控件的内容是否在行尾自动换行

**语法**

**express.WordWrap**

\*express是一个代表 OptionButton 对象的变量。

**说明**

指定控件的内容是否在行尾自动换行

**示例**

``` JavaScript
/*将OptionButton1设置为控件的内容在行尾自动换行*/
function func1()
{
    UserForm1.OptionButton1.WordWrap = true
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/OptionButton](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
