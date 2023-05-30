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
