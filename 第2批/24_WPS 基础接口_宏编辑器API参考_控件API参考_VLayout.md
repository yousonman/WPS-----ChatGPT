# VLayout 对象

## 简介

将内部的控件按照垂直方向排布，一行一个。

## 说明

将内部的控件按照垂直方向排布，一行一个。

## 方法

| 序号 | 名称                  | 说明                       |
|:----:|:----------------------|:---------------------------|
|  1   | [Move](#VLayout.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                              | 说明                                                 |
|:----:|:--------------------------------------------------|:-----------------------------------------------------|
|  1   | [Enabled](#VLayout.Enabled)                       | 您可以通过该属性设置或返回是否启用该控件             |
|  2   | [Height](#VLayout.Height)                         | 获取或设置指定对象的高度                             |
|  3   | [LayoutBottomMargin](#VLayout.LayoutBottomMargin) | 获取或设置容器内部的控件与容器底部的距离             |
|  4   | [LayoutLeftMargin](#VLayout.LayoutLeftMargin)     | 获取或设置容器内部的控件与容器左边缘的距离           |
|  5   | [LayoutRightMargin](#VLayout.LayoutRightMargin)   | 获取或设置容器内部的控件与容器右边缘的距离           |
|  6   | [LayoutSpacing](#VLayout.LayoutSpacing)           | 获取或设置容器内部多个控件之间的水平间距             |
|  7   | [LayoutTopMargin](#VLayout.LayoutTopMargin)       | 获取或设置容器内部的控件与容器顶部的距离             |
|  8   | [Left](#VLayout.Left)                             | 您可以通过该属性指定控件在对象或报表中的位置         |
|  9   | [Name](#VLayout.Name)                             | 您可以通过该属性指定或确定表示对象名称的字符串，只读 |
|  10  | [Top](#VLayout.Top)                               | 您可以通过该属性指定控件在对象或报表中的位置         |
|  11  | [Visible](#VLayout.Visible)                       | 返回或设置该对象是否可见                             |
|  12  | [Width](#VLayout.Width)                           | 获取或设置指定对象的宽度                             |

## 成员方法

### VLayout.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 VLayout 对象的变量。

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
/*将VLayout1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.VLayout1.Move(100,100,100,100)
}
```

## 成员属性

### VLayout.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 VLayout 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.VLayout1.Enabled = true
}
```

### VLayout.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将VLayout1控件的高度设置为100*/
function func1()
{
    UserForm1.VLayout1.Height = 100
}
```

### VLayout.LayoutBottomMargin

获取或设置容器内部的控件与容器底部的距离

**语法**

**express.LayoutBottomMargin**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器底部的距离

**示例**

``` JavaScript
/*将VLayout1中控件到容器底部的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutBottomMargin = 10
}
```

### VLayout.LayoutLeftMargin

获取或设置容器内部的控件与容器左边缘的距离

**语法**

**express.LayoutLeftMargin**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器左边缘的距离

**示例**

``` JavaScript
/*将VLayout1中控件到容器左边缘的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutLeftMargin = 10
}
```

### VLayout.LayoutRightMargin

获取或设置容器内部的控件与容器右边缘的距离

**语法**

**express.LayoutRightMargin**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器右边缘的距离

**示例**

``` JavaScript
/*将VLayout1中控件到容器右边缘的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutRightMargin = 10
}
```

### VLayout.LayoutSpacing

获取或设置容器内部多个控件之间的水平间距

**语法**

**express.LayoutSpacing**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部多个控件之间的水平间距

**示例**

``` JavaScript
/*将VLayout1中多个控件之间的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutSpacing = 10
}
```

### VLayout.LayoutTopMargin

获取或设置容器内部的控件与容器顶部的距离

**语法**

**express.LayoutTopMargin**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器顶部的距离

**示例**

``` JavaScript
/*将VLayout1中控件到容器顶部的距离设为10*/
function func1()
{
    UserForm1.VLayout1.LayoutTopMargin = 10
}
```

### VLayout.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 VLayout 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将VLayout1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.VLayout1.Left = 100
}
```

### VLayout.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 VLayout 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### VLayout.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 VLayout 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将VLayout1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.VLayout1.Top = 100
}
```

### VLayout.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 VLayout 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将VLayout1设置为可见*/
function func1()
{
    UserForm1.VLayout1.Visible = true
}
```

### VLayout.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 VLayout 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将VLayout1的宽度设置为100*/
function func1()
{
    UserForm1.VLayout1.Width = 100
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/VLayout](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
