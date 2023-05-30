# HLayout 对象

## 简介

将内部的控件按照水平方向排布，一列一个。

## 说明

将内部的控件按照水平方向排布，一列一个。

## 方法

| 序号 | 名称                  | 说明                       |
|:----:|:----------------------|:---------------------------|
|  1   | [Move](#HLayout.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                              | 说明                                                 |
|:----:|:--------------------------------------------------|:-----------------------------------------------------|
|  1   | [Enabled](#HLayout.Enabled)                       | 您可以通过该属性设置或返回是否启用该控件             |
|  2   | [Height](#HLayout.Height)                         | 获取或设置指定对象的高度                             |
|  3   | [LayoutBottomMargin](#HLayout.LayoutBottomMargin) | 获取或设置容器内部的控件与容器底部的距离             |
|  4   | [LayoutLeftMargin](#HLayout.LayoutLeftMargin)     | 获取或设置容器内部的控件与容器左边缘的距离           |
|  5   | [LayoutRightMargin](#HLayout.LayoutRightMargin)   | 获取或设置容器内部的控件与容器右边缘的距离           |
|  6   | [LayoutSpacing](#HLayout.LayoutSpacing)           | 获取或设置容器内部多个控件之间的水平间距             |
|  7   | [LayoutTopMargin](#HLayout.LayoutTopMargin)       | 获取或设置容器内部的控件与容器顶部的距离             |
|  8   | [Left](#HLayout.Left)                             | 您可以通过该属性指定控件在对象或报表中的位置         |
|  9   | [Name](#HLayout.Name)                             | 您可以通过该属性指定或确定表示对象名称的字符串，只读 |
|  10  | [Top](#HLayout.Top)                               | 您可以通过该属性指定控件在对象或报表中的位置         |
|  11  | [Visible](#HLayout.Visible)                       | 返回或设置该对象是否可见                             |
|  12  | [Width](#HLayout.Width)                           | 获取或设置指定对象的宽度                             |

## 成员方法

### HLayout.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 HLayout 对象的变量。

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
/*将HLayout1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.HLayout1.Move(100,100,100,100)
}
```

## 成员属性

### HLayout.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 HLayout 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.HLayout1.Enabled = true
}
```

### HLayout.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将HLayout1控件的高度设置为100*/
function func1()
{
    UserForm1.HLayout1.Height = 100
}
```

### HLayout.LayoutBottomMargin

获取或设置容器内部的控件与容器底部的距离

**语法**

**express.LayoutBottomMargin**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器底部的距离

**示例**

``` JavaScript
/*将HLayout1中控件到容器底部的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutBottomMargin = 10
}
```

### HLayout.LayoutLeftMargin

获取或设置容器内部的控件与容器左边缘的距离

**语法**

**express.LayoutLeftMargin**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器左边缘的距离

**示例**

``` JavaScript
/*将HLayout1中控件到容器左边缘的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutLeftMargin = 10
}
```

### HLayout.LayoutRightMargin

获取或设置容器内部的控件与容器右边缘的距离

**语法**

**express.LayoutRightMargin**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器右边缘的距离

**示例**

``` JavaScript
/*将HLayout1中控件到容器右边缘的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutRightMargin = 10
}
```

### HLayout.LayoutSpacing

获取或设置容器内部多个控件之间的水平间距

**语法**

**express.LayoutSpacing**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部多个控件之间的水平间距

**示例**

``` JavaScript
/*将HLayout1中多个控件之间的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutSpacing = 10
}
```

### HLayout.LayoutTopMargin

获取或设置容器内部的控件与容器顶部的距离

**语法**

**express.LayoutTopMargin**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器顶部的距离

**示例**

``` JavaScript
/*将HLayout1中控件到容器顶部的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutTopMargin = 10
}
```

### HLayout.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 HLayout 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将HLayout1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.HLayout1.Left = 100
}
```

### HLayout.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 HLayout 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### HLayout.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 HLayout 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将HLayout1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.HLayout1.Top = 100
}
```

### HLayout.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 HLayout 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将HLayout1设置为可见*/
function func1()
{
    UserForm1.HLayout1.Visible = true
}
```

### HLayout.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将HLayout1的宽度设置为100*/
function func1()
{
    UserForm1.HLayout1.Width = 100
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/HLayout](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
